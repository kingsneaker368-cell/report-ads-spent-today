// index.js
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile } = require('child_process');
const axios = require('axios');
const FormData = require('form-data');
const { google } = require('googleapis');
const sharp = require('sharp');

// === chống Google 429 ===
async function fetchPdfWithRetry(url, headers, attempt = 1) {
  try {
    return await axios.get(url, {
      responseType: 'arraybuffer',
      headers,
      maxContentLength: Infinity,
      maxBodyLength: Infinity
    });
  } catch (err) {
    if (err.response?.status === 429 && attempt < 5) {
      const delay = 3000 + Math.floor(Math.random() * 3000);
      await new Promise(r => setTimeout(r, delay));
      return fetchPdfWithRetry(url, headers, attempt + 1);
    }
    throw err;
  }
}

// === PDF → PNG ===
function convertPdfToPng(pdfPath, outPrefix) {
  return new Promise((resolve, reject) => {
    execFile(
      'pdftoppm',
      ['-png', '-singlefile', '-r', '150', pdfPath, outPrefix],
      async err => {
        if (err) return reject(err);
        const pngPath = outPrefix + '.png';
        const buf = await sharp(pngPath).trim().toBuffer();
        await fs.promises.writeFile(pngPath, buf);
        resolve(pngPath);
      }
    );
  });
}

async function main() {
  try {
    const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
    const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
    const SHEET_NAMES = (process.env.SHEET_NAMES || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);

    const START_COL = 'F';
    const END_COL = 'AD';
    const MAX_ROWS_PER_FILE = Number(process.env.MAX_ROWS_PER_FILE || '40');

    const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
    const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;

    const creds = JSON.parse(serviceAccountJson);
    const jwtClient = new google.auth.JWT(
      creds.client_email,
      null,
      creds.private_key,
      [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/spreadsheets.readonly'
      ]
    );

    await jwtClient.authorize();
    const accessToken = (await jwtClient.getAccessToken()).token;
    const sheetsApi = google.sheets({ version: 'v4', auth: jwtClient });

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheetpdf-'));
    const metadata = await sheetsApi.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const allSheets = metadata.data.sheets || [];

    for (const sheetName of SHEET_NAMES) {
      const sheetInfo = allSheets.find(s => s.properties.title === sheetName);
      if (!sheetInfo) continue;
      const gid = sheetInfo.properties.sheetId;

      // === TIÊU ĐỀ F → J ===
      const headerRes = await sheetsApi.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!F5:J5`
      });

      const headerVals = headerRes.data.values?.[0] || [];
      const headerText = headerVals.filter(Boolean).join(' ');

      // === TEXT A + B ===
      const abRes = await sheetsApi.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A4:B20`
      });

      const abText = (abRes.data.values || [])
        .map(r => r.filter(Boolean).join(' : '))
        .filter(Boolean)
        .join('\n');

      // === LAST ROW ===
      const colRes = await sheetsApi.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!K1:K2000`
      });

      const colVals = colRes.data.values || [];
      let lastRow = 1;
      for (let i = colVals.length - 1; i >= 0; i--) {
        if (colVals[i]?.[0]) {
          lastRow = i + 1;
          break;
        }
      }

      const albumImages = [];

      for (let r = 1; r <= lastRow; r += MAX_ROWS_PER_FILE) {
        const end = Math.min(r + MAX_ROWS_PER_FILE - 1, lastRow);
        const rangeParam = `${sheetName}!${START_COL}${r}:${END_COL}${end}`;

        const exportUrl =
          `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=pdf` +
          `&portrait=false&size=A4&fitw=true&sheetnames=false&printtitle=false` +
          `&pagenumbers=false&gridlines=false&gid=${gid}` +
          `&range=${encodeURIComponent(rangeParam)}`;

        const pdfResp = await fetchPdfWithRetry(exportUrl, {
          Authorization: `Bearer ${accessToken}`
        });

        const pdfPath = path.join(tmpDir, `${sheetName}_${r}-${end}.pdf`);
        fs.writeFileSync(pdfPath, pdfResp.data);
        const pngPath = await convertPdfToPng(pdfPath, pdfPath.replace('.pdf', ''));

        albumImages.push({
          path: pngPath,
          fileName: path.basename(pngPath)
        });
      }

      // === SEND ẢNH (KHÔNG CAPTION) ===
      const form = new FormData();
      form.append('chat_id', TELEGRAM_CHAT_ID);
      form.append(
        'media',
        JSON.stringify(
          albumImages.map(img => ({
            type: 'photo',
            media: `attach://${img.fileName}`
          }))
        )
      );

      albumImages.forEach(img =>
        form.append(img.fileName, fs.createReadStream(img.path))
      );

      await axios.post(
        `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMediaGroup`,
        form,
        { headers: form.getHeaders() }
      );

      // === SEND TEXT (F→I + A&B) ===
      const finalText = [headerText, abText].filter(Boolean).join('\n\n');

      if (finalText) {
        await axios.post(
          `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`,
          {
            chat_id: TELEGRAM_CHAT_ID,
            text: finalText
          }
        );
      }

      albumImages.forEach(i => fs.unlinkSync(i.path));
    }

    fs.rmSync(tmpDir, { recursive: true, force: true });
    console.log('DONE');
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
}

main();
