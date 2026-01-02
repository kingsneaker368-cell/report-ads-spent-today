// index.js
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile } = require('child_process');
const axios = require('axios');
const FormData = require('form-data');
const { google } = require('googleapis');
const sharp = require('sharp');

// === chống Google 429: retry 5 lần ===
async function fetchPdfWithRetry(url, headers, attempt = 1) {
  try {
    return await axios.get(url, {
      responseType: 'arraybuffer',
      headers,
      maxContentLength: Infinity,
      maxBodyLength: Infinity
    });
  } catch (err) {
    if (err.response && err.response.status === 429 && attempt < 5) {
      const delay = 3000 + Math.floor(Math.random() * 3000);
      console.log(`⚠️ Google 429 — retry ${attempt}/5 after ${delay}ms`);
      await new Promise(r => setTimeout(r, delay));
      return fetchPdfWithRetry(url, headers, attempt + 1);
    }
    throw err;
  }
}

// Convert PDF → PNG + trim
function convertPdfToPng(pdfPath, outPrefix) {
  return new Promise((resolve, reject) => {
    execFile(
      'pdftoppm',
      ['-png', '-singlefile', '-r', '150', pdfPath, outPrefix],
      async (err) => {
        if (err) return reject(err);
        const pngPath = outPrefix + '.png';
        if (!fs.existsSync(pngPath)) return reject(new Error('PNG conversion failed'));

        try {
          const img = sharp(pngPath);
          const trimmedBuffer = await img.trim().toBuffer();
          await fs.promises.writeFile(pngPath, trimmedBuffer);
          resolve(pngPath);
        } catch (e) {
          reject(e);
        }
      }
    );
  });
}

async function main() {
  try {
    const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
    const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
    const SHEET_NAMES = (process.env.SHEET_NAMES || '').split(',').map(s => s.trim()).filter(Boolean);
    const START_COL = process.env.START_COL || 'F';
    const END_COL = process.env.END_COL || 'AD';
    const MAX_ROWS_PER_FILE = Number(process.env.MAX_ROWS_PER_FILE || '40');
    const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
    const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;

    if (!serviceAccountJson || !SPREADSHEET_ID || !TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
      throw new Error('Missing required environment variables');
    }

    const creds = JSON.parse(serviceAccountJson);

    // === Authorize ===
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
    const accessToken = (await jwtClient.getAccessToken())?.token;
    if (!accessToken) throw new Error('Failed to obtain access token');

    const sheetsApi = google.sheets({ version: 'v4', auth: jwtClient });
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheetpdf-'));

    const metadata = await sheetsApi.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const allSheets = metadata.data.sheets || [];
    console.log(`📄 Loaded metadata for ${allSheets.length} sheets`);

    for (const sheetName of SHEET_NAMES) {
      console.log('--- Processing sheet:', sheetName);

      const sheetInfo = allSheets.find(s => s.properties?.title === sheetName);
      if (!sheetInfo) continue;
      const gid = sheetInfo.properties.sheetId;

      // ===== TIÊU ĐỀ CŨ: CHỈ LẤY F → I =====
      const headerRes = await sheetsApi.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!F5:K6`
      });
      const hVals = headerRes.data.values || [];

      const captionText = [
        hVals[0]?.[0], // F
        hVals[0]?.[1], // G
        hVals[0]?.[2], // H
        hVals[0]?.[3]  // I
      ].filter(Boolean).join('    ');

      // ===== TEXT CHÚ THÍCH: CỘT A & B =====
      const abRes = await sheetsApi.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A4:B20`
      });
      const abVals = abRes.data.values || [];
      const extraText = abVals
        .map(r => r.filter(Boolean).join(' : '))
        .filter(Boolean)
        .join('\n');

      // ===== LAST ROW (COL K) =====
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

      // ===== BUILD CHUNKS =====
      const chunks = [];
      for (let r = 1; r <= lastRow; r += MAX_ROWS_PER_FILE) {
        chunks.push({
          startRow: r,
          endRow: Math.min(r + MAX_ROWS_PER_FILE - 1, lastRow)
        });
      }

      const albumImages = [];

      for (const chunk of chunks) {
        const rangeParam = `${sheetName}!${START_COL}${chunk.startRow}:${END_COL}${chunk.endRow}`;
        const exportUrl =
          `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=pdf` +
          `&portrait=false&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false` +
          `&gridlines=false&fzr=false&gid=${gid}&range=${encodeURIComponent(rangeParam)}`;

        const pdfResp = await fetchPdfWithRetry(exportUrl, {
          Authorization: `Bearer ${accessToken}`
        });

        const pdfPath = path.join(tmpDir, `${sheetName}_${chunk.startRow}-${chunk.endRow}.pdf`);
        fs.writeFileSync(pdfPath, Buffer.from(pdfResp.data));

        const pngPath = await convertPdfToPng(pdfPath, pdfPath.replace('.pdf', ''));
        albumImages.push({
          path: pngPath,
          fileName: path.basename(pngPath)
        });
      }

      // ===== SEND ALBUM =====
      const form = new FormData();
      form.append('chat_id', TELEGRAM_CHAT_ID);
      form.append(
        'media',
        JSON.stringify(
          albumImages.map((img, i) => ({
            type: 'photo',
            media: `attach://${img.fileName}`,
            caption: i === 0 ? captionText : undefined
          }))
        )
      );
      albumImages.forEach(img => form.append(img.fileName, fs.createReadStream(img.path)));

      await axios.post(
        `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMediaGroup`,
        form,
        { headers: form.getHeaders() }
      );

      // ===== SEND TEXT =====
      if (extraText) {
        await axios.post(
          `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`,
          {
            chat_id: TELEGRAM_CHAT_ID,
            text: `${captionText}\n\n${extraText}`
          }
        );
      }

      albumImages.forEach(i => fs.unlinkSync(i.path));
    }

    fs.rmSync(tmpDir, { recursive: true, force: true });
    console.log('🎉 All sheets processed successfully');
  } catch (err) {
    console.error('ERROR:', err?.message || err);
    process.exit(1);
  }
}

main();
