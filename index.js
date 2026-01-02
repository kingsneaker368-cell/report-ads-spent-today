// index.js
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile } = require('child_process');
const axios = require('axios');
const FormData = require('form-data');
const { google } = require('googleapis');
const sharp = require('sharp');

// ===== Retry Google 429 =====
async function fetchPdfWithRetry(url, headers, attempt = 1) {
  try {
    return await axios.get(url, {
      responseType: 'arraybuffer',
      headers
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

// ===== PDF → PNG =====
function convertPdfToPng(pdfPath, outPrefix) {
  return new Promise((resolve, reject) => {
    execFile(
      'pdftoppm',
      ['-png', '-singlefile', '-r', '150', pdfPath, outPrefix],
      async err => {
        if (err) return reject(err);
        const pngPath = outPrefix + '.png';
        const trimmed = await sharp(pngPath).trim().toBuffer();
        await fs.promises.writeFile(pngPath, trimmed);
        resolve(pngPath);
      }
    );
  });
}

async function main() {
  const {
    GOOGLE_SERVICE_ACCOUNT_JSON,
    SPREADSHEET_ID,
    SHEET_NAMES,
    TELEGRAM_BOT_TOKEN,
    TELEGRAM_CHAT_ID
  } = process.env;

  const creds = JSON.parse(GOOGLE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.JWT(
    creds.client_email,
    null,
    creds.private_key,
    [
      'https://www.googleapis.com/auth/drive.readonly',
      'https://www.googleapis.com/auth/spreadsheets.readonly'
    ]
  );
  await auth.authorize();

  const accessToken = (await auth.getAccessToken()).token;
  const sheetsApi = google.sheets({ version: 'v4', auth });

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheetpdf-'));
  const sheetNames = SHEET_NAMES.split(',').map(s => s.trim());

  const meta = await sheetsApi.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });

  for (const sheetName of sheetNames) {
    const sheet = meta.data.sheets.find(s => s.properties.title === sheetName);
    if (!sheet) continue;

    const gid = sheet.properties.sheetId;

    // ===== TIÊU ĐỀ F5 → J5 =====
    const headerRes = await sheetsApi.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!F5:J5`
    });

    const titleText = (headerRes.data.values?.[0] || [])
      .filter(Boolean)
      .join(' | ');

    // ===== NỘI DUNG A → B =====
    const abRes = await sheetsApi.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A4:B20`
    });

    const bodyText = (abRes.data.values || [])
      .map(r => r.filter(Boolean).join(': '))
      .filter(Boolean)
      .join('\n');

    const captionText = `${titleText}\n\n${bodyText}`;

    // ===== LAST ROW (cột K) =====
    const colRes = await sheetsApi.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!K1:2000`
    });

    const colVals = colRes.data.values || [];
    let lastRow = colVals.length;
    while (lastRow > 1 && !colVals[lastRow - 1]?.[0]) lastRow--;

    // ===== EXPORT 1 ẢNH DUY NHẤT =====
    const range = `${sheetName}!F1:AD${lastRow}`;
    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=pdf` +
      `&gid=${gid}&portrait=false&fitw=true&gridlines=false` +
      `&range=${encodeURIComponent(range)}`;

    const pdfResp = await fetchPdfWithRetry(exportUrl, {
      Authorization: `Bearer ${accessToken}`
    });

    const pdfPath = path.join(tmpDir, `${sheetName}.pdf`);
    fs.writeFileSync(pdfPath, pdfResp.data);

    const pngPath = await convertPdfToPng(pdfPath, pdfPath.replace('.pdf', ''));

    // ===== SEND 1 ẢNH + CAPTION (BÊN DƯỚI ẢNH) =====
    const form = new FormData();
    form.append('chat_id', TELEGRAM_CHAT_ID);
    form.append('caption', captionText);
    form.append('photo', fs.createReadStream(pngPath));

    await axios.post(
      `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendPhoto`,
      form,
      { headers: form.getHeaders() }
    );

    fs.unlinkSync(pdfPath);
    fs.unlinkSync(pngPath);
  }

  fs.rmSync(tmpDir, { recursive: true, force: true });
  console.log('✅ Hoàn tất – mỗi sheet gửi 1 ảnh + caption bên dưới');
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
