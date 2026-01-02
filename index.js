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
      headers,
      maxContentLength: Infinity,
      maxBodyLength: Infinity
    });
  } catch (err) {
    if (err.response?.status === 429 && attempt < 5) {
      const delay = 3000 + Math.random() * 3000;
      console.log(`⚠️ Google 429 retry ${attempt}/5 after ${delay}ms`);
      await new Promise(r => setTimeout(r, delay));
      return fetchPdfWithRetry(url, headers, attempt + 1);
    }
    throw err;
  }
}

// ===== PDF → PNG + trim =====
function convertPdfToPng(pdfPath, outPrefix) {
  return new Promise((resolve, reject) => {
    execFile(
      'pdftoppm',
      ['-png', '-singlefile', '-r', '150', pdfPath, outPrefix],
      async err => {
        if (err) return reject(err);
        const pngPath = `${outPrefix}.png`;
        if (!fs.existsSync(pngPath)) return reject(new Error('PNG not found'));

        try {
          const buf = await sharp(pngPath).trim().toBuffer();
          await fs.promises.writeFile(pngPath, buf);
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
    const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
    const SHEET_NAMES = (process.env.SHEET_NAMES || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);

    const MAX_ROWS_PER_CHUNK = Number(process.env.MAX_ROWS_PER_CHUNK || '40');

    const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
    const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;

    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    console.log(`📄 Loaded metadata for ${meta.data.sheets.length} sheets`);

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-'));

    for (const sheetName of SHEET_NAMES) {
      console.log(`--- Processing sheet: ${sheetName}`);

      // ===== CAPTION A + B =====
      const capRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A4:B20`
      });

      const captionAB = (capRes.data.values || [])
        .map(r => r.filter(Boolean).join(' '))
        .filter(Boolean)
        .join('\n');

      // ===== DATA F → AD =====
      const dataRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!F1:AD`
      });

      const rows = dataRes.data.values || [];
      if (rows.length < 2) continue;

      const header = rows[0].slice(0, 4); // F → I
      const body = rows.slice(1);

      const images = [];

      for (let i = 0; i < body.length; i += MAX_ROWS_PER_CHUNK) {
        const chunk = body.slice(i, i + MAX_ROWS_PER_CHUNK);
        const content = [header, ...chunk];

        const ex
