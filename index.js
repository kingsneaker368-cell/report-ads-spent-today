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
    if (err.response && err.response.status === 429 && attempt < 5) {
      const delay = 3000 + Math.floor(Math.random() * 3000);
      console.log(`⚠️ Google 429 — retry ${attempt}/5 after ${delay}ms`);
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
      async (err) => {
        if (err) return reject(err);
        const pngPath = outPrefix + '.png';
        if (!fs.existsSync(pngPath)) return reject(new Error('PNG conversion failed'));

        try {
          const img = sharp(pngPath);
          const trimmed = await img.trim().toBuffer();
          await fs.promises.writeFile(pngPath, trimmed);
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
    const SHEET_NAMES = (process.env.SHEET_NAMES || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);

    const START_COL = 'F';
    const END_COL = 'AD';
    const MAX_ROWS_PER_FILE = Number(process.env.MAX_ROWS_PER_FILE || '40');

    const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
    const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;

    if (!serviceAccountJson) throw new Error('Missing GOOGLE_SERVICE_ACCOUNT_JSON');

    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(serviceAccountJson),
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    console.log(`📄 Loaded metadata for ${meta.data.sheets.length} sheets`);

    for (const sheetName of SHEET_NAMES) {
      console.log(`--- Processing sheet: ${sheetName}`);

      // ===== LẤY CAPTION A + B =====
      const captionRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A4:B20`
      });

      const captionText = (captionRes.data.values || [])
        .map(r => r.filter(Boolean).join(' '))
        .filter(Boolean)
        .join('\n');

      // ===== LẤY DATA F → AD =====
      const dataRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!${START_COL}1:${END_COL}`
      });

      const rows = dataRes.data.values || [];
      if (rows.length < 2) continue;

      const header = rows[0].slice(0, 4); // F → I
      const body = rows.slice(1);

      console.log(`➡ Export PDF for ${body.length} rows`);

      const images = [];
      let isFirstImage = true;

      for (let i = 0; i < body.length; i += MAX_ROWS_PER_FILE) {
        const chunk = body.slice(i, i + MAX_ROWS_PER_FILE);
