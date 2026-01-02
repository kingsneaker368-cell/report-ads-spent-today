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
    const MAX_ROWS_PER_FILE = Number(proces_
