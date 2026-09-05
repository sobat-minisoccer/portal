# Sobat Minisoccer — Portal

Portal manajemen match dan keuangan komunitas minisoccer **Sobat Minisoccer (SMS)**.

## Setup

### 1. Google Apps Script
1. Buka Google Sheets `Minisoccer_DB`
2. Extensions → Apps Script
3. Paste isi `Code.gs`
4. Isi `SS_ID` dengan Spreadsheet ID (dari URL Sheets)
5. Deploy → New deployment → Web App → Execute as: Me → Access: Anyone
6. Copy URL deployment

### 2. Portal (GitHub Pages)
1. Buka `index.html`
2. Isi `GAS_URL` dengan URL Apps Script dari langkah 1
3. Ganti `PASSWORD` sesuai keinginan
4. Upload ke repo ini

## URL
`https://sobat-minisoccer.github.io/portal`

## Stack
- Frontend: HTML / CSS / JS (GitHub Pages)
- Backend: Google Apps Script
- Database: Google Sheets
