// ============================================================
// Code.gs — Google Apps Script untuk Sobat Minisoccer Portal
// Deploy sebagai: Web App | Execute as: Me | Access: Anyone
// ============================================================

const SS_ID   = '1z6nbrxSHzDnW3WNVBWX4mP_f5rqJSh3Z27GJWftuq84'; // ← ISI dengan Spreadsheet ID dari Google Sheets
const SS      = () => SpreadsheetApp.openById(SS_ID);

const SHEET = {
  PLAYER   : 'DB_Player',
  JERSEY   : 'Stok_Jersey',
  SHTM     : 'SHTM_Log',
  REKAP    : 'Rekap_Keuangan',
  DEPOSIT  : 'Deposit_Log',
  INVENTARIS: 'Inventaris',
};

// ── CORS & Router ──────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  let result;
  try {
    switch(action) {
      case 'getPlayers':       result = getPlayers(); break;
      case 'getMatches':       result = getMatches(); break;
      case 'getMatch':         result = getMatch(e.parameter.id); break;
      case 'getStokJersey':    result = getStokJersey(); break;
      case 'getShtm':          result = getShtm(); break;
      case 'getDeposit':       result = getDeposit(); break;
      case 'getDashboard':     result = getDashboard(); break;
      case 'getConfig':        result = getConfig(); break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const action  = payload.action || '';
  let result;
  try {
    switch(action) {
      case 'createMatch':       result = createMatch(payload); break;
      case 'addPlayer':         result = addPlayerToMatch(payload); break;
      case 'updatePlayer':      result = updatePlayerInMatch(payload); break;
      case 'deletePlayer':      result = deletePlayerFromMatch(payload); break;
      case 'addBiaya':          result = addBiaya(payload); break;
      case 'updateBiaya':       result = updateBiaya(payload); break;
      case 'deleteBiaya':       result = deleteBiaya(payload); break;
      case 'addGol':            result = addGol(payload); break;
      case 'deleteGol':         result = deleteGol(payload); break;
      case 'flagShtm':          result = flagShtm(payload); break;
      case 'useShtm':           result = useShtm(payload); break;
      case 'addDeposit':        result = addDeposit(payload); break;
      case 'useDeposit':        result = useDeposit(payload); break;
      case 'closeMatch':        result = closeMatch(payload); break;
      case 'updateStokJersey':  result = updateStokJersey(payload); break;
      case 'addPlayer':         result = addNewPlayer(payload); break;
      case 'updateConfig':      result = updateConfig(payload); break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Helpers ────────────────────────────────────────────────
function getSheet(name) {
  return SS().getSheetByName(name);
}

function sheetToObjects(sheetName, headerRow) {
  const ws   = getSheet(sheetName);
  const data = ws.getDataRange().getValues();
  const hdrs = data[headerRow - 1].map(h => String(h).trim());
  const rows = [];
  for (let r = headerRow; r < data.length; r++) {
    const row = data[r];
    if (!row[0] && !row[1]) continue; // skip empty rows
    const obj = {};
    hdrs.forEach((h, i) => { obj[h] = row[i] === '' ? null : row[i]; });
    obj._row = r + 1; // 1-based row index for updates
    rows.push(obj);
  }
  return rows;
}

function generateId(prefix) {
  const d  = new Date();
  const ts = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyyMMdd');
  return prefix + '-' + ts + '-' + Math.floor(Math.random()*1000);
}

function today() {
  return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM yyyy');
}

// ── GET: Dashboard aggregates ───────────────────────────────
function getDashboard() {
  const players = getPlayers();
  const matches = getMatches();
  const shtm    = getShtm();
  const jersey  = getStokJersey();

  // Kumulatif dari baris BASELINE / baris terakhir
  const closed = matches.filter(m => m['Status Match'] === 'Ditutup');
  const last   = closed[closed.length - 1] || {};

  // Top scorer all-time
  const scorerMap = {};
  closed.forEach(m => {
    if (m._scorers) {
      m._scorers.forEach(s => {
        scorerMap[s.nama] = (scorerMap[s.nama] || 0) + (s.gol || 0);
      });
    }
  });
  const topScorer = Object.entries(scorerMap)
    .sort((a,b) => b[1]-a[1])
    .slice(0,10)
    .map(([nama,gol]) => ({nama,gol}));

  // Frequent players
  const freqMap = {};
  players.forEach(p => {
    freqMap[p['Nama Player']] = p['Total\nHadir'] || p['Total Hadir'] || 0;
  });
  const frequent = Object.entries(freqMap)
    .sort((a,b) => b[1]-a[1])
    .slice(0,9)
    .map(([nama,cnt]) => ({nama,cnt}));

  return {
    totalMatch  : closed.length,
    totalPlayer : players.length,
    kumulatifMalik: last['Kumul Malik\n(Rp)'] || last['Kumul Malik'] || 0,
    kumulatifFilan: last['Kumul Filan\n(Rp)'] || last['Kumul Filan'] || 0,
    topScorer,
    frequent,
    matches: closed.slice(-12).map(m => ({
      tgl   : m['Match / Tanggal'],
      jenis : m['Jenis\nGame'] || m['Jenis Game'],
      tipe  : m['Tipe Event'],
      margin: m['Margin\n(Rp)'] || m['Margin'],
      lap   : m['Lap.'],
    }))
  };
}

// ── GET: Players ────────────────────────────────────────────
function getPlayers() {
  return sheetToObjects(SHEET.PLAYER, 2);
}

// ── GET: Matches (dari sheet Rekap_Keuangan + sheet per-match jika ada) ──
function getMatches() {
  return sheetToObjects(SHEET.REKAP, 2)
    .filter(m => m['No'] && String(m['No']).match(/^\d+$/));
}

function getMatch(matchId) {
  // Cari sheet dengan nama match (misal: "Match-20260920-001")
  const ss = SS();
  const ws = ss.getSheetByName(matchId);
  if (!ws) return { error: 'Match sheet not found: ' + matchId };
  const data = ws.getDataRange().getValues();
  // Parse struktur sheet match
  return parseMatchSheet(data, matchId);
}

function parseMatchSheet(data, matchId) {
  // Info event: row 3-13 col C
  const info = {};
  const INFO_LABELS = ['Match ID','Tanggal','Venue','Jenis Game','Tipe Event',
                       'Tipe Lapangan','HTM Player','HTM GK','Kuota Kiper','Kuota Field','Status Match'];
  INFO_LABELS.forEach((lbl, i) => { info[lbl] = data[2+i] ? data[2+i][2] : null; });

  // Estimasi margin: row setelah info
  const margin = {
    salesPasti    : data[16] ? data[16][2] : 0,
    salesEstimasi : data[17] ? data[17][2] : 0,
    totalCost     : data[18] ? data[18][2] : 0,
    marginProyeksi: data[19] ? data[19][2] : 0,
    splitMalik    : data[20] ? data[20][2] : 0,
    splitFilan    : data[21] ? data[21][2] : 0,
  };

  // Daftar player: cari header row "No | Nama Player | ..."
  let playerStartRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (data[r][0] === 'No' && data[r][1] === 'Nama Player') {
      playerStartRow = r + 1; break;
    }
  }
  const players = [];
  if (playerStartRow > 0) {
    const hdrs = data[playerStartRow-1];
    for (let r = playerStartRow; r < data.length; r++) {
      if (!data[r][1]) break;
      const obj = {};
      hdrs.forEach((h,i) => { obj[h] = data[r][i]; });
      obj._row = r + 1;
      players.push(obj);
    }
  }

  return { matchId, info, margin, players };
}

// ── GET: Stok Jersey ────────────────────────────────────────
function getStokJersey() {
  return sheetToObjects(SHEET.JERSEY, 2);
}

// ── GET: SHTM ───────────────────────────────────────────────
function getShtm() {
  return sheetToObjects(SHEET.SHTM, 2)
    .filter(s => s['Nama Player']);
}

// ── GET: Deposit ────────────────────────────────────────────
function getDeposit() {
  return sheetToObjects(SHEET.DEPOSIT, 2);
}

// ── GET: Config ─────────────────────────────────────────────
function getConfig() {
  // Ambil dari sheet Rekap baris BASELINE untuk kumul
  const rekap = getMatches();
  const last  = rekap[rekap.length - 1] || {};
  return {
    htmPlayer  : 95000,
    htmGk      : 35000,
    htmShtm    : 50000,
    splitMalik : 60,
    kumulMalik : last['Kumul Malik\n(Rp)'] || last['Kumul Malik'] || 8147730,
    kumulFilan : last['Kumul Filan\n(Rp)'] || last['Kumul Filan'] || 5372762,
  };
}

// ── POST: Create Match ──────────────────────────────────────
function createMatch(p) {
  const ss      = SS();
  const tmpl    = ss.getSheetByName(SHEET.TEMPLATE || 'TEMPLATE_Match');
  if (!tmpl) return { error: 'TEMPLATE_Match sheet not found' };

  const matchId = generateId('Match');
  const newWs   = tmpl.copyTo(ss);
  newWs.setName(matchId);

  // Isi info event (col C, row 3-13)
  const vals = [
    matchId, p.tanggal, p.venue, p.jenisGame, p.tipeEvent,
    p.tipeLapangan, p.htmPlayer || 95000, p.htmGk || 35000,
    p.kuotaKiper || 4, p.kuotaField || 24, 'Aktif'
  ];
  vals.forEach((v, i) => { newWs.getRange(3+i, 3).setValue(v); });

  return { ok: true, matchId };
}

// ── POST: Add Player to Match ───────────────────────────────
function addPlayerToMatch(p) {
  const ws   = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found: ' + p.matchId };

  // Cari baris pertama kosong setelah header player
  const data = ws.getDataRange().getValues();
  let insertRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (data[r][0] === 'No' && data[r][1] === 'Nama Player') {
      // Cari baris kosong setelah header
      for (let rr = r+1; rr < data.length; rr++) {
        if (!data[rr][1]) { insertRow = rr+1; break; }
      }
      if (insertRow < 0) insertRow = data.length + 1;
      break;
    }
  }
  if (insertRow < 0) return { error: 'Player table not found in sheet' };

  const rowData = [
    p.no || '', p.nama, p.statusBayar || 'Sementara',
    p.htm, p.labelHarga || 'Normal',
    p.depositDipakai || 0, p.metodeBayar || '— (belum)',
    p.jumlahBayar || 0, p.tglBayar || '',
    p.posisi || 'Field', p.jerseyWarna || '', p.ukuran || '',
    p.tim || '', p.formasi || '', p.shtmFlag || '—', p.catatan || ''
  ];
  ws.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);

  // Kurangi stok jersey jika warna dipilih
  if (p.jerseyWarna && p.jerseyWarna !== 'Pribadi') {
    updateJerseyStok(p.jerseyWarna, p.ukuran, 1);
  }
  // Pakai deposit jika metode = Deposit
  if (p.metodeBayar === 'Deposit' && p.depositDipakai > 0) {
    useDeposit({ namaPemain: p.nama, jumlah: p.depositDipakai, matchId: p.matchId });
  }
  return { ok: true, row: insertRow };
}

// ── POST: Update Player ─────────────────────────────────────
function updatePlayerInMatch(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  const rowData = [
    p.no, p.nama, p.statusBayar, p.htm, p.labelHarga,
    p.depositDipakai || 0, p.metodeBayar, p.jumlahBayar || 0, p.tglBayar || '',
    p.posisi, p.jerseyWarna, p.ukuran, p.tim, p.formasi, p.shtmFlag, p.catatan || ''
  ];
  ws.getRange(p.rowIdx, 1, 1, rowData.length).setValues([rowData]);
  return { ok: true };
}

// ── POST: Delete Player ─────────────────────────────────────
function deletePlayerFromMatch(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  // Kembalikan stok jersey
  const data = ws.getRange(p.rowIdx, 1, 1, 16).getValues()[0];
  const warna = data[10]; const uk = data[11];
  if (warna && warna !== 'Pribadi') updateJerseyStok(warna, uk, -1);
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Biaya ─────────────────────────────────────────────
function addBiaya(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  // Cari section biaya di sheet (label "Item Biaya" dst.)
  // Append di baris kosong pertama di section biaya
  const data = ws.getDataRange().getValues();
  let insertRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).toLowerCase().includes('item biaya') ||
        String(data[r][0]).toLowerCase().includes('nama item')) {
      for (let rr = r+1; rr < data.length; rr++) {
        if (!data[rr][0]) { insertRow = rr+1; break; }
      }
      break;
    }
  }
  if (insertRow < 0) insertRow = ws.getLastRow() + 1;
  const rowData = [p.nama, p.kategori, p.tipe, p.rencana || 0, p.realisasi || 0,
                   p.skema || '', p.jmlCicilan || '', p.catatan || ''];
  ws.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);
  return { ok: true, row: insertRow };
}

function updateBiaya(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  const rowData = [p.nama, p.kategori, p.tipe, p.rencana||0, p.realisasi||0,
                   p.skema||'', p.jmlCicilan||'', p.catatan||''];
  ws.getRange(p.rowIdx, 1, 1, rowData.length).setValues([rowData]);
  return { ok: true };
}

function deleteBiaya(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Goal ──────────────────────────────────────────────
function addGol(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  // Simpan gol di section goal scorer
  const data = ws.getDataRange().getValues();
  let insertRow = ws.getLastRow() + 1;
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).toLowerCase().includes('pencetak') ||
        String(data[r][0]).toLowerCase().includes('goal scorer')) {
      for (let rr = r+1; rr < data.length; rr++) {
        if (!data[rr][0]) { insertRow = rr+1; break; }
      }
      break;
    }
  }
  ws.getRange(insertRow, 1, 1, 3).setValues([[p.nama, p.gol, p.keterangan || '']]);
  return { ok: true };
}

function deleteGol(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Flag SHTM ─────────────────────────────────────────
function flagShtm(p) {
  const ws  = getSheet(SHEET.SHTM);
  const lastRow = ws.getLastRow() + 1;
  const no  = lastRow - 2; // nomor urut
  ws.getRange(lastRow, 1, 1, 7).setValues([[
    no, p.nama, p.tglMatch, null, 'Belum dipakai',
    p.matchLabel || p.tglMatch, p.catatan || 'Flag dari Slot Player'
  ]]);
  return { ok: true };
}

// ── POST: Use SHTM (saat player hadir dgn SHTM di match berikutnya) ──
function useShtm(p) {
  const ws   = getSheet(SHEET.SHTM);
  const data = ws.getDataRange().getValues();
  for (let r = 2; r < data.length; r++) {
    if (data[r][1] === p.nama && data[r][4] === 'Belum dipakai') {
      ws.getRange(r+1, 4).setValue(p.tglDipakai);
      ws.getRange(r+1, 5).setValue('Sudah dipakai');
      ws.getRange(r+1, 6).setValue(p.matchLabel || p.tglDipakai);
      return { ok: true, row: r+1 };
    }
  }
  return { error: 'SHTM aktif tidak ditemukan untuk: ' + p.nama };
}

// ── POST: Deposit ───────────────────────────────────────────
function addDeposit(p) {
  const ws      = getSheet(SHEET.DEPOSIT);
  const lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    p.tgl || today(), p.nama, p.matchId || '—',
    p.sumber || 'Cancel → Deposit', p.keterangan || '',
    p.jumlah, 0, p.jumlah
  ]]);
  return { ok: true };
}

function useDeposit(p) {
  const ws   = getSheet(SHEET.DEPOSIT);
  const data = ws.getDataRange().getValues();
  // Hitung saldo aktif player
  let saldo = 0;
  for (let r = 2; r < data.length; r++) {
    if (data[r][1] === p.namaPemain) {
      saldo += (data[r][5] || 0) - (data[r][6] || 0);
    }
  }
  if (saldo < p.jumlah) return { error: 'Saldo deposit tidak cukup' };
  const lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    today(), p.namaPemain, p.matchId,
    'Digunakan', 'Bayar HTM', 0, p.jumlah, saldo - p.jumlah
  ]]);
  return { ok: true, saldoSisa: saldo - p.jumlah };
}

// ── POST: Stok Jersey ───────────────────────────────────────
function updateJerseyStok(warna, ukuran, delta) {
  // delta: +1 = tambah dipakai, -1 = kembalikan
  const ws   = getSheet(SHEET.JERSEY);
  const data = ws.getDataRange().getValues();
  for (let r = 2; r < data.length; r++) {
    if (data[r][0] === warna && data[r][2] === ukuran) {
      const dipakai = (data[r][4] || 0) + delta;
      ws.getRange(r+1, 5).setValue(Math.max(0, dipakai));
      ws.getRange(r+1, 6).setValue((data[r][3] || 0) - Math.max(0, dipakai));
      const sisa = (data[r][3] || 0) - Math.max(0, dipakai);
      ws.getRange(r+1, 7).setValue(sisa === 0 ? 'HABIS' : sisa <= 1 ? 'LOW' : 'OK');
      return { ok: true };
    }
  }
  return { error: 'Jersey tidak ditemukan: ' + warna + ' ' + ukuran };
}

function updateStokJersey(p) {
  const ws   = getSheet(SHEET.JERSEY);
  const data = ws.getDataRange().getValues();
  // Cari baris yang ada
  for (let r = 2; r < data.length; r++) {
    if (data[r][0] === p.warna && data[r][1] === p.tipe && data[r][2] === p.ukuran) {
      const newStok = (data[r][3] || 0) + (p.tambah || 0);
      ws.getRange(r+1, 4).setValue(newStok);
      const sisa = newStok - (data[r][4] || 0);
      ws.getRange(r+1, 6).setValue(sisa);
      ws.getRange(r+1, 7).setValue(sisa === 0 ? 'HABIS' : sisa <= 1 ? 'LOW' : 'OK');
      if (p.keterangan) ws.getRange(r+1, 8).setValue(p.keterangan);
      return { ok: true, newStok };
    }
  }
  // Warna/ukuran baru — tambah baris
  const lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    p.warna, p.tipe || 'Field', p.ukuran,
    p.tambah || 1, 0, p.tambah || 1,
    p.tambah > 1 ? 'OK' : 'LOW',
    p.keterangan || ''
  ]]);
  return { ok: true, newRow: lastRow };
}

// ── POST: Close Match ───────────────────────────────────────
function closeMatch(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };

  // 1. Set status match = Ditutup
  const data = ws.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    if (data[r][0] === 'Status Match') {
      ws.getRange(r+1, 3).setValue('Ditutup');
      break;
    }
  }

  // 2. Hitung sales, cost, margin dari data player & biaya
  const sales  = p.totalSales  || 0;
  const cost   = p.totalCost   || 0;
  const margin = sales - cost;
  const sm     = Math.round(margin * 0.6);
  const sf     = Math.round(margin * 0.4);
  const inv    = p.totalInv    || 0;
  const nm     = sm - Math.round(inv * 0.6);
  const nf     = sf - Math.round(inv * 0.4);

  // 3. Tambah baris ke Rekap_Keuangan
  const rekap  = getSheet(SHEET.REKAP);
  // Cari baris BASELINE atau last row
  const rdata  = rekap.getDataRange().getValues();
  let insertRow = rekap.getLastRow() + 1;
  for (let r = rdata.length-1; r >= 0; r--) {
    if (rdata[r][0] === 'BASELINE') { insertRow = r+1; break; }
  }

  // Kumul dari baris sebelumnya
  const prevRow  = rdata[insertRow-2] || [];
  const prevKumM = (prevRow[14] && typeof prevRow[14] === 'number') ? prevRow[14] : 8147730;
  const prevKumF = (prevRow[15] && typeof prevRow[15] === 'number') ? prevRow[15] : 5372762;
  const kumM     = prevKumM + nm;
  const kumF     = prevKumF + nf;

  const matchNo = (insertRow - 3); // nomor match baru
  rekap.getRange(insertRow, 1, 1, 22).setValues([[
    matchNo, p.tglMatch, p.jenisGame, p.tipeEvent,
    p.venue, p.tipeLapangan,
    sales, cost, margin, sm, sf,
    inv, nm, nf, kumM, kumF,
    p.bca||0, p.bsi||0, p.mandiri||0, p.bri||0, p.cash||0,
    ''
  ]]);

  // 4. Kembalikan semua stok jersey dari match ini
  // (data player sudah dikurangi saat addPlayer, dikembalikan jika deletePlayer)
  // Saat close, jersey dianggap dikembalikan (stok dipakai → 0 lagi)
  if (p.jerseyUsed && Array.isArray(p.jerseyUsed)) {
    p.jerseyUsed.forEach(j => updateJerseyStok(j.warna, j.uk, -1));
  }

  return {
    ok: true, matchNo, sales, cost, margin,
    splitMalik: sm, splitFilan: sf,
    kumulMalik: kumM, kumulFilan: kumF
  };
}

// ── POST: Update Config ─────────────────────────────────────
function updateConfig(p) {
  // Simpan config ke PropertiesService (persisten, tidak di Sheets)
  const props = PropertiesService.getScriptProperties();
  if (p.htmPlayer)  props.setProperty('HTM_PLAYER',  String(p.htmPlayer));
  if (p.htmGk)      props.setProperty('HTM_GK',      String(p.htmGk));
  if (p.htmShtm)    props.setProperty('HTM_SHTM',    String(p.htmShtm));
  if (p.splitMalik) props.setProperty('SPLIT_MALIK', String(p.splitMalik));
  return { ok: true };
}

// ── POST: Add New Player to DB ──────────────────────────────
function addNewPlayer(p) {
  const ws      = getSheet(SHEET.PLAYER);
  const lastRow = ws.getLastRow() + 1;
  const no      = lastRow - 2;
  ws.getRange(lastRow, 1, 1, 17).setValues([[
    no, p.nama, p.kategori || '—', 0, 0, 0,
    p.ukuran || '—', '—', 0, '—', '—', '—', p.jerseyPribadi || '—',
    '—', 'Aktif', p.catatan || '', today()
  ]]);
  return { ok: true, row: lastRow };
}
