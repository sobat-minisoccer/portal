// ============================================================
// Code.gs — Google Apps Script untuk Sobat Minisoccer Portal
// Deploy sebagai: Web App | Execute as: Me | Access: Anyone
// ============================================================

const SS_ID   = '1z6nbrxSHzDnW3WNVBWX4mP_f5rqJSh3Z27GJWftuq84'; // Sheet "Minisoccer_DB" ASLI (akun sbt.minisoccer) — udah di-share view ke intanktrav, strukturnya udah dicek cocok sama kode di bawah
const SS      = () => SpreadsheetApp.openById(SS_ID);

const SHEET = {
  PLAYER   : 'DB_Player',
  JERSEY   : 'Stok_Jersey',
  SHTM     : 'SHTM_Log',
  REKAP    : 'Rekap_Keuangan',
  DEPOSIT  : 'Deposit_Log',
  INVENTARIS: 'Inventaris',
  SCORER   : 'Rekap_Skorer', // tab baru — BELUM ada di sheet data riil kamu juga,
                              // dibikin otomatis sama setupSheets()
  DRAFTS   : 'Match_Drafts', // tab baru — draft match Transaksi yang masih
                              // berjalan (belum ditutup), dipakai supaya
                              // "Simpan" nyambung ke browser/device lain
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
      case 'getScorer':        result = getScorer(); break;
      case 'getMatchDrafts':   result = getMatchDrafts(); break;
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
      case 'deleteDeposit':     result = deleteDeposit(payload); break;
      case 'closeMatch':        result = closeMatch(payload); break;
      case 'deleteRekapMatch':  result = deleteRekapMatch(payload); break;
      case 'updateStokJersey':  result = updateStokJersey(payload); break;
      case 'addPlayerDb':       result = addNewPlayer(payload); break;
      case 'updateConfig':      result = updateConfig(payload); break;
      case 'addScorer':         result = addScorer(payload); break;
      case 'updateScorer':      result = updateScorer(payload); break;
      case 'deleteScorer':      result = deleteScorer(payload); break;
      case 'saveMatchDraft':    result = saveMatchDraft(payload); break;
      case 'deleteMatchDraft':  result = deleteMatchDraft(payload); break;
      case 'updatePlayerDb':    result = updatePlayerDb(payload); break;
      case 'deletePlayerDb':    result = deletePlayerDb(payload); break;
      case 'updateShtmRow':     result = updateShtmRow(payload); break;
      case 'deleteShtmRow':     result = deleteShtmRow(payload); break;
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

// Kolom "tanggal-ish" di sheet ini (mis. "Match / Tanggal") ditulis sbg
// STRING teks biasa ("6 September 2026" — lihat closeMatch()), TAPI Google
// Sheets suka diam2 "membantu" ndeteksi string yg keliatan kayak tanggal
// trus nyimpen selnya jadi tipe Date ASLI (tergantung locale spreadsheet)
// walau ditulis via setValues() sbg string. Begitu kebaca lagi di sini,
// baliknya jadi objek Date JS, lalu pas di-JSON.stringify() (jsonResponse)
// otomatis jadi ISO mentah kayak "2026-09-06T07:00:00.000Z" — bukan lagi
// "6 September 2026" — begitu nyampe ke frontend (ini yg bikin chart/tabel
// Dashboard & Reporting keluar kode aneh, bukan nama tanggal/event). Jaga2
// di SATU tempat ini (dipakai semua sheetToObjects()) -- kalau ternyata
// selnya kebaca sbg Date asli, format ulang jadi teks Indonesia biasa dulu
// SEBELUM keluar dari fungsi ini, jangan biarin objek Date mentah lolos.
function formatSheetDateForOutput(val) {
  if (val instanceof Date) {
    const bulan = ['Januari','Februari','Maret','April','Mei','Juni','Juli',
                   'Agustus','September','Oktober','November','Desember'];
    return val.getDate() + ' ' + bulan[val.getMonth()] + ' ' + val.getFullYear();
  }
  return val;
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
    hdrs.forEach((h, i) => { obj[h] = row[i] === '' ? null : formatSheetDateForOutput(row[i]); });
    obj._row = r + 1; // 1-based row index for updates
    rows.push(obj);
  }
  return rows;
}

function toNum(v) {
  if (typeof v === 'number') return v;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

// ── Header-safe write helpers ───────────────────────────────
// Dipakai buat sheet yang struktur ASLINYA belum 100% confirmed cocok
// sama urutan kolom yang diasumsikan di kode ini — DB_Player & SHTM_Log
// terutama (DB_Player punya sheet real punya Sam sendiri, sebelumnya
// dari sebelumnya, header text-nya kebukti beda dari yang diasumsikan
// kode, misal "Saldo Deposit\n(Rp)" vs "Deposit Saldo (Rp)"). Daripada
// nulis ke kolom berdasarkan posisi/index tetap (yang beresiko salah
// kolom & ngerusak data Sam kalau urutan asli beda), fungsi-fungsi ini
// nyari kolom berdasarkan NAMA header dulu, baru nulis ke situ.
function getHeaderMap(sheetName, headerRow) {
  const ws   = getSheet(sheetName);
  const hdrs = ws.getRange(headerRow, 1, 1, ws.getLastColumn()).getValues()[0]
                 .map(h => String(h).trim());
  const map  = {};
  hdrs.forEach((h, i) => { if (h) map[h] = i + 1; }); // 1-based kolom
  return { ws, map, hdrs };
}
function findCol(map, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    if (map[aliases[i]] !== undefined) return map[aliases[i]];
  }
  return -1;
}
// Varian header kolom saldo deposit di DB_Player — dipakai bareng di semua
// tempat yang baca/tulis kolom ini (addNewPlayer, updatePlayerDb,
// getPlayerDepositSaldo) supaya daftarnya nggak ke-duplikat & gampang
// beda sendiri-sendiri kalau nanti ada varian baru.
const DEPOSIT_COL_ALIASES = ['Deposit Saldo (Rp)','Saldo Deposit\n(Rp)','Deposit Saldo\n(Rp)','Saldo Deposit (Rp)'];

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

  // Setiap baris di Rekap_Keuangan MEMANG match yang udah ditutup — cuma
  // closeMatch() yang nulis ke sheet ini, dan sheet ini nggak punya kolom
  // "Status Match" (itu cuma ada di TEMPLATE_Match/per-match sheet). Kode
  // lama nge-filter pakai m['Status Match'], yang selalu undefined utk
  // baris dari sini — jadi `closed` selalu kosong dan totalMatch selalu 0.
  const closed = matches;
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

  // Match terbaik (margin tertinggi, all-time — bukan cuma 12 terakhir)
  let bestMatch = null;
  closed.forEach(m => {
    const margin = toNum(m['Margin\n(Rp)'] || m['Margin']);
    if (!bestMatch || margin > bestMatch.margin) {
      bestMatch = { tgl: m['Match / Tanggal'] || '', tipe: m['Tipe Event'] || '', margin };
    }
  });

  // Delta kumulatif Malik/Filan dari N match terakhir (N=4, atau kurang
  // kalau match yang tercatat belum sampai 5) — buat subtext "Shared profit".
  const deltaN = Math.min(4, closed.length > 0 ? closed.length - 1 : 0);
  const refMatch = deltaN > 0 ? closed[closed.length - 1 - deltaN] : null;
  const kumulMalikNow  = toNum(last['Kumul Malik\n(Rp)'] || last['Kumul Malik']);
  const kumulFilanNow  = toNum(last['Kumul Filan\n(Rp)'] || last['Kumul Filan']);
  const deltaMalik = refMatch ? (kumulMalikNow - toNum(refMatch['Kumul Malik\n(Rp)'] || refMatch['Kumul Malik'])) : kumulMalikNow;
  const deltaFilan = refMatch ? (kumulFilanNow - toNum(refMatch['Kumul Filan\n(Rp)'] || refMatch['Kumul Filan'])) : kumulFilanNow;

  return {
    totalMatch  : closed.length,
    totalPlayer : players.length,
    lastMatchTgl: last['Match / Tanggal'] || '',
    kumulatifMalik: kumulMalikNow,
    kumulatifFilan: kumulFilanNow,
    deltaMalik, deltaFilan, deltaN,
    bestMatch,
    topScorer,
    frequent,
    // Riwayat SEMUA match (bukan cuma 12 terakhir) — biar chart "Margin per
    // match" bisa switch 12/24/Semua tanpa perlu fetch ulang.
    matches: closed.map(m => ({
      tgl   : m['Match / Tanggal'],
      jenis : m['Jenis\nGame'] || m['Jenis Game'],
      tipe  : m['Tipe Event'],
      margin: toNum(m['Margin\n(Rp)'] || m['Margin']),
      lap   : m['Lap.'],
      tahun : m['Tahun'] || '',
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

  // Estimasi margin: row 16-21 di TEMPLATE_Match asli (row 15 = judul section
  // "ESTIMASI MARGIN", row 16-21 = 6 field-nya) — sebelumnya kode ini baca
  // row 17-22 (kegeser 1 baris), dibenerin biar cocok sama sheet aslinya.
  const margin = {
    salesPasti    : data[15] ? data[15][2] : 0,
    salesEstimasi : data[16] ? data[16][2] : 0,
    totalCost     : data[17] ? data[17][2] : 0,
    marginProyeksi: data[18] ? data[18][2] : 0,
    splitMalik    : data[19] ? data[19][2] : 0,
    splitFilan    : data[20] ? data[20][2] : 0,
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
  // HTM/split disimpan di PropertiesService (ditulis updateConfig()).
  // Dulu fungsi ini selalu balikin angka hardcoded & GAK PERNAH baca
  // baliknya — jadi "Simpan konfigurasi" nulis, tapi ke-load-nya tetep
  // angka lama terus. Dibenerin: baca PropertiesService dulu, fallback
  // ke default kalau belum pernah di-set.
  const props = PropertiesService.getScriptProperties();
  // Kumul tetap dari baris terakhir Rekap_Keuangan (itu sumber kebenarannya)
  const rekap = getMatches();
  const last  = rekap[rekap.length - 1] || {};
  return {
    htmPlayer  : toNum(props.getProperty('HTM_PLAYER'))  || 95000,
    htmGk      : toNum(props.getProperty('HTM_GK'))      || 35000,
    htmShtm    : toNum(props.getProperty('HTM_SHTM'))    || 50000,
    splitMalik : toNum(props.getProperty('SPLIT_MALIK')) || 60,
    kumulMalik : last['Kumul Malik\n(Rp)'] || last['Kumul Malik'] || 8147730,
    kumulFilan : last['Kumul Filan\n(Rp)'] || last['Kumul Filan'] || 5372762,
  };
}

// ── GET/SET: Match Drafts ───────────────────────────────────
// Draft match Transaksi yang masih berjalan (belum ditutup) — supaya
// tombol "Simpan" di satu browser/device kelihatan lagi kalau portal
// dibuka di browser/device lain, bukan cuma localStorage browser itu
// doang. 1 baris per match; kolom "Data" isinya JSON blob snapshot match
// itu (persis bentuk yang tadinya cuma disimpan ke localStorage lewat
// msPersist() di index.html — title, tabel slot/biaya/deposit sbg HTML
// string, dll). Match yang sudah ditutup TETAP ikut tersimpan di sini apa
// adanya (sama kayak localStorage) — Match_Drafts bukan sumber kebenaran
// finansial (itu tetap Rekap_Keuangan), cuma cerminan draft/riwayat lokal
// Transaksi biar nyambung antar browser.
//
// Match yang dihapus TIDAK langsung dibuang barisnya (lihat deleteMatchDraft)
// -- kalau langsung dibuang, browser/device LAIN yang localStorage-nya masih
// punya salinan lama match itu tidak akan pernah tahu match itu sudah
// dihapus (baris di sheet ini kosong terlihat sama persis dengan "belum
// pernah dipakai sama sekali", dan frontend sengaja tidak membersihkan
// draft lokal kalau sheet ini kosong -- supaya draft yang belum sempat
// ke-"Simpan" tidak ikut kehapus kalau kebetulan sheet lagi kosong/offline).
// Solusinya: baris diubah jadi TOMBSTONE (Data = {"__deleted":true, ...})
// bukan dihapus fisik, supaya browser lain yang narik getMatchDrafts() bisa
// bedakan "memang belum pernah ada draft sama sekali" vs "match ini pernah
// ada tapi sudah sengaja dihapus" -- lihat msSyncFromBackend di index.html.
// Tombstone lama (lebih dari MATCH_DRAFT_TOMBSTONE_TTL_DAYS) dibersihkan
// otomatis tiap kali getMatchDrafts() dipanggil, setelah cukup waktu buat
// nyebar ke semua device yang mungkin lagi dibuka.
const MATCH_DRAFT_TOMBSTONE_TTL_DAYS = 14;
// Google Sheets nolak isi 1 SEL lebih dari 50.000 karakter. Match yang
// pemainnya udah banyak (jersey+No Show+formasi dst per baris) gampang
// nembus itu (match 28 pemain terisi penuh bisa 60rb+ karakter) -- begitu
// kena limit, setValues() di saveMatchDraft() gagal diam2 dari sisi Sheets,
// browser cuma lihat "gagal kirim ke database" (CORS/network error) tanpa
// tahu sebabnya beneran limit karakter. Solusi: JSON-nya dipecah rata ke
// beberapa KOLOM (bukan 1 kolom), disambung lagi otomatis pas dibaca.
const MATCH_DRAFT_CHUNK_SIZE = 45000;
const MATCH_DRAFT_CHUNK_COLS = 8; // kolom B..I -> kapasitas total ~360rb karakter, kolom J = UpdatedAt
function chunkStringForSheet_(str, size) {
  const chunks = [];
  for (let i = 0; i < str.length; i += size) chunks.push(str.slice(i, i + size));
  while (chunks.length < MATCH_DRAFT_CHUNK_COLS) chunks.push('');
  return chunks;
}
// Baca semua baris Match_Drafts, otomatis sambung ulang isi kolom Data yang
// ke-pecah, TERMASUK baris lama yang masih format sebelum fitur ini ada
// (1 kolom "Data" polos, kolom C = UpdatedAt) -- supaya draft yang sudah
// kepenuhin (mis. sudah diisi 20+ pemain) TIDAK hilang cuma gara2 migrasi
// struktur kolom ini. Baris lama otomatis ke-tulis ulang ke format baru
// begitu match itu di-"Simpan"/disentuh lagi lewat saveMatchDraft().
function readMatchDraftRows_() {
  const ws = getSheet(SHEET.DRAFTS);
  const data = ws.getDataRange().getValues();
  const rows = [];
  for (let r = 2; r < data.length; r++) {
    const row = data[r];
    if (!row[0]) continue;
    const newFormatUpdatedAt = row[1 + MATCH_DRAFT_CHUNK_COLS]; // kolom J
    let json, updatedAt;
    if (newFormatUpdatedAt) {
      json = row.slice(1, 1 + MATCH_DRAFT_CHUNK_COLS).join('');
      updatedAt = newFormatUpdatedAt;
    } else {
      // Format lama: Data cuma di kolom B, UpdatedAt di kolom C
      json = row[1] || '';
      updatedAt = row[2];
    }
    rows.push({ MatchId: row[0], Data: json, UpdatedAt: updatedAt, _row: r + 1 });
  }
  return rows;
}
function getMatchDrafts() {
  pruneOldMatchDraftTombstones();
  return readMatchDraftRows_();
}
function saveMatchDraft(p) {
  const id = p.matchId;
  if (!id) return { error: 'matchId wajib diisi' };
  const ws = getSheet(SHEET.DRAFTS);
  const data = ws.getDataRange().getValues();
  let targetRow = -1;
  for (let r = 2; r < data.length; r++) { // data mulai row 3 (index 2)
    if (String(data[r][0]) === String(id)) { targetRow = r + 1; break; }
  }
  const now = new Date();
  const json = JSON.stringify(p.data || {});
  if (json.length > MATCH_DRAFT_CHUNK_SIZE * MATCH_DRAFT_CHUNK_COLS) {
    return { error: 'Data match ini kepanjangan buat disimpan (' + json.length + ' karakter) — hubungi admin.' };
  }
  const row = [id, ...chunkStringForSheet_(json, MATCH_DRAFT_CHUNK_SIZE), now];
  if (targetRow === -1) {
    const lastRow = Math.max(ws.getLastRow(), 2);
    ws.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
  } else {
    ws.getRange(targetRow, 1, 1, row.length).setValues([row]);
  }
  return { ok: true, matchId: id, updatedAt: now.toISOString() };
}
function deleteMatchDraft(p) {
  const id = p.matchId;
  if (!id) return { error: 'matchId wajib diisi' };
  const ws = getSheet(SHEET.DRAFTS);
  const data = ws.getDataRange().getValues();
  const now = new Date();
  const json = JSON.stringify({ __deleted: true, deletedAt: now.toISOString() });
  const tombstoneRow = [id, ...chunkStringForSheet_(json, MATCH_DRAFT_CHUNK_SIZE), now];
  let targetRow = -1;
  for (let r = 2; r < data.length; r++) {
    if (String(data[r][0]) === String(id)) { targetRow = r + 1; break; }
  }
  if (targetRow === -1) {
    const lastRow = Math.max(ws.getLastRow(), 2);
    ws.getRange(lastRow + 1, 1, 1, tombstoneRow.length).setValues([tombstoneRow]);
  } else {
    ws.getRange(targetRow, 1, 1, tombstoneRow.length).setValues([tombstoneRow]);
  }
  return { ok: true, matchId: id };
}
function pruneOldMatchDraftTombstones() {
  const ws = getSheet(SHEET.DRAFTS);
  const rows = readMatchDraftRows_();
  const cutoff = new Date().getTime() - MATCH_DRAFT_TOMBSTONE_TTL_DAYS * 24 * 60 * 60 * 1000;
  // hapus dari bawah ke atas biar _row baris yang belum diproses gak geser
  rows.sort((a, b) => b._row - a._row);
  rows.forEach(row => {
    let parsed;
    try { parsed = JSON.parse(row.Data); } catch (e) { return; }
    if (!parsed || !parsed.__deleted) return;
    const updatedAt = row.UpdatedAt ? new Date(row.UpdatedAt).getTime() : 0;
    if (updatedAt && updatedAt < cutoff) ws.deleteRow(row._row);
  });
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
// Kolom "Saldo Deposit" di DB_Player adalah SATU-SATUNYA angka saldo yang
// dipercaya (bukan dihitung ulang dari jumlah transaksi Deposit_Log) —
// soalnya sebagian saldo player itu asalnya dari data lama/manual (diisi
// langsung ke DB_Player, bukan lewat transaksi yang tercatat), jadi kalau
// saldo dihitung ulang dari NOL berdasarkan ledger doang, saldo lama itu
// bisa keitung 0 padahal beneran ada — malah bikin "Saldo deposit tidak
// cukup" muncul padahal saldonya ada. Deposit_Log tetap jalan sbg BUKU
// CATATAN/riwayat transaksi (kapan, berapa, match mana — buat tab Deposit
// di Admin), tapi bukan lagi sumber penghitungan saldo aktif.
// addDeposit = +jumlah ke saldo DB_Player. useDeposit = -jumlah (ditolak
// kalau saldo DB_Player SAAT INI kurang dari yang mau dipakai). Dua-duanya
// nulis baris riwayat ke Deposit_Log DULU baru nge-update saldo DB_Player,
// biar riwayatnya tetep lengkap konsisten sama saldo akhirnya.
function getPlayerDepositSaldo(nama) {
  const { ws, map } = getHeaderMap(SHEET.PLAYER, 2);
  const nameCol    = findCol(map, ['Nama Player','Nama']);
  const depositCol = findCol(map, DEPOSIT_COL_ALIASES);
  if (nameCol < 1 || depositCol < 1) return { ws, rowIdx: -1, depositCol: -1, saldo: 0 };
  const lastRow = ws.getLastRow();
  if (lastRow >= 3) {
    const names = ws.getRange(3, nameCol, lastRow - 2, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() === String(nama).trim()) {
        const rowIdx = i + 3;
        const saldo  = toNum(ws.getRange(rowIdx, depositCol).getValue());
        return { ws, rowIdx, depositCol, saldo };
      }
    }
  }
  // Player belum ada row-nya di DB_Player (mis. walk-in yg blm pernah
  // didaftarin ke DB) — dianggap saldo 0, nggak ada tempat nyimpen cache.
  return { ws, rowIdx: -1, depositCol, saldo: 0 };
}
function adjustPlayerDepositSaldo(nama, delta) {
  const info = getPlayerDepositSaldo(nama);
  if (info.rowIdx < 0) {
    // Belum ada row DB_Player buat nama ini. Kalau ini nambah saldo (delta
    // positif — dari addDeposit/topup, mis. hasil "Cancel → Deposit" buat
    // player yg blm pernah didaftarin manual ke DB), bikinin row baru
    // otomatis — biar player ini muncul di Kelola Player & saldo depositnya
    // nggak "gantung" (ada transaksinya tapi nggak nongol di mana pun).
    // Kalau ngurangin (delta negatif — dari useDeposit/deleteDeposit),
    // nggak ada row utk ditulis, nggak mungkin ada apa2 yg dikurangin, jadi
    // dibiarin aja (useDeposit sendiri udah nolak duluan krn saldo 0).
    if (delta > 0) {
      writeNewPlayerRow({ nama, deposit: delta,
        catatan: 'Otomatis dibuat dari transaksi deposit — blm pernah didaftarkan manual' });
      return delta;
    }
    return null;
  }
  if (info.depositCol < 1) return null;
  const next = Math.max(0, info.saldo + delta);
  info.ws.getRange(info.rowIdx, info.depositCol).setValue(next);
  return next;
}
// ── Perkakas SEKALI-JALAN (dari editor Apps Script: pilih function ini di
// dropdown "Run", klik Run) — nyisir Deposit_Log nyari nama yang punya
// saldo aktif (dari akumulasi transaksi lama) tapi BELUM ada row-nya sama
// sekali di DB_Player (mis. walk-in yang keburu dikonversi jadi deposit
// lewat "Cancel → Deposit" tapi nggak pernah didaftarin manual ke DB) —
// makanya nggak nongol di Kelola Player padahal beneran punya saldo. Bikin
// row baru buat tiap nama begitu, isi Saldo Deposit-nya dari total riwayat
// ledger. Aman dijalanin berkali-kali — nama yang udah ada row-nya dilewatin.
function syncMissingDepositPlayers() {
  const ws   = getSheet(SHEET.DEPOSIT);
  const data = ws.getDataRange().getValues();
  const totals = {}; // nama -> total saldo (masuk - keluar) dari ledger
  for (let r = 2; r < data.length; r++) {
    const nama = data[r][1];
    if (!nama) continue;
    totals[nama] = (totals[nama] || 0) + (toNum(data[r][5]) - toNum(data[r][6]));
  }
  const dibuat = [];
  Object.keys(totals).forEach(nama => {
    if (totals[nama] <= 0) return; // saldo abis/nol — nggak perlu row baru
    const info = getPlayerDepositSaldo(nama);
    if (info.rowIdx >= 0) return; // udah ada row-nya, lewatin
    writeNewPlayerRow({ nama, deposit: totals[nama],
      catatan: 'Otomatis dibuat — punya saldo deposit di riwayat tapi blm pernah didaftarkan manual' });
    dibuat.push(nama + ' (Rp ' + totals[nama] + ')');
  });
  Logger.log('Player DB_Player dibuatkan baru: ' + (dibuat.length ? dibuat.join(', ') : '(tidak ada yang perlu dibuat)'));
  return { ok: true, dibuat };
}
function addDeposit(p) {
  const ws      = getSheet(SHEET.DEPOSIT);
  const lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    p.tgl || today(), p.nama, p.matchId || '—',
    p.sumber || 'Cancel → Deposit', p.keterangan || '',
    p.jumlah, 0, p.jumlah
  ]]);
  const saldoBaru = adjustPlayerDepositSaldo(p.nama, p.jumlah);
  return { ok: true, saldoDbPlayer: saldoBaru };
}

function useDeposit(p) {
  const info = getPlayerDepositSaldo(p.namaPemain);
  if (info.saldo < p.jumlah) return { error: 'Saldo deposit tidak cukup' };
  const saldoSisa = info.saldo - p.jumlah;
  const ws      = getSheet(SHEET.DEPOSIT);
  const lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    today(), p.namaPemain, p.matchId,
    'Digunakan', 'Bayar HTM', 0, p.jumlah, saldoSisa
  ]]);
  adjustPlayerDepositSaldo(p.namaPemain, -p.jumlah);
  return { ok: true, saldoSisa };
}

// Hapus 1 baris deposit dari daftar "Tambah Deposit" (frontend, tab
// Deposit) — dulu tombol hapus di situ cuma ngilangin baris di layar
// (row.remove() doang), nggak pernah ngirim apa2 ke backend, jadi saldo
// yang sempat ditambahin (addDeposit) tetap nempel di database selama-
// lamanya walau udah "dihapus" di tampilan. Sekarang penghapusan itu
// ngirim ke sini, ditulis sbg baris pembalik di Deposit_Log (riwayatnya
// tetap append-only, bukan ngedit/ngehapus baris yg lama — konsisten sama
// pola "batal pakai deposit" yg juga nulis baris pembalik, bukan nge-undo
// baris asalnya) sekaligus ngurangin saldo DB_Player-nya beneran.
function deleteDeposit(p) {
  const ws      = getSheet(SHEET.DEPOSIT);
  const lastRow = ws.getLastRow() + 1;
  const saldoBaru = adjustPlayerDepositSaldo(p.nama, -p.jumlah);
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    today(), p.nama, p.matchId || '—',
    'Dihapus', p.keterangan || 'Dihapus dari daftar Tambah Deposit',
    0, p.jumlah, saldoBaru != null ? saldoBaru : ''
  ]]);
  return { ok: true, saldoDbPlayer: saldoBaru };
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
// CATATAN ARSITEKTUR: versi lama fungsi ini butuh sheet per-match
// (SS().getSheetByName(p.matchId)) yang cuma dibuat oleh createMatch()
// — tapi Transaksi di frontend TIDAK PERNAH pakai createMatch()/sheet
// per-match itu (semua data match cuma ada di localStorage + Match_Drafts,
// bentuknya HTML snapshot). Jadi versi lama ini nggak akan pernah jalan
// (selalu balikin {error:'Sheet not found'}). Ditulis ulang: semua angka
// (sales/cost/inventaris/breakdown bank) dikirim langsung dari frontend
// (dihitung dari data match yang aktif di layar Transaksi — sumbernya
// computeFinancialsFromLiveDom() di index.html, itu rumus yang sama persis
// dipakai buat kartu ringkasan "Rekap akhir"), bukan dihitung ulang dari
// sheet per-match yang nggak pernah ada.
function closeMatch(p) {
  const sales  = p.totalSales  || 0;
  const cost   = p.totalCost   || 0;
  const margin = sales - cost;
  const sm     = Math.round(margin * 0.6);
  const sf     = Math.round(margin * 0.4);
  const inv    = p.totalInv    || 0;
  const nm     = sm - Math.round(inv * 0.6);
  const nf     = sf - Math.round(inv * 0.4);

  // Tambah baris BARU ke Rekap_Keuangan — SELALU di baris paling bawah
  // (setelah baris terakhir yang udah kepake, entah itu BASELINE doang
  // atau udah ada puluhan match historis Sam di situ). Versi lama cari
  // baris 'BASELINE' dan selalu nyisip TEPAT SETELAHNYA — itu bug: kalau
  // udah ada match lain di bawah baseline (kasus nyata di sheet Sam),
  // insert akan NIMPA baris match yang udah ada, bukan nambah baris baru.
  const rekap   = getSheet(SHEET.REKAP);
  const rdata   = rekap.getDataRange().getValues();
  const insertRow = rekap.getLastRow() + 1;

  // Kumul dari baris tepat di atasnya (baris terakhir yang udah kepake —
  // entah BASELINE atau match sebelumnya)
  const prevRow  = rdata[insertRow-2] || [];
  const prevKumM = (typeof prevRow[14] === 'number') ? prevRow[14] : 8147730;
  const prevKumF = (typeof prevRow[15] === 'number') ? prevRow[15] : 5372762;
  const kumM     = prevKumM + nm;
  const kumF     = prevKumF + nf;

  const matchNo = (insertRow - 3); // nomor match baru (baris 3 = BASELINE = match "0")
  // Paksa kolom "Match / Tanggal" (kolom 2) jadi format TEKS biasa DULU
  // sebelum ditulis -- kalau nggak, Sheets suka diam2 ndeteksi string
  // "6 September 2026" ini sbg tanggal beneran & convert sel-nya jadi tipe
  // Date asli. Baru ketauan pas dibaca lagi (getMatches/getDashboard),
  // keluar jadi ISO mentah "2026-09-06T07:00:00.000Z" di chart/tabel
  // Dashboard & Reporting (bukan "6 September 2026" lagi). Cegah dari sini,
  // langsung pas ditulis -- lihat juga formatSheetDateForOutput() di
  // sheetToObjects() sbg jaring pengaman baca-ulang kalau ada baris LAMA yg
  // udah kadung ke-convert sebelum fix ini ada.
  rekap.getRange(insertRow, 2, 1, 1).setNumberFormat('@');
  rekap.getRange(insertRow, 1, 1, 22).setValues([[
    matchNo, p.tglMatch, p.jenisGame, p.tipeEvent,
    p.venue, p.tipeLapangan,
    sales, cost, margin, sm, sf,
    inv, nm, nf, kumM, kumF,
    p.bca||0, p.bsi||0, p.mandiri||0, p.bri||0, p.cash||0,
    p.lap || ''
  ]]);
  // Kolom 23 "Tahun" — dipakai grafik "Profit per match" biar bisa
  // dikelompokin per kuartal/tahun. Match lama (sebelum kolom ini ada)
  // diisi manual sama Sam; match baru yang ditutup lewat sistem ini
  // pakai tahun match itu sendiri kalau dikirim dari frontend (dari
  // tanggal match, bukan tanggal ditutup), fallback ke tahun berjalan.
  rekap.getRange(insertRow, 23).setValue(p.tahun || new Date().getFullYear());

  // Kembalikan semua stok jersey dari match ini (dianggap dipakai selesai)
  if (p.jerseyUsed && Array.isArray(p.jerseyUsed)) {
    p.jerseyUsed.forEach(j => { if (j && j.warna && j.uk) updateJerseyStok(j.warna, j.uk, -1); });
  }

  return {
    ok: true, matchNo, sales, cost, margin,
    splitMalik: sm, splitFilan: sf,
    kumulMalik: kumM, kumulFilan: kumF
  };
}

// ── POST: Hapus 1 baris match dari Rekap_Keuangan (mis. hasil dobel-klik
// tutup match sebelum bug-nya dibenerin) + hitung ulang Kumul Malik/Kumul
// Filan semua baris SESUDAHNYA dari awal (bukan cuma "kurangin" baris yg
// dihapus) — biar sekalian self-heal kalau kumulatifnya kebetulan udah
// keburu ngaco duluan. Kolom "No" match dirapikan ulang jadi 1,2,3,...
// tapi CUMA utk baris yg "No"-nya emang udah angka bersih (baris lama/beda
// format dibiarin apa adanya, nggak ikut disentuh) — konsisten sama filter
// yg dipakai getMatches().
function deleteRekapMatch(p) {
  const row = parseInt(p.row, 10);
  if (!row || row < 4) return { error: 'Baris tidak valid.' };
  const ws = getSheet(SHEET.REKAP);
  const data = ws.getDataRange().getValues();
  const idx = row - 1; // 0-based index ke data[]
  if (idx >= data.length) return { error: 'Baris tidak ditemukan (mungkin sudah dihapus sebelumnya).' };
  const target = data[idx];

  // Verifikasi baris yang mau dihapus BENERAN masih persis sama dengan yg
  // ditampilkan di layar Reporting pas tombol hapus diklik -- jaga2 kalau
  // sheet-nya berubah (mis. ada yg nutup match baru) di antara waktu
  // Reporting terakhir di-fetch & tombol hapus diklik, supaya nggak salah
  // hapus baris.
  const targetTglFormatted = formatSheetDateForOutput(target[1]);
  const targetSales = toNum(target[6]);
  if (String(targetTglFormatted) !== String(p.confirmTgl || '') ||
      String(target[3] || '') !== String(p.confirmTipe || '') ||
      Math.round(targetSales) !== Math.round(toNum(p.confirmSales))) {
    return { error: 'Data baris ini sudah berubah dari yang ditampilkan di layar — refresh dulu halaman Reporting-nya, lalu coba hapus lagi (biar nggak salah hapus baris).' };
  }

  ws.deleteRow(row);

  // Hitung ulang Kumul Malik/Kumul Filan dari baris BASELINE (row 3) sampai
  // baris terakhir yang tersisa.
  const data2 = ws.getDataRange().getValues();
  let kumM = toNum(data2[2] ? data2[2][14] : 0); // baseline, kolom O
  let kumF = toNum(data2[2] ? data2[2][15] : 0); // baseline, kolom P
  let matchNo = 0;
  for (let r = 3; r < data2.length; r++) {
    const rowData = data2[r];
    if (!rowData[0] && !rowData[1]) continue; // baris kosong, lewatin
    const isNumericNo = /^\d+$/.test(String(rowData[0]).trim());
    const nm = toNum(rowData[12]); // Net Malik, kolom M
    const nf = toNum(rowData[13]); // Net Filan, kolom N
    kumM += nm;
    kumF += nf;
    ws.getRange(r + 1, 15).setValue(kumM); // Kumul Malik, kolom O
    ws.getRange(r + 1, 16).setValue(kumF); // Kumul Filan, kolom P
    if (isNumericNo) {
      matchNo++;
      ws.getRange(r + 1, 1).setValue(matchNo); // No, kolom A
    }
  }

  return { ok: true, deletedRow: row, newLastKumulMalik: kumM, newLastKumulFilan: kumF };
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

// ============================================================
// ── SETUP: jalankan SEKALI aja dari editor Apps Script ──────
// Cara pakai: buka Extensions > Apps Script di Sheet "Minisoccer_DB",
// pilih function "setupSheets" di dropdown atas, klik Run (▶),
// izinkan akses saat diminta. Ini otomatis bikin semua tab + header
// yang dibutuhin backend ini (DB_Player, Stok_Jersey, SHTM_Log,
// Rekap_Keuangan, Deposit_Log, Inventaris, TEMPLATE_Match).
// Aman dijalankan berkali-kali — kalau tab udah ada, dilewatin
// (isinya TIDAK ditimpa), cuma tab yang belum ada yang dibikin.
// ============================================================
function setupSheets() {
  const ss = SS();
  const made = [];
  const skipped = [];

  function ensureSheet(name, rows, cols) {
    let ws = ss.getSheetByName(name);
    if (ws) { skipped.push(name); return ws; }
    ws = ss.insertSheet(name);
    if (rows || cols) ws.getRange(1,1, rows||1, cols||1); // no-op touch, avoids edge cases on some accounts
    made.push(name);
    return ws;
  }

  function writeHeader(ws, title, headers) {
    ws.getRange(1,1).setValue(title);
    ws.getRange(2,1,1,headers.length).setValues([headers]);
    ws.getRange(2,1,1,headers.length).setFontWeight('bold');
    ws.setFrozenRows(2);
  }

  // ── DB_Player ──
  let ws = ensureSheet(SHEET.PLAYER);
  if (ws.getLastRow() < 2) writeHeader(ws, 'DB_Player — master data player', [
    'No','Nama Player','Kategori','Total Hadir','Total Gol','Total SHTM',
    'Ukuran Jersey','Warna Favorit','Deposit Saldo (Rp)','Metode Favorit',
    'Bank','Kontak','Jersey Pribadi','Alias','Status','Catatan','Tgl Daftar'
  ]);

  // ── Stok_Jersey ── (diisi angka stok yang sekarang dipakai di index.html
  // supaya begitu tersambung, angkanya nyambung — bukan mulai dari 0)
  ws = ensureSheet(SHEET.JERSEY);
  if (ws.getLastRow() < 2) {
    writeHeader(ws, 'Stok_Jersey — stok jersey Field & GK', [
      'Warna','Tipe','Ukuran','Stok Total','Dipakai','Sisa','Status','Keterangan'
    ]);
    const stokField = {
      'Merah': {M:2,L:2,XL:0,'2XL':1}, 'Kuning': {M:1,L:3,XL:1,'2XL':1},
      'Biru': {M:2,L:3,XL:3,'2XL':1}, 'Hijau': {M:2,L:1,XL:4,'2XL':1}
    };
    const stokGk = { 'GK': {M:2,L:2,XL:2} };
    const rows = [];
    function pushRows(map, tipe) {
      Object.keys(map).forEach(warna => {
        Object.keys(map[warna]).forEach(uk => {
          const stok = map[warna][uk];
          const status = stok===0 ? 'HABIS' : stok<=1 ? 'LOW' : 'OK';
          rows.push([warna, tipe, uk, stok, 0, stok, status, '']);
        });
      });
    }
    pushRows(stokField, 'Field');
    pushRows(stokGk, 'GK');
    ws.getRange(3,1,rows.length,8).setValues(rows);
  }

  // ── SHTM_Log ──
  ws = ensureSheet(SHEET.SHTM);
  if (ws.getLastRow() < 2) writeHeader(ws, 'SHTM_Log — riwayat Sistem Half-price Top Match', [
    'No','Nama Player','Tgl Mendapat SHTM','Tgl SHTM Dipakai','Status','Match','Catatan'
  ]);

  // ── Rekap_Keuangan ──
  ws = ensureSheet(SHEET.REKAP);
  if (ws.getLastRow() < 2) {
    writeHeader(ws, 'Rekap_Keuangan — 1 baris per match yang ditutup', [
      'No','Match / Tanggal','Jenis Game','Tipe Event','Venue','Tipe Lapangan',
      'Sales (Rp)','Cost (Rp)','Margin (Rp)','Split Malik (Rp)','Split Filan (Rp)',
      'Inventaris (Rp)','Net Malik (Rp)','Net Filan (Rp)','Kumul Malik (Rp)','Kumul Filan (Rp)',
      'BCA (Rp)','BSI (Rp)','Mandiri (Rp)','BRI (Rp)','Cash (Rp)','Lap.'
    ]);
    // Baris BASELINE — titik awal hitungan kumulatif sebelum match pertama
    // ditutup lewat sistem ini. Ganti angka O/P di baris ini (Kumul Malik /
    // Kumul Filan) ke saldo riil kamu sebelum mulai pakai backend ini.
    const baseline = new Array(22).fill('');
    baseline[0] = 'BASELINE';
    baseline[14] = 8147730; // Kumul Malik (Rp) — sesuaikan ke saldo riil
    baseline[15] = 5372762; // Kumul Filan (Rp) — sesuaikan ke saldo riil
    ws.getRange(3,1,1,22).setValues([baseline]);
  } else {
    // Tab udah ada isinya (kasus real kamu) — kolom "Match / Tanggal" cuma
    // simpan tanggal singkat tanpa tahun ("21 Des", "12 April"), jadi grafik
    // yg butuh dikelompokin per kuartal/tahun (Profit per match) nggak bisa
    // dihitung akurat. Tambahin kolom "Tahun" di ujung (kolom 23) — ADDITIF,
    // nggak nimpa apa pun yang udah ada. Header row = row 2, sama kayak
    // header lainnya di sheet ini.
    const rHdr = ws.getRange(2,1,1,ws.getLastColumn()).getValues()[0];
    const hasTahun = rHdr.some(h => String(h).trim() === 'Tahun');
    if (!hasTahun) {
      ws.getRange(2, 23).setValue('Tahun');
    }
  }

  // ── Deposit_Log ──
  ws = ensureSheet(SHEET.DEPOSIT);
  if (ws.getLastRow() < 2) writeHeader(ws, 'Deposit_Log — riwayat deposit player', [
    'Tanggal','Nama Player','Match ID','Sumber','Keterangan',
    'Jumlah Masuk (Rp)','Jumlah Keluar (Rp)','Saldo (Rp)'
  ]);

  // ── Inventaris ──
  ws = ensureSheet(SHEET.INVENTARIS);
  if (ws.getLastRow() < 2) writeHeader(ws, 'Inventaris — belanja alat, potong kumulatif Malik/Filan', [
    'Tanggal','Item','Jumlah (Rp)','Keterangan','Split Malik (Rp)','Split Filan (Rp)'
  ]);

  // ── TEMPLATE_Match — sheet ini DI-COPY tiap kali "Buat match baru" ──
  // Kalau tab ini BELUM ada sama sekali: bikin dari nol, layout persis
  // sama kayak TEMPLATE_Match yang udah ada di sheet data riil kamu (info
  // row 3-13, judul "ESTIMASI MARGIN" row 15, 6 field margin row 16-21,
  // "DAFTAR PLAYER" row 23, header tabel player row 24) — supaya
  // parseMatchSheet() di kode ini baca row yang bener baik di sheet baru
  // maupun yang lama.
  ws = ss.getSheetByName('TEMPLATE_Match');
  if (!ws) {
    ws = ss.insertSheet('TEMPLATE_Match');
    made.push('TEMPLATE_Match');

    ws.getRange(1,1).setValue('TEMPLATE MATCH — jangan diisi manual, disalin otomatis tiap "Buat match baru"');
    ws.getRange(1,1).setFontWeight('bold');
    ws.getRange(2,1).setValue('INFO EVENT');
    ws.getRange(2,1).setFontWeight('bold');

    const infoLabels = ['Match ID','Tanggal','Venue','Jenis Game','Tipe Event',
      'Tipe Lapangan','HTM Player','HTM GK','Kuota Kiper','Kuota Field','Status Match'];
    infoLabels.forEach((lbl,i) => ws.getRange(3+i,1).setValue(lbl));

    ws.getRange(15,1).setValue('ESTIMASI MARGIN');
    ws.getRange(15,1).setFontWeight('bold');
    const marginLabels = ['Sales Pasti (Rp)','Sales Estimasi (Rp)','Total Cost (Rp)',
      'Margin Proyeksi (Rp)','Split Malik (Rp)','Split Filan (Rp)'];
    marginLabels.forEach((lbl,i) => ws.getRange(16+i,1).setValue(lbl));

    ws.getRange(23,1).setValue('DAFTAR PLAYER');
    ws.getRange(23,1).setFontWeight('bold');
    ws.getRange(24,1,1,16).setValues([[
      'No','Nama Player','Status Bayar','HTM','Label Harga','Deposit Dipakai',
      'Metode Bayar','Jumlah Bayar','Tgl Bayar','Posisi','Jersey Warna','Ukuran',
      'Tim','Formasi','Status SHTM','Catatan'
    ]]);
    ws.getRange(24,1,1,16).setFontWeight('bold');

    ws.getRange(60,1).setValue('BIAYA OPERASIONAL');
    ws.getRange(60,1).setFontWeight('bold');
    ws.getRange(61,1,1,8).setValues([[
      'Item Biaya','Kategori','Tipe','Rencana (Rp)','Realisasi (Rp)','Skema Split','Jml Cicilan','Catatan'
    ]]);
    ws.getRange(61,1,1,8).setFontWeight('bold');

    ws.getRange(85,1).setValue('PENCETAK GOL');
    ws.getRange(85,1).setFontWeight('bold');
    ws.getRange(86,1,1,3).setValues([['Nama','Gol','Keterangan']]);
    ws.getRange(86,1,1,3).setFontWeight('bold');
  } else {
    // Tab ini udah ada (kasus sheet data riil kamu) — JANGAN disentuh isi
    // row 1-24-nya sama sekali. Cuma cek: apa udah ada section "Biaya"
    // dan "Pencetak Gol"? Kalau belum (dan memang belum, di sheet riil
    // kamu TEMPLATE_Match cuma sampai row 24 = header Daftar Player),
    // ditambahin di BAWAH baris terakhir yang udah kepake — addBiaya()/
    // addGol() di kode ini butuh baris berlabel itu buat tau mau nulis
    // baris baru di mana.
    skipped.push('TEMPLATE_Match (isi lama tidak diubah)');
    const data = ws.getDataRange().getValues();
    const hasBiaya = data.some(row => {
      const c0 = String(row[0] || '').toLowerCase();
      return c0.includes('item biaya') || c0.includes('biaya operasional');
    });
    const hasGol = data.some(row => {
      const c0 = String(row[0] || '').toLowerCase();
      return c0.includes('pencetak') || c0.includes('goal scorer');
    });
    if (!hasBiaya || !hasGol) {
      // JANGAN taruh section baru cuma "3 baris di bawah baris terakhir
      // yang kepake" — baris terakhir yang kepake itu justru HEADER tabel
      // player (row 24), dan addPlayerToMatch() nyari baris kosong per
      // KOLOM B buat nyisipin player baru; kalau section baru ditaruh
      // kepepet deket situ, player ke-3/4 bisa numpuk nimpa label section.
      // Jadi kasih jarak cukup dari baris HEADER PLAYER (bukan dari baris
      // terakhir), muat sampai puluhan player per match.
      let playerHeaderRow = -1;
      for (let r = 0; r < data.length; r++) {
        if (data[r][0] === 'No' && data[r][1] === 'Nama Player') { playerHeaderRow = r + 1; break; }
      }
      let nextRow = playerHeaderRow > 0 ? playerHeaderRow + 40 : ws.getLastRow() + 3;
      if (!hasBiaya) {
        ws.getRange(nextRow,1).setValue('BIAYA OPERASIONAL');
        ws.getRange(nextRow,1).setFontWeight('bold');
        nextRow += 1;
        ws.getRange(nextRow,1,1,8).setValues([[
          'Item Biaya','Kategori','Tipe','Rencana (Rp)','Realisasi (Rp)','Skema Split','Jml Cicilan','Catatan'
        ]]);
        ws.getRange(nextRow,1,1,8).setFontWeight('bold');
        made.push('TEMPLATE_Match: section Biaya Operasional (ditambah di row ' + (nextRow-1) + ')');
        nextRow += 24; // headroom baris biaya
      }
      if (!hasGol) {
        nextRow += 2;
        ws.getRange(nextRow,1).setValue('PENCETAK GOL');
        ws.getRange(nextRow,1).setFontWeight('bold');
        nextRow += 1;
        ws.getRange(nextRow,1,1,3).setValues([['Nama','Gol','Keterangan']]);
        ws.getRange(nextRow,1,1,3).setFontWeight('bold');
        made.push('TEMPLATE_Match: section Pencetak Gol (ditambah di row ' + (nextRow-1) + ')');
      }
    }
  }

  // ── Rekap_Skorer ── (top skor all time, manual — field-nya sama persis
  // dgn tabel "Kelola Rekap Skorer" di Admin: Nama/Kategori/Tahun/Gol/
  // Uraian. BELUM ada di sheet data riil kamu juga, jadi ini tab baru,
  // aman dibikin kapan aja — nggak nyentuh 6 tab lain yang udah ada.
  // Diisi otomatis dari 7 data yang udah kamu masukin manual di Admin.
  ws = ensureSheet(SHEET.SCORER);
  if (ws.getLastRow() < 2) {
    writeHeader(ws, 'Rekap_Skorer — top skor all time (tampil di Dashboard & Goal scorer)', [
      'No','Nama','Kategori','Tahun','Total Gol','Uraian'
    ]);
    const seedScorer = [
      [1,'Annas','BMR','2026',19,''],
      [2,'Bambang','SMS','2026',13,''],
      [3,'Jayadi','Umum','2026',10,''],
      [4,'Ibnu','Umum','2026',7,''],
      [5,'Wasis','Umum','2026',7,''],
      [6,'Arsya','Umum','2026',4,''],
      [7,'Eri','SMS','2026',4,'']
    ];
    ws.getRange(3,1,seedScorer.length,6).setValues(seedScorer);
  }

  // ── Match_Drafts ── (draft match Transaksi yang masih berjalan, dipakai
  // supaya "Simpan" nyambung ke browser/device lain — lihat catatan di
  // getMatchDrafts()/saveMatchDraft() di atas. Tab baru, aman dibikin
  // kapan aja, nggak nyentuh tab lain.)
  ws = ensureSheet(SHEET.DRAFTS);
  if (ws.getLastRow() < 2) {
    writeHeader(ws, 'Match_Drafts — draft match Transaksi aktif, buat sinkron antar browser/device', [
      'MatchId','Data1','Data2','Data3','Data4','Data5','Data6','Data7','Data8','UpdatedAt'
    ]);
  }

  // Hapus "Sheet1" bawaan Google kalau masih ada dan kosong
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) {
    ss.deleteSheet(def);
  }

  const msg = 'Selesai. Tab dibuat: ' + (made.join(', ') || '(tidak ada, semua udah ada)') +
    ' | Tab dilewatin (udah ada sebelumnya): ' + (skipped.join(', ') || '(tidak ada)');
  Logger.log(msg);
  return msg;
}

// ── GET: Rekap Skorer (top skor all time, manual) ───────────
function getScorer() {
  return sheetToObjects(SHEET.SCORER, 2);
}

// ── POST: Rekap Skorer ───────────────────────────────────────
function addScorer(p) {
  const ws      = getSheet(SHEET.SCORER);
  const lastRow = ws.getLastRow() + 1;
  const no      = lastRow - 2;
  ws.getRange(lastRow, 1, 1, 6).setValues([[
    no, p.nama, p.kategori || '—', p.tahun || '', p.gol || 0, p.uraian || ''
  ]]);
  return { ok: true, row: lastRow };
}

function updateScorer(p) {
  const ws = getSheet(SHEET.SCORER);
  if (!ws) return { error: 'Sheet not found' };
  ws.getRange(p.rowIdx, 1, 1, 6).setValues([[
    p.no, p.nama, p.kategori || '—', p.tahun || '', p.gol || 0, p.uraian || ''
  ]]);
  return { ok: true };
}

function deleteScorer(p) {
  const ws = getSheet(SHEET.SCORER);
  if (!ws) return { error: 'Sheet not found' };
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Add New Player to DB ──────────────────────────────
// Header-safe (lihat catatan getHeaderMap di atas) — DB_Player itu sheet
// ASLI punya Sam sendiri (bukan dibikin baru sama setupSheets()), dan
// urutan kolomnya udah kebukti beda dari asumsi lama kode ini. Nulis
// berdasarkan NAMA kolom, jadi walau urutan asli beda, tetep masuk ke
// kolom yang benar (bukan positional/asal-nempel-index).
// Ditarik jadi function terpisah (writeNewPlayerRow) krn dipakai juga sama
// createPlayerDbRowMinimal (row baru yang dibuat OTOMATIS krn ada transaksi
// deposit atas nama yang blm terdaftar) — biar dua-duanya nulis kolom yg
// sama persis, nggak ada yg beda sendiri kalau nanti header sheet berubah.
function writeNewPlayerRow(opts) {
  const { ws, map, hdrs } = getHeaderMap(SHEET.PLAYER, 2);
  const lastRow = ws.getLastRow() + 1;
  const no      = lastRow - 2;
  const rowArr  = new Array(hdrs.length).fill('');
  function set(aliases, val) {
    const c = findCol(map, aliases);
    if (c > 0) rowArr[c-1] = val;
  }
  set(['No'], no);
  set(['Nama Player','Nama'], opts.nama);
  set(['Kategori'], opts.kategori || '—');
  set(['Total Hadir','Total\nHadir'], 0);
  set(['Total Gol','Total\nGol'], 0);
  set(['Total SHTM','Total\nSHTM'], 0);
  set(['Ukuran Jersey','Ukuran\nDominan','Ukuran'], opts.ukuran || '—');
  set(['Warna Favorit','Warna\nFavorit'], '—');
  set(DEPOSIT_COL_ALIASES, opts.deposit || 0);
  set(['Metode Favorit','Metode\nFavorit'], '—');
  set(['Bank'], '—');
  set(['Kontak'], opts.kontak || '—');
  set(['Jersey Pribadi','Jersey\nPribadi'], opts.jerseyPribadi || '—');
  set(['Alias'], opts.alias || '—');
  set(['Status'], 'Aktif');
  set(['Status SHTM','Status\nSHTM'], '—');
  set(['Catatan'], opts.catatan || '');
  set(['Tgl Daftar','Tgl\nDaftar'], today());
  ws.getRange(lastRow, 1, 1, rowArr.length).setValues([rowArr]);
  return lastRow;
}
function addNewPlayer(p) {
  const lastRow = writeNewPlayerRow({
    nama: p.nama, kategori: p.kategori, ukuran: p.ukuran, kontak: p.kontak,
    jerseyPribadi: p.jerseyPribadi, alias: p.alias, catatan: p.catatan
  });
  return { ok: true, row: lastRow };
}

// ── POST: Update / Delete Player (Kelola Player, header-safe) ──────
function updatePlayerDb(p) {
  if (!p.rowIdx) return { error: 'rowIdx wajib diisi' };
  const { ws, map } = getHeaderMap(SHEET.PLAYER, 2);
  function set(aliases, val) {
    if (val === undefined || val === null) return;
    const c = findCol(map, aliases);
    if (c > 0) ws.getRange(p.rowIdx, c).setValue(val);
  }
  set(['Nama Player','Nama'], p.nama);
  set(['Kategori'], p.kategori);
  set(['Total Hadir','Total\nHadir'], p.hadir);
  set(['Ukuran Jersey','Ukuran\nDominan','Ukuran'], p.ukuran);
  set(DEPOSIT_COL_ALIASES, p.deposit);
  set(['Status SHTM','Status\nSHTM'], p.statusShtm);
  set(['Status'], p.status);
  set(['Kontak'], p.kontak);
  set(['Jersey Pribadi','Jersey\nPribadi'], p.jerseyPribadi);
  set(['Catatan'], p.catatan);
  return { ok: true };
}

function deletePlayerDb(p) {
  if (!p.rowIdx) return { error: 'rowIdx wajib diisi' };
  const ws = getSheet(SHEET.PLAYER);
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Update / Delete SHTM row (Daftar SHTM, header-safe) ──────
function updateShtmRow(p) {
  if (!p.rowIdx) return { error: 'rowIdx wajib diisi' };
  const { ws, map } = getHeaderMap(SHEET.SHTM, 2);
  function set(aliases, val) {
    if (val === undefined || val === null) return;
    const c = findCol(map, aliases);
    if (c > 0) ws.getRange(p.rowIdx, c).setValue(val);
  }
  set(['Nama Player'], p.nama);
  set(['Tgl Mendapat SHTM'], p.tglMendapat);
  set(['Tgl SHTM Dipakai'], p.tglDipakai);
  set(['Status'], p.status);
  set(['Match'], p.match);
  set(['Catatan'], p.catatan);
  return { ok: true };
}

function deleteShtmRow(p) {
  if (!p.rowIdx) return { error: 'rowIdx wajib diisi' };
  const ws = getSheet(SHEET.SHTM);
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}
