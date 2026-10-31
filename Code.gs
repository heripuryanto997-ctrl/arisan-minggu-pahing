/** ARISAN - Backend Google Apps Script. Jalankan setup() sekali. Login awal: admin / admin123 */
const SS = () => SpreadsheetApp.getActiveSpreadsheet();
const H = {
  Users: ['username', 'password', 'role', 'anggotaId'],
  Anggota: ['id', 'nama', 'alamat', 'hp'],
  Perolehan: ['id', 'lot', 'anggotaId', 'tanggal', 'jumlah'],
  Pembayaran: ['id', 'perolehanId', 'anggotaId', 'nominal']
};

function setup() {
  for (const n in H) {
    const s = SS().getSheetByName(n) || SS().insertSheet(n);
    if (!s.getLastRow()) s.appendRow(H[n]);
    s.getRange(1, 1, s.getMaxRows(), H[n].length).setNumberFormat('@'); // teks agar 0 di depan No HP tidak hilang
  }
  const u = SS().getSheetByName('Users');
  if (u.getLastRow() < 2) u.appendRow(['admin', 'admin123', 'admin', '']);
}
const hash = s => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s)));
const uid = () => Utilities.getUuid().slice(0, 8);

function rows(n) {
  const v = SS().getSheetByName(n).getDataRange().getValues(), h = v.shift().map(x => H[n].find(k => k.toLowerCase() === String(x).trim().toLowerCase()) || String(x).trim());
  return v.map(r => Object.fromEntries(h.map((k, i) => [k,
    r[i] instanceof Date ? Utilities.formatDate(r[i], 'GMT+7', 'yyyy-MM-dd') : r[i]])));
}
function put(n, list) {
  const s = SS().getSheetByName(n), c = H[n].length;
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, c).clearContent();
  if (list.length) s.getRange(2, 1, list.length, c).setValues(list.map(o => H[n].map(k => o[k] ?? '')));
}

// ---------- Web entry ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Arisan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function doPost(e) {   // dipakai saat frontend di Vercel
  const b = JSON.parse(e.postData.contents);
  return ContentService.createTextOutput(JSON.stringify(rpc(b.fn, b.token, b.args)))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- RPC ----------
function rpc(fn, t, a) {
  try {
    const C = CacheService.getScriptCache();
    if (fn === 'login') {
      if (!SS().getSheetByName('Users')) setup();
      ensureAdmin();
      const u = rows('Users').find(x => String(x.username).trim() === String(a.u).trim() && (String(x.password) === String(a.p) || String(x.password) === hash(a.p)));
      if (!u) return { err: 'Username/password salah' };
      const token = uid() + uid(), me = { role: u.role, aid: String(u.anggotaId), name: String(u.username) };
      C.put(token, JSON.stringify(me), 21600);
      return { token, me };
    }
    const raw = C.get(t);
    if (!raw) return { err: 'Sesi habis', expired: 1 };
    const me = JSON.parse(raw), adm = me.role === 'admin';
    if (fn === 'logout') { C.remove(t); return { ok: 1 }; }
    if (fn === 'init') return snapshot(me);          // 1x baca semua data -> menu pindah instan
    if (fn === 'saveAkun') {
      const U = rows('Users'), k = U.find(u => String(u.username) === me.name);
      if (!k || String(k.password) !== String(a.lama)) return { err: 'Password lama salah' };
      k.password = a.baru; put('Users', U); return { ok: 1 };
    }
    if (!adm) return { err: 'Hanya admin' };
    const L = LockService.getScriptLock(); L.waitLock(20000);
    try {
      if (fn === 'saveAnggota') {
        const AG = rows('Anggota'), U = rows('Users'), un = String(a.username || a.hp).trim(), pw = String(a.password || a.hp);
        if (U.some(u => String(u.username) === un && u.anggotaId != a.id)) return { err: 'Username sudah dipakai' };
        const rec = { id: a.id, nama: a.nama, alamat: a.alamat, hp: a.hp };
        if (a.id) {
          AG.forEach(x => { if (String(x.id) === String(a.id)) Object.assign(x, rec); });
          const k = U.find(u => u.anggotaId == a.id);
          if (k) { k.username = un; k.password = pw; } else U.push({ username: un, password: pw, role: 'anggota', anggotaId: a.id });
        } else { rec.id = uid(); AG.push(rec); U.push({ username: un, password: pw, role: 'anggota', anggotaId: rec.id }); }
        put('Anggota', AG); put('Users', U); return { ok: 1 };
      }
      if (fn === 'delAnggota') {
        put('Anggota', rows('Anggota').filter(x => x.id != a.id));
        put('Users', rows('Users').filter(x => x.anggotaId != a.id)); return { ok: 1 };
      }
      if (fn === 'savePerolehan') {
        const PR = rows('Perolehan'), PB = rows('Pembayaran');
        const id = a.id || uid(), pay = a.payments.filter(p => Number(p.nominal) > 0);
        if (PR.some(p => p.lot == a.lot && p.id != id)) return { err: 'Lot ke-' + a.lot + ' sudah terisi' };
        const rest = PR.filter(p => p.id != id);
        rest.push({ id, lot: a.lot, anggotaId: a.anggotaId, tanggal: a.tanggal, jumlah: pay.reduce((s, p) => s + Number(p.nominal), 0) });
        put('Perolehan', rest);
        put('Pembayaran', PB.filter(b => b.perolehanId != id).concat(pay.map(p => ({ id: uid(), perolehanId: id, anggotaId: p.anggotaId, nominal: p.nominal }))));
        return { ok: 1 };
      }
      if (fn === 'delPerolehan') {
        put('Perolehan', rows('Perolehan').filter(p => p.id != a.id)); put('Pembayaran', rows('Pembayaran').filter(b => b.perolehanId != a.id)); return { ok: 1 };
      }
    } finally { L.releaseLock(); }
    return { err: 'Fungsi tidak dikenal' };
  } catch (e) { return { err: String(e) }; }
}

// Semua data untuk dashboard, anggota, perolehan dalam sekali baca
function snapshot(me) {
  const adm = me.role === 'admin', AG = rows('Anggota'), PR = rows('Perolehan'), PB = rows('Pembayaran');
  const U = adm ? rows('Users') : [], nm = id => (AG.find(x => String(x.id) === String(id)) || {}).nama || '-';
  const got = [...new Set(PR.map(p => String(p.anggotaId)))];
  PR.sort((x, y) => x.lot - y.lot);
  const dashboard = { total: AG.length, sudah: got.length, belum: AG.length - got.length,
    jumlah: PR.reduce((s, p) => s + Number(p.jumlah), 0), perLot: PR.map(p => ({ lot: p.lot, jumlah: Number(p.jumlah) })) };
  let anggota = (adm ? AG : AG.filter(x => String(x.id) === me.aid)).map(x => {
    const u = U.find(k => k.anggotaId == x.id) || {};
    return { ...x, username: u.username || '', password: adm ? (u.password || '') : '' };
  });
  let perolehan = PR.map(p => ({ ...p, nama: nm(p.anggotaId),
    bayar: PB.filter(b => b.perolehanId == p.id).map(b => ({ anggotaId: b.anggotaId, nominal: Number(b.nominal), nama: nm(b.anggotaId) })) }));
  // Anggota: hanya perolehan atas namanya sendiri, tetapi detail pembayaran dari semua anggota tetap lengkap
  if (!adm) perolehan = perolehan.filter(p => String(p.anggotaId) === me.aid);
  const saya = { role: me.role, user: me.name, nama: adm ? 'Administrator' : nm(me.aid) };
  return { saya, dashboard, anggota, perolehan, semua: AG.map(x => ({ id: x.id, nama: x.nama })) };
}

// Pastikan ada minimal satu akun admin (admin / admin123) agar tidak terkunci
function ensureAdmin() {
  const s = SS().getSheetByName('Users');
  if (!s.getLastRow()) s.appendRow(H.Users);
  if (!rows('Users').some(u => u.role === 'admin' && String(u.username).trim() && String(u.password).trim()))
    s.appendRow(['admin', 'admin123', 'admin', '']);
}
// Jalankan manual dari editor jika lupa password admin: reset jadi admin / admin123
function resetAdmin() {
  const s = SS().getSheetByName('Users') || (setup(), SS().getSheetByName('Users'));
  const v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][2]).trim() === 'admin') { s.getRange(i + 1, 1, 1, 2).setValues([['admin', 'admin123']]); return 'ok'; }
  s.appendRow(['admin', 'admin123', 'admin', '']); return 'ok';
}
