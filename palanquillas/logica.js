/* ===== Lógica: lectura de archivos y cálculo ===== */
const DENS = 7.81; // t/m³ efectiva (130x130 → 132 kg/m, igual que el programa)

function norm(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/[^A-Z0-9]/g, '');
}
function clean(s) { return String(s == null ? '' : s).replace(/ /g, ' ').replace(/\s+/g, ' ').trim(); }
function num(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (v == null || v === '') return null;
  const n = Number(String(v).replace(/\./g, '').replace(',', '.'));
  return isFinite(n) ? n : null;
}
function code(v) { const s = clean(v).replace(/\.0$/, ''); return s.replace(/^0+/, ''); }
// fechas Excel como serial → ms "UTC de pared"
function xlDate(v) {
  if (typeof v === 'number' && v > 20000 && v < 80000) return Math.round((v - 25569) * 864e5);
  if (v instanceof Date) return Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes());
  return null;
}

/* Descripción de palanquilla → sección, calidad, largo */
function parseBillet(desc) {
  const d = clean(desc).toUpperCase();
  const out = { section: '', grade: '', length: null, origin: '' };
  const sm = d.match(/(\d{2,3})\s*X\s*(\d{2,3})\s*MM/);
  let rest = d;
  if (sm) { out.section = sm[1] + 'x' + sm[2]; out.a = +sm[1]; out.b = +sm[2]; rest = d.slice(sm.index + sm[0].length); }
  const om = rest.match(/\(([A-Z])\)?\s*$/) || rest.match(/\(([A-Z])\b/);
  if (om) out.origin = om[1];
  const lm = [...rest.matchAll(/(\d+(?:[.,]\d+)?)\s*M\b/g)].pop();
  if (lm) { out.length = +lm[1].replace(',', '.'); out.grade = clean(rest.slice(0, lm.index)); }
  else out.grade = clean(rest.replace(/\(.*$/, ''));
  out.grade = out.grade.replace(/^\(T\)\s*/, '');
  return out;
}
function kgPerM(p) { return p.a && p.b ? DENS * p.a * p.b / 1000 : 132; }
function lenLabel(l) { return l == null ? '—' : String(l).replace('.', ',') + ' m'; }

/* ---------- Programa ---------- */
function findHeader(aoa, must) {
  for (let r = 0; r < Math.min(aoa.length, 60); r++) {
    const row = aoa[r] || [];
    if (row.some(c => norm(c) === must)) return r;
  }
  return -1;
}
function parseProgramWB(wb) {
  let sheetName = wb.SheetNames.find(n => norm(n) === 'PROGLAMREN');
  const tryOrder = sheetName ? [sheetName, ...wb.SheetNames] : wb.SheetNames;
  for (const name of tryOrder) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    const h = findHeader(aoa, 'CODIGOPALANQUILLA');
    if (h < 0) continue;
    const hdr = aoa[h].map(norm);
    const col = (pred) => hdr.findIndex(pred);
    const C = {
      of: col(x => x === 'CLASIFICADA' || x === 'OF'),
      sec: col(x => x === 'SECN' || x === 'SEC'),
      prod: col(x => x === 'CODIGO'),
      desc: col(x => x === 'DESCRIPCION'),
      corte: col(x => x.startsWith('CORTEENT4') || x.startsWith('CORTE')),
      progr: col(x => x === 'PROGR'),
      ton: col(x => x === 'TON'),
      cu: col(x => x === 'CU'),
      ini: col(x => x === 'INICIO'),
      real: col(x => x === 'REAL'),
      pc: col(x => x === 'CODIGOPALANQUILLA'),
      pd: col(x => x.startsWith('DESCRIPCIONMATERIAPRIMA')),
      cal: col(x => x === 'CALIDAD'),
      lib: col(x => x.startsWith('FECHALIBERACION')),
    };
    const g = (row, k) => (C[k] >= 0 ? row[C[k]] : null);
    const rows = [];
    for (let r = h + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const pc = code(g(row, 'pc')), pd = clean(g(row, 'pd'));
      const ton = num(g(row, 'ton'));
      if ((!pc && !pd) || !ton) continue;
      rows.push({
        n: rows.length, of: clean(g(row, 'of')), sec: num(g(row, 'sec')), prod: code(g(row, 'prod')),
        desc: clean(g(row, 'desc')), corte: num(g(row, 'corte')), progr: num(g(row, 'progr')) || 0,
        ton, cu: num(g(row, 'cu')) || 0, ini: xlDate(g(row, 'ini')), real: num(g(row, 'real')),
        pc, pd, cal: clean(g(row, 'cal')), lib: xlDate(g(row, 'lib')),
      });
    }
    if (rows.length) return { sheet: name, rows };
  }
  throw new Error('No encontré la columna "CÓDIGO PALANQUILLA" en ninguna hoja. Revisa que sea el programa de laminación.');
}

/* ---------- Stock ---------- */
function parseStockWB(wb) {
  for (const name of wb.SheetNames) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    const h = findHeader(aoa, 'MATERIAL');
    if (h < 0) continue;
    const hdr = aoa[h].map(norm);
    const col = (pred) => hdr.findIndex(pred);
    const C = {
      mat: col(x => x === 'MATERIAL'), desc: col(x => x.startsWith('TEXTOBREVE')),
      libre: col(x => x.startsWith('LIBREUTILIZACION')), insp: col(x => x.startsWith('INSPECC')),
      bloq: col(x => x.startsWith('BLOQUEADO')), lote: col(x => x === 'LOTE'),
      centro: col(x => x === 'CENTRO'), alm: col(x => x.startsWith('ALMACEN')),
    };
    if (C.libre < 0) continue;
    const g = (row, k) => (C[k] >= 0 ? row[C[k]] : null);
    const rows = [];
    for (let r = h + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const mat = code(g(row, 'mat'));
      if (!mat) continue; // filas de subtotal
      rows.push({ mat, desc: clean(g(row, 'desc')), libre: num(g(row, 'libre')) || 0, insp: num(g(row, 'insp')) || 0,
        bloq: num(g(row, 'bloq')) || 0, lote: clean(g(row, 'lote')), centro: clean(g(row, 'centro')), alm: clean(g(row, 'alm')) });
    }
    if (rows.length) return { sheet: name, rows };
  }
  throw new Error('No encontré las columnas "Material" y "Libre utilización". Revisa que sea el reporte de stock.');
}
function dateFromName(name) {
  const m = String(name).match(/(\d{1,2})[_\-.](\d{1,2})[_\-.](\d{2,4})/);
  if (!m) return null;
  let y = +m[3]; if (y < 100) y += 2000;
  const d = Date.UTC(y, +m[2] - 1, +m[1]);
  return isNaN(d) ? null : d;
}

/* ---------- Cálculo ---------- */
function bkey(p) { return p.section + '|' + norm(p.grade) + '|' + p.length; }

function analyze(prog, stock, opt) {
  // stock por material
  const S = new Map();
  for (const s of stock) {
    let x = S.get(s.mat);
    if (!x) { x = { mat: s.mat, desc: s.desc, p: parseBillet(s.desc), libre: 0, insp: 0, bloq: 0, lotes: 0, centros: new Set() }; S.set(s.mat, x); }
    x.libre += s.libre; x.insp += s.insp; x.bloq += s.bloq; x.lotes++; if (s.centro) x.centros.add(s.centro);
  }
  for (const x of S.values()) x.disp = (x.libre + (opt.insp ? x.insp : 0)) / 1000; // t
  const byKey = new Map();
  for (const x of S.values()) { const k = bkey(x.p); if (!byKey.has(k)) byKey.set(k, x); }

  // filas pendientes
  const rows = prog.map(r => {
    let status = 'pend', factor = 1;
    if (opt.cut != null && r.ini != null && r.ini < opt.cut) status = 'antes';
    else if (r.real != null && r.real > 0) {
      if (opt.realMode === 'resta' && r.progr > 0) { factor = Math.max(0, 1 - r.real / r.progr); status = factor > 0 ? 'pend' : 'hecha'; }
      else { status = 'hecha'; factor = 0; }
    }
    if (status !== 'pend') factor = 0;
    return { ...r, status, need: r.ton * factor, needCu: r.cu * factor, p: parseBillet(r.pd) };
  }).sort((a, b) => (a.ini ?? 9e15) - (b.ini ?? 9e15) || a.n - b.n);

  // necesidad por código de palanquilla
  const N = new Map();
  for (const r of rows) {
    let x = N.get(r.pc);
    if (!x) {
      let st = S.get(r.pc), via = 'código';
      if (!st) { st = byKey.get(bkey(r.p)); via = st ? 'descripción' : ''; }
      x = { pc: r.pc, pd: r.pd, p: r.p, st, via, need: 0, cu: 0, orders: [], first: null, last: null, cum: 0, runOut: null };
      N.set(r.pc, x);
    }
    r.billet = x;
    if (r.need <= 0) continue;
    const avail = x.st ? x.st.disp : 0;
    const before = x.cum; x.cum += r.need;
    r.covered = Math.max(0, Math.min(r.need, avail - before));
    r.cov = r.covered >= r.need - 1e-6 ? 'ok' : r.covered > 0 ? 'parcial' : 'falta';
    if (r.cov !== 'ok' && !x.runOut) x.runOut = r;
    x.need += r.need; x.cu += r.needCu; x.orders.push(r);
    if (r.ini != null) { x.first = x.first == null ? r.ini : Math.min(x.first, r.ini); x.last = Math.max(x.last ?? 0, r.ini); }
  }
  const billets = [...N.values()].filter(x => x.need > 0);
  for (const x of billets) {
    x.stock = x.st ? x.st.disp : 0;
    x.bloq = x.st ? x.st.bloq / 1000 : 0;
    x.insp = x.st ? x.st.insp / 1000 : 0;
    x.saldo = x.stock - x.need;
    x.kgm = kgPerM(x.p);
    x.pieceT = x.p.length ? x.kgm * x.p.length / 1000 : null;
    x.stockCu = x.pieceT ? x.stock / x.pieceT : null;
    x.state = x.saldo >= 0 ? (x.saldo < x.need * opt.margin ? 'justo' : 'ok') : 'falta';
  }
  // saldo por material de stock (después del programa) para alternativas
  const usedBy = new Map(billets.filter(b => b.st).map(b => [b.st.mat, b]));
  const saldoMat = m => { const s = S.get(m); const b = usedBy.get(m); return s.disp - (b ? b.need : 0); };
  for (const x of billets) {
    if (x.state === 'ok') { x.alts = []; continue; }
    x.alts = [...S.values()]
      .filter(s => s !== x.st && s.p.section === x.p.section && norm(s.p.grade) === norm(x.p.grade))
      .map(s => ({ s, saldo: saldoMat(s.mat), longer: (s.p.length || 0) >= (x.p.length || 0) }))
      .filter(a => a.saldo > 0.05)
      .sort((a, b) => Math.abs((a.s.p.length || 0) - (x.p.length || 0)) - Math.abs((b.s.p.length || 0) - (x.p.length || 0)));
  }
  const order = { falta: 0, justo: 1, ok: 2 };
  billets.sort((a, b) => order[a.state] - order[b.state] || (a.runOut?.ini ?? a.first ?? 0) - (b.runOut?.ini ?? b.first ?? 0));

  // agrupaciones
  const group = (keyFn, labelFn) => {
    const G = new Map();
    for (const b of billets) {
      const k = keyFn(b.p);
      let g = G.get(k);
      if (!g) { g = { k, label: labelFn(b.p), need: 0, cu: 0, stockUsed: 0, falta: 0, codes: [], stockAll: 0 }; G.set(k, g); }
      g.need += b.need; g.cu += b.cu; g.stockUsed += b.stock; g.falta += Math.max(0, -b.saldo); g.codes.push(b);
    }
    for (const g of G.values()) {
      const others = [...S.values()].filter(s => keyFn(s.p) === g.k);
      g.stockAll = others.reduce((t, s) => t + s.disp, 0);
      g.stockItems = others;
    }
    return [...G.values()].sort((a, b) => b.falta - a.falta || b.need - a.need);
  };
  const byGrade = group(p => p.section + '|' + norm(p.grade), p => p.grade + ' · ' + p.section.replace('x', '×'));
  const byLength = group(p => p.section + '|' + p.length, p => lenLabel(p.length) + ' · ' + p.section.replace('x', '×'));

  const usedMats = new Set(billets.filter(b => b.st).map(b => b.st.mat));
  const unused = [...S.values()].filter(s => !usedMats.has(s.mat) && (s.disp > 0 || s.bloq > 0))
    .sort((a, b) => b.disp - a.disp);

  const pend = rows.filter(r => r.need > 0);
  const tot = {
    need: billets.reduce((t, b) => t + b.need, 0),
    cu: billets.reduce((t, b) => t + b.cu, 0),
    covered: pend.reduce((t, r) => t + r.covered, 0),
    falta: billets.reduce((t, b) => t + Math.max(0, -b.saldo), 0),
    codesFalta: billets.filter(b => b.state === 'falta').length,
    codesJusto: billets.filter(b => b.state === 'justo').length,
    codes: billets.length,
    ordersPend: pend.length,
    ordersFalta: pend.filter(r => r.cov !== 'ok').length,
    firstFalta: pend.find(r => r.cov !== 'ok') || null,
    excluded: rows.filter(r => r.status !== 'pend').length,
  };
  return { rows, billets, byGrade, byLength, unused, tot };
}

if (typeof module !== 'undefined') module.exports = { parseProgramWB, parseStockWB, analyze, parseBillet, dateFromName };
