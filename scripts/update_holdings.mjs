// 每日自動下載 ETF 持股檔並寫入 holdings/<帳戶>.json（只新增還沒有的日期，不動舊資料）
// 本機執行：  npm i xlsx  &&  node scripts/update_holdings.mjs
import fs from 'node:fs';
import path from 'node:path';
import XLSX from 'xlsx';

const ROOT = process.env.REPO_ROOT || process.cwd();
const HOLD_DIR = path.join(ROOT, 'holdings');
const EZ = process.env.EZMONEY_BASE || 'https://www.ezmoney.com.tw/ETF/Fund/AssetExcelNPOI?fundCode=';
const FH = process.env.FHTRUST_BASE || 'https://www.fhtrust.com.tw/api/assetsExcel/ETF23/';

/* ===== 要自動更新的 ETF（帳戶名稱留空 = 自動找「名稱含該代號」的既有帳戶，找不到就用代號新建） ===== */
const ETFS = [
  { code:'00403A', source:'ezmoney', fundCode:'63YTW', accountName:'' },
  { code:'00981A', source:'ezmoney', fundCode:'49YTW', accountName:'' },
  { code:'00988A', source:'ezmoney', fundCode:'61YTW', accountName:'' },
  { code:'00991A', source:'fhtrust', accountName:'' },   // 復華：依日期逐日下載（會回補最近 7 天沒抓到的日期）
];

const HEADER = {
  code: /代碼|代號|代码|symbol|ticker/i,
  name: /名稱|名称|company|name/i,
  shares: /股數|股数|股份|持股數|持股数|quantity|shares|張數|张数/i,
  weight: /權重|权重|比重|weight|佔比|占比/i,
};
const SKIP_CODE = /^(合計|總計|总计|小計|小计|total|sum|平均|備註|备注)$/i;
const CODE_OK = /^[0-9A-Za-z][0-9A-Za-z .\-\/]*$/;
const pad2 = n => String(n).padStart(2, '0');
const toNum = v => { if (v == null || v === '') return 0; if (typeof v === 'number') return isFinite(v) ? v : 0; const n = parseFloat(String(v).replace(/[,\s]/g, '')); return isNaN(n) ? 0 : n; };
const sanitize = n => (String(n || '').trim().replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '_') || 'unnamed').slice(0, 120);

export function validIso(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso); if (!m) return false;
  const y = +m[1], mo = +m[2], d = +m[3];
  if (y < 2000 || y > 2100 || mo < 1 || mo > 12 || d < 1) return false;
  return new Date(y, mo - 1, d).getMonth() === mo - 1;
}
export function findDate(text) {
  const s = String(text ?? '');
  let m = /(20\d{2})\s*[\/\-_.年]\s*(\d{1,2})\s*[\/\-_.月]\s*(\d{1,2})/.exec(s);
  if (m) { const iso = `${m[1]}-${pad2(+m[2])}-${pad2(+m[3])}`; if (validIso(iso)) return iso; }
  m = /(?:^|[^0-9])(20\d{2})(\d{2})(\d{2})(?![0-9])/.exec(s);
  if (m) { const iso = `${m[1]}-${m[2]}-${m[3]}`; if (validIso(iso)) return iso; }
  m = /(?:^|[^0-9])(1[01]\d)\s*[\/\-_.年]\s*(\d{1,2})\s*[\/\-_.月]\s*(\d{1,2})/.exec(s);
  if (m) { const iso = `${+m[1] + 1911}-${pad2(+m[2])}-${pad2(+m[3])}`; if (validIso(iso)) return iso; }
  m = /(?:^|[^0-9])(1[01]\d)(\d{2})(\d{2})(?![0-9])/.exec(s);   // 民國 7 碼，如 1151106
  if (m) { const iso = `${+m[1] + 1911}-${m[2]}-${m[3]}`; if (validIso(iso)) return iso; }
  return '';
}
function findHeader(rows) {
  for (let r = 0; r < Math.min(rows.length, 80); r++) {
    const row = rows[r] || [];
    let codeCol = -1, nameCol = -1, sharesCol = -1, weightCol = -1;
    row.forEach((c, i) => {
      const cell = String(c ?? '').trim();
      if (!cell || cell.length > 16) return;
      if (codeCol === -1 && HEADER.code.test(cell)) codeCol = i;
      else if (nameCol === -1 && HEADER.name.test(cell)) nameCol = i;
      else if (sharesCol === -1 && HEADER.shares.test(cell)) sharesCol = i;
      else if (weightCol === -1 && HEADER.weight.test(cell)) weightCol = i;
    });
    if (codeCol !== -1 && sharesCol !== -1) return { headerRow: r, codeCol, nameCol, sharesCol, weightCol };
  }
  return null;
}
function extractRows(rows, mp) {
  const map = new Map();
  for (let r = mp.headerRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (row.every(c => c === '' || c == null)) { if (map.size) break; else continue; }
    const code = String(row[mp.codeCol] ?? '').trim();
    if (map.size && HEADER.code.test(code)) break;
    if (!code || SKIP_CODE.test(code) || HEADER.code.test(code) || !CODE_OK.test(code)) continue;
    const shares = toNum(row[mp.sharesCol]);
    if (!(shares > 0)) continue;
    const name = mp.nameCol >= 0 ? String(row[mp.nameCol] ?? '').trim() : '';
    const w = (mp.weightCol >= 0 && row[mp.weightCol] != null && row[mp.weightCol] !== '') ? row[mp.weightCol] : null;
    if (map.has(code)) map.get(code)[1].shares += shares;
    else map.set(code, [code, { name, shares, dValue: w }]);
  }
  return [...map.values()];
}
/* 解析一份下載回來的 xlsx → { iso, rows:[[code,{name,shares,dValue}],...] } */
export function parseEtfWorkbook(buf) {
  const wb = XLSX.read(buf, { type: 'buffer' });
  let best = null;
  for (const sn of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: true });
    const mp = findHeader(rows); if (!mp) continue;
    const out = extractRows(rows, mp);
    if (out.length && (!best || out.length > best.out.length)) best = { sn, rows, mp, out };
  }
  if (!best) throw new Error('找不到含「代碼」與「股數」欄位標題的工作表');
  let iso = '';
  const scan = [];
  for (const sn of wb.SheetNames) {
    const rows = sn === best.sn ? best.rows : XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: true });
    const lim = sn === best.sn ? best.mp.headerRow : Math.min(rows.length, 12);
    for (let r = 0; r < lim; r++) (rows[r] || []).forEach(c => { if (c !== '' && c != null) scan.push(String(c)); });
  }
  for (const c of scan) { iso = findDate(c); if (iso) break; }
  return { iso, rows: best.out };
}

/* ===== 讀寫 holdings/*.json（格式與網頁的 serializeOneAccount 相同） ===== */
function loadAccounts() {
  fs.mkdirSync(HOLD_DIR, { recursive: true });
  return fs.readdirSync(HOLD_DIR).filter(f => f.endsWith('.json')).map(f => {
    try { return { file: f, data: JSON.parse(fs.readFileSync(path.join(HOLD_DIR, f), 'utf8')) }; } catch { return null; }
  }).filter(Boolean);
}
function resolveAccount(etf, accounts) {
  if (etf.accountName) {
    const hit = accounts.find(a => a.data.name === etf.accountName);
    return hit || { file: sanitize(etf.accountName) + '.json', data: null, name: etf.accountName };
  }
  const hits = accounts.filter(a => String(a.data.name || '').toUpperCase().includes(etf.code));
  if (hits.length === 1) return hits[0];
  if (hits.length > 1) { console.warn(`⚠️ ${etf.code}：有多個帳戶名稱含此代號（${hits.map(h => h.data.name).join('、')}），請在 ETFS 設定 accountName。本次略過。`); return null; }
  return { file: sanitize(etf.code) + '.json', data: null, name: etf.code };
}
function rocOf(iso) { return String(+iso.slice(0, 4) - 1911) + iso.slice(5, 7) + iso.slice(8, 10); }
function addSheet(acc, iso, rows) {
  const roc = rocOf(iso);
  const data = acc.data || { id: 'a' + Date.now().toString(36), name: acc.name, color: '#1F6F78', shape: 'circle', included: true, sheets: [] };
  if (data.sheets.some(s => s.roc === roc || s.iso === iso)) return { added: false, acc: { ...acc, data } };
  data.sheets.push({ roc, iso, rows });
  data.sheets.sort((a, b) => (a.roc < b.roc ? -1 : a.roc > b.roc ? 1 : 0));
  fs.writeFileSync(path.join(HOLD_DIR, acc.file), JSON.stringify(data));
  return { added: true, acc: { ...acc, data } };
}

/* ===== 下載 ===== */
const BASE_HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
  'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*;q=0.8',
  'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
  'Sec-Fetch-Dest': 'document', 'Sec-Fetch-Mode': 'navigate', 'Sec-Fetch-Site': 'same-origin',
};
async function download(url, referer) {
  let lastErr;
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const res = await fetch(url, { headers: { ...BASE_HEADERS, Referer: referer || new URL(url).origin + '/' }, redirect: 'follow', signal: AbortSignal.timeout(60000) });
      const buf = Buffer.from(await res.arrayBuffer());
      const ct = res.headers.get('content-type') || '';
      const isXlsx = buf.length > 100 && buf[0] === 0x50 && buf[1] === 0x4B;
      if (res.ok && isXlsx) return buf;
      const body = buf.toString('utf8', 0, 200).replace(/\s+/g, ' ').trim();
      const blocked = /cloudflare|attention required|access denied|captcha|just a moment|forbidden/i.test(body + ' ' + (res.headers.get('server') || ''));
      throw new Error(`HTTP ${res.status}｜content-type: ${ct || '無'}｜大小 ${buf.length} bytes｜${blocked ? '疑似被網站防護機制擋下｜' : ''}回應開頭：${body || '(空)'}`);
    } catch (e) {
      // Node 的 fetch 連線失敗時只會說 "fetch failed"，真正原因藏在 e.cause 裡，這裡把它挖出來
      const c = e && e.cause;
      if (c && !/原因：/.test(e.message)) e.message += `｜原因：${c.code || c.name || ''} ${c.message || ''}`.trimEnd();
      lastErr = e;
      if (attempt < 2) await new Promise(r => setTimeout(r, 3000));
    }
  }
  throw lastErr;
}
const twToday = () => { const p = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei' }).format(new Date()); return p; };

export async function run() {
  let accounts = loadAccounts();
  let changed = 0, failed = 0;
  const ONLY = (process.env.ONLY_SOURCE || '').split(',').map(x => x.trim()).filter(Boolean);   // 例：ONLY_SOURCE=ezmoney 只更新 ezmoney 那幾檔
  for (const etf of ETFS) {
    if (ONLY.length && !ONLY.includes(etf.source)) continue;
    try {
      const acc = resolveAccount(etf, accounts);
      if (!acc) continue;
      let candidates = [];   // [{iso, buf}]
      if (etf.source === 'ezmoney') {
        const parsed = parseEtfWorkbook(await download(EZ + etf.fundCode, 'https://www.ezmoney.com.tw/ETF/Fund/Info?fundCode=' + etf.fundCode));
        let iso = parsed.iso;
        if (!iso) { iso = twToday(); console.warn(`⚠️ ${etf.code}：檔案內找不到日期，改用今天（台北）${iso}`); }
        candidates.push({ iso, rows: parsed.rows });
      } else {   // fhtrust：依日期逐日下載，往回補最近 7 天
        const have = new Set((acc.data?.sheets || []).map(s => s.iso));
        const base = new Date(twToday() + 'T00:00:00Z');
        for (let i = 0; i < 8; i++) {
          const d = new Date(base.getTime() - i * 86400000);
          const iso = d.toISOString().slice(0, 10);
          if ([0, 6].includes(d.getUTCDay()) || have.has(iso)) continue;
          try {
            const parsed = parseEtfWorkbook(await download(FH + iso.replace(/-/g, '')));
            candidates.push({ iso: parsed.iso || iso, rows: parsed.rows });
          } catch (e) { console.log(`  ${etf.code} ${iso}：${e.message}（可能為非交易日）`); }
        }
      }
      for (const c of candidates) {
        if (c.rows.length < 5) { console.warn(`⚠️ ${etf.code} ${c.iso}：只解析到 ${c.rows.length} 檔持股，疑似異常，略過`); continue; }
        const r = addSheet(acc, c.iso, c.rows);
        if (r.added) { changed++; console.log(`✅ ${etf.code} → ${r.acc.data.name} ${c.iso} 新增（${c.rows.length} 檔持股）`); acc.data = r.acc.data; }
        else console.log(`➖ ${etf.code} ${c.iso} 已存在，略過`);
      }
      accounts = loadAccounts();
    } catch (e) { failed++; console.error(`❌ ${etf.code}：${e.message}`); }
  }
  console.log(`完成：新增 ${changed} 個日期，失敗 ${failed} 檔。`);
  return { changed, failed };
}
if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith('update_holdings.mjs')) {
  run().then(r => { if (r.failed > 0) process.exitCode = 1; });   // 有任何一檔失敗就讓 Actions 顯示紅燈（GitHub 會寄信通知）
}
