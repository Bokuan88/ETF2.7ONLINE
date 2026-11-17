// 每日自動下載 ETF 持股檔並寫入 holdings/<帳戶>.json（只新增還沒有的日期，不動舊資料）
// 本機執行：  npm i xlsx  &&  node scripts/update_holdings.mjs
import fs from 'node:fs';
import path from 'node:path';
import XLSX from 'xlsx';

const ROOT = process.env.REPO_ROOT || process.cwd();
const HOLD_DIR = path.join(ROOT, 'holdings');
const EZ = process.env.EZMONEY_BASE || 'https://www.ezmoney.com.tw/ETF/Fund/AssetExcelNPOI?fundCode=';
const FH = process.env.FHTRUST_BASE || 'https://www.fhtrust.com.tw/api/assetsExcel/ETF23/';
const AZ = process.env.ALLIANZ_BASE || 'https://etf.allianzgi.com.tw';
const NM = process.env.NOMURA_API || 'https://www.nomurafunds.com.tw/API/ETFAPI/api/Fund/GetFundAssets';

/* ===== 要自動更新的 ETF（帳戶名稱留空 = 自動找「名稱含該代號」的既有帳戶，找不到就用代號新建） ===== */
const ETFS = [
  { code:'00403A', source:'ezmoney', fundCode:'63YTW', accountName:'' },
  { code:'00981A', source:'ezmoney', fundCode:'49YTW', accountName:'' },
  { code:'00988A', source:'ezmoney', fundCode:'61YTW', accountName:'' },
  { code:'00991A', source:'fhtrust', accountName:'' },   // 復華：依日期逐日下載（會回補最近 7 天沒抓到的日期）
  { code:'00985A', source:'nomura', accountName:'' },    // 野村：直接呼叫網站的 JSON 介面（不是 Excel）
  { code:'00980A', source:'nomura', accountName:'' },
  { code:'00402A', source:'allianz', fundId:'E0003', dateField:'NavDate', accountName:'' },   // 安聯：內部代號 E0003；資料日期欄位可改成 'PCFDate'
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
/* 手動追蹤跳轉並自動帶 cookie（有些網站要先發 cookie 才放行，不帶就會一直跳轉）；失敗時回報完整跳轉路徑 */
async function fetchFollow(url, referer) {
  const jar = new Map();
  const chain = [];
  let cur = url;
  for (let hop = 0; hop < 10; hop++) {
    const headers = { ...BASE_HEADERS, Referer: referer || new URL(cur).origin + '/' };
    if (jar.size) headers.Cookie = [...jar].map(([k, v]) => `${k}=${v}`).join('; ');
    const res = await fetch(cur, { headers, redirect: 'manual', signal: AbortSignal.timeout(60000) });
    const setCookies = typeof res.headers.getSetCookie === 'function' ? res.headers.getSetCookie() : [];
    setCookies.forEach(c => { const kv = c.split(';')[0]; const i = kv.indexOf('='); if (i > 0) jar.set(kv.slice(0, i).trim(), kv.slice(i + 1).trim()); });
    if (res.status >= 300 && res.status < 400 && res.headers.get('location')) {
      const next = new URL(res.headers.get('location'), cur).toString();
      chain.push(`${res.status} ${cur.replace(/^https?:\/\//, '').slice(0, 70)} → ${next.replace(/^https?:\/\//, '').slice(0, 70)}${setCookies.length ? '（發了 ' + setCookies.length + ' 個 cookie）' : ''}`);
      cur = next;
      continue;
    }
    return { res, chain };
  }
  const err = new Error('跳轉超過 10 次｜跳轉路徑：' + chain.slice(0, 4).join(' ‖ '));
  throw err;
}
async function download(url, referer) {
  let lastErr;
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const { res, chain } = await fetchFollow(url, referer);
      const buf = Buffer.from(await res.arrayBuffer());
      const ct = res.headers.get('content-type') || '';
      const isXlsx = buf.length > 100 && buf[0] === 0x50 && buf[1] === 0x4B;
      if (res.ok && isXlsx) return buf;
      const body = buf.toString('utf8', 0, 200).replace(/\s+/g, ' ').trim();
      const blocked = /cloudflare|attention required|access denied|captcha|just a moment|forbidden/i.test(body + ' ' + (res.headers.get('server') || ''));
      throw new Error(`HTTP ${res.status}｜content-type: ${ct || '無'}｜大小 ${buf.length} bytes｜${blocked ? '疑似被網站防護機制擋下｜' : ''}${chain.length ? '曾跳轉：' + chain.join(' ‖ ') + '｜' : ''}回應開頭：${body || '(空)'}`);
    } catch (e) {
      const c = e && e.cause;   // Node 連線失敗只會說 "fetch failed"，真正原因在 e.cause
      if (c && !/原因：/.test(e.message)) e.message += `｜原因：${c.code || c.name || ''} ${c.message || ''}`.trimEnd();
      lastErr = e;
      if (attempt < 2) await new Promise(r => setTimeout(r, 3000));
    }
  }
  throw lastErr;
}
const twToday = () => { const p = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei' }).format(new Date()); return p; };

/* ===== 野村投信：網站持股頁背後的 JSON 介面（回傳最新一個資料日的完整股票清單） ===== */
export function parseNomura(json) {
  const d = json?.Entries?.Data;
  const navDate = d?.FundAsset?.NavDate;                      // 例如 2026/11/16
  const iso = findDate(navDate);
  if (!iso) throw new Error('野村回應裡找不到資料日期（NavDate）：' + JSON.stringify(navDate));
  const table = (d?.Table || []).find(t => t.TableTitle === '股票');
  if (!table) throw new Error('野村回應裡找不到標題為「股票」的表格');
  const cols = (table.Columns || []).map(c => String(c.Name || ''));
  const ic = cols.findIndex(c => /代號|代碼/.test(c)), iname = cols.findIndex(c => /名稱/.test(c));
  const ish = cols.findIndex(c => /股數/.test(c)), iw = cols.findIndex(c => /權重|比重/.test(c));
  if (ic < 0 || ish < 0) throw new Error('野村表格欄位名稱不符預期：' + cols.join('、'));
  const map = new Map();
  for (const r of table.Rows || []) {
    const code = String(r[ic] ?? '').trim();
    const shares = toNum(r[ish]);
    if (!code || !(shares > 0) || SKIP_CODE.test(code) || !CODE_OK.test(code)) continue;
    const w = iw >= 0 && r[iw] !== '' && r[iw] != null ? toNum(r[iw]) : null;
    if (map.has(code)) map.get(code)[1].shares += shares;
    else map.set(code, [code, { name: iname >= 0 ? String(r[iname] ?? '').trim() : '', shares, dValue: w }]);
  }
  return { iso, rows: [...map.values()] };
}
async function fetchNomura(code) {
  let lastErr;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const res = await fetch(NM, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json, text/plain, */*',
          'Origin': 'https://www.nomurafunds.com.tw',
          'Referer': `https://www.nomurafunds.com.tw/ETFWEB/product-description?fundNo=${code}&tab=Shareholding`,
          'User-Agent': BASE_HEADERS['User-Agent'],
          'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
        },
        body: JSON.stringify({ FundID: code, SearchDate: null }),
        signal: AbortSignal.timeout(60000),
      });
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}｜回應開頭：${text.slice(0, 200).replace(/\s+/g, ' ')}`);
      try { return JSON.parse(text); } catch { throw new Error('回應不是 JSON｜開頭：' + text.slice(0, 200).replace(/\s+/g, ' ')); }
    } catch (e) {
      const c = e && e.cause;
      if (c && !/原因：/.test(e.message)) e.message += `｜原因：${c.code || c.name || ''} ${c.message || ''}`.trimEnd();
      lastErr = e;
      if (attempt < 3) await new Promise(r => setTimeout(r, 3000 * attempt));
    }
  }
  throw lastErr;
}

/* ===== 安聯投信：先開持股頁取得 cookie 與 X-XSRF-TOKEN，再呼叫 JSON 介面 ===== */
export function parseAllianz(json, dateField = 'NavDate') {
  const d = json?.Entries?.Data;
  const rawDate = d?.FundAsset?.[dateField];
  const iso = findDate(rawDate);
  if (!iso) throw new Error(`安聯回應裡找不到資料日期（${dateField}）：` + JSON.stringify(rawDate));
  const table = (d?.Table || []).find(t => /^股票/.test(String(t.TableTitle || '')) && (t.Columns || []).some(c => /代號|代碼/.test(String(c?.Name || ''))));
  if (!table) throw new Error('安聯回應裡找不到「股票」明細表格');
  const cols = (table.Columns || []).map(c => String(c?.Name || ''));
  const ic = cols.findIndex(c => /代號|代碼/.test(c)), iname = cols.findIndex(c => /名稱/.test(c));
  const ish = cols.findIndex(c => /股數/.test(c)), iw = cols.findIndex(c => /權重|比重/.test(c));
  if (ic < 0 || ish < 0) throw new Error('安聯表格欄位名稱不符預期：' + cols.join('、'));
  const map = new Map();
  for (const r of table.Rows || []) {
    const code = String(r[ic] ?? '').trim();          // 例如 "AAPL US"
    const shares = toNum(r[ish]);                      // "39,400" → 39400
    if (!code || !(shares > 0) || SKIP_CODE.test(code) || !CODE_OK.test(code)) continue;
    const w = iw >= 0 && r[iw] !== '' && r[iw] != null ? toNum(r[iw]) : null;   // "9.23%" → 9.23
    if (map.has(code)) map.get(code)[1].shares += shares;
    else map.set(code, [code, { name: iname >= 0 ? String(r[iname] ?? '').trim() : '', shares, dValue: w }]);
  }
  return { iso, rows: [...map.values()] };
}
async function fetchAllianz(fundId) {
  const pageUrl = `${AZ}/etf-info/${fundId}?tab=4`;
  const jar = new Map();
  const absorb = res => (typeof res.headers.getSetCookie === 'function' ? res.headers.getSetCookie() : []).forEach(c => {
    const kv = c.split(';')[0]; const i = kv.indexOf('='); if (i > 0) jar.set(kv.slice(0, i).trim(), kv.slice(i + 1).trim());
  });
  const cookie = () => [...jar].map(([k, v]) => `${k}=${v}`).join('; ');
  let lastErr;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      jar.clear();
      let cur = pageUrl, res;
      for (let hop = 0; hop < 6; hop++) {
        const h = { ...BASE_HEADERS, Accept: 'text/html,application/xhtml+xml,*/*;q=0.8', 'Sec-Fetch-Site': 'none' };
        if (jar.size) h.Cookie = cookie();
        res = await fetch(cur, { headers: h, redirect: 'manual', signal: AbortSignal.timeout(60000) });
        absorb(res);
        if (res.status >= 300 && res.status < 400 && res.headers.get('location')) { cur = new URL(res.headers.get('location'), cur).toString(); continue; }
        break;
      }
      if (!res.ok) throw new Error(`開啟持股頁失敗 HTTP ${res.status}`);
      let token = jar.get('X-XSRF-TOKEN');
      if (!token) throw new Error('網站沒有發 X-XSRF-TOKEN cookie（收到的 cookie：' + ([...jar.keys()].join('、') || '無') + '）｜可能被防護機制擋下或網站改版');
      if (token.includes('%')) { try { token = decodeURIComponent(token); } catch {} }
      const r = await fetch(`${AZ}/webapi/api/Fund/GetFundAssets`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json', 'Accept': 'application/json, text/plain, */*',
          'Origin': AZ, 'Referer': pageUrl, 'X-XSRF-TOKEN': token, 'Cookie': cookie(),
          'User-Agent': BASE_HEADERS['User-Agent'], 'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
          'Sec-Fetch-Dest': 'empty', 'Sec-Fetch-Mode': 'cors', 'Sec-Fetch-Site': 'same-origin',
        },
        body: JSON.stringify({ FundID: fundId }),
        signal: AbortSignal.timeout(60000),
      });
      const text = await r.text();
      if (!r.ok) throw new Error(`HTTP ${r.status}｜回應開頭：${text.slice(0, 200).replace(/\s+/g, ' ')}`);
      try { return JSON.parse(text); } catch { throw new Error('回應不是 JSON｜開頭：' + text.slice(0, 200).replace(/\s+/g, ' ')); }
    } catch (e) {
      const c = e && e.cause;
      if (c && !/原因：/.test(e.message)) e.message += `｜原因：${c.code || c.name || ''} ${c.message || ''}`.trimEnd();
      lastErr = e;
      if (attempt < 3) await new Promise(r => setTimeout(r, 3000 * attempt));
    }
  }
  throw lastErr;
}

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
      } else if (etf.source === 'nomura') {
        const parsed = parseNomura(await fetchNomura(etf.code));
        candidates.push({ iso: parsed.iso, rows: parsed.rows });
      } else if (etf.source === 'allianz') {
        const parsed = parseAllianz(await fetchAllianz(etf.fundId), etf.dateField || 'NavDate');
        candidates.push({ iso: parsed.iso, rows: parsed.rows });
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
