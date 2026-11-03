/* facelift 開題指令產生器（GitLab 版）
 * 參數池讀同一層的「參數池.xlsx」（SheetJS 解析）；讀不到時可手動選檔，選過的檔會存在這台瀏覽器。
 * 防撞題紀錄、複製紀錄、當天卡片都只存在這台瀏覽器（localStorage）。
 * 本頁不呼叫任何 AI API，只產生指令文字。
 *
 * 區塊：常數與寫法規則 → 讀 Excel 與檢查 → 抽籤 roll() → 組指令 promptText() → 卡片 → 紀錄 → 初始化
 */
'use strict';

/* ================= 常數與寫法規則（不在 Excel 裡，改這裡要動程式） ================= */
const POOL_FILE = '參數池.xlsx';
const KEY = 'facelift-'; // localStorage 前綴
const CATS = ['電音波', '針劑', '醫美診所', '空包彈', '閒聊趣味'];
const SHORT = {'電音波': '電音波', '針劑': '針劑', '醫美診所': '診所', '空包彈': '空包彈', '閒聊趣味': '閒聊'};
const R = '__rand';
// 今日配比預設（工作流程表的每日配比）
const PRESETS = {
  an: {'閒聊趣味': 1, '醫美診所': 1, '電音波': 1, '針劑': 0, '空包彈': 2},
  yy: {'閒聊趣味': 1, '醫美診所': 0, '電音波': 0, '針劑': 1, '空包彈': 3},
  all: {'閒聊趣味': 2, '醫美診所': 1, '電音波': 1, '針劑': 1, '空包彈': 5},
};
// 地區熱區與每週輪替
const CITY = {'台北': ['台北 大安區', '台北 中山區', '台北 中正區', '台北 信義區', '台北 松山區'], '新北': ['新北 板橋'], '台中': ['台中 西屯', '台中 西區'], '高雄': ['高雄 左營'], '桃園': ['桃園'], '新竹': ['新竹'], '台南': ['台南']};
const ROTATE = ['台北', '新北', '台中', '高雄'];
// 發文者選項（tone、trigger 由 Excel 填入）
const P = {
  gender: ['女', '女', '女', '女', '男'],
  who: ['自己', '自己', '自己', '自己', '自己', '陪朋友去諮詢', '幫伴侶問', '幫媽媽問', '跟姊妹一起考慮'],
  exp: ['完全沒做過醫美', '只做過雷射除斑', '打過一兩次肉毒', '每年固定打一次電音波', '做過不少療程，算老手', '做過一次但效果很差', '只在美容院做過臉'],
  budget: ['一萬以內', '一到三萬', '三到六萬', '六萬以上', '沒概念，想先了解行情'],
  len: ['35–90', '90–130', '130–200'],
  tone: [],
  trigger: [],
};
const TEAM_PHRASES = ['看到有人說', '想問大家', '有沒有人', '可以分享', '爬文看', '不知道', '是不是', '真的有', '會不會'];
// 閒聊趣味、醫美診所分頁：情境有列真實品類時，配到品類（並指定療程）的機率，依舊版實際比例
const REAL_GROUP_RATE = {'閒聊趣味': .35, '醫美診所': .20};
// 心得情境合起來在抽籤時算幾個情境的份量（舊版 3 個心得情境）
const SHARE_SLOTS = 3;
// 心得結論分配：紀錄、比較各 15%，其餘 70% 照滿意 45／普通 35／後悔 20
const MOOD_W = {'紀錄': 15, '比較': 15, '滿意': 31.5, '普通': 24.5, '後悔': 14};
const SCENE_TYPES = ['一般', '心得', '求推', '診所主軸'];
const MOODS = ['滿意', '普通', '後悔', '紀錄', '比較'];
const FEMALE_RE = /產後|陰道|小陰唇|私密|乳|胸|漏尿|孕斑|生理期|比基尼|伴娘|婚紗|妊娠/;
const TOPIC_FIELDS = ['f-group', 'f-obj', 'f-pain', 'f-scene', 'f-tag'];
const SEASON_NAME = m => m >= 3 && m <= 5 ? '春天' : m >= 6 && m <= 8 ? '夏天' : m >= 9 && m <= 11 ? '秋天' : '冬天';

/* ---------- 由 Excel 填入的資料 ---------- */
let DATA = {groups: {}, tab: {}, tabPain: {}, scene: []};
let TONE_INFO = {};      // 語氣 → {notShare, notCC}
let TRIGGERS = [];       // [{text, m:[月份]}]
let SEASON_RULES = [];   // [{re, m, label}]
let SHEET_BAN = [];
let LIB = [];            // 範例文 [{t,g,s,y,n,x}]
let ready = false;
let MONTH = new Date().getMonth() + 1;

/* ---------- 小工具 ---------- */
const $ = id => document.getElementById(id);
const store = {
  get(k) { try { return localStorage.getItem(KEY + k) } catch (e) { return null } },
  set(k, v) { try { localStorage.setItem(KEY + k, v); return true } catch (e) { return false } },
};
const pick = a => a[Math.floor(Math.random() * a.length)];
const rint = (a, b) => a + Math.floor(Math.random() * (b - a + 1));
function sample(a, n) { const c = [...a]; for (let i = c.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [c[i], c[j]] = [c[j], c[i]] } return c.slice(0, n) }
function wpick(items, w) { const t = items.reduce((s, x) => s + w(x), 0); let r = Math.random() * t; for (const x of items) { r -= w(x); if (r <= 0) return x } return items[items.length - 1] }
function esc(s) { return String(s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c])) }
const escRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const list = s => String(s || '').split(/[、,，\n]/).map(x => x.trim()).filter(Boolean);
const monthsOf = s => list(s).map(x => +x.replace(/月$/, '')).filter(n => Number.isInteger(n) && n >= 1 && n <= 12);

/* ---------- 季節 ---------- */
function inSeason(s) { return SEASON_RULES.every(r => !r.re.test(s) || r.m.includes(MONTH)) }
function seasonPool(arr, f) { const a = arr.filter(x => inSeason(f ? f(x) : x)); return a.length ? a : arr }

/* ---------- 防撞題計數 ---------- */
const TAB_FIELDS = new Set(['scene', 'group', 'obj', 'tone', 'title', 'open', 'tt']);
const REC_FIELDS = ['scene', 'pain', 'group', 'obj', 'tone', 'title', 'open', 'trigger', 'layout', 'tt'];
function mkU() {
  const m = new Map(); const key = (f, t, v) => f + '|' + (TAB_FIELDS.has(f) ? t : '*') + '|' + v;
  return {n: (f, t, v) => m.get(key(f, t, v)) || 0, add: (f, t, v, w = 1) => { if (v == null || v === '') return; const k = key(f, t, v); m.set(k, (m.get(k) || 0) + w) }};
}
function pickLeast(pool, f, t, U) { if (!pool.length) return undefined; let min = Infinity; pool.forEach(v => { const c = U.n(f, t, v); if (c < min) min = c }); const c = pool.filter(v => U.n(f, t, v) === min); return pick(c.length ? c : pool) }

/* ================= 讀 Excel 與檢查 ================= */
// 依表頭文字開頭找欄位；找不到必填欄就丟錯
function sheetRows(wb, name, cols, warns, required) {
  const ws = wb.Sheets[name];
  if (!ws) {
    if (required) throw new Error(`找不到「${name}」分頁`);
    warns.push(`找不到「${name}」分頁，這部分視為空白`);
    return [];
  }
  const start = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0;
  const rows = XLSX.utils.sheet_to_json(ws, {header: 1, defval: '', raw: false, blankrows: true});
  const head = (rows[0] || []).map(h => String(h).trim());
  const ix = {};
  for (const [key, label] of Object.entries(cols)) {
    const i = head.findIndex(h => h.startsWith(label));
    if (i < 0) throw new Error(`「${name}」分頁找不到「${label}」欄`);
    ix[key] = i;
  }
  const out = [];
  rows.slice(1).forEach((r, k) => {
    const o = {line: start + k + 2}; let any = false;
    for (const key in ix) { const v = String(r[ix[key]] ?? '').replace(/\r/g, '').trim(); o[key] = v; if (v && key !== 'on') any = true }
    if (any) out.push(o);
  });
  return out;
}

function buildPool(wb) {
  const warns = [];
  const W = (sheet, line, msg) => warns.push(`${sheet} 第 ${line} 列：${msg}`);
  const enabled = (sheet, o) => { if (/^y$/i.test(o.on)) return true; if (!o.on) W(sheet, o.line, '「啟用」沒填，已略過'); return false };

  // 品類
  const allG = {};
  sheetRows(wb, '品類', {name: '品類名稱', tabs: '出現在哪些分頁', surg: '手術類', w: '抽中比重', on: '啟用'}, warns, true).forEach(o => {
    if (!enabled('品類', o)) return;
    if (!o.name) { W('品類', o.line, '品類名稱空白'); return }
    if (allG[o.name]) { W('品類', o.line, `品類「${o.name}」重複，只用第一列`); return }
    const tabs = list(o.tabs).filter(t => { if (CATS.includes(t)) return true; W('品類', o.line, `出現在哪些分頁「${t}」不是五個分頁之一`); return false });
    if (!tabs.length) W('品類', o.line, `品類「${o.name}」沒有填出現在哪些分頁，任何分頁都抽不到`);
    let w = null;
    if (o.w !== '') { w = parseFloat(o.w); if (!(w >= 0)) { W('品類', o.line, `抽中比重「${o.w}」不是數字，改用 2`); w = null } }
    allG[o.name] = {tabs, surgery: /^y$/i.test(o.surg), weight: w, objs: [], objInfo: {}, pains: []};
  });

  // 療程、困擾先讀進來
  const oRows = sheetRows(wb, '療程', {g: '品類', name: '療程名稱', nick: '板上也常叫', only: '只配這些困擾', on: '啟用'}, warns, true).filter(o => enabled('療程', o));
  const pRows = sheetRows(wb, '困擾', {g: '品類', pain: '困擾', on: '啟用'}, warns, true).filter(o => enabled('困擾', o));
  oRows.forEach(o => {
    if (!o.g) { W('療程', o.line, '品類空白'); return }
    if (!allG[o.g]) { W('療程', o.line, `品類「${o.g}」不在品類分頁（或已停用）`); return }
    if (!o.name) { W('療程', o.line, '療程名稱空白'); return }
    const G = allG[o.g];
    if (G.objs.includes(o.name)) { W('療程', o.line, `療程「${o.name}」在同一品類重複`); return }
    G.objs.push(o.name); G.objInfo[o.name] = {nick: list(o.nick).filter(n => n !== o.name), only: list(o.only), line: o.line};
  });

  // 沒有療程、只出現在診所／閒聊分頁的品類（「診所」「閒聊」）＝分頁專用的困擾池
  const groups = {}, pseudo = {};
  Object.entries(allG).forEach(([g, v]) => {
    if (!v.objs.length && v.tabs.length && v.tabs.every(t => t === '醫美診所' || t === '閒聊趣味')) pseudo[g] = v.tabs;
    else groups[g] = v;
  });
  const tab = {}; CATS.forEach(c => tab[c] = Object.keys(groups).filter(g => groups[g].tabs.includes(c)));
  const tabPain = {};
  Object.values(pseudo).forEach(ts => ts.forEach(t => tabPain[t] = tabPain[t] || []));
  pRows.forEach(o => {
    if (!o.g) { W('困擾', o.line, '品類空白'); return }
    if (!o.pain) { W('困擾', o.line, '困擾空白'); return }
    if (groups[o.g]) { if (!groups[o.g].pains.includes(o.pain)) groups[o.g].pains.push(o.pain) }
    else if (pseudo[o.g]) pseudo[o.g].forEach(t => { if (!tabPain[t].includes(o.pain)) tabPain[t].push(o.pain) });
    else W('困擾', o.line, `品類「${o.g}」不在品類分頁（或已停用）`);
  });
  // 只配這些困擾：要對得到同品類的困擾
  Object.entries(groups).forEach(([g, G]) => Object.entries(G.objInfo).forEach(([o, info]) => {
    const bad = info.only.filter(p => !G.pains.includes(p));
    if (bad.length) W('療程', info.line, `「${o}」只配的困擾「${bad.join('、')}」不在「${g}」的困擾裡，已忽略`);
    info.only = info.only.filter(p => G.pains.includes(p));
    if (!info.only.length) delete info.only;
  }));

  // 情境
  const scene = []; const seen = {};
  sheetRows(wb, '情境', {name: '情境名稱', dir: '給 AI 看的方向', desc: '情境說明', gs: '適用品類', noSurg: '不配手術', type: '類型', mood: '心得結論', on: '啟用'}, warns, true).forEach(o => {
    if (!enabled('情境', o)) return;
    if (!o.name) { W('情境', o.line, '情境名稱空白'); return }
    let name = o.name;
    if (seen[name]) { seen[name]++; name = `${o.name}-${seen[o.name]}`; W('情境', o.line, `情境名稱「${o.name}」重複（第 ${seen[o.name + '#line']} 列也有），這列暫時當成「${name}」，建議在 Excel 改名`) }
    else { seen[name] = 1; seen[name + '#line'] = o.line }
    const entries = list(o.gs);
    if (!entries.length) { W('情境', o.line, `情境「${name}」沒有填適用品類，抽不到`); return }
    const tabs = new Set(); const real = []; let all = false;
    entries.forEach(e => {
      if (e === '全部') { all = true; CATS.forEach(t => tabs.add(t)); return }
      if (groups[e]) { real.push(e); groups[e].tabs.forEach(t => tabs.add(t)) }
      else if (pseudo[e]) pseudo[e].forEach(t => tabs.add(t));
      else W('情境', o.line, `適用品類「${e}」不在品類分頁（或已停用），已忽略`);
    });
    let type = o.type || '一般';
    if (!SCENE_TYPES.includes(type)) { W('情境', o.line, `類型「${type}」看不懂，當成一般`); type = '一般' }
    let mood = o.mood || null;
    if (mood && !MOODS.includes(mood)) { W('情境', o.line, `心得結論「${mood}」看不懂，改成隨機`); mood = null }
    if (type === '心得' && !mood) W('情境', o.line, `心得類情境「${name}」沒填心得結論，會隨機`);
    if (!o.dir) W('情境', o.line, `情境「${name}」沒填給 AI 看的方向，改用情境名稱`);
    const s = {name, dir: o.dir || name, desc: o.desc || o.dir || name, cats: CATS.filter(t => tabs.has(t)), type, share: type === '心得', mood: type === '心得' ? mood : null};
    if (!all) s.only = real;
    if (/^y$/i.test(o.noSurg)) s.noSurg = true;
    if (!s.cats.length) { W('情境', o.line, `情境「${name}」對不到任何分頁，抽不到`); return }
    scene.push(s);
  });

  // 語氣
  const tones = [];
  sheetRows(wb, '語氣', {text: '語氣', not: '不要用在', on: '啟用'}, warns, true).forEach(o => {
    if (!enabled('語氣', o)) return;
    if (!o.text) { W('語氣', o.line, '語氣空白'); return }
    if (o.not && !/心得|閒聊|診所/.test(o.not)) W('語氣', o.line, `不要用在「${o.not}」看不懂，可填「不用在心得文」或「不用在閒聊和診所」`);
    tones.push({text: o.text, notShare: /心得/.test(o.not), notCC: /閒聊|診所/.test(o.not)});
  });
  // 觸發事件
  const triggers = [];
  sheetRows(wb, '觸發事件', {text: '事件', m: '只在這些月份', on: '啟用'}, warns, true).forEach(o => {
    if (!enabled('觸發事件', o)) return;
    if (!o.text) { W('觸發事件', o.line, '事件空白'); return }
    const m = monthsOf(o.m);
    if (o.m && !m.length) W('觸發事件', o.line, `月份「${o.m}」看不懂，當成不限月份`);
    triggers.push({text: o.text, m});
  });
  // 禁用詞
  const ban = [];
  sheetRows(wb, '禁用詞', {text: '禁用詞', on: '啟用'}, warns, false).forEach(o => { if (enabled('禁用詞', o) && !ban.includes(o.text)) ban.push(o.text) });
  // 季節規則
  const season = [];
  sheetRows(wb, '季節規則', {kw: '關鍵字', m: '只在這些月份出現', label: '說明', on: '啟用'}, warns, false).forEach(o => {
    if (!enabled('季節規則', o)) return;
    const kw = list(o.kw), m = monthsOf(o.m);
    if (!kw.length) { W('季節規則', o.line, '關鍵字空白'); return }
    if (!m.length) { W('季節規則', o.line, `月份「${o.m}」看不懂，這條規則略過`); return }
    season.push({re: new RegExp(kw.map(escRe).join('|')), m, label: o.label || kw.join('、')});
  });
  // 範例文
  const lib = []; const badG = {};
  sheetRows(wb, '範例文', {t: '標籤', g: '品類', s: '跟診所有關', y: '年份', x: '全文', on: '啟用'}, warns, true).forEach(o => {
    if (!enabled('範例文', o)) return;
    const x = o.x.replace(/\n{3,}/g, '\n\n').trim();
    if (!x) { W('範例文', o.line, '全文空白'); return }
    const tm = x.match(/^\[([^\]]+)\]/);
    if (!tm) W('範例文', o.line, '第一行不是「[標籤] 標題」的格式');
    const g = list(o.g).filter(v => { if (groups[v] || pseudo[v]) return true; (badG[v] = badG[v] || []).push(o.line); return false });
    lib.push({t: o.t || (tm ? tm[1] : ''), g, s: /^y$/i.test(o.s) ? 1 : 0, y: o.y || '', n: exBody(x).replace(/\s/g, '').length, x});
  });
  Object.entries(badG).forEach(([v, lines]) => warns.push(`範例文 第 ${lines.slice(0, 5).join('、')}${lines.length > 5 ? ' 等' : ''} 列：品類「${v}」不在品類分頁，已忽略`));

  // 整體檢查
  Object.entries(groups).forEach(([g, v]) => {
    if (!v.objs.length) warns.push(`品類「${g}」沒有任何療程`);
    if (!v.pains.length) warns.push(`品類「${g}」沒有任何困擾，抽到時改用通用困擾`);
  });
  CATS.forEach(c => { if (!scene.some(s => s.cats.includes(c))) warns.push(`分頁「${c}」沒有任何情境，這個分頁暫時不能產生`) });
  if (!tones.length) throw new Error('「語氣」分頁沒有任何啟用的語氣');
  if (!triggers.length) throw new Error('「觸發事件」分頁沒有任何啟用的事件');
  if (lib.length < 2) throw new Error('「範例文」分頁至少要有 2 篇啟用的範例');
  return {data: {groups, tab, tabPain, scene}, tones, triggers, ban, season, lib, warns};
}

function applyPool(res) {
  DATA = res.data;
  TONE_INFO = {}; res.tones.forEach(t => TONE_INFO[t.text] = t);
  P.tone = res.tones.map(t => t.text);
  TRIGGERS = res.triggers; P.trigger = res.triggers.map(t => t.text);
  SEASON_RULES = res.season; SHEET_BAN = res.ban; LIB = res.lib;
  poolWarns = res.warns;
  opts('f-tone', P.tone); opts('f-trigger', P.trigger);
  syncTopic();
}

/* ---------- 讀檔：先讀同一層，讀不到用手動選的檔（存在 IndexedDB） ---------- */
let poolWarns = [], poolSig = '', poolSource = '';
const IDB = {
  open() { return new Promise((ok, no) => { try { const q = indexedDB.open('facelift-prompt', 1); q.onupgradeneeded = () => q.result.createObjectStore('pool'); q.onsuccess = () => ok(q.result); q.onerror = () => no(q.error) } catch (e) { no(e) } }) },
  async get() { try { const db = await this.open(); return await new Promise(ok => { const q = db.transaction('pool').objectStore('pool').get('last'); q.onsuccess = () => ok(q.result || null); q.onerror = () => ok(null) }) } catch (e) { return null } },
  async put(v) { try { const db = await this.open(); await new Promise(ok => { const tx = db.transaction('pool', 'readwrite'); tx.objectStore('pool').put(v, 'last'); tx.oncomplete = ok; tx.onerror = ok }) } catch (e) {} },
};
function fmtDate(ts) { const d = new Date(ts); return `${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}` }
function setStatus(kind, text) { const s = $('pool-status'); s.className = 'status' + (kind ? ' ' + kind : ''); s.textContent = text }
function setMsg(t) { $('pool-msg').textContent = t || '' }

// 回傳 true 表示成功套用
function useBuffer(buf, meta) {
  if (typeof XLSX === 'undefined') { setStatus('warn', '少了 SheetJS'); setMsg('找不到 vendor/xlsx.full.min.js，請確認 vendor 資料夾跟 index.html 放在一起。'); return false }
  let res;
  try { res = buildPool(XLSX.read(new Uint8Array(buf), {type: 'array'})) }
  catch (e) { setStatus('warn', 'Excel 有誤'); setMsg(`${meta.name}：${e.message}。${ready ? '沿用目前的資料。' : '請修正後按「重新讀取」，或選另一個檔。'}`); return false }
  const sig = buf.byteLength + ':' + (meta.mtime || '') + ':' + res.lib.length + ':' + res.data.scene.length;
  const changed = sig !== poolSig; poolSig = sig; poolSource = meta.source;
  if (changed) applyPool(res); else poolWarns = res.warns;
  const label = {server: `已讀取 ${POOL_FILE}`, manual: `已讀取 ${meta.name}`, cache: `使用 ${fmtDate(meta.pickedAt)} 選的檔`}[meta.source];
  setStatus(res.warns.length ? 'warn' : 'ok', label + (res.warns.length ? `（${res.warns.length} 個問題）` : ''));
  $('pool-time').textContent = meta.mtime ? '檔案修改於 ' + fmtDate(meta.mtime) : '';
  const first = !ready; ready = true; $('go').disabled = false;
  drawPool();
  if (first && !touched) generate();
  return true;
}
async function loadPool() {
  setStatus('', '讀取中…'); setMsg('');
  try {
    if (location.protocol === 'file:') throw new Error('file');
    const res = await fetch('./' + encodeURIComponent(POOL_FILE), {cache: 'no-cache'});
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const buf = await res.arrayBuffer();
    const lm = res.headers.get('Last-Modified');
    if (useBuffer(buf, {source: 'server', name: POOL_FILE, mtime: lm ? Date.parse(lm) : null})) return;
    return;
  } catch (e) { /* 直接開檔（file://）或檔案不在，改用快取或手動選檔 */ }
  const fileMode = location.protocol === 'file:';
  const why = fileMode ? '直接開檔時瀏覽器不讓頁面讀取 Excel' : `讀不到同一層的 ${POOL_FILE}`;
  if (ready && poolSource !== 'server') { setStatus('', poolSource === 'cache' ? '沿用上次選的檔' : '沿用手動選的檔'); setMsg(`${why}。Excel 有更新的話，請按「選擇 Excel 檔」重新選。`); return }
  const cached = await IDB.get();
  if (cached && cached.buf && useBuffer(cached.buf, {source: 'cache', name: cached.name, mtime: cached.mtime, pickedAt: cached.pickedAt})) {
    setMsg(`${why}，改用上次在這台瀏覽器選的「${cached.name}」。Excel 有更新的話，請按「選擇 Excel 檔」重新選。`);
    return;
  }
  setStatus('warn', '還沒有參數池');
  setMsg(`${why}。請按「選擇 Excel 檔」，選 ${POOL_FILE}；選過一次之後，這台瀏覽器會記住。`);
  $('pick-box').hidden = false;
  drawPool();
}
async function onPickFile(file) {
  if (!file) return;
  const buf = await file.arrayBuffer();
  if (useBuffer(buf, {source: 'manual', name: file.name, mtime: file.lastModified})) {
    $('pick-box').hidden = true;
    setMsg('這台瀏覽器已記住這個檔，下次直接開頁面也會用它。Excel 有更新時，再按「選擇 Excel 檔」選一次。');
    IDB.put({buf, name: file.name, mtime: file.lastModified, pickedAt: Date.now()});
  }
}

/* ================= 選單 ================= */
function opts(id, items, randLabel, extra) {
  const el = $(id); const keep = el.value; el.innerHTML = '';
  const add = (v, t) => { const x = document.createElement('option'); x.value = v; x.textContent = t; el.appendChild(x) };
  add(R, randLabel || '隨機'); (extra || []).forEach(([v, t]) => add(v, t)); items.forEach(v => add(v, v));
  el.value = [...el.options].some(o => o.value === keep) ? keep : R; mark(el);
}
function mark(el) { el.classList.toggle('rand', el.value === R); el.classList.toggle('fixed', el.value !== R && el.id !== 'f-region'); updateBadge() }
function updateBadge() {
  const n = [...document.querySelectorAll('#adv select')].filter(s => s.value !== R).length;
  const b = $('fixed-badge'); b.hidden = !n; b.textContent = '已指定 ' + n + ' 項';
}
const groupsOf = t => DATA.tab[t] || [];
const scenesOf = t => DATA.scene.filter(s => s.cats.includes(t));
function sceneOk(s, g) { if (!g) return true; if (s.only && !s.only.includes(g)) return false; if (s.noSurg && DATA.groups[g] && DATA.groups[g].surgery) return false; return true }
const isAsk = s => s.type === '求推';

function syncTopic() {
  const gs = groupsOf(tab);
  const gEl = $('f-group');
  opts('f-group', gs, gs.length ? '隨機' : '依情境自動');
  gEl.disabled = !gs.length;
  const g = gEl.value !== R ? gEl.value : null;
  const objPool = g ? DATA.groups[g].objs : [...new Set(gs.flatMap(k => DATA.groups[k].objs))];
  opts('f-obj', objPool, gs.length ? '隨機' : '依情境自動', gs.length ? [['__none', '不指定療程']] : []);
  $('f-obj').disabled = !gs.length;
  const painPool = g ? DATA.groups[g].pains : (gs.length ? [...new Set(gs.flatMap(k => DATA.groups[k].pains))] : (DATA.tabPain[tab] || []));
  opts('f-pain', painPool);
  opts('f-scene', scenesOf(tab).filter(s => sceneOk(s, g)).map(s => s.name));
  opts('f-tag', ['討論', '閒聊', '問題', '詢問', '心得'], '依情境自動');
}
function syncTabs() { CATS.forEach(c => $('tab-' + c).setAttribute('aria-selected', c === tab)) }
function syncMode() {
  $('mode-single').setAttribute('aria-pressed', mode === 'single'); $('mode-mix').setAttribute('aria-pressed', mode === 'mix');
  $('single-box').hidden = mode !== 'single'; $('mix-box').hidden = mode !== 'mix';
  TOPIC_FIELDS.forEach(id => { $(id).closest('.row').style.opacity = mode === 'mix' ? .45 : 1 });
}
function weekCity() {
  const d = new Date(); const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1)); const wk = Math.ceil(((t - y0) / 86400000 + 1) / 7);
  return {city: ROTATE[wk % 4], wk};
}
function syncRegion() {
  const {city, wk} = weekCity();
  opts('f-region', Object.keys(CITY), '本週輪替：' + city, [['__none', '不帶地區']]);
  const sel = $('f-region').value;
  const c = sel === R ? city : sel;
  $('region-now').textContent = sel === '__none' ? '這批不會出現地區。' : `第 ${wk} 週，部分篇數會帶入：${CITY[c].join('、')}`;
}

/* ================= 抽籤 ================= */
function fixed(id, t) {
  if (mode === 'mix' && TOPIC_FIELDS.includes(id)) return null;
  if (TOPIC_FIELDS.includes(id) && t !== tab) return null;
  const v = $(id).value; return v === R ? null : v;
}
function groupOfObj(o, t) { return groupsOf(t).find(k => DATA.groups[k].objs.includes(o)) }
function groupOfPain(p, t) { return groupsOf(t).filter(k => DATA.groups[k].pains.includes(p)) }
function objFits(g, o, p) { const info = DATA.groups[g] && DATA.groups[g].objInfo[o]; return !info || !info.only || info.only.includes(p) }

/* ---------- 範例挑選 ---------- */
const TAGMAP = {'疑問': '問題', '請益': '問題', '分享': '心得'};
const usedEx = new Set();
const FILLER = new Set(['同問', '謝謝分享', '我也想問', '是', '對', '推', '推推', '蹲', '蹲蹲', '卡', '卡個', '卡一下', '可以', '可吧', '沒差', '想知道', '好奇', '我也想知道', '同樓上', '樓上', '幫推', '先推', '哈哈', '哈哈哈', '哈哈哈哈', '推一個', '真的', '對啊', '是喔', '喔喔', '+1', '同感', '我也是', '我也覺得']);
// 提到具體診所或醫師名字的推文（例如「XX診所」「王醫師」），範例清理時刪掉
const DOCTOR_NAME = /(^|[^這那哪每各某好的任])[王李張陳林黃吳劉蔡楊許鄭謝郭洪曾邱廖賴周徐蘇葉莊呂江蕭羅潘簡朱鍾彭游詹胡沈余趙盧梁顏柯孫魏翁戴范宋鄧杜傅侯曹薛丁馬蔣唐卓藍馮姚董紀歐汪湯姜鄒涂巫韓龔嚴袁鐘阮陸倪邵][一-鿿]?(醫師|醫生)/;
// 「診所」前一個字是這些的，視為泛稱（好診所、很多診所、不同診所、皮膚科診所…），不刪
const GENERIC_CLINIC = /[多好有同原科下些間家種個的去到在跟問找推是麼大小新舊別每各這那哪其般鎖美線該此]$/;
function namesClinic(b) {
  if (DOCTOR_NAME.test(b)) return true;
  const m = b.match(/([一-鿿A-Za-z]{2,6})診所/);
  return !!(m && !GENERIC_CLINIC.test(m[1]));
}
function cleanEx(x) {
  let inC = false;
  return x.split('\n').map(l => {
    if (/^(推|噓)\|/.test(l)) inC = true;
    if (!inC || !l.trim()) return l;
    const m = l.match(/^((?:推|噓)\|)?\s*(.*)$/);
    const b = m[2].replace(/\s*\+1\s*/g, '').trim();
    if (FILLER.has(b.replace(/[!！~～.。]+$/, '')) || b.length <= 1) return null;
    if (namesClinic(b)) return null;
    return (m[1] || '') + b;
  }).filter(l => l !== null).map((l, i) => i ? l.replace(/\s*\+1\s*$/, '') : l).join('\n').replace(/\n{3,}/g, '\n\n').trim();
}
const HELP_RE = /幫問|代問|代PO|代po|代Po|代友|幫朋友問|幫.{0,3}問|代.{0,4}發|[\(（]代/;
function exBody(x) { const L = x.split('\n'); const i = L.findIndex((l, k) => k > 0 && /^(推|噓)\|/.test(l)); return L.slice(1, i < 0 ? L.length : i).join('\n').trim() }
function exStyle(b) { if (/[，、！；]/.test(b)) return /[一-鿿] [一-鿿]/.test(b) ? 'fullspace' : 'full'; if (/[一-鿿][,!]/.test(b)) return 'half'; if (/[一-鿿] [一-鿿]/.test(b)) return 'space'; return 'line' }
function exConflict(r, x) {
  const b = exBody(x); let s = 0;
  if (r.noOpenWord && /^(最近|如題)/.test(b)) s -= 8;
  if (/[（(](?!幫問|代)[^）)]{1,30}[）)]/.test(b)) s -= 6;
  if (!r.emo && /XD|@@|QQ|==|><|\^\^/.test(b)) s -= 4;
  if (!r.dots && /\.\./.test(b)) s -= 2;
  const st = exStyle(b); if (st === r.punct) s += 4; else if ((st === 'full' || st === 'fullspace') && (r.punct === 'full' || r.punct === 'fullspace')) s += 2;
  return s;
}
function pickExamples(r, avoid) {
  const tagOf = e => TAGMAP[e.t] || e.t;
  const score = e => {
    let s = 0;
    if (r.group && e.g.includes(r.group)) s += 6; else if (!r.group && r.tab === '醫美診所' && e.s) s += 4;
    if (tagOf(e) === r.tag) s += 3;
    if (Math.abs(e.n - r.len) < 35) s += 2;
    if (e.y === '2026') s += 1;
    if (avoid.has(e.x)) s -= 20;
    if (!inSeason(e.x)) s -= 12;
    const eh = HELP_RE.test(e.x);
    if (eh && !(r.helping && r.who !== '跟姊妹一起考慮')) s -= 40;
    if (eh && r.helping) s += 3;
    if (r.share && tagOf(e) === '心得') s += 6;
    s += exConflict(r, e.x);
    return s + Math.random() * 2.5;
  };
  // 每篇只評一次分再排序（隨機加分固定在這次）
  const ranked = LIB.map(e => [score(e), e]).sort((a, b) => b[0] - a[0]).map(x => x[1]);
  const first = ranked[0];
  const okHelp = e => !HELP_RE.test(e.x) || (r.helping && r.who !== '跟姊妹一起考慮');
  const second = ranked.find(e => e !== first && okHelp(e) && tagOf(e) === r.tag && !(r.group && e.g.includes(r.group)) && !avoid.has(e.x) && exConflict(r, e.x) >= 0) || ranked.find(e => e !== first && okHelp(e));
  const out = [first, second].filter(Boolean).map(e => e.x); out.forEach(x => avoid.add(x)); return out.map(cleanEx);
}

function roll(t, used) {
  const r = {tab: t};
  const fObj = fixed('f-obj', t), fPain = fixed('f-pain', t), fScene = fixed('f-scene', t);
  let fGroup = fixed('f-group', t);
  if (!fGroup && fObj && fObj !== '__none') fGroup = groupOfObj(fObj, t);
  const painGroups = fPain ? groupOfPain(fPain, t) : null;
  const gs = groupsOf(t);

  // 1 情境（只抽有至少一個品類能配的）
  let sPool = scenesOf(t).filter(s => sceneOk(s, fGroup));
  if (painGroups && painGroups.length) { const f = sPool.filter(s => painGroups.some(g => sceneOk(s, g))); if (f.length) sPool = f }
  if (gs.length) { const f = sPool.filter(s => gs.some(g => sceneOk(s, g))); if (f.length) sPool = f }
  if (!sPool.length) sPool = scenesOf(t);
  sPool = seasonPool(sPool, s => s.name + s.desc);
  if (t !== '醫美診所' && !fScene && Math.random() < .85) { const f = sPool.filter(s => !isAsk(s)); if (f.length) sPool = f }
  if (fScene) r.scene = DATA.scene.find(s => s.name === fScene);
  else {
    // 心得情境合起來只算 SHARE_SLOTS 個情境的份量；抽中心得時先依結論比例決定結論
    let pool = sPool;
    const shares = sPool.filter(s => s.share), plain = sPool.filter(s => !s.share);
    if (shares.length && plain.length) { const k = Math.min(SHARE_SLOTS, shares.length); pool = Math.random() < k / (plain.length + k) ? shares : plain }
    if (pool.length && pool.every(s => s.share)) {
      const moods = [...new Set(pool.map(s => s.mood || ''))];
      const m = wpick(moods, x => MOOD_W[x] || 20);
      pool = pool.filter(s => (s.mood || '') === m);
    }
    const nm = pickLeast(pool.map(s => s.name), 'scene', t, used); r.scene = pool.find(s => s.name === nm);
  }
  used.add('scene', t, r.scene.name, 5);

  // 2 品類（依比重，避開最近用過）
  const weightOf = g => { const w = DATA.groups[g].weight; return w == null ? 2 : w };
  const pickGroup = cand => { const gmin = Math.min(...cand.map(g => used.n('group', t, g))); const fresh = cand.filter(g => used.n('group', t, g) <= gmin + 1); const p = fresh.length ? fresh : cand; return p.some(g => weightOf(g) > 0) ? wpick(p, weightOf) : pick(p) };
  if (fGroup) r.group = fGroup;
  else if (gs.length) {
    let cand = gs.filter(g => sceneOk(r.scene, g));
    if (painGroups && painGroups.length) cand = cand.filter(g => painGroups.includes(g));
    if (!cand.length) cand = gs;
    r.group = pickGroup(cand);
  } else if (Math.random() < (REAL_GROUP_RATE[t] || 0)) {
    // 閒聊、診所分頁：有時從情境列出的品類抽一個，讓指令指定療程
    const cand = (r.scene.only || Object.keys(DATA.groups)).filter(g => DATA.groups[g] && sceneOk(r.scene, g));
    r.group = cand.length ? pickGroup(cand) : null;
  } else r.group = null;
  if (r.group) used.add('group', t, r.group, 5);
  const G = r.group ? DATA.groups[r.group] : null;

  // 3 困擾（先決定困擾）
  const fallbackPain = DATA.tabPain[t] && DATA.tabPain[t].length ? DATA.tabPain[t] : ['容貌焦慮'];
  if (fPain) r.pain = fPain;
  else if (t === '醫美診所') r.pain = pickLeast(seasonPool(fallbackPain), 'pain', t, used);
  else if (G) {
    let pp = G.pains.length ? G.pains : fallbackPain;
    if (fObj && fObj !== '__none' && G.objInfo[fObj] && G.objInfo[fObj].only) { const f = pp.filter(x => objFits(r.group, fObj, x)); if (f.length) pp = f }
    r.pain = pickLeast(seasonPool(pp), 'pain', t, used);
  } else r.pain = pickLeast(seasonPool(fallbackPain), 'pain', t, used);
  used.add('pain', t, r.pain, 5);
  // 4 療程：只挑適合這個困擾的
  if (fObj === '__none') r.obj = null;
  else if (fObj) r.obj = fObj;
  else if (G && G.objs.length) { const fit = G.objs.filter(o => objFits(r.group, o, r.pain)); r.obj = !fit.length ? null : pickLeast(fit, 'obj', t, used) }
  else r.obj = null;
  if (r.obj) used.add('obj', t, r.obj, 5);
  r.nick = r.obj && G && G.objInfo[r.obj] ? G.objInfo[r.obj].nick : [];

  // 5 發文者
  const femaleNeeded = FEMALE_RE.test(r.pain) || FEMALE_RE.test(r.scene.name) || FEMALE_RE.test(r.scene.desc) || /私密|胸/.test(r.group || '');
  const gv = $('f-gender').value;
  r.gender = gv !== R ? gv : (femaleNeeded && Math.random() < .85 ? '女' : pick(P.gender));
  const wv = $('f-who').value;
  r.who = wv !== R ? wv : (r.gender === '男' && femaleNeeded ? pick(['幫伴侶問', '幫媽媽問']) : (Math.random() < .07 ? pick(P.who.filter(x => x !== '自己')) : '自己'));
  r.conflict = r.gender === '男' && femaleNeeded && r.who === '自己';
  r.exp = $('f-exp').value !== R ? $('f-exp').value : pick(P.exp);
  r.budget = $('f-budget').value !== R ? $('f-budget').value : (G && G.surgery ? pick(P.budget.filter(b => b !== '一萬以內')) : pick(P.budget));
  let trigPool = TRIGGERS.filter(x => !x.m.length || x.m.includes(MONTH)).map(x => x.text);
  if (!trigPool.length) trigPool = P.trigger;
  trigPool = seasonPool(trigPool);
  r.trigger = $('f-trigger').value !== R ? $('f-trigger').value : pickLeast(trigPool, 'trigger', t, used); used.add('trigger', t, r.trigger, 5);

  // 6 標籤與心得判定
  const tv = fixed('f-tag', t);
  r.tag = tv || (r.scene.share ? '心得' : isAsk(r.scene) ? pick(['問題', '詢問']) : t === '閒聊趣味' ? pick(['閒聊', '閒聊', '閒聊', '討論']) : wpick(['閒聊', '討論', '問題', '詢問'], x => ({'閒聊': 30, '討論': 35, '問題': 18, '詢問': 17})[x]));
  r.share = !!r.scene.share || r.tag === '心得';

  // 7 地區（只在診所類或求推類出現）
  const rv = $('f-region').value;
  const city = rv === R ? weekCity().city : (rv === '__none' ? null : rv);
  if (isAsk(r.scene) && /地區/.test(r.scene.name)) r.region = pick(CITY[city || weekCity().city]);
  else if (!city || r.share) r.region = null;
  else r.region = (t === '醫美診所' || isAsk(r.scene)) && Math.random() < .5 ? pick(CITY[city]) : null;

  // 8 寫法（依真實板友文）
  r.fear = /恐懼/.test(r.scene.name) || /^(擔心|怕)/.test(r.scene.dir);
  const treat = !!(r.obj || G);
  // 語氣：依心得、類別排除不搭的
  let tonePool = P.tone;
  if (r.share) tonePool = tonePool.filter(x => !(TONE_INFO[x] && TONE_INFO[x].notShare));
  if (t === '閒聊趣味' || t === '醫美診所' || !treat) tonePool = tonePool.filter(x => !(TONE_INFO[x] && TONE_INFO[x].notCC));
  if (!tonePool.length) tonePool = P.tone;
  r.tone = $('f-tone').value !== R ? $('f-tone').value : pickLeast(tonePool, 'tone', t, used);
  const novice = /沒做過|只做過|只在美容院/.test(r.exp), veteran = /老手|固定/.test(r.exp);
  if ($('f-tone').value === R && ((novice && /老手/.test(r.tone)) || (veteran && /新手/.test(r.tone)))) { const f = tonePool.filter(x => !/老手|新手/.test(x)); r.tone = pick(f.length ? f : tonePool) }
  used.add('tone', t, r.tone, 5);
  r.open = ''; r.layout = '';
  // 字數：一般 75% 為 35–90、25% 為 90–130；心得 70% 為 150–260、30% 為 260–360
  const lv = $('f-len').value; let lo, hi;
  if (lv !== R) { const m = lv.match(/(\d+)–(\d+)/); lo = +m[1]; hi = +m[2] }
  else if (r.share) { [lo, hi] = Math.random() < .7 ? [150, 260] : [260, 360] }
  else { [lo, hi] = Math.random() < .75 ? [35, 90] : [90, 130] }
  r.len = Math.round((lo + Math.random() * (hi - lo)) / 5) * 5;
  r.lines = Math.max(2, Math.min(r.share ? 18 : 9, Math.round(r.len / (r.share ? 18 : 14))));
  r.blank = r.share ? Math.random() < .75 : Math.random() < .5;
  r.askEnd = !r.share && (isAsk(r.scene) || Math.random() < (r.tag === '閒聊' ? .4 : .65));
  r.push = t !== '空包彈';
  // 心得內容
  if (r.share) {
    r.mood = r.scene.mood || wpick(['滿意', '普通', '後悔'], k => ({滿意: 45, 普通: 35, 後悔: 20})[k]);
    const items = [['做了什麼、在什麼時候做的', 100], ['花了多少錢，或打了幾發、幾劑、幾cc', 80], ['效果跟做之前比起來的差別', 65], ['一個不滿意的地方或但書', 65], ['做的時候的痛感', 58], ['恢復期的狀況', 32], ['之後會不會再做', 20]];
    r.shareItems = [items[0][0], ...items.slice(1).filter(x => Math.random() * 100 < x[1]).map(x => x[0])].slice(0, 5);
  }
  // 除了主題，最多再帶一件自己的事；起因只在跟方向有關時才帶
  const sc = r.scene.name + r.scene.desc;
  const TRIG_OK = /重要場合|容貌焦慮|上妝|起效|心得|術後紀錄|分享|預算|CP/.test(sc) || r.share;
  const extras = [];
  if (r.region) r.extra = {k: 'region'};
  else if (r.share) r.extra = null;
  else {
    if (/預算|CP|價格|發數|花費|投資|便宜|貴/.test(sc + r.pain)) extras.push({k: 'budget', v: r.budget});
    if (/失敗|重修|跨越|術後|複合|頻率|維持|狀態確認|起效/.test(sc)) extras.push({k: 'exp', v: r.exp});
    if ($('f-trigger').value !== R || TRIG_OK) extras.push({k: 'trigger', v: r.trigger});
    r.extra = extras.length && Math.random() < .5 ? pick(extras) : null;
  }
  r.helping = r.who !== '自己';
  if (r.helping && r.extra && (r.extra.k === 'exp' || r.extra.k === 'budget')) r.extra = null;
  r.priceInBody = r.share || (r.extra && r.extra.k === 'budget') || /價格|預算|CP|發數|花費|便宜|貴/.test(r.scene.name + plainDesc(r.scene.desc) + r.pain);
  // 推文配比：固定 10 則，噓 0–2、中立 0–2
  r.cTotal = 10;
  if ($('f-pushrand').checked) { r.cBoo = pick([0, 0, 1, 1, 1, 2]); r.cNeu = pick([0, 1, 1, 2]) }
  else { r.cBoo = Math.min(4, Math.max(0, +$('f-b').value || 0)); r.cNeu = Math.min(4, Math.max(0, +$('f-n').value || 0)) }
  r.cPush = 10 - r.cBoo - r.cNeu; { const x = Math.random(); r.wave1 = x < .4 ? 5 : x < .8 ? 6 : x < .92 ? 7 : 4 }
  // 標題風格
  const vsScene = /AB|比較|選擇|vs|對決/.test(r.scene.name + r.scene.dir);
  if (r.share) r.titleStyle = r.mood === '比較' ? 'compare' : wpick(['result', 'feel', 'what'], k => ({result: 45, feel: 35, what: 20})[k]);
  else if (isAsk(r.scene)) r.titleStyle = 'ask';
  else r.titleStyle = wpick(['ask', 'state', 'feel', 'name', 'vs'], k => ({ask: r.askEnd ? 45 : 20, state: 30, feel: r.tag === '閒聊' ? 25 : 12, name: r.obj ? 10 : 0, vs: vsScene ? 25 : 0})[k]);
  r.title = r.titleStyle; used.add('title', t, r.title, 5);
  // 標題長度：依過往標題分布（不含標籤）
  r.tLen = r.titleStyle === 'name' ? [3, 8] : wpick([[5, 8], [9, 12], [13, 16], [17, 19]], x => ({5: 11, 9: 33, 13: 36, 17: 20})[x[0]]);
  r.avoidPhr = sample(TEAM_PHRASES, 3);
  r.noOpenWord = Math.random() < .7;
  r.thanks = r.askEnd && Math.random() < .15;
  r.paren = false;
  r.emo = Math.random() < .35 ? pick(['XD', '@@', 'QQ', '==', '><']) : null;
  r.dots = Math.random() < .3;
  r.punct = wpick(['space', 'line', 'full', 'fullspace', 'half'], k => ({space: 33, line: 26, full: 24, fullspace: 12, half: 5})[k]);
  if (r.share && r.punct === 'line' && Math.random() < .6) r.punct = 'full';
  r.qmark = r.punct === 'half' ? '?' : (r.punct === 'full' || r.punct === 'fullspace') ? (Math.random() < .5 ? '？' : '?') : '?';
  r.titleRepeat = Math.random() < .4;
  { const x = Math.random(); r.pMax = x < .5 ? 0 : x < .85 ? 1 : 2 }
  r.cAsk = Math.random() < (r.share ? .2 : .3);
  // 推文逐則安排：比例依過往推文（多數是附和、經驗、小建議；唱反調很少）
  {
    const ST = [
      {k: 'pos', w: 18, d: r.share ? '講自己做過同一個療程，感覺跟發文者差不多或不太一樣' : '講自己或身邊的人做過的一個具體感受，不一定要跟發文者擔心的事有關'},
      {k: 'tip', w: 15, d: '補一個跟主題有關、正確的小知識或注意事項'},
      {k: 'short', w: 12, d: '很短地講一個自己的結論'},
      {k: 'follow', w: 10, d: '接前一則推文的話，補一句自己的看法'},
      {k: 'want', w: 8, d: r.share ? '說自己也想做，講出看完這篇在意的點' : '說自己也在考慮，講出自己在意的點'},
      {k: 'heard', w: 7, d: '講在網路上或聽別人說過的說法'},
      {k: 'side', w: 5, d: treat ? '聊同一個療程的其他事，例如恢復期、多久做一次' : '聊跟主文同一件事的另一面'},
      {k: 'alt', w: treat ? 7 : 0, d: '提一個也能處理這個困擾、而且合理的其他療程'},
      {k: 'price', w: r.priceInBody ? 5 : 0, d: '講一個自己的價格、堂數或發數，跟主文比較'},
      {k: 'joke', w: 4, d: '開個小玩笑或吐槽一下'},
      {k: 'neg', w: 4, d: '講自己或身邊的人做過，有一個不滿意的地方'},
      {k: 'contra', w: 3, d: '不同意主文或前面某則推文的看法，講出理由'},
      {k: 'doubt', w: treat ? 3 : 0, d: '懷疑這個療程或某個說法到底有沒有用'},
    ].filter(s => s.w > 0);
    const BOO = ['neg', 'contra', 'doubt', 'joke'], NEGK = ['neg', 'contra', 'doubt'];
    const slots = [...Array(r.cBoo).fill('噓|'), ...Array(r.cNeu).fill('→|'), ...Array(r.cPush).fill('推|')];
    for (let i = slots.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [slots[i], slots[j]] = [slots[j], slots[i]] }
    const cnt = {}; let prev = '', negPush = 0; const askAt = r.cAsk ? rint(2, 9) : -1;
    r.cPlan = slots.map((pf, i) => {
      if (i === askAt && pf !== '噓|') { prev = 'ask'; return {pf, d: r.share ? '問發文者一個做的細節，用問句' : '問一個具體細節，用問句'} }
      const cap = k => ['side', 'joke', 'follow', 'short', 'heard', 'price', 'tip'].includes(k) ? 1 : 2;
      let pool = ST.filter(s => (cnt[s.k] || 0) < cap(s.k) && s.k !== prev && (i > 0 || s.k !== 'follow'));
      if (pf === '噓|') { pool = pool.filter(s => BOO.includes(s.k)); if (!pool.length) pool = ST.filter(s => BOO.includes(s.k)) }
      else { if (negPush >= 1) pool = pool.filter(s => !NEGK.includes(s.k)) }
      if (pf === '→|') pool = pool.filter(s => ['tip', 'side', 'heard', 'price', 'short'].includes(s.k));
      if (!pool.length) pool = ST.filter(s => !NEGK.includes(s.k) && s.k !== 'follow');
      const s = wpick(pool, x => x.w); cnt[s.k] = (cnt[s.k] || 0) + 1; prev = s.k; if (pf !== '噓|' && NEGK.includes(s.k)) negPush++; return {pf, d: s.d};
    });
  }
  // 範例：同品類、同標籤、字數相近，且跟這篇的寫法規則不衝突
  r.examples = pickExamples(r, usedEx);
  return r;
}

/* ================= 組指令 ================= */
function clinicFocus(r) { return r.tab === '醫美診所' || r.scene.type === '求推' || r.scene.type === '診所主軸' }
function stripClinicAsk(d) {
  const parts = d.split(/(?<=[，。；？?！!])/);
  const bad = /(求推|推薦|名單|找).{0,20}(診所|醫師|醫生)|(診所|醫師|醫生).{0,8}(推薦|求推)|安心診所|挑選醫生|挑醫生|的醫師|的診所/;
  const kept = parts.filter(s => !bad.test(s));
  return (kept.join('').replace(/[，、；]$/, '。') || d);
}
const VIVID = /阿凡達|毛毛蟲|石頭奶|釋迦|充氣娃娃|饅化臉|烙鐵|陰影面積|求生欲|社死|打水漂|心酸|血淚|推坑|生火|滅火|眉毛壓眼|皮笑肉不笑|菜市場|呼吸困難|滿心期待|驚豔|連根拔起/;
function stripVivid(d) { return d.split(/(?<=[、，。；])/).filter(s => !VIVID.test(s)).join('').replace(/[、，]$/, '。').replace(/[、，]([。；])/g, '$1') || d }
function plainDesc(d) {
  d = stripVivid(d);
  return d.replace(/⚠️/g, '').replace(/[（(][^）)]{0,80}[）)]/g, '').replace(/(例如|如：|舉例如|舉例可描述|可直接舉例)[^。；？?]*[。；]?/g, '').replace(/【([^】]*)】/g, '$1').replace(/「([^」]*)」/g, '$1').replace(/\s{2,}/g, ' ').trim();
}
function promptText(r) {
  const ban = [...new Set([...SHEET_BAN, ...$('f-ban').value.split(/[,，、\n]/).map(s => s.trim()).filter(Boolean)])];
  const avoid = $('f-avoid').value.split('\n').map(s => s.trim()).filter(Boolean);
  const kw = $('f-kw').value.split(/[,，、\n]/).map(s => s.trim()).filter(Boolean);
  const note = $('f-note').value.trim();
  const push = r.push && r.cTotal > 0;
  const G = r.group;
  const L = [];
  L.push(`你是台灣 PTT 醫美板 facelift 的板友，照下面的主題發一篇文${push ? '，再寫底下的推文' : ''}。`);
  L.push('整篇放在一個 ``` 程式碼區塊裡輸出，保留所有換行和空行，區塊外不要加任何說明。');
  L.push('這篇和這個對話裡之前寫過的任何文章都無關：不要沿用前面文章的開頭、標記、格式、句型或推文寫法。');
  L.push('');
  L.push('【主題】');
  if (G) L.push(`品類：${G}`);
  if (r.obj) { const nk = (r.nick || []).length ? `；${r.obj}板上也常叫${r.nick.join('、')}` : ''; L.push(`療程：${r.obj}${nk}`) }
  else if (r.tab === '醫美診所') L.push('療程：不提具體療程，聊診所本身');
  else if (G) L.push(`療程：用「${G.replace(/^.*-/, '')}」這個通稱就好，不要自己編療程或儀器名稱`);
  const showPain = !(r.tab === '閒聊趣味' && !G) && !(r.fear && !r.share);
  if (r.share) L.push(`做這個療程的原因：${r.pain}，用一般人的講法帶過`);
  else if (showPain) L.push(`困擾：${r.pain}。用一般人的講法描述；它是這個方向裡的具體例子，不是另一個題目`);
  if (r.share) {
    const M = {滿意: '整體滿意，但不是全部都好', 普通: '效果普通，有好有壞，說不上值不值得', 後悔: '不太滿意，有點後悔，但不是在罵', 紀錄: '照時間順序記錄做完到現在的狀況，不下定論', 比較: '自己做過兩種，比較兩者的差別'};
    L.push(`這是一篇心得文。${r.mood === '比較' || r.mood === '紀錄' ? '寫法' : '心得結論'}：${M[r.mood]}`);
    L.push(`要寫到這幾點：${r.shareItems.join('；')}`);
  } else {
    L.push(`方向：${r.scene.dir}。說明：${plainDesc(clinicFocus(r) ? r.scene.desc : stripClinicAsk(r.scene.desc))}`);
    L.push('說明只是參考，不要照抄用詞，也不要寫得這麼戲劇化；整篇只講這個方向，方向名稱不要直接寫進標題');
    if (r.fear) L.push('擔心的事用平常的講法講一句就好，不要描述想像出來的畫面，也不要自己造新詞');
  }
  if (!clinicFocus(r)) L.push('重點不是找診所：不要問或請大家推薦哪間診所、哪位醫生');
  if (r.extra) {
    if (r.extra.k === 'region') L.push(`地區：${r.region}，帶一句想找這一區的就好`);
    else if (r.extra.k === 'budget') L.push(`可以帶到一句：預算大概${r.extra.v}`);
    else if (r.extra.k === 'exp') L.push(`可以帶到一句自己的狀況：${r.extra.v}`);
    else if (r.extra.k === 'trigger') L.push(`可以帶到一句起因：${r.extra.v}`);
  }
  if (kw.length) L.push(`必提關鍵字：${kw.join('、')}，自然放進標題或正文`);
  if (note) L.push(`補充要求：${note}`);
  L.push('');
  L.push('【寫法】照真實板友發文的樣子寫');
  const TS = {ask: '標題用問句直接問想知道的事', state: '標題用陳述句講自己的狀況或發現，不用問號', feel: '標題用一句感想或吐槽', name: '標題很短，只寫療程名稱或關鍵字', vs: '標題寫成兩個選項的比較', result: '標題直接講做完的結果', what: '標題寫做了什麼跟過了多久', compare: '標題寫兩個療程的比較'};
  L.push(`標籤用 [${r.tag}]，${TS[r.titleStyle]}；標籤後面的標題 ${r.tLen[0]} 到 ${r.tLen[1]} 個字，英數字算半個字${r.titleRepeat ? '' : '；第一行不要重述標題'}`);
  L.push(`正文約 ${r.len} 字、${r.lines} 行左右，每行 8 到 25 字，一行一句話${r.blank ? (r.share ? '，分成兩三段，段落之間空一行' : '，中間可以空一行') : '，行跟行之間不空行'}；前後句要接得起來，像一個人在講同一件事`);
  if (r.share) L.push('心得文要講清楚自己做了什麼、過程跟結果，數字和感受要具體；可以有一句一開始的動機，不要寫成問題文，也不要向大家發問');
  else L.push(`第一行就進入主題，不要鋪陳背景故事${r.noOpenWord ? '，也不要用「最近」「如題」開頭' : ''}；除了上面「可以帶到一句」的內容，不要交代年齡、職業、預算、經歷或發生了什麼事`);
  if (!r.share) L.push(r.askEnd ? `最後一行把想問的事用自己的話丟給大家，不要用「你們覺得呢」「大家怎麼看」這種套話${r.thanks ? '，最後可以補一句謝謝大家' : ''}` : '結尾不用硬問問題，用自己的一句感想或碎念收掉就好');
  if (r.helping) L.push(r.who === '跟姊妹一起考慮' ? '這篇是跟姊妹一起考慮：可以提到是兩個人一起在看，但重點放在要問的事' : `這篇是${r.who}：第一行寫「(幫問)」，整篇用幫別人問的角度，講的是對方的狀況`);
  else L.push('這篇是發文者自己的事，不要寫「(幫問)」「代問」「代PO」');
  L.push(`語氣：${r.tone}；這只影響講話方式，不要把語氣寫成文章內容`);
  L.push(`用詞平淡口語，療程用台灣醫美板常見的中文俗稱或簡稱，不用英文原廠全名、學名或中國叫法；不要比喻、誇飾和作文用語；少用${r.avoidPhr.join('、')}`);
  L.push(/替代|比較|選擇|AB|vs|CP|複合|侵入|一起做/.test(r.scene.name + r.scene.dir) || r.mood === '比較' ? '可以提到用來比較的其他療程，但只講跟這個困擾有關的' : '不要提到主題以外的療程名稱');
  const PUNCT = {space: '這篇句子裡不加逗號，要停頓就空一個半形空白', line: '這篇每行都是一個短句，行內不加逗號也不空格，靠換行斷句', full: '這篇用全形「，」「、」斷句，可以有一處「！」', fullspace: '這篇混著用，有些句子用全形「，」斷，有些用半形空白斷', half: '這篇用半形「,」斷句，可以有一處半形「!」'};
  L.push(`標點：${PUNCT[r.punct]}；句尾不用句號；問號用「${r.qmark}」；${r.dots ? '可以有一處 .. 或 ...' : '不要用 .. 或 ...'}；${r.emo ? `可以用一次「${r.emo}」也可以不用` : '不要用 XD、@@、QQ 這類表情'}；${r.helping ? '除了(幫問)以外不要用括號' : '不要用括號'}`);
  { const off = SEASON_RULES.filter(x => !x.m.includes(MONTH)).map(x => x.label); L.push(`現在是 ${MONTH} 月，${SEASON_NAME(MONTH)}，開題和推文都不要出現不符時間的事${off.length ? '，例如' + sample(off, Math.min(4, off.length)).join('、') : ''}`) }
  L.push('療程原理和效果要符合一般醫美常識，不要把不同品類的效果混在一起');
  L.push('PTT 用繁體 Big5 編碼：標題、正文、推文都不要用 emoji、簡體字或特殊符號，會變成問號');
  if (ban.length) L.push(`禁用詞：${ban.join('、')}`);
  if (r.conflict) L.push('困擾偏女性，請寫成幫伴侶或家人問');
  if (avoid.length) { L.push(''); L.push('【避免重複】最近寫過這些標題，這篇的切角和標題不能跟它們相似：'); avoid.slice(-30).forEach(a => L.push(a)) }
  L.push('');
  if (push) {
    const w2 = r.cTotal - r.wave1;
    L.push('【推文】');
    L.push(`剛好 10 則：正文後空一行寫第一批 ${r.wave1} 則，再空一行寫第二批 ${w2} 則，第二批可以比較有資訊；每則獨立一行，不要編號或粗體，寫完自己數一次`);
    L.push(`照下面的安排寫：「推|」「噓|」照寫${r.cBoo ? '，噓文是不認同發文者的想法或對這個做法有負面看法，語氣可以直但不要人身攻擊' : ''}；「→|」是中立推文，輸出時那一行前面不加任何符號；冒號後面的說明不要寫進推文`);
    r.cPlan.forEach((c, i) => { if (i === r.wave1) L.push('（這裡空一行，接第二批）'); L.push(`第${i + 1}則 ${c.pf}：${c.d}`) });
    const rules = [
      '10 則講的事都要不一樣；直接回應發文者擔心或想問的那件事最多兩則，其他照安排寫',
      `每則推文就是一行，不要把一則拆成兩行寫；沒有加「推|」「噓|」的行整串剛好 ${r.cNeu} 行`,
      '推文是一般網友隨手回，不要像診所衛教或業配；講原理的推文整串最多一則，不要用「專業評估」「建議諮詢」這種說法',
      '講療程效果、其他療程或注意事項要符合一般醫美常識，不確定就講自己的感受；不要提具體診所或醫師名字',
      `${r.cAsk ? '除了安排好的那一則，其他' : ''}推文不要用問句；不要用「原PO」稱呼發文者；不要用括號；不要寫「+1」「同問」「推推」「卡」「蹲」「想知道」「謝謝分享」「我也是」這種看不出在回什麼的推文`,
      `每則 30 字以內，${pick(['大部分 8 到 22 字', '長短交錯，有幾則只有幾個字', '多數一行內講完'])}；像隨手打的，句中停頓用空格，幾乎不用而且、但是、不過、所以`,
      r.pMax === 0 ? '句尾直接結束，整串都不要加啦、喔、欸、耶、吧、XD 這類語助詞' : `大部分句尾直接結束；整串最多 ${r.pMax} 則用語助詞，「啦」只用在把話講淡或安撫人，「吧」只用在猜測`,
      '「我朋友」「我覺得」「醫生技術」「看醫生」「真的」「其實」整串各最多一次',
      '推文只根據主文出現過的資訊回應，不要自己補主文沒提到的價格或恢復期',
    ];
    L.push(...sample(rules, rules.length));
  } else L.push('【推文】不要寫推文，正文寫完就結束。');
  L.push('');
  L.push('【範例】以下是過往實際發過的文，只學長度、斷行、語氣和推文的感覺，不要抄內容、標題或句子：');
  r.examples.forEach((x, i) => { L.push(''); L.push(`--- 範例 ${i + 1} ---`); L.push(x) });
  return L.filter((x, i, a) => x !== '' || a[i - 1] !== '').join('\n');
}

/* ================= 卡片 ================= */
let mode = store.get('mode') === 'mix' ? 'mix' : 'single';
let tab = CATS.includes(store.get('cat')) ? store.get('cat') : '電音波';
let mix = {...PRESETS.an};
try { const m = JSON.parse(store.get('mix') || 'null'); if (m) mix = {...mix, ...m} } catch (e) {}
let cards = []; // {id,uid,tab,r,text,pinned,copiedAt}
let RECS = []; try { RECS = JSON.parse(store.get('hist') || '[]') } catch (e) { RECS = [] }
if (!Array.isArray(RECS)) RECS = [];
let touched = false;
let windowDays = [7, 14, 30].includes(+store.get('win')) ? +store.get('win') : 14;
function recentRecs() { const since = Date.now() - windowDays * 864e5; return RECS.filter(x => x.ts >= since) }
let seq = 0;

function cardRec(c) { const r = c.r; return {tt: r.tone + '/' + r.title, tab: c.tab, scene: r.scene.name, group: r.group || '', obj: r.obj || '', pain: r.pain, tone: r.tone, title: r.title, open: r.open, trigger: r.trigger, layout: r.layout, tag: r.tag} }
function newUsed(except) {
  const U = mkU(); const now = Date.now();
  recentRecs().forEach(x => { const age = (now - x.ts) / 864e5; const y = {...x, tt: x.tone + '/' + x.title}; REC_FIELDS.forEach(f => U.add(f, x.tab, y[f], age <= 3 ? 2 : 1)) });
  cards.forEach(c => { if (c === except) return; const x = cardRec(c); REC_FIELDS.forEach(f => U.add(f, x.tab, x[f], 5)) });
  return U;
}
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
function make(t, used) { const r = roll(t, used); return {id: ++seq, uid: uid(), tab: t, r, text: promptText(r), pinned: false, fresh: true} }
function plan() {
  if (mode === 'single') { const n = Math.max(1, Math.min(10, +$('f-count').value || 1)); return Array(n).fill(tab) }
  const out = []; CATS.forEach(c => { for (let i = 0; i < (+mix[c] || 0); i++) out.push(c) }); return out;
}
const todayKey = () => { const d = new Date(); return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate() };
function saveCards() {
  const data = JSON.stringify({day: todayKey(), cards: cards.map(c => ({uid: c.uid, tab: c.tab, r: c.r, text: c.text, pinned: c.pinned, copiedAt: c.copiedAt || null}))});
  if (!store.set('cards', data)) $('cards-msg').textContent = '瀏覽器儲存空間滿了，卡片沒有存起來，重新整理會不見。可以到複製紀錄按「歸零」清掉舊紀錄。';
  else $('cards-msg').textContent = '';
}
function restoreCards() {
  try {
    const s = JSON.parse(store.get('cards') || 'null');
    if (!s || s.day !== todayKey() || !Array.isArray(s.cards) || !s.cards.length) return false;
    cards = s.cards.map(c => ({...c, id: ++seq, fresh: false})); touched = true; draw(); return true;
  } catch (e) { return false }
}
function fmtTime(ts) { const d = new Date(ts); return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0') }
function drawProgress() {
  const el = $('progress'); if (!cards.length) { el.hidden = true; return }
  const done = cards.filter(c => c.copiedAt).length, total = cards.length, left = total - done;
  const start = new Date(); start.setHours(0, 0, 0, 0);
  const mine = RECS.filter(x => x.ts >= start.getTime()).length;
  let html = `這批 <b>${done} / ${total}</b> 篇已複製` + (left ? `<span class="left">還有 ${left} 篇沒複製</span>` : '<span class="pc full">這批都複製完了</span>');
  if (mode === 'mix') { CATS.forEach(c => { const tot = cards.filter(x => x.tab === c).length; if (!tot) return; const d = cards.filter(x => x.tab === c && x.copiedAt).length; html += `<span class="pc${d === tot ? ' full' : ''}">${esc(SHORT[c])} ${d}/${tot}</span>` }) }
  html += `<span class="hint">今天這台瀏覽器一共複製了 ${mine} 篇${$('f-test').checked ? '（測試模式不計）' : ''}</span>`;
  el.innerHTML = html; el.hidden = false;
}
function generate() {
  if (!ready) return;
  touched = true;
  usedEx.clear();
  const kept = cards.filter(c => c.pinned || c.copiedAt);
  cards = [...kept];
  const want = plan();
  // 已保留、已複製的卡片抵掉對應類別的名額
  kept.forEach(p => { const i = want.indexOf(p.tab); if (i > -1) want.splice(i, 1) });
  const skipped = [...new Set(want.filter(t => !scenesOf(t).length))];
  const used = newUsed();
  want.filter(t => scenesOf(t).length).forEach(t => cards.push(make(t, used)));
  $('cards-note').textContent = skipped.length ? `「${skipped.join('、')}」在 Excel 裡沒有可用的情境，這批沒有產生這些分頁。` : '';
  draw(); saveCards();
}
function regenOne(card) {
  if (!ready) return;
  touched = true;
  // 已複製的卡片重抽＝這篇不用了，從複製紀錄移除，不再影響防撞題
  if (card.copiedAt && RECS.some(x => x.id === card.uid)) {
    RECS = RECS.filter(x => x.id !== card.uid); saveRecs(); drawHist(); drawPool();
  }
  const used = newUsed(card);
  const nc = make(card.tab, used);
  cards[cards.indexOf(card)] = nc; draw(); saveCards();
}
function highlight(t) { return esc(t).replace(/^(【[^】]+】)/gm, '<span class="k">$1</span>').replace(/^(推\||噓\|)/gm, '<span class="d">$1</span>') }
function draw() {
  const box = $('cards'); box.innerHTML = '';
  if (!cards.length) { box.innerHTML = `<p class="note">${ready ? '今日配比全部是 0，請至少填一篇。' : '參數池讀進來之後就會產生指令。'}</p>`; drawProgress(); return }
  cards.forEach((c, i) => {
    const r = c.r;
    const el = document.createElement('article'); el.className = 'card' + (c.pinned ? ' pinned' : '') + (c.copiedAt ? ' done' : '') + (c.fresh ? ' new' : ''); c.fresh = false;
    const chips = [['cat', SHORT[c.tab]], ['tag', '[' + r.tag + ']'], ['', r.scene.name], ['', r.obj || r.group || '—'], ['', r.pain], ['', r.gender + (r.who === '自己' ? '' : ' · ' + r.who)], ['', r.len + ' 字']];
    if (r.region) chips.push(['', r.region]);
    if (c.copiedAt) chips.unshift(['donechip', '已複製 ' + fmtTime(c.copiedAt)]);
    const head = document.createElement('div'); head.className = 'card-head';
    head.innerHTML = '<span class="no">#' + String(i + 1).padStart(2, '0') + '</span>' + chips.map(([k, v]) => '<span class="chip ' + k + '">' + esc(v) + '</span>').join('');
    const tools = document.createElement('div'); tools.className = 'card-tools';
    const msg = document.createElement('span'); msg.className = 'copied';
    const bPin = document.createElement('button'); bPin.type = 'button'; bPin.className = 'ghost small'; bPin.textContent = c.pinned ? '已保留' : '保留'; bPin.setAttribute('aria-pressed', c.pinned);
    bPin.onclick = () => { touched = true; c.pinned = !c.pinned; draw(); saveCards() };
    const bRe = document.createElement('button'); bRe.type = 'button'; bRe.className = 'ghost small'; bRe.textContent = '重抽'; bRe.disabled = c.pinned || !ready;
    if (c.copiedAt) bRe.title = '重抽後，這篇會從複製紀錄移除';
    bRe.onclick = () => {
      // 已複製的卡片要按兩次，避免誤按
      if (c.copiedAt && bRe.dataset.confirm !== '1') { bRe.dataset.confirm = '1'; bRe.textContent = '確定重抽？'; msg.textContent = '這篇會從紀錄移除'; setTimeout(() => { bRe.dataset.confirm = ''; bRe.textContent = '重抽'; msg.textContent = '' }, 3000); return }
      regenOne(c);
    };
    const bCp = document.createElement('button'); bCp.type = 'button'; bCp.className = c.copiedAt ? 'ghost small' : 'copy-main small'; bCp.textContent = c.copiedAt ? '再複製一次' : '複製指令';
    bCp.onclick = () => copy(c, msg, el);
    tools.append(msg, bPin, bRe, bCp); head.appendChild(tools);
    const det = document.createElement('details'); det.className = 'prompt'; if (c.open) det.open = true; det.ontoggle = () => { c.open = det.open };
    const sm = document.createElement('summary'); sm.textContent = '查看指令內容';
    const pre = document.createElement('pre'); pre.className = 'term'; pre.innerHTML = highlight(c.text);
    det.append(sm, pre); el.append(head, det); box.appendChild(el);
  });
  drawProgress();
}
function copy(c, msg, el) {
  const markDone = () => { if (!c.copiedAt) c.copiedAt = Date.now(); log(c); saveCards() };
  const ok = () => { msg.textContent = '已複製'; markDone(); setTimeout(draw, 700) };
  const fb = () => {
    const dd = el.querySelector('details'); if (dd) dd.open = true;
    const pre = el.querySelector('pre'); const rg = document.createRange(); rg.selectNodeContents(pre); const s = getSelection(); s.removeAllRanges(); s.addRange(rg);
    msg.textContent = '已選取，請按 Ctrl/⌘+C'; markDone();
  };
  try { navigator.clipboard.writeText(c.text).then(ok, fb) } catch (e) { fb() }
}

/* ================= 紀錄（只存在這台瀏覽器） ================= */
function saveRecs() { if (!store.set('hist', JSON.stringify(RECS))) { RECS = RECS.slice(0, 150); store.set('hist', JSON.stringify(RECS)) } }
function log(c) {
  touched = true;
  if ($('f-test').checked) return;
  if (RECS.some(x => x.id === c.uid)) return; // 同一張卡再複製不重複記
  const r = c.r;
  const rec = {id: c.uid, ...cardRec(c), ts: Date.now(), sum: [SHORT[c.tab], '[' + r.tag + ']', r.scene.name, r.obj || r.group || '', r.pain].filter(Boolean).join(' / '), text: c.text};
  RECS.unshift(rec); RECS = RECS.slice(0, 300);
  saveRecs(); drawHist(); drawPool(); drawProgress();
}
function drawHist() {
  const items = RECS.slice(0, 120);
  $('hist-title').textContent = `複製紀錄（${RECS.length}）`;
  const box = $('hist-list'); box.innerHTML = '';
  items.forEach(h => {
    const d = new Date(h.ts); const ts = `${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
    const row = document.createElement('div'); row.className = 'hist-item';
    const tm = document.createElement('time'); tm.textContent = ts;
    const sp = document.createElement('span'); sp.textContent = h.sum || ''; sp.title = sp.textContent;
    const tools = document.createElement('div'); tools.className = 'inline';
    if (h.text) {
      const b = document.createElement('button'); b.type = 'button'; b.className = 'ghost small'; b.textContent = '再複製';
      b.onclick = () => { try { navigator.clipboard.writeText(h.text).then(() => { b.textContent = '已複製'; setTimeout(() => b.textContent = '再複製', 1400) }, () => { b.textContent = '無法複製' }) } catch (e) { b.textContent = '無法複製' } };
      tools.appendChild(b);
    }
    const del = document.createElement('button'); del.type = 'button'; del.className = 'ghost small'; del.textContent = '移除';
    del.onclick = () => {
      if (del.dataset.confirm !== '1') { del.dataset.confirm = '1'; del.textContent = '確定移除？'; setTimeout(() => { del.dataset.confirm = ''; del.textContent = '移除' }, 2500); return }
      RECS = RECS.filter(x => x !== h); saveRecs(); drawHist(); drawPool(); drawProgress();
    };
    tools.appendChild(del);
    row.append(tm, sp, tools); box.appendChild(row);
  });
  if (!items.length) { const p = document.createElement('p'); p.className = 'hint'; p.textContent = '還沒有紀錄。按卡片上的「複製指令」就會記一筆。'; box.appendChild(p) }
}
function resetRecords() {
  const range = $('reset-range').value, btn = $('reset-go'), msg = $('reset-msg');
  const start = new Date(); start.setHours(0, 0, 0, 0);
  const targets = RECS.filter(x => range === 'all' || x.ts >= start.getTime());
  if (!targets.length) { msg.textContent = '沒有符合的紀錄'; return }
  if (btn.dataset.confirm !== '1') { btn.dataset.confirm = '1'; btn.textContent = `確定刪除 ${targets.length} 筆？`; msg.textContent = '刪除後無法復原，再按一次才會執行'; setTimeout(() => { btn.dataset.confirm = ''; btn.textContent = '清除本機紀錄' }, 4000); return }
  btn.dataset.confirm = ''; btn.textContent = '清除本機紀錄';
  const ids = new Set(targets.map(x => x.id)); RECS = RECS.filter(x => !ids.has(x.id)); saveRecs(); drawHist(); drawPool(); drawProgress();
  msg.textContent = `已刪除 ${targets.length} 筆，抽籤已重新計算`;
  if (!cards.some(c => c.pinned)) generate();
}

/* ---------- 參數池統計 ---------- */
function drawPool() {
  const d = DATA; const allObj = new Set(), allPain = new Set();
  Object.values(d.groups).forEach(g => { g.objs.forEach(o => allObj.add(o)); g.pains.forEach(p => allPain.add(p)) });
  Object.values(d.tabPain).forEach(a => a.forEach(p => allPain.add(p)));
  const rec = recentRecs();
  const st = [[`近 ${windowDays} 天已用`, rec.length], ['品類', Object.keys(d.groups).length], ['療程', allObj.size], ['困擾', allPain.size], ['情境', d.scene.length], ['語氣', P.tone.length], ['觸發事件', TRIGGERS.length], ['範例文', LIB.length], ['禁用詞', SHEET_BAN.length]];
  CATS.forEach(c => { const tot = d.scene.filter(s => s.cats.includes(c)).length; const usedN = new Set(rec.filter(x => x.tab === c).map(x => x.scene)).size; st.push([c + ' 情境 已用/總數', usedN + ' / ' + tot]) });
  $('pool-stats').innerHTML = st.map(([k, v]) => `<div class="stat"><b>${v}</b><span>${esc(k)}</span></div>`).join('');
  const rows = Object.entries(d.groups).map(([g, v]) => {
    const tabs = CATS.filter(c => (d.tab[c] || []).includes(g)).map(c => SHORT[c]).join('、') || '—';
    const sc = d.scene.filter(s => sceneOk(s, g) && (!s.only ? true : s.only.includes(g))).length;
    const un = rec.filter(x => x.group === g).length;
    return `<tr><td>${esc(g)}${v.surgery ? '（手術）' : ''}</td><td>${esc(tabs)}</td><td class="n">${v.weight == null ? '—' : v.weight}</td><td class="n${v.objs.length < 2 ? ' low' : ''}">${v.objs.length}</td><td class="n${v.pains.length < 3 ? ' low' : ''}">${v.pains.length}</td><td class="n${sc < 5 ? ' low' : ''}">${sc}</td><td class="n">${un}</td></tr>`;
  }).join('');
  const tp = Object.entries(d.tabPain).map(([c, a]) => `<tr><td>${esc(SHORT[c])}（分頁困擾）</td><td>${esc(SHORT[c])}</td><td class="n">—</td><td class="n">—</td><td class="n${a.length < 3 ? ' low' : ''}">${a.length}</td><td class="n">${d.scene.filter(s => s.cats.includes(c)).length}</td><td class="n">${rec.filter(x => x.tab === c).length}</td></tr>`).join('');
  $('pool-table').innerHTML = '<thead><tr><th>品類</th><th>出現在</th><th class="n">比重</th><th class="n">療程</th><th class="n">困擾</th><th class="n">可配情境</th><th class="n">近期已用</th></tr></thead><tbody>' + rows + tp + '</tbody>';
  const wb = $('pool-warn-box');
  wb.hidden = !poolWarns.length;
  $('pool-warn-title').textContent = `Excel 有 ${poolWarns.length} 個地方要檢查`;
  $('pool-warns').innerHTML = poolWarns.map(w => '<li>' + esc(w) + '</li>').join('');
}

/* ================= 初始化 ================= */
function init() {
  const tabs = $('tabs');
  CATS.forEach(c => {
    const b = document.createElement('button'); b.type = 'button'; b.className = 'tab'; b.id = 'tab-' + c; b.setAttribute('role', 'tab'); b.textContent = c;
    b.onclick = () => { tab = c; store.set('cat', c); syncTabs(); syncTopic(); generate() }; tabs.appendChild(b);
  });
  const mixBox = $('mix');
  CATS.forEach(c => {
    const l = document.createElement('label'); l.htmlFor = 'mix-' + c; l.innerHTML = esc(SHORT[c]) + '<input type="number" min="0" max="10" id="mix-' + c + '">'; mixBox.appendChild(l);
    const inp = l.querySelector('input'); inp.value = mix[c]; inp.oninput = () => { mix[c] = Math.max(0, +inp.value || 0); store.set('mix', JSON.stringify(mix)) };
  });
  document.querySelectorAll('[data-preset]').forEach(b => b.onclick = () => { mix = {...PRESETS[b.dataset.preset]}; CATS.forEach(c => $('mix-' + c).value = mix[c]); store.set('mix', JSON.stringify(mix)); generate() });
  $('mode-single').onclick = () => { mode = 'single'; store.set('mode', mode); syncMode(); generate() };
  $('mode-mix').onclick = () => { mode = 'mix'; store.set('mode', mode); syncMode(); generate() };

  opts('f-gender', ['女', '男']); opts('f-who', [...new Set(P.who)]); opts('f-exp', P.exp); opts('f-budget', P.budget); opts('f-trigger', P.trigger);
  opts('f-tone', P.tone);
  opts('f-len', P.len.map(n => n + ' 字'), '隨機（多數 35–90 字）');
  [['f-ban', 'ban'], ['f-avoid', 'avoid'], ['f-kw', 'kw'], ['f-note', 'note']].forEach(([i, k]) => { $(i).value = store.get(k) || ''; $(i).addEventListener('input', () => store.set(k, $(i).value)) });
  $('form').addEventListener('change', e => {
    if (e.target.tagName !== 'SELECT') return; mark(e.target);
    if (e.target.id === 'f-group') syncTopic();
    if (e.target.id === 'f-region') syncRegion();
  });
  $('form').addEventListener('submit', e => { e.preventDefault(); generate() });
  $('regen').onclick = generate;
  $('clearpin').onclick = () => { cards.forEach(c => c.pinned = false); draw(); saveCards() };
  $('reset').onclick = () => { document.querySelectorAll('#adv select').forEach(s => { s.value = R }); syncTopic(); document.querySelectorAll('#adv select').forEach(mark) };
  const togglePush = () => { const manual = !$('f-pushrand').checked; $('push-counts').style.opacity = manual ? 1 : .4; ['f-b', 'f-n'].forEach(i => $(i).disabled = !manual) };
  $('f-pushrand').addEventListener('change', togglePush); togglePush();
  $('reset-go').onclick = resetRecords;
  {
    const tst = $('f-test'); try { tst.checked = sessionStorage.getItem(KEY + 'test') === '1' } catch (e) {}
    const sync = () => { $('testbar').hidden = !tst.checked; drawProgress(); try { sessionStorage.setItem(KEY + 'test', tst.checked ? '1' : '0') } catch (e) {} };
    tst.addEventListener('change', sync); $('test-off').onclick = () => { tst.checked = false; sync() }; sync();
  }
  $('f-win').value = String(windowDays);
  $('f-win').addEventListener('change', () => { windowDays = +$('f-win').value; store.set('win', String(windowDays)); drawPool() });
  {
    const cur = new Date().getMonth() + 1; const el = $('f-month');
    for (let m = 1; m <= 12; m++) { const o = document.createElement('option'); o.value = String(m); o.textContent = m + ' 月' + (m === cur ? '（本月）' : ''); el.appendChild(o) }
    el.value = String(cur);
    el.addEventListener('change', () => { MONTH = +el.value; generate() });
  }
  $('pool-refresh').onclick = loadPool;
  const fileIn = $('pool-file');
  [$('pool-pick'), $('pick-big')].forEach(b => b.onclick = () => fileIn.click());
  fileIn.addEventListener('change', () => { onPickFile(fileIn.files[0]); fileIn.value = '' });
  $('go').disabled = true;
  syncTabs(); syncTopic(); syncMode(); syncRegion(); drawHist();
  if (!restoreCards()) draw();
  drawPool(); loadPool();
}
init();
