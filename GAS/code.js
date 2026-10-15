const SHEET_LOG = '記録';
const SHEET_WEIGHT = '体重';
const SHEET_QA = '質問';
const BIRTHDAY = '2026/07/21';   // 誕生日（ヘッダーの月齢表示に使う）

/* ===== Webアプリのエントリポイント ===== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('記録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===== 共通ユーティリティ ===== */
function ss_() { return SpreadsheetApp.getActive(); }
function tz_() { return ss_().getSpreadsheetTimeZone() || 'Asia/Tokyo'; }
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name);
  return sh;
}
function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

const DOW = ['日', '月', '火', '水', '木', '金', '土'];
function dow_(d) { return DOW[d.getDay()]; }

function fmtDate_(d) { return Utilities.formatDate(d, tz_(), 'yyyy/MM/dd'); }
function fmtTimeNow_() { return Utilities.formatDate(new Date(), tz_(), 'H:mm'); }
function fmtStamp_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd HH:mm'); }

/** 'yyyy/MM/dd' でも 'yyyy-MM-dd' でも Date を返す */
function parseDate_(s) {
  if (s instanceof Date) return s;
  const m = String(s).match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!m) return new Date();
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** セル値（Date or 文字列）を 'yyyy/MM/dd' 文字列へ正規化 */
function normDate_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return fmtDate_(v);
  return String(v).trim();
}
/** セル値（Date or 文字列）を 'H:mm' 文字列へ正規化 */
function normTime_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'H:mm');
  return String(v).trim();
}
/** 'H:mm' を分に。並べ替え用。解釈不能は末尾へ */
function timeKey_(t) {
  const m = String(t).match(/(\d{1,2}):(\d{2})/);
  if (!m) return 99999;
  return Number(m[1]) * 60 + Number(m[2]);
}

/** 日付文字列('yyyy/MM/dd' 等)と時刻文字列('H:mm')から、並べ替え用の実際の Date を組み立てる */
function combineDateTime_(dateStr, timeStr) {
  const d = parseDate_(dateStr);
  const m = String(timeStr || '').match(/(\d{1,2}):(\d{2})/);
  const hh = m ? Number(m[1]) : 0;
  const mm = m ? Number(m[2]) : 0;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), hh, mm, 0);
}

/** 期間（毎月10日〜翌9日） 例: 2026/09/10〜2026/10/09 */
function periodOf_(d) {
  const y = d.getFullYear(), mo = d.getMonth(), day = d.getDate();
  let s, e;
  if (day >= 10) {
    s = new Date(y, mo, 10);
    e = new Date(y, mo + 1, 9);
  } else {
    s = new Date(y, mo - 1, 10);
    e = new Date(y, mo, 9);
  }
  return fmtDate_(s) + '〜' + fmtDate_(e);
}

/**
 * 【一度だけ手動実行】記録シートの「期間」列を、「日時」列から新ルール（10日〜翌9日）で全行書き直す。
 * 期間ルールを21日〜翌20日から変更したため、過去の記録の期間列を揃えたいときに使う。
 * アプリからは呼ばれない。不要なら実行しなくてよい。
 */
function recalcPeriods() {
  const sh = sheet_(SHEET_LOG);
  if (sh.getLastRow() < 2) return;
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iDatetime = h.indexOf('日時'), iPeriod = h.indexOf('期間');
  if (iDatetime < 0 || iPeriod < 0) throw new Error('「日時」または「期間」列が見つかりません');
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const v = values[r][iDatetime];
    // 日時が空の行は元の期間をそのまま残す
    out.push([v ? periodOf_(v instanceof Date ? v : parseDate_(v)) : values[r][iPeriod]]);
  }
  sh.getRange(2, iPeriod + 1, out.length, 1).setValues(out);
}

function genId_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss') +
    '-' + Math.floor(Math.random() * 1e12);
}

/* ===== ショートカット（マスタ） ===== */
const SHEET_SHORTCUT = 'マスタ';

/** マスタシートが無い・空のときに使う初期値 */
const DEFAULT_SHORTCUTS = [
  { name: 'ワン' }, { name: 'ツー' }, { name: 'ごはん' }, { name: '水' },
  { name: '散歩', ask: true, max: 120 },
  { name: '夢中作り', ask: true, max: 60 },
  { name: '仰向け抱っこ', ask: true, max: 60 },
  { name: '歯磨き' }, { name: '耳掃除' }, { name: '爪切り' }, { name: 'ブラッシング' },
  { name: '病院' }, { name: 'ネクスガード' }, { name: 'カルドメック' }
];

/**
 * 「マスタ」シートからショートカット一覧を読む。
 * 見出し行（「項目名」または「表示名」を含む行）を上から探し、その次の行からをデータとして読む。
 * 列: 項目名 / 分数を聞く（○なら分数選択を出す） / 最大分（5分刻みの上限。空欄は60）
 * ※ シートの上のタイトル行や、右側のメモ列は無視される。
 * シートが無い・読めない・有効な行が1つも無いときは DEFAULT_SHORTCUTS を返す。
 */
function getShortcuts_() {
  try {
    const sh = ss_().getSheetByName(SHEET_SHORTCUT);
    if (!sh || sh.getLastRow() < 2) return DEFAULT_SHORTCUTS;
    const values = sh.getDataRange().getValues();
    const isNameHeader = x => ['項目名', '表示名'].indexOf(String(x).trim()) >= 0;
    let hr = -1;
    for (let r = 0; r < Math.min(values.length, 10); r++) {
      if (values[r].some(isNameHeader)) { hr = r; break; }
    }
    if (hr < 0) return DEFAULT_SHORTCUTS;
    const h = values[hr].map(x => String(x).trim());
    const iName = h.findIndex(isNameHeader);
    const iAsk = h.indexOf('分数を聞く'), iMax = h.indexOf('最大分');
    const out = [], seen = {};
    for (let r = hr + 1; r < values.length; r++) {
      const name = String(values[r][iName] || '').trim();
      if (!name || seen[name]) continue; // 空行・重複は無視
      seen[name] = true;
      const askVal = iAsk >= 0 ? values[r][iAsk] : '';
      const ask = askVal === true || ['○', '〇', 'TRUE', 'true', '1'].indexOf(String(askVal).trim()) >= 0;
      const max = iMax >= 0 ? Number(values[r][iMax]) : NaN;
      out.push({ name: name, ask: ask, max: (max >= 5) ? Math.floor(max / 5) * 5 : 60 });
    }
    return out.length ? out : DEFAULT_SHORTCUTS;
  } catch (e) {
    return DEFAULT_SHORTCUTS;
  }
}

/* ===== クライアントへ渡す初期情報 ===== */
function getInit() {
  const shortcuts = getShortcuts_();
  const durations = {};
  shortcuts.forEach(s => { if (s.ask) durations[s.name] = s.max || 60; });
  // ヘッダーに出す最新の体重（体重シートが読めなくても起動は止めない）
  let latestWeight = null;
  try { const w = getWeights(); if (w.length) latestWeight = w[w.length - 1].kg; } catch (e) {}
  return {
    birthday: BIRTHDAY,                // yyyy/MM/dd
    latestWeight: latestWeight,
    today: fmtDate_(new Date()),       // yyyy/MM/dd
    nowTime: fmtTimeNow_(),            // H:mm
    presets: shortcuts.map(s => s.name),
    durations: durations               // 分数選択を出すショートカット名 → 最大分
  };
}

/* ===== 記録 ===== */

/**
 * 指定日の記録を時刻順（新しい順）で返す。dateStr 省略時は当日。
 *
 * 記録シートは追加・更新のたびに「日時」列で降順（新しい順）に自動で並べ替えられるため、
 * まずは先頭の直近数十行だけを読み、対象日の行が見つかればそこで打ち切る
 * （＝シート全体を毎回読み込まない）。見つからない場合だけ範囲を広げて読み直す。
 * これにより、記録件数が増えても「今日の記録」の取得速度がほぼ一定に保たれる。
 */
function getLogs(dateStr) {
  return getLogsByDate_(dateStr);
}

/**
 * 先頭（＝最新）から少しずつ範囲を広げて対象日を探す内部共通処理。
 * getLogs（記録タブ）と getLogsRange の単日指定（一覧タブの「日付」モード）の
 * 両方から使う、共通の高速ルート。
 */
function getLogsByDate_(dateStr) {
  const sh = sheet_(SHEET_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sh.getLastColumn();
  const h = headers_(sh);
  const iDatetime = h.indexOf('日時'), iDow = h.indexOf('曜日'),
        iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考'), iId = h.indexOf('ID');
  const target = dateStr || fmtDate_(new Date());

  let chunk = 30; // まず先頭30行から探す
  let values = [];
  let firstPass = true;
  while (true) {
    const numRows = Math.min(chunk, lastRow - 1);
    values = sh.getRange(2, 1, numRows, lastCol).getValues();
    // シートは日時降順。読んだ範囲の最後の行が対象日より古ければ、その日の分は全部読めている。
    // （対象日の行が1行でも見つかった時点で止めると、1日分がチャンクをまたぐとき古い時間帯が欠ける）
    const lastDate = normDate_(values[values.length - 1][iDatetime]);
    if (numRows === lastRow - 1 || (lastDate && lastDate < target)) break;

    if (firstPass) {
      // シートは日時降順（上ほど新しい）に保たれているので、先頭行の日付より
      // 新しい日付を探している場合はこれ以上読んでも見つからない（＝まだデータなし）
      const topDate = normDate_(values[0][iDatetime]);
      if (topDate && target > topDate) return [];
      firstPass = false;
    }
    chunk *= 4; // 見つからなければ範囲を広げて読み直す
  }

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const d = normDate_(values[i][iDatetime]);
    if (d !== target) continue;
    out.push({
      id: String(values[i][iId] || ''),
      date: d,
      dow: String(values[i][iDow] || ''),
      time: normTime_(values[i][iDatetime]),
      action: String(values[i][iAct] || ''),
      note: String(values[i][iNote] || '')
    });
  }
  out.sort((a, b) => timeKey_(b.time) - timeKey_(a.time));
  return out;
}

/**
 * 記録がある直近 n 日分を、日付ごとに途中で切れないよう全件返す（記録タブ用）。
 * 戻り値: { logs: [...新しい順], hasMore: それより前の日の記録がまだあるか }
 * 記録シートは日時降順に保たれているので、先頭から読み、(n+1)日目の行が出たところで打ち切る。
 */
function getLogsDays(n) {
  n = Math.max(1, Math.floor(Number(n)) || 2);
  const sh = sheet_(SHEET_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { logs: [], hasMore: false };
  const lastCol = sh.getLastColumn();
  const h = headers_(sh);
  const iDatetime = h.indexOf('日時'), iDow = h.indexOf('曜日'),
        iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考'), iId = h.indexOf('ID');

  let chunk = 60, values = [], dates = [];
  while (true) {
    const numRows = Math.min(chunk, lastRow - 1);
    values = sh.getRange(2, 1, numRows, lastCol).getValues();
    const seen = {};
    dates = [];
    values.forEach(r => {
      const d = normDate_(r[iDatetime]);
      if (d && !seen[d]) { seen[d] = true; dates.push(d); }
    });
    dates.sort().reverse();
    // (n+1)日目の日付が見えていれば、n日目までは途中で切れていない
    if (dates.length > n || numRows === lastRow - 1) break;
    chunk *= 4;
  }
  const hasMore = dates.length > n;
  const keep = {};
  dates.slice(0, n).forEach(d => { keep[d] = true; });

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const d = normDate_(values[i][iDatetime]);
    if (!d || !keep[d]) continue;
    out.push({
      id: String(values[i][iId] || ''),
      date: d,
      dow: String(values[i][iDow] || ''),
      time: normTime_(values[i][iDatetime]),
      action: String(values[i][iAct] || ''),
      note: String(values[i][iNote] || '')
    });
  }
  out.sort((a, b) => {
    const diff = parseDate_(b.date) - parseDate_(a.date);
    return diff !== 0 ? diff : (timeKey_(b.time) - timeKey_(a.time));
  });
  return { logs: out, hasMore: hasMore };
}

/**
 * 日付条件なしで、最新の記録を新しい順に n 件返す（記録タブの「今日の記録」用）。
 * 記録シートは常に「日時」列で降順に保たれているため、先頭から n 件だけ読めばよく高速。
 */
function getRecentLogs(n) {
  n = Number(n) || 20;
  const sh = sheet_(SHEET_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sh.getLastColumn();
  const h = headers_(sh);
  const iDatetime = h.indexOf('日時'), iDow = h.indexOf('曜日'),
        iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考'), iId = h.indexOf('ID');

  const numRows = Math.min(n, lastRow - 1);
  const values = sh.getRange(2, 1, numRows, lastCol).getValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const d = normDate_(values[i][iDatetime]);
    if (!d) continue; // 空行はスキップ
    out.push({
      id: String(values[i][iId] || ''),
      date: d,
      dow: String(values[i][iDow] || ''),
      time: normTime_(values[i][iDatetime]),
      action: String(values[i][iAct] || ''),
      note: String(values[i][iNote] || '')
    });
  }
  // 追記順ではなく、実際の日付＋時刻で新しい順（降順）に並べ替える
  out.sort((a, b) => {
    const diff = parseDate_(b.date) - parseDate_(a.date);
    return diff !== 0 ? diff : (timeKey_(b.time) - timeKey_(a.time));
  });
  return out.slice(0, n);
}

/**
 * 記録を追加
 * p = { date:'yyyy/MM/dd'|'', time:'H:mm'|'', action:'やったこと（、区切りテキスト）', note:'' }
 *
 * ※ 「やったこと」は画面側でチップ選択と自由入力を統合した1本のテキストとして
 *    渡されてくるので、ここではそのまま保存する（サーバー側での結合はしない）。
 */
function addLog(p) {
  p = p || {};
  const sh = sheet_(SHEET_LOG);
  const h = headers_(sh);
  const dateObj = p.date ? parseDate_(p.date) : new Date();
  const actionText = String(p.action || '').trim();
  if (!actionText) {
    return { ok: false, error: 'やったことを入力してください' };
  }
  const timeStr = (p.time && p.time.trim()) ? normTime_(p.time.trim()) : fmtTimeNow_();
  const rowObj = {};
  rowObj['曜日'] = dow_(dateObj);
  rowObj['やったこと'] = actionText;
  rowObj['備考'] = p.note || '';
  rowObj['期間'] = periodOf_(dateObj);
  rowObj['ID'] = genId_();
  rowObj['日時'] = combineDateTime_(fmtDate_(dateObj), timeStr);
  const row = h.map(name => (name in rowObj) ? rowObj[name] : '');
  sh.appendRow(row);
  sortLogSheet_();
  return { ok: true };
}

/** 記録を更新（ID一致行）。p = { time, action, note } */
function updateLog(id, p) {
  p = p || {};
  const sh = sheet_(SHEET_LOG);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iId = h.indexOf('ID'), iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考'),
        iDatetime = h.indexOf('日時');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iId]) === String(id)) {
      if (p.time !== undefined) {
        const newTime = normTime_(p.time);
        const dStr = normDate_(values[r][iDatetime]);
        sh.getRange(r + 1, iDatetime + 1).setValue(combineDateTime_(dStr, newTime));
      }
      if (p.action !== undefined) sh.getRange(r + 1, iAct + 1).setValue(p.action);
      if (p.note !== undefined) sh.getRange(r + 1, iNote + 1).setValue(p.note);
      sortLogSheet_();
      return { ok: true };
    }
  }
  return { ok: false, error: '対象が見つかりませんでした' };
}

/** 記録を削除（ID一致行） */
function deleteLog(id) {
  const sh = sheet_(SHEET_LOG);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iId = h.indexOf('ID');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iId]) === String(id)) {
      sh.deleteRow(r + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: '対象が見つかりませんでした' };
}

/** 記録シートを「日時」列で新しい順（降順）に並べ替える。日時列が無ければ何もしない */
function sortLogSheet_() {
  const sh = sheet_(SHEET_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return; // データが0〜1行なら並べ替え不要
  const lastCol = sh.getLastColumn();
  const h = headers_(sh);
  const iDatetime = h.indexOf('日時');
  if (iDatetime < 0) return; // 日時列がまだ無ければ何もしない（後方互換）
  sh.getRange(2, 1, lastRow - 1, lastCol).sort({ column: iDatetime + 1, ascending: false });
}


/**
 * 指定期間（from〜to、ともに 'yyyy/MM/dd'）の記録を日付→時刻順で返す。
 * from / to を省略（空文字）すると、その端は無制限として扱う。
 * 両方省略すると全期間を返す。
 */
function getLogsRange(fromStr, toStr) {
  // 単日指定（一覧タブの「日付」モード）は、getLogsByDate_ と同じ「先頭から探す」
  // 高速ルートに乗せる。範囲指定・全期間のときだけシート全体を読む。
  if (fromStr && toStr && fromStr === toStr) {
    return getLogsByDate_(fromStr);
  }
  const sh = sheet_(SHEET_LOG);
  if (sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iDatetime = h.indexOf('日時'), iDow = h.indexOf('曜日'),
        iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考'), iId = h.indexOf('ID');
  const fromD = fromStr ? parseDate_(fromStr) : null;
  const toD = toStr ? parseDate_(toStr) : null;
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const dStr = normDate_(values[r][iDatetime]);
    if (!dStr) continue;
    const d = parseDate_(dStr);
    if (fromD && d < fromD) continue;
    if (toD && d > toD) continue;
    out.push({
      id: String(values[r][iId] || ''),
      date: dStr,
      dow: String(values[r][iDow] || ''),
      time: normTime_(values[r][iDatetime]),
      action: String(values[r][iAct] || ''),
      note: String(values[r][iNote] || '')
    });
  }
  out.sort((a, b) => {
    const diff = parseDate_(b.date) - parseDate_(a.date);
    return diff !== 0 ? diff : (timeKey_(b.time) - timeKey_(a.time));
  });
  return out;
}

/* ===== 体重 ===== */
function getWeights() {
  const sh = sheet_(SHEET_WEIGHT);
  if (sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iDate = h.indexOf('計測日');
  let iKg = h.findIndex(x => String(x).indexOf('体重') >= 0); if (iKg < 0) iKg = 1;
  let iNote = h.indexOf('備考'); if (iNote < 0) iNote = 2;
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const d = normDate_(values[r][iDate]);
    const kg = values[r][iKg];
    if (!d || kg === '' || kg === null) continue;
    out.push({ date: d, kg: Number(kg), note: String(values[r][iNote] || ''), row: r + 1 });
  }
  out.sort((a, b) => parseDate_(a.date) - parseDate_(b.date));
  return out;
}

/** p = { date, kg, note } */
function addWeight(p) {
  p = p || {};
  if (!p.kg || isNaN(Number(p.kg))) return { ok: false, error: '体重(数値)を入力してください' };
  const sh = sheet_(SHEET_WEIGHT);
  const h = headers_(sh);
  const dateObj = p.date ? parseDate_(p.date) : new Date();
  const rowObj = {};
  rowObj['計測日'] = fmtDate_(dateObj);
  const kgHeader = h.find(x => String(x).indexOf('体重') >= 0) || '体重(kg)';
  rowObj[kgHeader] = Number(p.kg);
  rowObj['備考'] = p.note || '';
  const row = h.map(name => (name in rowObj) ? rowObj[name] : '');
  sh.appendRow(row);
  return { ok: true };
}

/**
 * 体重の更新・削除は「シートの行番号」で対象を特定する（体重シートにはIDがないため）。
 * 並び替えや行の挿入で行がずれていた場合に別の行を書き換えないよう、
 * expect = { date, kg } （画面で見えていた元の値）と一致するか確認してから実行する。
 */
function findWeightRow_(row, expect) {
  const sh = sheet_(SHEET_WEIGHT);
  row = Number(row);
  if (!row || row < 2 || row > sh.getLastRow()) return { error: '対象が見つかりませんでした。再読み込みしてください' };
  const h = headers_(sh);
  const iDate = h.indexOf('計測日');
  let iKg = h.findIndex(x => String(x).indexOf('体重') >= 0); if (iKg < 0) iKg = 1;
  let iNote = h.indexOf('備考'); if (iNote < 0) iNote = 2;
  const cur = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  if (expect) {
    const sameDate = normDate_(cur[iDate]) === normDate_(expect.date);
    const sameKg = Number(cur[iKg]) === Number(expect.kg);
    if (!sameDate || !sameKg) return { error: 'データが変更されています。再読み込みしてください' };
  }
  return { sh, row, iDate, iKg, iNote };
}

/** 体重を更新。p = { date, kg, note } */
function updateWeight(row, p, expect) {
  p = p || {};
  if (p.kg === '' || p.kg === null || p.kg === undefined || isNaN(Number(p.kg))) {
    return { ok: false, error: '体重(数値)を入力してください' };
  }
  const t = findWeightRow_(row, expect);
  if (t.error) return { ok: false, error: t.error };
  const dateObj = p.date ? parseDate_(p.date) : new Date();
  t.sh.getRange(t.row, t.iDate + 1).setValue(fmtDate_(dateObj));
  t.sh.getRange(t.row, t.iKg + 1).setValue(Number(p.kg));
  t.sh.getRange(t.row, t.iNote + 1).setValue(p.note || '');
  return { ok: true };
}

/** 体重の記録を行ごと削除 */
function deleteWeight(row, expect) {
  const t = findWeightRow_(row, expect);
  if (t.error) return { ok: false, error: t.error };
  t.sh.deleteRow(t.row);
  return { ok: true };
}

/* ===== 質問 ===== */
function getQA() {
  const sh = sheet_(SHEET_QA);
  if (sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iNo = h.indexOf('No'), iDt = h.indexOf('起票日時'),
        iQ = h.indexOf('質問'), iA = h.indexOf('回答');
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const q = String(values[r][iQ] || '').trim();
    if (!q) continue; // 空の予約行はスキップ
    out.push({
      no: values[r][iNo],
      date: normDateTime_(values[r][iDt]),
      q: q,
      a: String(values[r][iA] || '')
    });
  }
  // No降順（新しい質問を上に）
  out.sort((a, b) => (Number(b.no) || 0) - (Number(a.no) || 0));
  return out;
}

function normDateTime_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy/MM/dd HH:mm');
  return String(v).trim();
}

/** 質問を追加（空の予約行があればそこへ、なければ末尾に追加） p = { q } */
function addQA(p) {
  p = p || {};
  const q = String(p.q || '').trim();
  if (!q) return { ok: false, error: '質問を入力してください' };
  const sh = sheet_(SHEET_QA);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iNo = h.indexOf('No'), iDt = h.indexOf('起票日時'), iQ = h.indexOf('質問');
  let maxNo = 0, targetRow = -1;
  for (let r = 1; r < values.length; r++) {
    const no = Number(values[r][iNo]);
    if (!isNaN(no) && no > maxNo) maxNo = no;
    if (targetRow < 0 && String(values[r][iQ] || '').trim() === '') targetRow = r;
  }
  const stamp = fmtStamp_();
  if (targetRow < 0) {
    const rowObj = {};
    rowObj['No'] = maxNo + 1;
    rowObj['起票日時'] = stamp;
    rowObj['質問'] = q;
    rowObj['回答'] = '';
    sh.appendRow(h.map(n => (n in rowObj) ? rowObj[n] : ''));
  } else {
    if (String(values[targetRow][iNo] || '').trim() === '') {
      sh.getRange(targetRow + 1, iNo + 1).setValue(maxNo + 1);
    }
    sh.getRange(targetRow + 1, iDt + 1).setValue(stamp);
    sh.getRange(targetRow + 1, iQ + 1).setValue(q);
  }
  return { ok: true };
}

/** 回答を更新（No一致行） */
function updateAnswer(no, answer) {
  const sh = sheet_(SHEET_QA);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iNo = h.indexOf('No'), iA = h.indexOf('回答');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iNo]) === String(no)) {
      sh.getRange(r + 1, iA + 1).setValue(answer || '');
      return { ok: true };
    }
  }
  return { ok: false, error: '対象が見つかりませんでした' };
}

/** 質問と回答をまとめて更新（No一致行）。p = { q, a } */
function updateQA(no, p) {
  p = p || {};
  const q = String(p.q || '').trim();
  if (!q) return { ok: false, error: '質問を入力してください' };
  const sh = sheet_(SHEET_QA);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iNo = h.indexOf('No'), iQ = h.indexOf('質問'), iA = h.indexOf('回答');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iNo]) === String(no)) {
      sh.getRange(r + 1, iQ + 1).setValue(q);
      sh.getRange(r + 1, iA + 1).setValue(p.a || '');
      return { ok: true };
    }
  }
  return { ok: false, error: '対象が見つかりませんでした' };
}

/** 質問を行ごと削除（No一致行） */
function deleteQA(no) {
  const sh = sheet_(SHEET_QA);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iNo = h.indexOf('No'), iQ = h.indexOf('質問');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iNo]) === String(no) && String(values[r][iQ] || '').trim() !== '') {
      sh.deleteRow(r + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: '対象が見つかりませんでした' };
}
