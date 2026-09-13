
const SHEET_LOG = '記録';
const SHEET_WEIGHT = '体重';
const SHEET_QA = '質問';

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

/** 期間（毎月21日〜翌20日） 例: 2026/05/21〜2026/06/20 */
function periodOf_(d) {
  const y = d.getFullYear(), mo = d.getMonth(), day = d.getDate();
  let s, e;
  if (day >= 21) {
    s = new Date(y, mo, 21);
    e = new Date(y, mo + 1, 20);
  } else {
    s = new Date(y, mo - 1, 21);
    e = new Date(y, mo, 20);
  }
  return fmtDate_(s) + '〜' + fmtDate_(e);
}

function genId_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss') +
    '-' + Math.floor(Math.random() * 1e12);
}

/* ===== クライアントへ渡す初期情報 ===== */
function getInit() {
  return {
    today: fmtDate_(new Date()),       // yyyy/MM/dd
    nowTime: fmtTimeNow_(),            // H:mm
    presets: ['ワン', 'ツー', 'ごはん', '水', '散歩', 'カート散歩',
              '夢中作り', '仰向け抱っこ', '歯磨き', '耳掃除', 'ブラッシング',
              '病院', 'ネクスガード', 'カルドメック']
  };
}

/* ===== 記録 ===== */

/**
 * 指定日の記録を時刻順で返す。dateStr 省略時は当日。
 *
 * 記録は常にシート末尾に追記されていくため、まずは末尾の直近数百行だけを読み、
 * 対象日の行が見つかればそこで打ち切る（＝シート全体を毎回読み込まない）。
 * 見つからない場合だけ範囲を広げて読み直す。これにより、記録件数が増えても
 * 「今日の記録」の取得速度がほぼ一定に保たれる。
 */
function getLogs(dateStr) {
  return getLogsByDate_(dateStr);
}

/**
 * 末尾から少しずつ範囲を広げて対象日を探す内部共通処理。
 * getLogs（記録タブ）と getLogsRange の単日指定（一覧タブの「日付」モード）の
 * 両方から使う、共通の高速ルート。
 */
function getLogsByDate_(dateStr) {
  const sh = sheet_(SHEET_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sh.getLastColumn();
  const h = headers_(sh);
  const iDate = h.indexOf('日付'), iDow = h.indexOf('曜日'), iTime = h.indexOf('時刻'),
        iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考'), iId = h.indexOf('ID');
  const target = dateStr || fmtDate_(new Date());

  let chunk = 30; // まず直近30行から探す
  let values = [];
  let firstPass = true;
  while (true) {
    const startRow = Math.max(2, lastRow - chunk + 1);
    const numRows = lastRow - startRow + 1;
    values = sh.getRange(startRow, 1, numRows, lastCol).getValues();
    const found = values.some(row => normDate_(row[iDate]) === target);
    if (found || startRow === 2) break; // 見つかった、またはシート先頭まで読みきった

    if (firstPass) {
      // シートは追記順（下ほど新しい）なので、最終行の日付より新しい日付を
      // 探している場合はこれ以上遡っても見つからない（＝その日はまだデータなし）
      const lastDate = normDate_(values[values.length - 1][iDate]);
      if (lastDate && target > lastDate) return [];
      firstPass = false;
    }
    chunk *= 4; // 見つからなければ範囲を広げて読み直す
  }

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const d = normDate_(values[i][iDate]);
    if (d !== target) continue;
    out.push({
      id: String(values[i][iId] || ''),
      date: d,
      dow: String(values[i][iDow] || ''),
      time: normTime_(values[i][iTime]),
      action: String(values[i][iAct] || ''),
      note: String(values[i][iNote] || '')
    });
  }
  out.sort((a, b) => timeKey_(b.time) - timeKey_(a.time));
  return out;
}

/**
 * 日付条件なしで、最新の記録を新しい順に n 件返す（記録タブの「今日の記録」用）。
 * 記録は常にシート末尾に追記されるため、末尾から n 件だけ読めばよく高速。
 */
function getRecentLogs(n) {
  n = Number(n) || 20;
  const sh = sheet_(SHEET_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sh.getLastColumn();
  const h = headers_(sh);
  const iDate = h.indexOf('日付'), iDow = h.indexOf('曜日'), iTime = h.indexOf('時刻'),
        iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考'), iId = h.indexOf('ID');

  const startRow = Math.max(2, lastRow - n + 1);
  const numRows = lastRow - startRow + 1;
  const values = sh.getRange(startRow, 1, numRows, lastCol).getValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const d = normDate_(values[i][iDate]);
    if (!d) continue; // 空行はスキップ
    out.push({
      id: String(values[i][iId] || ''),
      date: d,
      dow: String(values[i][iDow] || ''),
      time: normTime_(values[i][iTime]),
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
  const rowObj = {};
  rowObj['日付'] = fmtDate_(dateObj);
  rowObj['曜日'] = dow_(dateObj);
  rowObj['時刻'] = (p.time && p.time.trim()) ? normTime_(p.time.trim()) : fmtTimeNow_();
  rowObj['やったこと'] = actionText;
  rowObj['備考'] = p.note || '';
  rowObj['期間'] = periodOf_(dateObj);
  rowObj['ID'] = genId_();
  const row = h.map(name => (name in rowObj) ? rowObj[name] : '');
  sh.appendRow(row);
  return { ok: true };
}

/** 記録を更新（ID一致行）。p = { time, action, note } */
function updateLog(id, p) {
  p = p || {};
  const sh = sheet_(SHEET_LOG);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iId = h.indexOf('ID'), iTime = h.indexOf('時刻'),
        iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iId]) === String(id)) {
      if (p.time !== undefined) sh.getRange(r + 1, iTime + 1).setValue(normTime_(p.time));
      if (p.action !== undefined) sh.getRange(r + 1, iAct + 1).setValue(p.action);
      if (p.note !== undefined) sh.getRange(r + 1, iNote + 1).setValue(p.note);
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

/**
 * 指定期間（from〜to、ともに 'yyyy/MM/dd'）の記録を日付→時刻順で返す。
 * from / to を省略（空文字）すると、その端は無制限として扱う。
 * 両方省略すると全期間を返す。
 */
function getLogsRange(fromStr, toStr) {
  // 単日指定（一覧タブの「日付」モード）は、getLogs と同じ「末尾から探す」
  // 高速ルートに乗せる。範囲指定・全期間のときだけシート全体を読む。
  if (fromStr && toStr && fromStr === toStr) {
    return getLogsByDate_(fromStr);
  }
  const sh = sheet_(SHEET_LOG);
  if (sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const iDate = h.indexOf('日付'), iDow = h.indexOf('曜日'), iTime = h.indexOf('時刻'),
        iAct = h.indexOf('やったこと'), iNote = h.indexOf('備考'), iId = h.indexOf('ID');
  const fromD = fromStr ? parseDate_(fromStr) : null;
  const toD = toStr ? parseDate_(toStr) : null;
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const dStr = normDate_(values[r][iDate]);
    if (!dStr) continue;
    const d = parseDate_(dStr);
    if (fromD && d < fromD) continue;
    if (toD && d > toD) continue;
    out.push({
      id: String(values[r][iId] || ''),
      date: dStr,
      dow: String(values[r][iDow] || ''),
      time: normTime_(values[r][iTime]),
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
    out.push({ date: d, kg: Number(kg), note: String(values[r][iNote] || '') });
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
