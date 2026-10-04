/**
 * 商談スケジュール Web API
 *
 * GitHub Pages 上の index.html から fetch(POST) で呼ばれる。
 * 既存の Code.gs と同じ Apps Script プロジェクトに「Api.gs」として追加する
 * （LOG_SHEET / COLOR_MAP / START_HOUR などの定数や getTeamMembers, timeValueToStr_ は Code.gs のものを共有）。
 *
 * 予約ログ: A:日付 B:担当者 C:開始時刻 D:終了時刻 E:商談内容 F:色(hex) G:予約ID
 *   - G列(予約ID)はこのAPIで追加。シート側フォームで登録されたIDなしの行には、次回読み込み時に自動でIDを振る。
 *   - A〜D列・G列は文字列(@)で保存する（Sheetsの日付/時刻自動変換対策）。
 */
const LOG_ID_COL = 7;
const LOG_COLS = 7;

/** 初回だけエディタから実行：合言葉(API_TOKEN)を発行して実行ログに表示する */
function setupApiToken() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('API_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('API_TOKEN', token);
  }
  Logger.log('合言葉(API_TOKEN): ' + token);
}

function doGet() {
  return json_({ ok: true, service: 'shodan-schedule-api' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が不正です' });
  }

  const token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (!token) return json_({ ok: false, code: 'auth', error: 'サーバー側の合言葉が未設定です（setupApiToken を実行してください）' });
  if (req.token !== token) return json_({ ok: false, code: 'auth', error: '合言葉が違います' });

  try {
    switch (req.action) {
      case 'bootstrap':
        return json_({
          ok: true,
          members: getTeamMembers(),
          colors: COLOR_MAP,
          startHour: START_HOUR,
          endHour: END_HOUR,
          intervalMin: INTERVAL_MIN
        });
      case 'list':
        return json_({ ok: true, version: getVersion_(), bookings: withLock_(function () { return listBookings_(req.from, req.to); }) });
      case 'sync': {
        // 画面の定期確認用：変更がなければシートを読まずにすぐ返す
        const version = getVersion_();
        if (!req.force && req.version === version) return json_({ ok: true, changed: false, version: version });
        return json_({ ok: true, changed: true, version: version, bookings: withLock_(function () { return listMonths_(req.months); }) });
      }
      case 'create':
      case 'update':
      case 'delete':
        // 保存後の一覧も同じ応答で返し、画面側の読み直し通信を省く
        return json_(withLock_(function () {
          const res = { ok: true };
          if (req.action === 'delete') deleteBooking_(req.id);
          else res.booking = saveBooking_(req.booking, req.action === 'create');
          res.version = bumpVersion_();
          if (req.months) res.bookings = listMonths_(req.months);
          return res;
        }));
      default:
        return json_({ ok: false, error: '不明な操作です: ' + req.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 同時に複数人が登録しても重複チェックが崩れないよう、読み書きは排他で行う
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('混み合っています。少し待ってから再度お試しください');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim();
}

function normTime_(v) {
  const s = timeValueToStr_(v);
  // "9:00" のほか、Dateを文字列化した "Sat Dec 30 1899 11:00:00 GMT+0900" のような値も拾う
  const m = s.match(/^(\d{1,2}):(\d{2})/) || s.match(/\b(\d{1,2}):(\d{2})(?::\d{2})?\b/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : s;
}

function toMin_(t) {
  const m = String(t).match(/^(\d{2}):(\d{2})$/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : NaN;
}

function resolveColor_(c) {
  c = String(c || '').trim();
  if (COLOR_MAP[c]) return COLOR_MAP[c];
  const hexes = Object.keys(COLOR_MAP).map(function (k) { return COLOR_MAP[k]; });
  return hexes.indexOf(c.toLowerCase()) !== -1 ? c.toLowerCase() : hexes[0];
}

// 予約ログ全件を読む。IDが無い行にはIDを振ってG列にまとめて書き戻す
function readLog_() {
  const log = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
  if (String(log.getRange(1, LOG_ID_COL).getValue()) === '') log.getRange(1, LOG_ID_COL).setValue('予約ID');
  const lastRow = log.getLastRow();
  if (lastRow < 2) return { log: log, rows: [] };

  const values = log.getRange(2, 1, lastRow - 1, LOG_COLS).getValues();
  const ids = [];
  const rows = [];
  let idChanged = false;
  values.forEach(function (v, i) {
    let id = String(v[6] || '').trim();
    if (!v[0] && !v[1]) { ids.push([id]); return; } // 空行
    if (!id) { id = Utilities.getUuid(); idChanged = true; }
    ids.push([id]);
    rows.push({
      row: i + 2,
      id: id,
      date: normDate_(v[0]),
      member: String(v[1]).trim(),
      start: normTime_(v[2]),
      end: normTime_(v[3]),
      title: String(v[4] || ''),
      color: String(v[5] || '') || resolveColor_('')
    });
  });
  if (idChanged) {
    log.getRange(2, LOG_ID_COL, ids.length, 1).setNumberFormat('@').setValues(ids);
  }
  return { log: log, rows: rows };
}

function toBooking_(r) {
  return { id: r.id, date: r.date, member: r.member, start: r.start, end: r.end, title: r.title, color: r.color };
}

function listBookings_(from, to) {
  return readLog_().rows
    .filter(function (r) { return (!from || r.date >= from) && (!to || r.date <= to); })
    .map(toBooking_);
}

function listMonths_(months) {
  const set = {};
  (months || []).forEach(function (m) { set[String(m)] = true; });
  return readLog_().rows
    .filter(function (r) { return set[r.date.slice(0, 7)]; })
    .map(toBooking_);
}

// 予約データの版数。書き込みのたびに更新し、画面は版数が変わったときだけ読み直す
function getVersion_() {
  return PropertiesService.getScriptProperties().getProperty('DATA_VERSION') || '0';
}

function bumpVersion_() {
  // 同じミリ秒に2回書いても必ず前より大きくなるようにする
  const v = String(Math.max(Date.now(), Number(getVersion_()) + 1));
  PropertiesService.getScriptProperties().setProperty('DATA_VERSION', v);
  return v;
}

function validate_(d) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) throw new Error('日付が不正です');
  if (getTeamMembers().indexOf(d.member) === -1) throw new Error('担当者「' + d.member + '」は担当者マスタにいません');
  if (!d.title) throw new Error('商談内容を入力してください');
  const s = toMin_(d.start);
  const e = toMin_(d.end);
  if (isNaN(s) || isNaN(e)) throw new Error('時刻が不正です');
  if (s < START_HOUR * 60 || e > END_HOUR * 60) throw new Error('時刻は ' + START_HOUR + ':00〜' + END_HOUR + ':00 の範囲で指定してください');
  if (s % INTERVAL_MIN || e % INTERVAL_MIN) throw new Error('時刻は ' + INTERVAL_MIN + '分刻みで指定してください');
  if (s >= e) throw new Error('終了時刻は開始時刻より後にしてください');
}

function saveBooking_(b, isNew) {
  if (!b) throw new Error('予約データがありません');
  const data = {
    date: String(b.date || '').trim(),
    member: String(b.member || '').trim(),
    start: normTime_(b.start),
    end: normTime_(b.end),
    title: String(b.title || '').trim(),
    color: resolveColor_(b.color)
  };
  validate_(data);

  const ctx = readLog_();
  let target = null;
  if (!isNew) {
    target = ctx.rows.filter(function (r) { return r.id === String(b.id); })[0];
    if (!target) throw new Error('この予約は削除されています。画面を更新してください');
  }
  // 新規登録は画面側で決めたIDを使う。通信の再試行で同じ依頼が2回届いても1件だけになる
  const clientId = /^[A-Za-z0-9-]{8,64}$/.test(String(b.id || '')) ? String(b.id) : '';
  if (isNew && clientId) {
    const existing = ctx.rows.filter(function (r) { return r.id === clientId; })[0];
    if (existing) return toBooking_(existing);
  }
  const id = isNew ? (clientId || Utilities.getUuid()) : target.id;

  // 重複チェック（同じ日・同じ担当者で時間が重なる予約）
  const conflict = ctx.rows.filter(function (r) {
    return r.id !== id && r.date === data.date && r.member === data.member && r.start < data.end && data.start < r.end;
  })[0];
  if (conflict) {
    throw new Error(conflict.member + 'さんは ' + conflict.start + '〜' + conflict.end + ' に「' + conflict.title + '」が入っています');
  }

  const rowNum = isNew ? ctx.log.getLastRow() + 1 : target.row;
  const range = ctx.log.getRange(rowNum, 1, 1, LOG_COLS);
  range.setNumberFormat('@');
  range.setValues([[data.date, data.member, data.start, data.end, data.title, data.color, id]]);

  data.id = id;
  return data;
}

function deleteBooking_(id) {
  const ctx = readLog_();
  const target = ctx.rows.filter(function (r) { return r.id === String(id); })[0];
  if (target) ctx.log.deleteRow(target.row);
}
