/**
 * 家計簿(割り勘)アプリ用 Apps Script
 * スプレッドシートの「拡張機能 > Apps Script」に貼り付けて使う。
 * デプロイ: ウェブアプリ / 実行ユーザー:自分 / アクセス:全員
 *
 * 支出シート列: id, 日付, 内容, 金額, 支払者, 精算済み, 全額請求
 * 買い物シート列: id, 内容, 購入済み, 購入日時
 */

const SHEET_NAME = '支出';
const SHOPPING_SHEET_NAME = '買い物';
const PASSPHRASE = 'bebichan';
const PAYERS = ['碧', '竜'];

function doGet(e) {
  if (e.parameter.passphrase !== PASSPHRASE) {
    return jsonResponse_({ error: '合言葉が違います' });
  }
  if (e.parameter.list === 'shopping') {
    return jsonResponse_({ items: readShoppingItems_() });
  }
  const items = readItems_();
  return jsonResponse_({ items: items, balance: calcBalance_(items) });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ error: 'invalid request' });
  }

  if (body.passphrase !== PASSPHRASE) {
    return jsonResponse_({ error: '合言葉が違います' });
  }

  // 同時編集対策: 書き込み処理を直列化する(完全後勝ちでよいので、
  // 衝突検出はせずロックで一件ずつ処理するだけ)
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return jsonResponse_({ error: '混み合っています。少し待って再試行してください' });
  }

  try {
    if (body.action === 'add') {
      const list = Array.isArray(body.items) ? body.items : [];
      if (!list.length) {
        return jsonResponse_({ error: '追加するデータがありません' });
      }
      const rows = [];
      for (let i = 0; i < list.length; i++) {
        const it = list[i];
        const amount = Number(it.amount);
        if (!amount || amount <= 0) {
          return jsonResponse_({ error: (i + 1) + '件目: 金額が不正です' });
        }
        if (PAYERS.indexOf(it.payer) === -1) {
          return jsonResponse_({ error: (i + 1) + '件目: 支払者が不正です' });
        }
        rows.push([
          Utilities.getUuid(),
          it.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          it.desc || '',
          amount,
          it.payer,
          false,
          !!it.fullCharge,
        ]);
      }
      const sheet = getSheet_();
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
      const items = readItems_();
      return jsonResponse_({ ok: true, items: items, balance: calcBalance_(items) });
    }

    if (body.action === 'delete') {
      const sheet = getSheet_();
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      const items = readItems_();
      return jsonResponse_({ ok: true, items: items, balance: calcBalance_(items) });
    }

    if (body.action === 'update') {
      const amount = Number(body.amount);
      if (!amount || amount <= 0) {
        return jsonResponse_({ error: '金額が不正です' });
      }
      if (PAYERS.indexOf(body.payer) === -1) {
        return jsonResponse_({ error: '支払者が不正です' });
      }
      const sheet = getSheet_();
      const data = sheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) {
          sheet.getRange(i + 1, 2, 1, 6).setValues([[
            body.date || data[i][1],
            body.desc || '',
            amount,
            body.payer,
            data[i][5], // 精算済みは変更しない
            !!body.fullCharge,
          ]]);
          found = true;
          break;
        }
      }
      if (!found) {
        return jsonResponse_({ error: '対象の記録が見つかりません' });
      }
      const items = readItems_();
      return jsonResponse_({ ok: true, items: items, balance: calcBalance_(items) });
    }

    if (body.action === 'shoppingAdd') {
      const desc = String(body.desc || '').trim();
      if (!desc) {
        return jsonResponse_({ error: '内容を入力してください' });
      }
      const sheet = getShoppingSheet_();
      sheet.appendRow([Utilities.getUuid(), desc, false, '']);
      return jsonResponse_({ ok: true, shoppingItems: readShoppingItems_() });
    }

    if (body.action === 'shoppingUpdate') {
      const desc = String(body.desc || '').trim();
      if (!desc) {
        return jsonResponse_({ error: '内容を入力してください' });
      }
      const sheet = getShoppingSheet_();
      const data = sheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) {
          sheet.getRange(i + 1, 2).setValue(desc);
          found = true;
          break;
        }
      }
      if (!found) {
        return jsonResponse_({ error: '対象が見つかりません' });
      }
      return jsonResponse_({ ok: true, shoppingItems: readShoppingItems_() });
    }

    if (body.action === 'shoppingDelete') {
      const sheet = getShoppingSheet_();
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse_({ ok: true, shoppingItems: readShoppingItems_() });
    }

    if (body.action === 'shoppingToggle') {
      const sheet = getShoppingSheet_();
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) {
          const purchased = !!body.purchased;
          sheet.getRange(i + 1, 3, 1, 2).setValues([[purchased, purchased ? new Date() : '']]);
          break;
        }
      }
      return jsonResponse_({ ok: true, shoppingItems: readShoppingItems_() });
    }

    if (body.action === 'settleMonth') {
      const yearMonth = String(body.yearMonth || '');
      if (!/^\d{4}-\d{2}$/.test(yearMonth)) {
        return jsonResponse_({ error: '年月が不正です' });
      }
      const sheet = getSheet_();
      const range = sheet.getDataRange();
      const data = range.getValues();
      for (let i = 1; i < data.length; i++) {
        const rowDate = formatDate_(data[i][1]);
        if (rowDate.indexOf(yearMonth) === 0) {
          sheet.getRange(i + 1, 6).setValue(true); // 精算済み列
        }
      }
      const items = readItems_();
      return jsonResponse_({ ok: true, items: items, balance: calcBalance_(items) });
    }

    return jsonResponse_({ error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}

function readItems_() {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // ヘッダー除去
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: r[0],
        date: formatDate_(r[1]),
        desc: r[2],
        amount: Number(r[3]),
        payer: r[4],
        settled: r[5] === true || r[5] === 'TRUE',
        fullCharge: r[6] === true || r[6] === 'TRUE',
      };
    })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; });
}

// 残高: 正なら「竜が碧に借りている」、負なら「碧が竜に借りている」
// 精算済みの記録は残高計算から除外する。全額請求は半額にせず全額を計上する
function calcBalance_(items) {
  let balance = 0;
  items.forEach(function (it) {
    if (it.settled) return;
    const sign = it.payer === '碧' ? 1 : -1;
    const portion = it.fullCharge ? it.amount : it.amount / 2;
    balance += sign * portion;
  });
  return Math.round(balance);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', '日付', '内容', '金額', '支払者', '精算済み', '全額請求']);
  }
  return sheet;
}

function readShoppingItems_() {
  const sheet = getShoppingSheet_();
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // ヘッダー除去
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: r[0],
        desc: r[1],
        purchased: r[2] === true || r[2] === 'TRUE',
      };
    });
}

function getShoppingSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHOPPING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHOPPING_SHEET_NAME);
    sheet.appendRow(['id', '内容', '購入済み', '購入日時']);
  }
  return sheet;
}

function formatDate_(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return d;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 1年半以上前の「精算済み」記録を削除する。未精算のまま古くなった記録は
// 誤って消さないよう残す(手動で精算するか確認してから消すこと)。
// installMonthlyCleanupTrigger() を一度だけ実行しておくと毎月自動で呼ばれる。
const RETENTION_MONTHS = 18;

function cleanupOldRecords() {
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - RETENTION_MONTHS);
  const cutoffStr = Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();
    let deleted = 0;
    // 後ろから消さないと行番号がずれる
    for (let i = data.length - 1; i >= 1; i--) {
      const rowDate = formatDate_(data[i][1]);
      const settled = data[i][5] === true || data[i][5] === 'TRUE';
      if (rowDate < cutoffStr && settled) {
        sheet.deleteRow(i + 1);
        deleted++;
      }
    }
    Logger.log('cleanupOldRecords: deleted ' + deleted + ' rows older than ' + cutoffStr);
  } finally {
    lock.releaseLock();
  }
}

// Apps Scriptエディタでこの関数を選んで一度だけ実行(▶)すると、
// 毎月1日の深夜にcleanupOldRecordsが自動実行されるようになる。
function installMonthlyCleanupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cleanupOldRecords') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('cleanupOldRecords')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();
}

// 購入済みになってから7日以上経った買い物リストの記録を削除する。
// 未購入の記録は件数が増えても削除しない。
const SHOPPING_RETENTION_DAYS = 7;

function cleanupOldShoppingItems() {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - SHOPPING_RETENTION_DAYS);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getShoppingSheet_();
    const data = sheet.getDataRange().getValues();
    let deleted = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      const purchased = data[i][2] === true || data[i][2] === 'TRUE';
      const purchasedAt = data[i][3];
      if (purchased && purchasedAt instanceof Date && purchasedAt < cutoff) {
        sheet.deleteRow(i + 1);
        deleted++;
      }
    }
    Logger.log('cleanupOldShoppingItems: deleted ' + deleted + ' rows');
  } finally {
    lock.releaseLock();
  }
}

// Apps Scriptエディタでこの関数を選んで一度だけ実行(▶)すると、
// 毎日深夜にcleanupOldShoppingItemsが自動実行されるようになる。
function installDailyShoppingCleanupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cleanupOldShoppingItems') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('cleanupOldShoppingItems')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();
}
