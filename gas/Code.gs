/**
 * 家計簿(割り勘)アプリ用 Apps Script
 * スプレッドシートの「拡張機能 > Apps Script」に貼り付けて使う。
 * デプロイ: ウェブアプリ / 実行ユーザー:自分 / アクセス:全員
 *
 * シート列: id, 日付, 内容, 金額, 支払者, 精算済み
 */

const SHEET_NAME = '支出';
const PASSPHRASE = 'bebichan';
const PAYERS = ['碧', '竜'];

function doGet(e) {
  if (e.parameter.passphrase !== PASSPHRASE) {
    return jsonResponse_({ error: '合言葉が違います' });
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
      const amount = Number(body.amount);
      if (!amount || amount <= 0) {
        return jsonResponse_({ error: '金額が不正です' });
      }
      if (PAYERS.indexOf(body.payer) === -1) {
        return jsonResponse_({ error: '支払者が不正です' });
      }
      const sheet = getSheet_();
      const id = Utilities.getUuid();
      sheet.appendRow([
        id,
        body.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        body.desc || '',
        amount,
        body.payer,
        false,
      ]);
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
      };
    })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; });
}

// 残高: 正なら「竜が碧に借りている」、負なら「碧が竜に借りている」
// 精算済みの記録は残高計算から除外する
function calcBalance_(items) {
  let balance = 0;
  items.forEach(function (it) {
    if (it.settled) return;
    const sign = it.payer === '碧' ? 1 : -1;
    balance += sign * (it.amount / 2);
  });
  return Math.round(balance);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', '日付', '内容', '金額', '支払者', '精算済み']);
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
