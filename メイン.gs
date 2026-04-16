// ===== 設定 =====
var CONFIG = {
  senderName: 'チケモ運営事務局',
  contactEmail: 'chikemo.info@gmail.com',
  resendFrom: 'チケモ運営事務局 <chikemo.info@chikemo.net>',
  shippingMethod: '日本郵便 レターパックライト',
};

// ===== メイン処理：入金列が変更されたら自動でメール送信 =====
function onEdit(e) {
  if (!e || !e.range || e.range.getRow() <= 1) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'シート1') return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = e.range.getRow();

  // 編集された行の運用列が空なら自動コピー
  autoCopyColumns_(sheet, row, headers);

  var paymentCol = headers.indexOf('入金') + 1;

  // 入金列以外の編集は無視
  if (paymentCol === 0 || e.range.getColumn() !== paymentCol) return;
  var value = String(e.range.getValue()).trim();

  if (value === 'OK') sendShippingNotification_(sheet, row, headers);
  if (value === 'NG') sendCancellationNotification_(sheet, row, headers);
}

// ===== 新規行の自動コピー（onEdit内で呼び出し） =====
function autoCopyColumns_(sheet, row, headers) {
  var copyMap = [
    { from: '商品お届け先名', to: '送付先名' },
    { from: '商品お届け先住所', to: '送付先住所' },
    { from: 'ご購入商品', to: 'ご購入商品' },
    { from: '購入枚数', to: '購入枚数' },
    { from: 'お名前（スペースなし）', to: 'お名前' },
    { from: 'フリガナ（スペースなし）', to: 'フリガナ' },
  ];

  copyMap.forEach(function(pair) {
    var fromCol = headers.indexOf(pair.from) + 1;
    var toCols = [];
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === pair.to) toCols.push(i + 1);
    }
    var toCol = toCols.length > 1 ? toCols[1] : toCols[0];
    if (fromCol > 0 && toCol > 0 && fromCol !== toCol) {
      if (String(sheet.getRange(row, toCol).getValue()).trim() === '') {
        var value = sheet.getRange(row, fromCol).getValue();
        sheet.getRange(row, toCol).setValue(value);
      }
    }
  });
}

// ===== 発送通知メール =====
function sendShippingNotification_(sheet, row, headers) {
  if (getCell_(sheet, row, headers, '発送通知済み') === '送信済み') return;

  var email = getCell_(sheet, row, headers, 'メールアドレス');
  var tracking = getCell_(sheet, row, headers, '追跡番号');
  if (!email) return setError_(sheet, row, headers, '発送通知', 'メールアドレスが空');
  if (!tracking) return setError_(sheet, row, headers, '発送通知', '追跡番号が空');

  var name = getCell_(sheet, row, headers, 'システム表示名');
  var item = getCell_(sheet, row, headers, 'ご購入商品');
  var qty = getCell_(sheet, row, headers, '購入枚数');
  var toName = getCell_(sheet, row, headers, '商品お届け先名');
  var toAddr = getCell_(sheet, row, headers, '商品お届け先住所');

  var body =
    name + '様\n\n' +
    'チケモをご利用いただき、誠にありがとうございます。\n' +
    'ご入金が確認できましたので、お知らせいたします。\n\n' +
    '【ご注文商品】\n' +
    '・商品名：' + item + '\n' +
    '・購入枚数：' + qty + '\n\n' +
    '【発送について】\n' +
    '商品の発送は3日以内に行います。\n' +
    '追跡番号のご連絡は原則当日中にいたします。\n' +
    '・送付先名：' + toName + '\n' +
    '・送付先住所：' + toAddr + '\n' +
    '・発送方法：' + CONFIG.shippingMethod + '\n' +
    '・到着予定：発送から1〜3日\n' +
    '・追跡番号：' + tracking + '\n' +
    '※ポスト投函でのお届けとなります（受取サイン不要）\n' +
    '※追跡番号の反映はポスト投函から半日程度時間を要します\n\n' +
    'ご不明な点がございましたら、お名前を添えて下記までお問い合わせください。\n' +
    CONFIG.contactEmail + '\n\n' +
    'この度はご利用いただき誠にありがとうございました。\n' +
    '今後とも、チケモをよろしくお願いいたします。\n\n' +
    CONFIG.senderName;

  sendEmail_(sheet, row, headers, '発送通知', email, '【追跡番号のお知らせ】ご入金ありがとうございます', body);
}

// ===== キャンセル通知メール =====
function sendCancellationNotification_(sheet, row, headers) {
  if (getCell_(sheet, row, headers, 'キャンセル通知済み') === '送信済み') return;

  var email = getCell_(sheet, row, headers, 'メールアドレス');
  if (!email) return setError_(sheet, row, headers, 'キャンセル通知', 'メールアドレスが空');

  var name = getCell_(sheet, row, headers, 'システム表示名');
  var item = getCell_(sheet, row, headers, 'ご購入商品');
  var qty = getCell_(sheet, row, headers, '購入枚数');
  var toName = getCell_(sheet, row, headers, '商品お届け先名');
  var toAddr = getCell_(sheet, row, headers, '商品お届け先住所');

  var body =
    name + '様\n\n' +
    'チケモをご利用いただき、誠にありがとうございます。\n' +
    '以下のご注文につきまして、大変恐れ入りますが、キャンセルとさせていただきます。\n\n' +
    '【ご注文内容】\n' +
    '・商品名：' + item + '\n' +
    '・購入枚数：' + qty + '\n' +
    '・送付先名：' + toName + '\n' +
    '・送付先住所：' + toAddr + '\n\n' +
    '商品をご希望の際は、改めてLINEよりご注文ください。\n' +
    'チケモLINE公式アカウント：https://lin.ee/nbdod08F\n\n' +
    'この度はご利用いただき誠にありがとうございました。\n' +
    '今後とも、チケモをよろしくお願いいたします。\n\n' +
    CONFIG.senderName;

  sendEmail_(sheet, row, headers, 'キャンセル通知', email, '【チケモ】ご注文キャンセルのお知らせ', body);
}

// ===== メール送信 & ステータス記録 =====
function sendEmail_(sheet, row, headers, type, to, subject, body) {
  try {
    GmailApp.sendEmail(to, subject, body, {
      name: CONFIG.senderName,
      replyTo: CONFIG.contactEmail,
    });
    setCell_(sheet, row, headers, type + '済み', '送信済み');
    setCell_(sheet, row, headers, type + '日時', new Date());
    setCell_(sheet, row, headers, type + 'エラー', '');
  } catch (gmailErr) {
    if (isQuotaError_(gmailErr)) {
      try {
        sendViaResend_(to, subject, body);
        setCell_(sheet, row, headers, type + '済み', '送信済み');
        setCell_(sheet, row, headers, type + '日時', new Date());
        setCell_(sheet, row, headers, type + 'エラー', '');
      } catch (resendErr) {
        setCell_(sheet, row, headers, type + '済み', 'エラー');
        setCell_(sheet, row, headers, type + '日時', new Date());
        setCell_(sheet, row, headers, type + 'エラー', 'Resend fallback失敗: ' + String(resendErr));
      }
    } else {
      setCell_(sheet, row, headers, type + '済み', 'エラー');
      setCell_(sheet, row, headers, type + '日時', new Date());
      setCell_(sheet, row, headers, type + 'エラー', String(gmailErr));
    }
  }
}

// ===== Gmail クォータエラー判定 =====
function isQuotaError_(err) {
  var msg = String(err).toLowerCase();
  return msg.indexOf('limit') !== -1
    || msg.indexOf('quota') !== -1
    || msg.indexOf('too many') !== -1;
}

// ===== Resend API によるメール送信 =====
function sendViaResend_(to, subject, body) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');
  if (!apiKey) throw new Error('RESEND_API_KEY が未設定');

  var res = UrlFetchApp.fetch('https://api.resend.com/emails', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify({
      from: CONFIG.resendFrom,
      to: [to],
      subject: subject,
      text: body,
      reply_to: CONFIG.contactEmail,
    }),
    muteHttpExceptions: true,
  });

  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Resend API error ' + code + ': ' + res.getContentText());
  }
}

function setError_(sheet, row, headers, type, message) {
  setCell_(sheet, row, headers, type + '済み', 'エラー');
  setCell_(sheet, row, headers, type + '日時', new Date());
  setCell_(sheet, row, headers, type + 'エラー', message);
}

// ===== セル読み書き =====
function getCell_(sheet, row, headers, name) {
  var col = headers.indexOf(name) + 1;
  return col > 0 ? String(sheet.getRange(row, col).getValue()).trim() : '';
}

function setCell_(sheet, row, headers, name, value) {
  var col = headers.indexOf(name) + 1;
  if (col > 0) sheet.getRange(row, col).setValue(value);
}

// ===== トリガー管理 =====
function setupTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  Logger.log('トリガー設定完了');
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
}

function backfillCopyColumns() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastRow = sheet.getLastRow();

  var copyMap = [
    { from: '商品お届け先名', to: '送付先名' },
    { from: '商品お届け先住所', to: '送付先住所' },
    { from: 'ご購入商品', to: 'ご購入商品' },
    { from: '購入枚数', to: '購入枚数' },
    { from: 'お名前（スペースなし）', to: 'お名前' },
    { from: 'フリガナ（スペースなし）', to: 'フリガナ' },
  ];

  var resolved = copyMap.map(function(pair) {
    var fromCol = headers.indexOf(pair.from) + 1;
    var toCols = [];
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === pair.to) toCols.push(i + 1);
    }
    var toCol = toCols.length > 1 ? toCols[1] : toCols[0];
    return { fromCol: fromCol, toCol: toCol };
  }).filter(function(r) { return r.fromCol > 0 && r.toCol > 0 && r.fromCol !== r.toCol; });

  var filled = 0;
  for (var row = 2; row <= lastRow; row++) {
    var hasFormData = String(sheet.getRange(row, 1).getValue()).trim() !== '';
    if (!hasFormData) continue;

    var needsFill = resolved.some(function(r) {
      return String(sheet.getRange(row, r.toCol).getValue()).trim() === '';
    });
    if (!needsFill) continue;

    resolved.forEach(function(r) {
      var value = sheet.getRange(row, r.fromCol).getValue();
      sheet.getRange(row, r.toCol).setValue(value);
    });
    filled++;
  }
  Logger.log('バックフィル完了: ' + filled + '行処理');
}

function setupPaymentDropdown() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf('入金') + 1;
  if (col === 0) { Logger.log('入金列が見つかりません'); return; }

  var range = sheet.getRange(2, col, sheet.getMaxRows() - 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OK', 'NG'], true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  Logger.log('入金列（' + col + '列目）にプルダウン設定完了');
}

function testResend() {
  sendViaResend_(
    CONFIG.contactEmail,
    '【テスト】Resend API 送信テスト',
    'このメールは Resend API のテスト送信です。\n受信できていれば正常に動作しています。'
  );
  Logger.log('テスト送信完了');
}

function checkStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var msg = 'トリガー数: ' + triggers.length + '\n';
  triggers.forEach(function(t) {
    msg += '- ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')\n';
  });
  Logger.log(msg);
}
