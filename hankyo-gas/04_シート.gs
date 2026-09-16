/**
 * 04_シート.gs … スプレッドシートへの書き込み
 *
 * 列は「列番号の決め打ち」ではなくヘッダー行の見出し名で探す。
 * 途中に列を挿しても壊れない。
 */

function シートを取得_() {
  const ss = SpreadsheetApp.openById(設定.スプレッドシートID);
  if (設定.シートGID) {
    const s = ss.getSheets().find(x => x.getSheetId() === 設定.シートGID);
    if (s) return s;
  }
  if (設定.シート名) {
    const s = ss.getSheetByName(設定.シート名);
    if (s) return s;
  }
  throw new Error('シートが見つかりません。設定.シートGID か 設定.シート名 を確認してください。'
    + ' 候補: ' + ss.getSheets().map(x => x.getName() + '(gid=' + x.getSheetId() + ')').join(', '));
}

/** 見出し名 → 列番号（1始まり）の辞書を作る。見出しの空白は無視する。 */
function 列番号を解決_(sheet) {
  const 幅 = Math.max(sheet.getLastColumn(), 1);
  const 見出し = sheet.getRange(設定.ヘッダー行, 1, 1, 幅).getValues()[0];
  const 辞書 = {};
  見出し.forEach((v, i) => {
    const 名 = String(v).replace(/\s/g, '');
    if (名 && !(名 in 辞書)) 辞書[名] = i + 1;
  });
  return 辞書;
}

/** 管理列（メールアドレス〜メッセージID）が無ければ右端に作る。 */
function 管理列を用意_(sheet) {
  let 辞書 = 列番号を解決_(sheet);
  const 不足 = 管理列.filter(名 => !(名 in 辞書));
  if (不足.length) {
    const 開始 = sheet.getLastColumn() + 1;
    if (開始 + 不足.length - 1 > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(),
        開始 + 不足.length - 1 - sheet.getMaxColumns());
    }
    sheet.getRange(設定.ヘッダー行, 開始, 1, 不足.length).setValues([不足]);
    Logger.log('管理列を追加しました: %s', 不足.join(', '));
    辞書 = 列番号を解決_(sheet);
  }
  return 辞書;
}

/** 既に取り込んだGmailメッセージIDの集合。二重取込の防止はここが要。 */
function 取込済IDを集める_(sheet, 列) {
  const 集合 = new Set();
  const 列番号 = 列['メッセージID'];
  const 最終行 = sheet.getLastRow();
  if (!列番号 || 最終行 < 設定.データ開始行) return 集合;
  sheet.getRange(設定.データ開始行, 列番号, 最終行 - 設定.データ開始行 + 1, 1)
    .getValues()
    .forEach(r => { const v = String(r[0]).trim(); if (v) 集合.add(v); });
  return 集合;
}

/** 反響1件を最終行の下に書き足す。 */
function 行を追加_(sheet, 列, 反響) {
  const tz = 設定.タイムゾーン;
  const 日付 = 設定.日付を文字列で書く
    ? Utilities.formatDate(反響.日時, tz, 'yyyy/MM/dd')
    : 反響.日時;
  const 時刻 = Utilities.formatDate(反響.日時, tz, 'HH:mm')
    .replace(':', 設定.時刻の区切り);

  const 値 = {
    '反響日': 日付,
    '反響時間': 時刻,
    '反響経路': 反響.反響経路,
    '名前': 反響.名前,
    'フリガナ': 反響.フリガナ,
    '電話番号': 反響.電話番号,
    '反響物件名': 反響.物件名,
    'メールアドレス': 反響.メールアドレス,
    '問合せ内容': 反響.問合せ内容,
    '賃料': 反響.賃料,
    '所在地': 反響.所在地,
    '間取り': 反響.間取り,
    '返信案': 反響.返信案,
    'Gmailリンク': 反響.Gmailリンク,
    'メッセージID': 反響.メッセージID,
  };

  const 行番号 = Math.max(sheet.getLastRow() + 1, 設定.データ開始行);
  const 幅 = sheet.getLastColumn();
  const 一行 = new Array(幅).fill('');
  let 書いた = false;

  Object.keys(値).forEach(名 => {
    const c = 列[名];
    if (!c || c > 幅) return;
    一行[c - 1] = 値[名] === undefined ? '' : 値[名];
    書いた = true;
  });
  if (!書いた) throw new Error('ヘッダー行（' + 設定.ヘッダー行 + '行目）に見出しが見つかりません。設定.ヘッダー行 を確認してください。');

  // 「2026/09/01」が日付値に自動変換されないよう、書き込む前に書式を文字列にする
  if (設定.日付を文字列で書く && 列['反響日']) {
    sheet.getRange(行番号, 列['反響日']).setNumberFormat('@');
  }
  if (列['反響時間']) sheet.getRange(行番号, 列['反響時間']).setNumberFormat('@');

  sheet.getRange(行番号, 1, 1, 幅).setValues([一行]);
}
