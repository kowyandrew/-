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

/**
 * 見出し行が何行目かを探す。
 *
 * 行番号を設定で固定していると、台帳の上に行を足したり消したりしたときに
 * データ行を見出しとして読んでしまい、まったく違う列に書き込む事故になる。
 * 「反響日」と「反響経路」が並んでいる行を毎回探す。
 */
let _見出し行の記憶 = null;

function 見出し行_(sheet) {
  if (設定.ヘッダー行) return 設定.ヘッダー行;   // 固定したいときだけ

  const id = sheet.getSheetId();
  if (_見出し行の記憶 && _見出し行の記憶.id === id) return _見出し行の記憶.行;

  const 幅 = Math.max(sheet.getLastColumn(), 1);
  const 探す行数 = Math.min(10, Math.max(sheet.getLastRow(), 1));
  const 値 = sheet.getRange(1, 1, 探す行数, 幅).getValues();

  for (let i = 0; i < 値.length; i++) {
    const 行 = 値[i].map(v => String(v).replace(/\s/g, ''));
    if (行.indexOf('反響日') >= 0 && 行.indexOf('反響経路') >= 0) {
      _見出し行の記憶 = { id: id, 行: i + 1 };
      return i + 1;
    }
  }
  throw new Error('見出し行が見つかりません。'
    + 'シート「' + sheet.getName() + '」の上から10行の中に'
    + '「反響日」と「反響経路」が並んだ行がありません。');
}

/** データが始まる行。見出しの次の行。 */
function データ開始行_(sheet) {
  return 設定.データ開始行 || 見出し行_(sheet) + 1;
}

/** 見出し名 → 列番号（1始まり）の辞書を作る。見出しの空白は無視する。 */
function 列番号を解決_(sheet) {
  const 幅 = Math.max(sheet.getLastColumn(), 1);
  const 見出し = sheet.getRange(見出し行_(sheet), 1, 1, 幅).getValues()[0];
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
    sheet.getRange(見出し行_(sheet), 開始, 1, 不足.length).setValues([不足]);
    SpreadsheetApp.flush();  // 直後の getLastColumn が古い値を返さないように
    Logger.log('管理列を追加しました: %s', 不足.join(', '));
    辞書 = 列番号を解決_(sheet);
  }
  return 辞書;
}

/**
 * 既に取り込んだGmailメッセージIDの集合。二重取込の防止はここが要。
 *
 * 「メッセージID」列と「Gmailリンク」列の末尾のID、両方を見る。
 * 行を消すつもりでセルの中身だけ消してしまっても、
 * 片方が残っていれば二重取込にならない。
 */
function 取込済IDを集める_(sheet, 列) {
  const 集合 = new Set();
  const 最終行 = sheet.getLastRow();
  if (最終行 < データ開始行_(sheet)) return 集合;
  const 行数 = 最終行 - データ開始行_(sheet) + 1;

  const 読む = [
    { 名: 'メッセージID', 取り出す: v => v },
    { 名: 'Gmailリンク',  取り出す: v => (v.match(/([0-9a-f]{8,})\s*$/i) || ['', ''])[1] },
  ];

  読む.forEach(({ 名, 取り出す }) => {
    const c = 列[名];
    if (!c) return;
    sheet.getRange(データ開始行_(sheet), c, 行数, 1).getValues().forEach(r => {
      const v = 取り出す(String(r[0]).trim());
      if (v) 集合.add(v);
    });
  });
  return 集合;
}

/**
 * 台帳として使っている列（見出しのある列）だけを見て、最終行を決める。
 *
 * sheet.getLastRow() はシート全体の最終行を返すので、
 * 見出しの無い列に関係ないデータが残っていると、
 * 新しい行がその下（見た目には空白のはるか下）まで飛ばされてしまう。
 * それを避けるため、見出しのある範囲だけを見る。
 */
function 最終データ行_(sheet, 列) {
  const 空 = データ開始行_(sheet) - 1;
  const 下端 = sheet.getLastRow();
  if (下端 < データ開始行_(sheet)) return 空;

  const 右端 = Object.keys(列).reduce((m, k) => Math.max(m, 列[k]), 0);
  if (!右端) return 空;

  const 値 = sheet.getRange(データ開始行_(sheet), 1, 下端 - データ開始行_(sheet) + 1, 右端).getValues();
  for (let i = 値.length - 1; i >= 0; i--) {
    if (値[i].some(v => String(v).trim() !== '')) return データ開始行_(sheet) + i;
  }
  return 空;
}

/**
 * 反響1件を最終行の下に書き足す。
 *
 * 手入力（飛込・紹介・リピート）との併用が前提なので、
 * 自動で値を入れるのは上の対応表にある列だけ。
 * 返信日・追客1〜3・来店日などの手運用の列や、
 * その行に人が仕込んだ数式・入力規則の値には一切触らない。
 */
function 行を追加_(sheet, 列, 反響) {
  const tz = 設定.タイムゾーン;
  const 日付 = 設定.日付を文字列で書く
    ? Utilities.formatDate(反響.日時, tz, 設定.日付の書式 || 'yyyy/MM/dd')
    : 反響.日時;
  const 時刻 = Utilities.formatDate(反響.日時, tz, 'HH:mm')
    .replace(':', 設定.時刻の区切り);

  // 見出しを読み違えたまま書き込むと、まったく違う列にデータが散る。
  // 台帳として最低限必要な列が揃っていなければ、書く前に止める。
  const 欠け = ['反響日', '反響経路', '名前', 'メッセージID'].filter(名 => !列[名]);
  if (欠け.length) {
    throw new Error('見出し行（' + 見出し行_(sheet) + '行目）に '
      + 欠け.join('・') + ' が見つかりません。見出しの文字が変わっていないか確認してください。');
  }

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
    'Gmailリンク': 反響.Gmailリンク,
    'メッセージID': 反響.メッセージID,
  };

  const 行番号 = Math.max(最終データ行_(sheet, 列) + 1, データ開始行_(sheet));
  if (行番号 > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 行番号 - sheet.getMaxRows() + 20);
  }
  const 幅 = sheet.getLastColumn();

  // その行に既に何か入っていれば残したまま、担当する列だけ差し替える
  const 一行 = sheet.getRange(行番号, 1, 1, 幅).getValues()[0];
  let 書いた = false;

  Object.keys(値).forEach(名 => {
    const c = 列[名];
    if (!c || c > 幅) return;
    一行[c - 1] = 値[名] === undefined ? '' : 値[名];
    書いた = true;
  });
  if (!書いた) throw new Error('見出し行に書き込める列が1つもありません。');

  // 「2026/09/01」が日付値に自動変換されないよう、書き込む前に書式を文字列にする
  if (設定.日付を文字列で書く && 列['反響日']) {
    sheet.getRange(行番号, 列['反響日']).setNumberFormat('@');
  }
  if (列['反響時間']) sheet.getRange(行番号, 列['反響時間']).setNumberFormat('@');

  sheet.getRange(行番号, 1, 1, 幅).setValues([一行]);
}
