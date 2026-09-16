/**
 * 02_取込.gs … 実行の入り口
 *
 *  初期セットアップ()  最初に1回だけ手で実行する
 *  取込実行()          5分ごとにトリガーが呼ぶ本体
 *  試し取込()          シートに書かずに、何が入るかログで確認する
 */

/** 最初に1回だけ手で実行する。列の準備とトリガー設置をまとめて行う。 */
function 初期セットアップ() {
  const sheet = シートを取得_();
  管理列を用意_(sheet);
  トリガーを設置_();
  Logger.log('セットアップ完了。シート「%s」に管理列を用意し、5分ごとのトリガーを設置しました。',
             sheet.getName());
}

/** 時間主導トリガーから呼ばれる本体。 */
function 取込実行() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('前回の実行が終わっていないため、今回は見送りました。');
    return;
  }
  try {
    const 件数 = 取り込む_(false);
    Logger.log('取込完了：%s件', 件数);
  } catch (e) {
    Logger.log('取込中にエラー：%s\n%s', e.message, e.stack);
    if (設定.失敗通知先) {
      MailApp.sendEmail(設定.失敗通知先,
        '【反響取込】エラーが発生しました',
        e.message + '\n\n' + e.stack);
    }
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** シートを一切変更せず、取り込まれる内容だけをログに出す。導入前の確認用。 */
function 試し取込() {
  const 件数 = 取り込む_(true);
  Logger.log('（試し取込）新規は %s件でした。シートは変更していません。', 件数);
}

/**
 * 本体。
 * @param {boolean} 試しだけ true ならシートに書かず下書きも作らない
 * @return {number} 新規に取り込んだ件数
 */
function 取り込む_(試しだけ) {
  const sheet = シートを取得_();
  const 列 = 試しだけ ? 列番号を解決_(sheet) : 管理列を用意_(sheet);
  const 既存ID = 取込済IDを集める_(sheet, 列);

  // 新しい順に返ってくるので、古い順に並べ直してから追記する
  const メール一覧 = [];
  GmailApp.search(設定.Gmail検索, 0, 200).forEach(thread => {
    thread.getMessages().forEach(m => メール一覧.push({ message: m, thread: thread }));
  });
  メール一覧.sort((a, b) => a.message.getDate() - b.message.getDate());

  let 件数 = 0;
  メール一覧.forEach(({ message, thread }) => {
    const id = message.getId();
    if (既存ID.has(id)) return;

    const 反響 = 反響を読む_(message);
    if (!反響) return;  // 反響メール以外（広告・返信など）は無視

    反響.返信案 = 返信案を作る_(反響);

    if (試しだけ) {
      Logger.log('【新規】%s', JSON.stringify(反響, null, 2));
    } else {
      行を追加_(sheet, 列, 反響);
      既存ID.add(id);
      if (設定.下書きを作る) 下書きを作る_(反響);
      if (設定.取込済ラベルを付ける) 取込済にする_(thread);
    }
    件数++;
  });
  return 件数;
}

/** 5分ごとの時間主導トリガーを設置する（重複して作らない）。 */
function トリガーを設置_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === '取込実行')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('取込実行').timeBased().everyMinutes(5).create();
}

/** トリガーを止めたいとき用。 */
function トリガーを解除() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === '取込実行')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('自動実行を止めました。');
}

function 取込済にする_(thread) {
  let label = GmailApp.getUserLabelByName(設定.取込済ラベル);
  if (!label) label = GmailApp.createLabel(設定.取込済ラベル);
  thread.addLabel(label);
}
