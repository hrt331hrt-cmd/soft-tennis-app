/**
 * ソフトテニス部 部活ノート用の簡易データ保存バックエンド
 *
 * 使い方（このファイルの内容をそのままコピーして使います）：
 * 1. Googleスプレッドシートを新規作成する（名前は何でもOK。例:「部活ノートDB」）
 * 2. メニューの「拡張機能」→「Apps Script」を開く
 * 3. エディタに最初から入っているコードを全部消して、このファイルの中身を
 *    すべてコピーして貼り付ける
 * 4. 上部の「保存」（フロッピーディスクのアイコン）を押す
 * 5. 右上の「デプロイ」→「新しいデプロイ」を押す
 * 6. 歯車アイコンから種類を選ぶ画面で「ウェブアプリ」を選ぶ
 * 7. 「次のユーザーとして実行」→「自分」
 *    「アクセスできるユーザー」→「全員」 を選ぶ
 * 8. 「デプロイ」を押す
 * 9. 初回は「承認が必要です」という画面が出るので、
 *    自分のGoogleアカウントを選び、「詳細」→
 *    「（プロジェクト名）に移動（安全ではないページ）」を押して進める
 *    ※これは自分で作ったスクリプトなので安全です
 * 10. 表示された「ウェブアプリのURL」をコピーして、Claudeに貼り付けて教えてください
 *
 * 注意：コードを後で修正した場合は、もう一度「デプロイ」→
 * 「デプロイを管理」→ 鉛筆アイコン →「新しいバージョン」を選んで
 * 再デプロイしないと、変更が反映されません（URLは変わりません）。
 */

const SHEET_NAME = "KV";

function doGet(e) {
  const key = e.parameter.key;
  const sheet = getSheet_();
  if (!key) {
    return jsonOutput_({ error: "key is required" });
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return jsonOutput_({ key: key, value: data[i][1] });
    }
  }
  return jsonOutput_({ key: key, value: null });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const key = body.key;
    const value = body.value;
    if (!key) {
      return jsonOutput_({ ok: false, error: "key is required" });
    }
    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, value]);
    }
    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["key", "value"]);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
