/**
 * 設問タイトル、Item ID、および entry.ID（URLパラメータ用）を一括取得する
 * 権限エラー(401)対策版
 */
function listFormDetailedIds() {
  const form = FormApp.getActiveForm();
  if (!form) {
    console.error("フォームが見つかりません。フォームの編集画面から実行するか、FormApp.openById('ID') を使用してください。");
    return;
  }
  
  const formUrl = form.getPublishedUrl();
  const items = form.getItems();
  
  // 自分自身のログイン権限（OAuthトークン）をヘッダーに付けてHTMLを取得
  const options = {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(formUrl, options);
    const code = response.getResponseCode();
    
    if (code !== 200) {
      console.error(`HTMLの取得に失敗しました。ステータスコード: ${code}`);
      console.log("フォームの「設定」で「回答を1回に制限する」や「組織限定」を一時的にオフにすると、より確実に取得できる場合があります。");
      return;
    }
    
    const html = response.getContentText();
    console.log("--- フォーム詳細IDリスト 開始 ---");

    items.forEach((item, index) => {
      const title = item.getTitle();
      const itemId = item.getId();
      const type = item.getType();
      
      let entryId = "---";
      // HTML内の構造解析用正規表現
      const regex = new RegExp(`"${title}".*?\\[(\\d+),`);
      const match = html.match(regex);
      
      if (match && match[1]) {
        entryId = "entry." + match[1];
      } else {
        // 別パターンの検索
        const fallbackRegex = new RegExp(`\\["${title}".*?,(\\d+)\\]`, "s");
        const fallbackMatch = html.match(fallbackRegex);
        if (fallbackMatch) entryId = "entry." + fallbackMatch[1];
      }

      console.log(`${index + 1}. 【${type}】`);
      console.log(`   タイトル : ${title}`);
      console.log(`   Item ID  : ${itemId}`);
      console.log(`   entry.ID : ${entryId}`);
      console.log("--------------------------------");
    });

    console.log("--- フォーム詳細IDリスト 終了 ---");
    
  } catch (e) {
    console.error("エラーが発生しました: " + e.message);
  }
}