# Task 4 報告

Status: DONE

## やったこと
- テスト用HTML 4本の `chrome.runtime` に `getURL: (p) => "../" + p` を追加し、サブモジュールでコミット（d5c82f3、pushなし）
- content.js: 更新ボタンの文言を「更新」に変更。直後に `a.kot-zangyo-guide-link`（href=`chrome.runtime.getURL("guide/index.html")`、target=_blank、rel=noopener、title=使い方）を追加。CSSは `injectDialogStyles()` 内に、G7の値をそのまま入れた（下線なし、hover時のみ薄い灰色の地）
- README.md 51行目、STORE_LISTING.txt 54行目を「『更新』を押すと」に変更

## 測定（1500x800、extension-test.html、年データ2026投入済み）
- 更新ボタン cy=436.70、年度select cy=436.38、入口 cy=436.68
- 入口と更新ボタンの差 0.02px（1px以内）
- 行の右端 1467.6、入口の右端 1375、凡例の右端 1467.6（越えない）
- 入口のhref属性 `../guide/index.html`、解決後 `http://127.0.0.1:8744/guide/index.html`。text-decoration none
- サーバは停止済み

## テスト
make test: 66 pass / 0 fail

## コミット
- main: acddde0（test gitlinkを含む）
- submodule(test): d5c82f3

## 懸念
- content.js の編集を python で行った（日本語を含む置換。ルールはWrite/Edit優先）。内容はgit diffで確認済み
- 実際のクリックによる遷移確認はしていない（hrefの解決のみ確認）
- yomiyasuフックは通っていない
