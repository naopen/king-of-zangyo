# Task 7 / 7b report

Status: DONE (manual extension checks pending for owner)

## What
- Task 7: popup.html に `<a id="open-guide" class="guide-link">使い方を見る</a>` を最後の .setting-item の後・.footer の前に追加。popup.css に .guide-link（中央, #1565c0, underline, 14px, margin 4px 0 16px）。popup.js に setupGuideLink()（DOMContentLoadedで呼ぶ。preventDefault後 chrome.tabs.create で guide/index.html）。フッターは変更なし。
- Task 7b: background.js 新規（importScripts("notices.js") と onInstalled で shouldOpenGuideOnInstalled(reason) のとき guide を開く）。manifest.json に background.service_worker を追加（permissions 不変）。Makefile の zip に background.js を追加。

## Verification
- make test: 66 tests, pass 66, fail 0
- node --check background.js / popup.js: OK
- manifest.json parses; permissions = ["storage","activeTab"]; background = {"service_worker":"background.js"}
- make clean && make zip && unzip -l: background.js, notices.js, guide/index.html が含まれる

## Owner manual checks
1. chrome://extensions で拡張を削除 -> zipを展開したフォルダーを「パッケージ化されていない拡張機能を読み込む」 -> 使い方ページが1回だけ開く
2. 拡張カードの再読み込みを押す -> 何も開かない。サービスワーカーのコンソールにエラーなし
3. ツールバーのアイコンからポップアップを開き「使い方を見る」を押す -> 新しいタブで使い方ページが開く（カグカグ表示、入口と吹き出しの「詳しく見る」の飛び先も確認）

## Files
popup.html popup.css popup.js background.js manifest.json Makefile

## Concerns
- 実機の拡張読み込みは未実施。
- advisor() は呼んでいない（小規模のため）。
- 未追跡の .superpowers/ はコミットしていない。
