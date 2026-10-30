# Task 2 報告: 使い方ページ

## 実装したこと
- `guide/index.html` `guide/guide.css` `guide/guide.js` をモックから分離して作成。Google Fontsのlinkを外し、書体はシステムフォントのスタックにした。カグカグ5枚は `../kagkag/<名前>.png` に変更。インラインscriptは `guide.js` に移した。文面・表・配置はモックのまま。
- `store-assets/guide-shots.html`: モックの `crop.html` を写し、参照先を `../test/extension-test.html` と `../popup.html` に変更。先頭に書き出し手順のコメントを追加。
- `guide/img/shot-*.png` 4枚を、一時フォルダー（rsync、ポート8741、Chrome headless 倍率2、窓1300x700）で再生成してマゼンタを切り落とした。
- `manifest.json` に `web_accessible_resources` を足した（`guide/index.html` のみ）。
- `tests/guide.test.js`: id 6つ、src無しscriptなし、参照ファイルの存在（`../kagkag/` は `kagkag/` が無ければ飛ばす）、`2.6.0` なし。

## TDD
- RED: `make test` で `ENOENT ... guide/index.html`、tests 63 / pass 62 / fail 1。
- GREEN: 実装後 `make test` で tests 66 / pass 66 / fail 0。

## 画像の照合
モックの既存PNGとのRMSE: monthly 0、annual 0、annual-unfetched 0（完全一致、サイズも同じ）、popup 0.04（サイズ同じ。入力欄内の数字の位置が数pxずれる程度）。popupは今回の生成物を採用。

## 表示確認（headless Chrome、iframe幅で測定）
- 幅1280: scrollWidth=1280、innerWidth=1280、年=2026、フッター「© 2026 かんなん」
- 幅375: scrollWidth=375、innerWidth=375、年=2026
- 1280pxのスクリーンショットを目視。レイアウトはモックどおり。カグカグはTask 3まで未表示。

## 変更ファイル
guide/index.html, guide/guide.css, guide/guide.js, guide/img/shot-{monthly,annual,annual-unfetched,popup}.png, store-assets/guide-shots.html, manifest.json, tests/guide.test.js

## 自己レビュー・懸念
- `guide-shots.html` と `guide.css` の変換にpython/sedを使った（日本語を含む `guide-shots.html` のコメントはpython経由で書いた。内容は確認済み）。
- `make zip` の対象に `guide/` は入っていない（Makefileは触っていない）。配布zipに含める作業は別Taskで必要になる可能性がある。
- popupのPNGはモックと完全一致ではない。
- 画像の `alt_06.png` 等は Task 3 待ち。
