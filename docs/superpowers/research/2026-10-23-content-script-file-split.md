# コンテンツスクリプトを複数ファイルに分けるときの書き方

2026-10-23 調査。年間残業の色分けの変更（`docs/superpowers/specs/2026-10-14-annual-pace-bands-design.md`）で `annual-pace.js` を新しく足すにあたり、ファイル間の名前の扱いを決めるための材料。

## 決めたいこと

`annual-pace.js` を、どう書けば「将来ほかの場所から使いたくなっても、名前がぶつかって動かなくなることがない」か。

## 背景

いまの拡張は `content.js` 1ファイル（2,386行）で、関数51個と定数・変数18個をすべてファイルの一番外側に並べている。構成の方針を決めた記録はリポジトリに無い。今回 `annual-pace.js` を足して `manifest.json` の `content_scripts.js` に `["annual-pace.js", "content.js"]` と並べると、2つのファイルの一番外側の名前が同じ場所に並ぶ。実際、`annual-pace.js` に作る予定の `parseDate` と同じ名前の関数（別物）が、`content.js` の679行目に既にある。

## 調べて分かったこと

### 同じ拡張のコンテンツスクリプトは、ファイルが違っても1つのグローバルを共有する

MDNの「Content scripts」に次の記述がある。Chromeの公式ドキュメントには、同じ拡張のファイル同士が共有するかの記述が無い。調査を任せたサブエージェントは、Chromiumのソース（`extensions/renderer/isolated_world_manager.h`）のコメントから、拡張のIDごとに1つの実行環境を作る作りだと読み取った（このソースは私は開いて確かめていない）。

> There is only one global scope per frame, per extension.
> This means that variables from a content script can be accessed by any other content scripts, regardless of how the content script was loaded.

https://developer.mozilla.org/en-US/docs/Mozilla/Add-ons/WebExtensions/Content_scripts

並べたファイルは配列の順に読み込まれる。

> An array of JavaScript file paths, injected in the order they appear in this array, after css files are injected.

https://developer.chrome.com/docs/extensions/reference/manifest/content-scripts

### 名前がぶつかったときの挙動（Chromeで実測）

2026-10-23、このセッションの組み込みブラウザ（Chromium）で、別々のscript要素を順に読み込んで確かめた。拡張の実機ではなく、通常のページで同じ条件を作った。

| 実験 | 結果 |
|---|---|
| 1つ目のファイルの関数が `parseDate()` を呼び、2つ目のファイルにも別の `function parseDate` がある | エラーは出ない。1つ目のファイルの関数が、2つ目のファイルの `parseDate` を呼んだ |
| 2つのファイルに同じ名前の `const` がある | 2つ目のファイルが `SyntaxError: Identifier 'LIMIT' has already been declared` で止まり、1行目も実行されなかった |
| 1つ目のファイルの中身を即時関数で閉じ込め、名前空間 `NS` だけを外に出す。2つ目のファイルに `function parseDate` がある | `NS` の中の関数は、中の `parseDate` をそのまま呼んだ |

前者2つは、ECMAScriptの仕様どおりの挙動。MDNの `let` の項に、別の `<script>` で `let` や `const` を再宣言するとエラーになると書かれている。

### コンテンツスクリプトでは、ES Modulesを素直には使えない

`manifest.json` の `content_scripts` には、モジュールとして読み込む指定が無い。指定できるのは `matches` `js` `css` `run_at` `world` などだけ。標準化の提案は2026-10-23時点でまだ開いている。

> Currently, content scripts do not support ES module format.

https://github.com/w3c/webextensions/issues/156 （状態はOpen、ラベルは `supportive: safari`、`needs-triage: chrome`、`needs-triage: firefox`）

同じ提案は、`import(chrome.runtime.getURL(...))` で読み込む回避策の欠点も挙げている。

> it requires adding imported script to the `web_accessible_resources` option in the manifest
> it breaks the `document_start` timing in `run_at` option

`web_accessible_resources` に登録したファイルは、そのサイトのページ側のスクリプトからも読めるようになる。

> Note that this also exposes the resources to any first-party or third-party scripts running on the same site.

https://developer.chrome.com/docs/extensions/develop/concepts/content-scripts

ほかの部品とは扱いが違う。service workerは `"type": "module"` で `import` 文が使える（ただし動的な `import()` は使えない）。popupは `<script type="module">` で読み込める。

> To use the `import` statement, add the `"type"` field to your manifest and specify `"module"`.
> Note that `import()`, often called a dynamic import, is not supported.

https://developer.chrome.com/docs/extensions/develop/concepts/service-workers/basics

### 一般のJavaScriptの考え方

グローバルに名前を並べないことは、ESLintが「ベストプラクティス」と書いている。対処として即時関数で包むか、モジュールにするかを挙げている。

> It is the best practice to avoid "polluting" the global scope with variables that are intended to be local to the script.
> Otherwise variables intended to be local to the script should be wrapped in an IIFE.
> Top-level declarations in ES modules and CommonJS modules create module-scoped variables.

https://eslint.org/docs/latest/rules/no-implicit-globals

今の主流はES Modulesで、即時関数はモジュールが使えない場面の手段という位置づけ。

> Always use modules (`import`/`export`) over a non-standard module system.
> Note that in a world with modules everywhere, you almost never need an IIFE.

https://github.com/airbnb/javascript （Airbnb JavaScript Style Guide 10.1 と 7.2）

Google JavaScript Style Guide も、新しいファイルは `goog.module` かES Modulesで書くとしている（https://google.github.io/styleguide/jsguide.html ）。

### 拡張の開発では、ES Modulesで書いて1ファイルにまとめるのが一般的

Chrome公式のサンプル（GoogleChrome/chrome-extensions-samples、2026-10-20のコミット）では、外部ライブラリを使うコンテンツスクリプトをRollupで1ファイルにまとめている（`functional-samples/ai.gemini-on-device-summarization/`）。サブエージェントが数えたところ、現行のサンプルのコンテンツスクリプトは12件とも1エントリ1ファイルで、複数ファイルを並べた例は旧版（MV2）のアーカイブに1件あるだけだった（件数は私は数え直していない）。

拡張のフレームワークのWXTも、コンテンツスクリプトをES Modulesのまま出力する機能はまだ無いと書いている。

> WXT does not yet include built-in support for bundling content scripts as ESM.

https://wxt.dev/guide/essentials/es-modules.html

Chrome Web Storeの規約は、ファイルを1つにまとめることを認めている（難読化は禁止）。

> Minification is allowed, including the following forms:

https://developer.chrome.com/docs/webstore/program-policies/code-readability

esbuildで `--format=iife --global-name=名前` を指定すると、中身を即時関数で包み、公開するものを1つのグローバル名に載せた形で出力する（https://esbuild.github.io/api/ ）。つまり下の案Bは、バンドラーが出力する形を手で書いたものに当たる。

### ビルド手順を入れることの是非

拡張に絞って比べた公式の資料は見つからなかった。一般論では両論ある。htmxのエッセイ「No Build Step」（2023-08-19）は、ビルドの無いコードは長く動き続けると主張する。Datastarのエッセイ「Yes, you want a build step」は、ビルドは一度払えばずっと効く費用だと主張する。どちらも特定プロジェクトの意見。

## 選択肢

「ほかの場所から使う」場面として、同じ拡張のpopupから使う、別のプロジェクトにファイルを写して使う、の2つを想定して比べた。

| 案 | 書き方 | 名前の衝突 | Nodeのテスト | ビルド | ほかの場所から使うとき |
|---|---|---|---|---|---|
| A | 今の `content.js` と同じく、一番外側に関数を並べる | 起きる。同じ名前の関数は黙って入れ替わる（実測） | `module.exports` の1行が別に要る | 不要 | 読み込み先の名前と1つでも重なると壊れる |
| B | 即時関数で閉じ込め、外に出す名前を1つだけにする。Nodeでは `module.exports` で渡す | 外に出る名前1つ以外は起きない（実測） | そのまま読める（2026-10-15 に計画のコードで46件通過） | 不要 | popupは `<script src>` で読めば使える。別のプロジェクトでも、ブラウザとNode（CommonJS）でそのまま読める。ES Modulesのプロジェクトからは読み込み方に工夫が要る |
| C | ES Modulesで書き、`content.js` から `import(chrome.runtime.getURL(...))` で読む | 起きない | そのまま読める | 不要 | どこからでも `import` できる。ただしファイルがKING OF TIMEのページ側からも読めるようになり、`content.js` の起動が非同期になる |
| D | ES Modulesで書き、esbuildなどで1ファイルにまとめる | 起きない | そのまま読める | 必要 | どこからでも `import` できる。ビルド手順と依存パッケージを保守し続けることになる |

案Bに残る弱点は2つある。1つ目は、外に出す名前そのもの（`KingOfZangyoAnnualPace`）がぶつかる可能性。拡張ごとに実行環境が分かれているので、ぶつかる相手は同じ拡張のファイルだけで、拡張に固有の長い名前にすれば実際には起きない。ただし `const` で宣言しているので、万一ぶつかると後ろのファイルが丸ごと止まる。2つ目は、閉じ込めた中の関数を、あとでグローバルに1つずつ出したくなった場合。たとえば `parseDate` を `window.parseDate` として出すと、`content.js` の `parseDate` とぶつかる。案Bが防げるのは、名前空間の中に置いている間だけ。

案CとDのES Modulesには、この2つ目の弱点が無い。名前を決めるのが出す側ではなく使う側だから。使う側が `import { parseDate as annualParseDate } from "./annual-pace.js"` のように、自分のファイルでぶつからない名前を付けて取り込める。「どこに出しても名前がぶつからない」を仕組みとして保証できるのは、ES Modulesだけ。

## 推奨

案Bを推奨する。

理由は3つ。1つ目に、Naokiの求める「外に出しても名前がぶつかって動かなくなることがない」を、ビルド手順を足さずに満たせる。中身の名前は外から見えないので、`content.js` に同じ名前の関数があっても影響しない（実測）。2つ目に、案CとDが前提にするES Modulesは、コンテンツスクリプトでは素直に使えない。案Cはファイルをページ側に公開することになり、案Dはこの拡張に初めてビルド手順を持ち込むことになる。3つ目に、案Bはバンドラーの出力と同じ形なので、将来案Dに移りたくなったときは、即時関数を外して `export` を付けるだけで済む。

`content.js` の書き方は今回は変えない。51個の名前は1ファイルの中で完結していて、今まで問題になっていないため。`annual-pace.js` の中身は外から見えないので、`content.js` の名前が増えても `annual-pace.js` は壊れない。

## 残っている不明点

- 拡張の実機（コンテンツスクリプトとして読み込んだ状態）での衝突の挙動は試していない。実装計画のTask 3で、拡張を読み込んだときに `content.js` から `KingOfZangyoAnnualPace` が見えるかを確かめる
- Chromeの公式ドキュメントには、同じ拡張のコンテンツスクリプトがグローバルを共有するという記述が無い。根拠はMDNとChromiumのソース
- 案Bのファイルを、`package.json` に `"type": "module"` を書いたプロジェクトに置くと、`.js` がES Modulesとして読まれて `module.exports` の分岐が働かない。いまのリポジトリに `package.json` は無いので影響しないが、案Dに移るときに併せて書き換える
