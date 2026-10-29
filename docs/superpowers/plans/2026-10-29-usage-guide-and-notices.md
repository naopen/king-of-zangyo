# 使い方の案内と変更の知らせ 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 年間残業時間を取得していない人と、v2.5.0の変更に気づいていない人に、タイムカード画面の吹き出しで知らせ、拡張に同梱する使い方ページへ案内する。

**Architecture:** どの吹き出しを出すかの判定は、`annual-pace.js` と同じ形の純粋なモジュール `notices.js`（グローバル名 `KingOfZangyoNotices`）に分け、`node:test` で試す。画面への描画と保存は `content.js` が行う。使い方ページは `guide/` に置く拡張内のページで、`web_accessible_resources` で公開し、タイムカード画面からはリンク、ポップアップからは `chrome.tabs.create` で開く。

**Tech Stack:** Chrome拡張 Manifest V3、素のJavaScript、`node --test`、テスト用HTML（`test/` サブモジュール）、ヘッドレスのChrome（画像の書き出し）。

**Spec:** `docs/superpowers/specs/2026-10-29-usage-guide-and-notices-design.md`

## Global Constraints

- バージョンは `2.6.0`（`manifest.json`）
- 更新ボタンのラベルは `更新`（今の `最新データで更新` から変える）
- 吹き出しの文面はカグカグの口調にしない
- 吹き出しの見た目は1pxの枠と淡い地色。枠 `#b7c7d6`、地 `#f4f8fb`、文字13px。赤い点は `#d32f2f`
- 使い方の吹き出しの「詳しく見る」は `guide/index.html#annual-fetch`、変更点の吹き出しは `guide/index.html#annual-colors` を開く
- 使い方ページのフッターは「KING OF TIME 公式とは関係のない、非公式の拡張機能です。」と「© （開いた時点の年） かんなん」。バージョン番号は出さない
- 拡張のページはMV3のCSPでインラインのスクリプトが動かないので、年を入れる処理は外部ファイル `guide/guide.js` に書く
- 使い方ページの書体は `-apple-system, BlinkMacSystemFont, "Hiragino Sans", "Noto Sans JP", sans-serif` のシステムの書体にする。Google Fontsは読まない（拡張のページから毎回外へ読みにいくのを避ける。モックの `<link>` は外す）
- `web_accessible_resources` に入れるのは `guide/index.html` だけ。ページの中の画像やCSS、`kagkag/` は、拡張のページから同じオリジンで読めるので入れない（入れるとKING OF TIMEのページからカグカグの画像を取れてしまう）
- 保存キーは既存に合わせて `kingOfZangyo` で始める。閉じた記録は `chrome.storage.sync` の `kingOfZangyoDismissedNotices`（文字列の配列）
- 変更点の吹き出しの識別子は `annual-pace-2.5.0`
- コミットメッセージは1行の日本語。`Co-Authored-By` は付けない。`git add` はパスを名指しする
- 非公開リポジトリの作成と、そこへのpushは、そのつどNaokiに確認してから行う。この計画のどのタスクも、作業ブランチを勝手にpushしない
- テストは `make test`（`node --test "tests/*.test.js"`）。ディレクトリを渡すとNode 22で失敗する

## Review Focus

- 残業時間表示がオフのとき：どちらの吹き出しも出さない（Task 1でテスト）
- 年間データの取得中（月ごとにページが切り替わっている途中）：吹き出しを出さない（Task 1でテスト）
- 前の年度だけ取得したことがあり、今年度は「未取得」の人：使い方の吹き出しを出さない（Task 1でテスト）
- 年間の凡例が無いとき（閾値凡例表示がオフ、または前の年度だけ取得していて今年度が「未取得」）：変更点の吹き出しは年間集計の表の下を指す（Task 6で確かめる）
- 年度のプルダウンを切り替えて年別データが描き直されたとき：吹き出しが古い位置に取り残されない（Task 5、Task 6で確かめる）

---

### Task 1: 吹き出しの出し分けを決める `notices.js`

**Files:**
- Create: `notices.js`
- Test: `tests/notices.test.js`
- Modify: `manifest.json`（`content_scripts.js` を `["annual-pace.js", "notices.js", "content.js"]` にする）

**Interfaces:**
- Produces: `KingOfZangyoNotices.CHANGE_NOTICE_ID`（`"annual-pace-2.5.0"`）、`KingOfZangyoNotices.pickNotice({ enabled: boolean, processing: boolean, savedYears: number[], dismissed: string[] }) -> "usage" | "change" | null`、`KingOfZangyoNotices.withDismissed(dismissed: string[], id: string) -> string[]`（重複させず足した新しい配列を返す）

- [ ] **Step 1: 失敗するテストを書く**

```js
const { test } = require("node:test");
const assert = require("node:assert/strict");
const { pickNotice, withDismissed, CHANGE_NOTICE_ID } = require("../notices.js");
const base = { enabled: true, processing: false, savedYears: [], dismissed: [] };

test("一度も取得していなければ使い方", () => assert.equal(pickNotice(base), "usage"));
test("使い方は閉じた記録があっても出す", () => assert.equal(pickNotice({ ...base, dismissed: [CHANGE_NOTICE_ID] }), "usage"));
// 判定は今の年度を見ない。前の年度だけ取得した人も、ここに入る
test("取得したことがあり、閉じていなければ変更点", () => assert.equal(pickNotice({ ...base, savedYears: [2025] }), "change"));
test("変更点を閉じていれば何も出さない", () => assert.equal(pickNotice({ ...base, savedYears: [2026], dismissed: [CHANGE_NOTICE_ID] }), null));
test("表示オフなら何も出さない", () => assert.equal(pickNotice({ ...base, enabled: false }), null));
test("取得中なら何も出さない", () => assert.equal(pickNotice({ ...base, processing: true }), null));
test("withDismissed は重複させない", () => assert.deepEqual(withDismissed(["a"], "a"), ["a"]));
test("withDismissed は元の配列を変えない", () => { const d = []; withDismissed(d, "a"); assert.deepEqual(d, []); });
```

- [ ] **Step 2: 失敗を確かめる**

Run: `make test`
Expected: `notices.test.js` が `Cannot find module '../notices.js'` で失敗する

- [ ] **Step 3: `notices.js` を書く**

`annual-pace.js` と同じく即時関数でオブジェクトを返し、末尾で `if (typeof module !== "undefined") module.exports = KingOfZangyoNotices;` とする。先頭のコメントに設計書のパスを書く。

- [ ] **Step 4: 通ることを確かめる**

Run: `make test`
Expected: 既存の49件と新しい8件がすべてPASS

- [ ] **Step 5: コミット**

```bash
git add notices.js tests/notices.test.js manifest.json
git commit -m "吹き出しの出し分けを決める notices.js を足した"
```

### Task 2: 使い方ページ

**Files:**
- Create: `guide/index.html`、`guide/guide.css`、`guide/guide.js`
- Create: `guide/img/shot-monthly.png`、`shot-annual.png`、`shot-annual-unfetched.png`、`shot-popup.png`
- Create: `store-assets/guide-shots.html`（切り抜き画像を作るページ）
- Modify: `manifest.json`（`web_accessible_resources` に `{ "resources": ["guide/index.html"], "matches": ["https://s2.ta.kingoftime.jp/*"] }` を足す）
- Test: `tests/guide.test.js`

**Interfaces:**
- Produces: ページ内の見出しの id は `about`、`monthly`、`annual-fetch`、`annual-colors`、`settings`、`patterns`。カグカグの画像は `../kagkag/<名前>.png` で参照する（Task 3で用意する）

- [ ] **Step 1: 失敗するテストを書く**

`tests/guide.test.js` で `guide/index.html` を文字列で読み、次を確かめる。
- 6つの id がすべてある
- `src` の無い `<script>` が無い（CSPで止まるため）
- `<img src>` と `<link href>` と `<script src>` が指すファイルがすべて存在する。ただし `../kagkag/` で始まるものは、`kagkag/` が無ければ飛ばす（サブモジュールを取得していない環境でも `make test` が通るように）
- 本文に `2.6.0` が無い

- [ ] **Step 2: 失敗を確かめる**

Run: `make test`
Expected: `guide/index.html` が無いので失敗する

- [ ] **Step 3: 切り抜き画像を作る**

壁打ちのモックの切り抜きページ（`docs/superpowers/mocks/2026-10-29-guide/crop.html`）を `store-assets/guide-shots.html` に写し、参照先を `../test/extension-test.html` と `../popup.html` に直す。先頭のコメントに、`screenshot-compose.html` と同じ書き出し手順（一時フォルダーに写して配信し、`--force-device-scale-factor=2` で撮り、周りのマゼンタを切り落とす）を書く。4枚を `guide/img/` に書き出す。ボタン名は Task 4 が終わるまで画面の上で「更新」に差し替える（`screenshot-compose.html` の `applyPlannedUi` と同じ扱い。Task 8 で外す）。

- [ ] **Step 4: ページを書く**

モック（`docs/superpowers/mocks/2026-10-29-guide/index.html`）を元に、`<style>` を `guide/guide.css` に、年を入れる処理を `guide/guide.js` に分ける。Google Fontsの `<link>` は外し、書体はGlobal Constraintsのとおりにする。文面と表はモックのまま使う。モックはNaokiが「すごく良さそう」と評価したものなので、文面を書き換えるときは設計書と `README.md` に照らして事実だけを直す。

- [ ] **Step 5: 通ることを確かめる**

Run: `make test`
Expected: すべてPASS

- [ ] **Step 6: 表示を確かめる**

一時フォルダーに写して配信し、`guide/index.html` を幅1280pxと375pxで開く。横にはみ出さない（`document.documentElement.scrollWidth === innerWidth`）。フッターに現在の年が出る。カグカグは Task 3 までは表示されなくてよい。

- [ ] **Step 7: コミット**

```bash
git add guide store-assets/guide-shots.html manifest.json tests/guide.test.js
git commit -m "拡張に同梱する使い方ページを足した"
```

### Task 3: カグカグの画像を非公開のサブモジュールにする

**Files:**
- Create: 非公開リポジトリ `naopen/king-of-zangyo-kagkag`（中身は使う5枚と、出所を書いた `README.md`）
- Modify: `.gitmodules`（`path = kagkag`）
- Modify: `Makefile`

**Interfaces:**
- Consumes: Task 2 の参照名
- Produces: `kagkag/new-pose10.png`（先生）、`kagkag/new-pose21.png`（飛び回る）、`kagkag/new-pose15.png`（びっくり）、`kagkag/alt_06.png`（首をかしげる）、`kagkag/new-pose04.png`（応援）

- [ ] **Step 1: 画像をそろえる**

5枚は Confluence「03_カグカグ利用レギュレーション」から落としたもの。壁打ちで `~/Downloads/` に落としてある（`new-pose10.png` `new-pose21.png` `new-pose15.png` `new-pose04.png` は `~/Downloads/` 直下、`alt_06.png` は `~/Downloads/kagkag_png/`）。無ければ、VPNにつないだ仕事用のChromeで `https://kddi-agdc.atlassian.net/wiki/download/attachments/245564266/03_/<ファイル名>` を開くと落ちる（`alt_06.png` はページの「kagkag_png.zip」に入っている）。`README.md` には、出所のページのURL、落とした日、ファイル名とポーズの対応、Naokiの判断（マーケには相談しない、拡張から取り出せることは許容する）を書く。

- [ ] **Step 2: Naokiに確認してから、非公開リポジトリを作ってpushする**

Run（確認のあとで）: `gh repo create naopen/king-of-zangyo-kagkag --private` のあと、画像と `README.md` をコミットしてpushする
Expected: `gh repo view naopen/king-of-zangyo-kagkag --json visibility` が `PRIVATE`

- [ ] **Step 3: サブモジュールとして取り込む**

Run: `git submodule add https://github.com/naopen/king-of-zangyo-kagkag.git kagkag`
Expected: `kagkag/` に5枚と `README.md`。`make test` の `guide.test.js` が `../kagkag/` の参照も含めてPASS

- [ ] **Step 4: `make zip` に足し、画像が欠けていれば止める**

`zip` の前に、`kagkag/` の5枚がそろっているかを見る `check-assets` を置き、1枚でも無ければ「kagkag サブモジュールを取得してください（git submodule update --init）」と出して失敗させる。`zip` に `notices.js guide/ kagkag/*.png` を足す。`kagkag/README.md` と `.git` は入れない。

- [ ] **Step 5: 止まることと、入ることを確かめる**

Run: `mv kagkag/new-pose10.png /tmp/ && make clean && make zip; mv /tmp/new-pose10.png kagkag/`
Expected: `make zip` が失敗し、zipができない
Run: `make clean && make zip && unzip -l king-of-zangyo.zip`
Expected: `notices.js`、`guide/` 一式、`kagkag/` の5枚が入り、`README.md` と `.git` は入っていない

- [ ] **Step 6: コミット**

```bash
git add .gitmodules kagkag Makefile
git commit -m "カグカグの画像を非公開のサブモジュールにし、make zip に入れた"
```

### Task 4: 更新ボタンの名前と、タイムカード画面の入口（G7）

**Files:**
- Modify: `content.js`（`injectAnnualDataSection` の更新ボタン、`injectDialogStyles`）
- Modify: `test/extension-test.html` ほかテスト用HTML（`chrome.runtime.getURL` のモックを足す。サブモジュール側でコミットする）
- Modify: `README.md`（51行目）、`store-assets/STORE_LISTING.txt`（54行目）

**Interfaces:**
- Produces: 入口の要素は `a.kot-zangyo-guide-link`。`href` は `chrome.runtime.getURL("guide/index.html")`、`target="_blank"`、`rel="noopener"`、`title="使い方"`。更新ボタンの直後に置く。Task 5 は更新ボタンを `#king-of-zangyo-annual-section button.htBlock-buttonPrimary` で探す

- [ ] **Step 1: テスト用HTMLのモックに `getURL` を足す**

`window.chrome.runtime` に `getURL: (p) => "../" + p` を足し、`test/` でコミットする（`git -C test commit`）。テスト用HTMLが複数あるので、すべてに足す。

- [ ] **Step 2: ボタン名を「更新」にし、入口を足す**

見た目はモックのG7（灰色の角丸の枠、塗りつぶした丸の中に「?」、その右に「使い方」）。年別データの行は `align-items: baseline` なので、入口も行の中央には寄せず、ベースラインでそろえたうえで `position: relative; top: -1px` にする。

- [ ] **Step 3: 位置を測る**

テスト用HTMLを開き、更新ボタン、年度の選択、入口の `getBoundingClientRect()` の上下の中心を比べる
Expected: 入口と更新ボタンの差が1px以内。年別データの行の右端が、年の凡例の右端を越えない

- [ ] **Step 4: 説明文を直す**

`README.md` 51行目と `STORE_LISTING.txt` 54行目の「『最新データで更新』を押すと」を「『更新』を押すと」にする。

- [ ] **Step 5: テストとコミット**

Run: `make test`
Expected: すべてPASS

```bash
git add content.js README.md store-assets/STORE_LISTING.txt test
git commit -m "更新ボタンの名前を「更新」にし、横に使い方ページへの入口を足した"
```

### Task 5: 使い方の吹き出し（C1）

**Files:**
- Modify: `content.js`、`injectDialogStyles`

**Interfaces:**
- Consumes: `KingOfZangyoNotices.pickNotice`（Task 1）、更新ボタンの探し方（Task 4）
- Produces: `renderNotices()`（引数なし、Promise）。`updateAnnualDataDisplay(annualData, fiscalYear)`（`content.js` 1244行目付近）の最後で呼ぶ。この関数は、初回の表示と年度の切り替えの両方で通り、年の凡例もここで描き直される。呼ぶたびに前の吹き出しと赤い点を消してから描く

- [ ] **Step 1: `renderNotices()` を書く**

保存データから `enabled`（`kingOfZangyoEnabled`、無ければ true）、`processing`（`loadProcessingState()` の `isProcessing`）、`savedYears`（`getSavedFiscalYears()`）、`dismissed`（`kingOfZangyoDismissedNotices`、無ければ `[]`）を読み、`pickNotice` に渡す。`"usage"` なら、更新ボタンに赤い点（直径10px、白い2pxの縁）を付け、ボタンの直下に吹き出しを出す。文面は「このボタンで、年度の残業時間を集計できます。年360hに収まりそうかを色で示します。」。下に「詳しく見る」（`guide/index.html#annual-fetch` を新しいタブで開く）と「閉じる」。「閉じる」と更新ボタンのクリックで、吹き出しと赤い点を消す。閉じたことは保存しない（次に開けばまた出る）。

- [ ] **Step 2: 確かめる**

テスト用HTMLで、保存データの `king-of-zangyo-mock-sync` を `{}` にして開く
Expected: 赤い点と吹き出しが出る。「閉じる」で消え、開き直すとまた出る。年度を切り替えても吹き出しが1つだけ出る。表示をオフにすると出ない
保存データに年度を1つ入れて開く
Expected: 使い方の吹き出しは出ない

- [ ] **Step 3: コミット**

```bash
git add content.js
git commit -m "年間残業時間を取得していない人に、更新ボタンを指す吹き出しを出す"
```

### Task 6: 変更点の吹き出し（D2）

**Files:**
- Modify: `content.js`（`renderNotices`、`completeFetchAnnualOvertime`）

**Interfaces:**
- Consumes: `pickNotice`、`withDismissed`、`CHANGE_NOTICE_ID`（Task 1）、`renderNotices`（Task 5）

- [ ] **Step 1: `"change"` のときの描画を足す**

年の凡例（`#king-of-zangyo-annual-legend`）の直下に、凡例を隠さずに吹き出しを出す。凡例が無いときは年間集計の表の直下を指す。凡例が無いのは、閾値凡例表示がオフのときと、表示中の年度が「未取得」のとき（前の年度だけ取得した人は、今年度を開くとこの状態になる）。どちらの場合も吹き出しは出す。出さないと、前の年度だけ取得した人には一度も届かないため。位置は要素の実寸から決め、幅を決め打ちしない（凡例は1行で約472px）。文面は「色の境目が、年度の残りに応じて動くようになりました。境目は「更新」を押した日を基準に計算します。」。「詳しく見る」は `guide/index.html#annual-colors`。「閉じる」で `withDismissed` した配列を `kingOfZangyoDismissedNotices` に保存して消す。

- [ ] **Step 2: 取得の完了で閉じた扱いにする**

`completeFetchAnnualOvertime` で年間データを保存したあと、同じく `CHANGE_NOTICE_ID` を `kingOfZangyoDismissedNotices` に足す。

- [ ] **Step 3: 確かめる**

テスト用HTMLで、年間データあり・閉じた記録なしで開く
Expected: 凡例の下に吹き出しが出て、凡例は隠れない。「閉じる」のあと開き直すと出ない
閾値凡例表示をオフにして開く
Expected: 年間集計の表の下に出る
保存データを `{ kingOfZangyoAnnualDataYears: [2025], kingOfZangyoAnnualData_2025: {…} }`（2026年度は無し）にして開く
Expected: 今年度（2026年度）は「未取得」の表で、その下に変更点の吹き出しが出る。使い方の吹き出しは出ない
一度も取得していない状態から取得を完了する
Expected: 完了のあと、変更点の吹き出しは出ない

- [ ] **Step 4: テストとコミット**

Run: `make test`
Expected: すべてPASS

```bash
git add content.js
git commit -m "取得したことがある人に、凡例を指す変更点の吹き出しを一度だけ出す"
```

### Task 7: ポップアップの「使い方を見る」

**Files:**
- Modify: `popup.html`、`popup.css`、`popup.js`

- [ ] **Step 1: リンクを足す**

設定の項目の下、フッターの上に、中央寄せの文字リンク「使い方を見る」を置く。クリックで `chrome.tabs.create({ url: chrome.runtime.getURL("guide/index.html") })`。フッターの2行（thanks to と developed by）は残す（消すのはストア画像の中だけ）。

- [ ] **Step 2: 確かめる**

`make zip` したものを「パッケージ化されていない拡張機能を読み込む」で入れ、アイコンを押す
Expected: 「使い方を見る」で使い方ページが新しいタブで開く。カグカグも表示される。タイムカード画面の入口と2つの吹き出しの「詳しく見る」も、それぞれの節に飛ぶ

- [ ] **Step 3: コミット**

```bash
git add popup.html popup.css popup.js
git commit -m "設定画面に使い方ページを開くリンクを足した"
```

### Task 8: バージョンとストア画像

**Files:**
- Modify: `manifest.json`（`"version": "2.6.0"`）
- Modify: `store-assets/screenshot-compose.html`、`store-assets/guide-shots.html`（画面の上で足していた「実装前の仮」を消す）
- Modify: `store-assets/screenshot.png`、`guide/img/shot-*.png`
- Modify: `store-assets/STORE_LISTING.txt`（v2.6.0の行の下書き）

- [ ] **Step 1: 仮の部分を消して撮り直す**

`screenshot-compose.html` の `applyPlannedUi` と、`loadPopup` の「使い方を見る」を足す部分を消す。`guide-shots.html` のボタン名の差し替えも消す。書き出し手順どおりに撮り直す
Expected: `store-assets/screenshot.png` が1280×800。下書き `store-assets/drafts/screenshot-off-on.png` と見比べて、配置が同じで、入口と「使い方を見る」が拡張そのものの表示で写っている

- [ ] **Step 2: ストア掲載文の下書きを書き、Naokiに見せる**

v2.6.0の行を1行で書く。v2.5.0のときと同じく、Naokiの言葉に合わせて直してもらう前提の下書きにする。

- [ ] **Step 3: 全体を確かめる**

Run: `make test && make clean && make zip`
Expected: すべてPASSし、zipに `manifest.json` のバージョン2.6.0が入っている

- [ ] **Step 4: コミット**

```bash
git add manifest.json store-assets guide/img
git commit -m "バージョンを2.6.0にし、ストア画像と使い方ページの画像を実物で撮り直した"
```
