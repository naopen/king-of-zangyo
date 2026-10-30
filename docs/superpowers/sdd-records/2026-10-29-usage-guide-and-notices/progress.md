# SDD ledger — plan: docs/superpowers/plans/2026-10-29-usage-guide-and-notices.md

Spec: docs/superpowers/specs/2026-10-29-usage-guide-and-notices-design.md
Branch: feature/usage-guide-and-notices（main から 2026-10-29 に切った。start HEAD 7a205b1）
注意: このワークスペースは最後に消さない。Naokiのルール（no-delete-unrecorded）により、台帳と報告はプロジェクトの中へ移してコミットする。

## Pre-flight scan

| 対象 | 作る側 → 使う側 | 結果 |
|---|---|---|
| T1 × T7b | `shouldOpenGuideOnInstalled` を T1 が作り、T7b が `importScripts("notices.js")` で使う | 一致。notices.js は `module` が無ければ export しないので、サービスワーカーでも動く |
| T1 × T5 × T6 | `pickNotice` `withDismissed` `CHANGE_NOTICE_ID` | 一致 |
| T5 × T6 | `renderNotices()` を T5 が作り、T6 が "change" を足す | 一致。呼び出し場所は `updateAnnualDataDisplay` の最後 |
| T4 × T5 | 更新ボタンの探し方 `#king-of-zangyo-annual-section button.htBlock-buttonPrimary` | 一致 |
| T2 × T3 | 使い方ページが `../kagkag/<名前>.png` を参照し、T3 が用意する | 一致。T2 のテストは kagkag/ が無ければ飛ばす |
| T2 × T4 × T5/T6/T7 | `guide/index.html` と節の id（annual-fetch, annual-colors） | 一致 |
| manifest.json（T1, T2, T7b, T8） | content_scripts, web_accessible_resources, background, version を順に足す | 衝突なし（順に別の項目） |
| Makefile（T3, T7b） | T3 が zip に notices.js guide/ kagkag/ を、T7b が background.js を足す | T1〜T2 の間は zip に notices.js が入らない。下の Ruling |
| T2 × T8 / screenshot-compose.html × T8 | 画面の上で足した「実装前の仮」を T8 で外す | 一致 |
| T1 自身 | テストは 8＋5＝13件、「既存49件と新しい13件」 | 一致 |
| T2 自身 | Create した guide/ をテストが読む。2.6.0 を本文に書かない | 一致 |
| T3 自身 | Step 2 は非公開リポジトリの作成とpush | 外へ出る操作。止まってNaokiに確認する（スキルの停止条件） |
| T4 自身 | test/ サブモジュールのHTMLを変えてサブモジュール側でコミット | push しない。親は未pushのサブモジュールのコミットを指すことになる |
| T7b 自身 | manifest に background を足し、permissions は増やさない | tabs.create が権限なしで動くかは Step 2 で実測 |

Ruling: T1〜T2 の間の `make zip` は notices.js を含まない — T3 で足す計画どおりにする。途中で zip を配らないので実害がない — 誤りなら、途中で zip を作った人が notices.js の無い拡張を得る
Ruling: T4 で test/ サブモジュールにコミットするが push しない — 計画の Global Constraints（push はNaokiに確認）に従う — 誤りなら、親のコミットが指すサブモジュールのコミットがリモートに無く、別の環境で取得できない（push時に解決）

## Progress
Task 1: minor (deferred): pickNotice は savedYears/dismissed が欠けると TypeError。Task 5 の呼び出し側で既定値 [] を必ず渡す
Task 1: complete (commits 7a205b1..60d3020, review clean)
Task 2: minor (deferred): guide/index.html の画像幅（style="width:320px" など）と色見本の色がインライン指定。モック由来で動作に影響なし
Task 2: minor (deferred): popup の切り抜き画像はモックの画像と画素単位では一致しない（RMSE 0.04、入力欄の数字が数px動く）。新しく撮ったほうを採用
Task 2: complete (commits 60d3020..99edfe8, review clean)
Task 3: Naoki承認 2026-10-30「作ってpushしてよい」— 非公開リポジトリ naopen/king-of-zangyo-kagkag を作り、画像5枚とREADMEをpushする（本体の作業ブランチはpushしない）
Task 3: gh repo create がClaude Codeの自動判定で止められたため、Naokiが自分で実行してpushした。controller が PRIVATE と中身6件を確認
Task 3: minor (deferred): make zip に -x の除外が無く、guide/ に .DS_Store があれば入る
Task 3: complete (commits 99edfe8..8d66f0d, review clean)
Task 4: 実装役が content.js の書き換えに python を使った（日本語ファイルは Edit で書く決まりに反する）。レビューで余計な変化・文字化け・改行の変動が無いことを確認。以後の指示で Edit を使うよう強調する
Task 4: test サブモジュールに d5c82f3（getURL のモック）をコミット。push していない
Task 4: minor (deferred): .kot-zangyo-guide-link:hover の text-decoration: none は重複（ページ側の a:hover に勝つ保険として残す）
Task 4: complete (commits 8d66f0d..acddde0, review clean)
Task 5: Naokiの指示 2026-10-30「『色で示します』よりは『色で表示します』のほうがいい」— 使い方の吹き出しの文面を「…年360hに収まりそうかを色で表示します。」に変える（修正の回で反映）
Task 5: minor (deferred): storage.sync.get のコールバックで lastError を見ていない（失敗は catch でログされるだけ）
Task 5: minor (deferred): 更新ボタンの once リスナーが描画のたびに積み重なる。Task 6 で変更点の吹き出しも更新ボタンで消える可能性
Task 5: minor (deferred): 古い呼び出しが後に完了すると古い状態で描かれうる（重複はしない）
Task 5: minor (deferred): 吹き出し幅250px固定で、右端に近い画面幅だとはみ出す可能性
Task 5: review Approved。Naokiの文言の直しを fix round 1/5（1 addressed, 0 open; commits fa3952f..de152f9）で反映
Ruling: fix round 1 の再レビューは、差分が文字列1行（「色で示します」→「色で表示します」）だけなので、レビュー役を立てず controller が差分を直接確かめた — 1行の文字列の置き換えで判断の余地が無い — 誤りなら、見落としがあれば最終レビューで拾う
Task 5: complete (commits acddde0..de152f9, review clean)
Task 5: Naokiの指示 2026-10-30 — 文面を「このボタンで年度の残業時間を集計できます。年360hに収まりそうかを色で表示します。」（読点を外す）。content.js を編集中の Task 6 の実装役に、別コミットで直すよう依頼
Task 6: minor (deferred): markChangeNoticeDismissed の storage.sync.set で lastError を見ていない。書き込み失敗でも resolve し、次回また出る可能性
Task 6: minor (deferred): 吹き出しはリサイズに追従しない
Task 6: 取得完了の処理（completeFetchAnnualOvertime）そのものは通していない。呼び出しを手で再現して確認
Task 6: complete (commits de152f9..9be4b9b, review clean。読点を外す edf2c22 もこのレビューで確認)
Ruling: Task 7 と 7b を1回の依頼にまとめる。拡張を実際に読み込んで確かめる手順（chrome://extensions）は自動で操作できないので、最後にNaokiの実機確認としてまとめる — 小さく別ファイルの作業で、確認の手段が同じため — 誤りなら、実機で動かない不具合の発見が最後まで遅れる
Task 7: minor (deferred): 設定画面のリンクは <a href="#"> で実質ボタン（Enter で動き、# へは飛ばない）
Task 7/7b: complete (commits ed6490d..02ad830, review clean)。拡張を実際に読み込む確認は、Naokiの実機確認に回した
Task 8: 実装役が、書き出し用の2ページで吹き出しと赤い点を隠す処理を足した（実物の拡張が描くため）。レビューで適切と判定
Task 8: minor (deferred): 書き出し用の2ページで吹き出しを隠す処理が重複（独立した道具なので許容）
Task 8: minor (deferred): 実物の設定画面が下書きより約10px高く、ストア画像の OFF/ON が約10px下がった（リンクの下の余白16pxのため）。Naokiに見てもらう
Task 8: complete (commits 02ad830..2825a21, review clean)
Task 6: Naokiの指示 2026-10-30 — 変更点の吹き出しの文面を「背景色の切り替わりが年度の残り日数に応じて変化するようになりました。」の1文にし、下に「詳しく見る」。直しを別の実装役（haiku）に依頼
Final review (opus, a3d0f2d..2a9787f): With fixes。Important 1件（表示・凡例の切り替えで変更点の吹き出しが取り残される）、Minor 3件
Final review: 台帳の「once リスナーが積み重なる」は誤り（同じ関数参照なので addEventListener が重複登録しない）
Ruling: 最終レビューの Minor のうち、guide.test.js の判定と zip の .DS_Store 除外は Important と一緒に直す。「閉じる」と描画の競合、lastError の未確認は leave — 害が小さく、最悪でも次回1回余計に出るだけ — 誤りなら、まれに閉じた吹き出しがもう1回出る
Ruling: Task 5 の renderNotices() は updateAnnualDataDisplay の「最後」ではなく、途中の return（年間データが無いとき）を含むすべての出口で呼ぶ — 計画の文面どおりだと、使い方の吹き出しが要る未取得の年度で描かれない — 誤りなら、出口が増えたときに呼び忘れが起きる（try/finally で1か所にまとめるよう指示）
