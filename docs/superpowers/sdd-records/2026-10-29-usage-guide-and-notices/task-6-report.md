# Task 6 報告: 変更点の吹き出し（D2）

## コミット
- edf2c22 使い方の吹き出しの文言から読点を外した（オーナー追加依頼。「このボタンで、」の読点だけ削除）
- 9be4b9b 取得したことがある人に、凡例を指す変更点の吹き出しを一度だけ出す

## 変更（content.js のみ）
- renderNotices の "change" 分岐で renderChangeNotice() を呼ぶ。
- renderChangeNotice: 凡例が幅・高さとも 0 より大きいときは凡例、そうでなければ年間集計の表の直下に出す。位置は getBoundingClientRect の実寸差から計算（幅の決め打ちなし）。Task 5 のCSSクラスをそのまま流用。赤い点なし。
- markChangeNoticeDismissed(): 保存済み配列を直前に読み直し withDismissed(CHANGE_NOTICE_ID) で保存。「閉じる」と completeFetchAnnualOvertime（saveAnnualData の直後、try/catch でログのみ）の両方が使う。
- Task 5 の更新ボタンclickリスナーは触っていない（同一関数参照なので重複登録もされない）。変更点の吹き出しはこのリスナーでは消えない（取得完了で既読になるため問題なし）。

## 確認（ブラウザ 1500x800、test/extension-test.html。モックはlocalStorage）
1. 2026年度データあり・既読なし: 吹き出し1件、文面一致。bubble top 595.1 >= legend bottom 587.1（8px下）、left 996 = 凡例 left。凡例は隠れない。
   「閉じる」→ 吹き出し0件、sync に ["annual-pace-2.5.0"] 保存。再読込後も0件。
2. 凡例表示オフ（kingOfZangyoLegendEnabled=false）: 凡例は display:none。bubble top 575.0 = 表 bottom 567.0 + 8、left 996 = 表 left。
3. 2025年度のみ保存（2026年度なし）: 表は「未取得」、凡例の高さ0。bubble top 575.0（表の下）、変更点の文面、使い方の吹き出し・赤い点なし。
4. 一度も取得していない状態から完了: 取得そのもの（KING OF TIMEのfetch）はモックで走らせられないため、完了処理と同じ呼び出しを手で実行した。saveAnnualData → markChangeNoticeDismissed → updateAnnualDataDisplay。開始時は使い方の吹き出し1件、完了後は0件（変更点も出ない）、dismissed は ["annual-pace-2.5.0"]。completeFetchAnnualOvertime 自体を通した実行はしていない。
- 横スクロール: 吹き出しを外しても外さなくても documentElement.scrollWidth は 1883（元からのテストページの幅）。吹き出し(996-1246px)による増加なし。

## スクリーンショット
- 凡例の下: /Users/na-kannan/.claude/projects/-Users-na-kannan-work/d43c5620-60e4-4ee2-8fe4-09a5f32f85d3/tool-results/mcp-Claude_Browser-blob-1793323673636-6wpntb.jpg
- 表の下（未取得）: /Users/na-kannan/.claude/projects/-Users-na-kannan-work/d43c5620-60e4-4ee2-8fe4-09a5f32f85d3/tool-results/mcp-Claude_Browser-blob-1793323698432-xbiaf0.jpg
- 凡例オフのシナリオ2は数値のみで撮っていない（表アンカーは3の画像と同じ位置）。

## テスト
make test: 66 件すべて PASS（2回、読点の変更後と最終）。テストファイルの変更なし。

## 懸念
- 凡例の幅(約470px)が表より広いので、アンカー左端はそろうが吹き出し(250px)は凡例の左寄りに出る。仕様どおり。
- 設計書 docs/superpowers/specs/...design.md は私の作業前から未コミットの変更あり（触っていない）。plans の236行目には旧文言（読点あり・「色で示します」）が残っている。
- advisor() は未実行（このセッションでは呼ばずに完了）。
