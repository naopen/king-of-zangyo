# Task 1 報告

## 実装
- `notices.js`: 即時関数で `KingOfZangyoNotices`（CHANGE_NOTICE_ID, shouldOpenGuideOnInstalled, pickNotice, withDismissed）を返す。先頭コメントに設計書パス。
- `tests/notices.test.js`: ブリーフどおり13件。
- `manifest.json`: content_scripts.js を `["annual-pace.js", "notices.js", "content.js"]` に。

## TDD
- RED: `make test` -> `tests/notices.test.js` が `Cannot find module '../notices.js'` で失敗（tests 50, pass 49, fail 1）。想定どおり。
- GREEN: `make test` -> tests 62, pass 62, fail 0（既存49 + 新規13）。

## 変更ファイル
notices.js, tests/notices.test.js, manifest.json

## コミット
60d3020 吹き出しの出し分けを決める notices.js を足した

## セルフレビュー
- pickNotice は enabled/processing を先に判定、取得履歴なしなら usage（dismissed 無視）、履歴ありで未dismissなら change。
- withDismissed は新しい配列を返し元を変えない。
- 懸念なし。
