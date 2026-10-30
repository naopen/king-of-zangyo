# Task 9 report

Status: DONE_WITH_CONCERNS

## Changed
- notices.js: pickNotice に latestLastUpdated / staleDismissedAt / now を追加（既定は null で、渡さなければ "stale" にならない）。優先順は usage > change > stale。STALE_DAYS=30 を公開。最終更新の読み取りと、複数年度から最新を選ぶ newestLastUpdated（時刻で比較）を追加して公開。
- content.js: renderNotices が全年度の lastUpdated を読み、kingOfZangyoStaleNoticeDismissedAt と Date.now() を渡す。renderUsageNotice の描画を renderUpdateButtonNotice(updateButton, message, onClose) に共通化し、renderStaleNotice を追加（閉じる = Date.now() を保存して消す。更新ボタンを押すと共通処理で消える。詳しく見る = ?section=annual-fetch のまま）。
- tests/notices.test.js: 18件追加（30日ちょうど/29日、閉じた日10日/31日、未取得は usage、change 優先、"abc"、null、既存呼び出しは stale にならない、表示オフ・取得中、newestLastUpdated）。

## Test
make test: pass 80 / fail 0（RED時は5件失敗を確認）。node --check content.js OK。

## Step 5
未実施。test/ のHTMLは実拡張と chrome.storage が要り、ヘッドレスで再現していない。実機確認が必要。

## Concerns
- 日付の文字列を new Date(y,m-1,d,h,mi) で読むのでタイムゾーンは手元時刻（保存側の formatDateTime と一致）。
- 閉じるの保存は sync.set のエラー(lastError)を見ていない（変更点側と同じ作り）。
