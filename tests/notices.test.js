const { test } = require("node:test");
const assert = require("node:assert/strict");
const { pickNotice, withDismissed, shouldOpenGuideOnInstalled, newestLastUpdated, CHANGE_NOTICE_ID, STALE_DAYS } = require("../notices.js");
const base = { enabled: true, processing: false, savedYears: [], dismissed: [] };

test("一度も取得していなければ使い方", () => assert.equal(pickNotice(base), "usage"));
test("使い方は閉じた記録があっても出す", () => assert.equal(pickNotice({ ...base, dismissed: [CHANGE_NOTICE_ID] }), "usage"));
// 判定は今の年度を見ない。前の年度だけ取得した人も、ここに入る
test("取得したことがあり、閉じていなければ変更点", () => assert.equal(pickNotice({ ...base, savedYears: [2025] }), "change"));
test("変更点を閉じていれば何も出さない", () => assert.equal(pickNotice({ ...base, savedYears: [2026], dismissed: [CHANGE_NOTICE_ID] }), null));
test("表示オフなら何も出さない", () => assert.equal(pickNotice({ ...base, enabled: false }), null));
test("取得中なら何も出さない", () => assert.equal(pickNotice({ ...base, processing: true }), null));
// 更新で使い方ページを開かないことを、ここで縛る（設計書「やらないこと」）
test("インストール直後は使い方ページを開く", () => assert.equal(shouldOpenGuideOnInstalled("install"), true));
for (const reason of ["update", "chrome_update", "shared_module_update", undefined]) {
  test(`${reason} では開かない`, () => assert.equal(shouldOpenGuideOnInstalled(reason), false));
}
test("withDismissed は重複させない", () => assert.deepEqual(withDismissed(["a"], "a"), ["a"]));
test("withDismissed は元の配列を変えない", () => { const d = []; withDismissed(d, "a"); assert.deepEqual(d, []); });

// 更新を促す吹き出し（設計書 3b）。日付は固定して、同じ入力なら同じ結果にする
const DAY = 24 * 60 * 60 * 1000;
const now = new Date(2026, 9, 30, 12, 0).getTime();
const fmt = (ms) => {
  const d = new Date(ms);
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
};
const seen = { ...base, savedYears: [2026], dismissed: [CHANGE_NOTICE_ID], now, staleDismissedAt: null };

test("STALE_DAYS は30", () => assert.equal(STALE_DAYS, 30));
test("最終更新が30日前ちょうどなら更新を促す", () =>
  assert.equal(pickNotice({ ...seen, latestLastUpdated: fmt(now - 30 * DAY) }), "stale"));
test("最終更新が29日前なら出さない", () =>
  assert.equal(pickNotice({ ...seen, latestLastUpdated: fmt(now - 29 * DAY) }), null));
test("閉じたのが10日前なら出さない", () =>
  assert.equal(pickNotice({ ...seen, latestLastUpdated: fmt(now - 60 * DAY), staleDismissedAt: now - 10 * DAY }), null));
test("閉じたのが31日前ならまた出す", () =>
  assert.equal(pickNotice({ ...seen, latestLastUpdated: fmt(now - 60 * DAY), staleDismissedAt: now - 31 * DAY }), "stale"));
test("一度も取得していなければ最終更新に関係なく使い方", () =>
  assert.equal(pickNotice({ ...base, now, staleDismissedAt: null, latestLastUpdated: fmt(now - 90 * DAY) }), "usage"));
test("変更点が未読なら、最終更新が古くても変更点が先", () =>
  assert.equal(pickNotice({ ...seen, dismissed: [], latestLastUpdated: fmt(now - 90 * DAY) }), "change"));
test("最終更新が読めない値なら更新を促さない", () =>
  assert.equal(pickNotice({ ...seen, latestLastUpdated: "abc" }), null));
test("最終更新が無ければ更新を促さない", () =>
  assert.equal(pickNotice({ ...seen, latestLastUpdated: null }), null));
test("追加の3つを渡さなければ stale にならない", () =>
  assert.equal(pickNotice({ ...base, savedYears: [2026], dismissed: [CHANGE_NOTICE_ID] }), null));
test("表示オフなら更新を促さない", () =>
  assert.equal(pickNotice({ ...seen, enabled: false, latestLastUpdated: fmt(now - 90 * DAY) }), null));
test("取得中なら更新を促さない", () =>
  assert.equal(pickNotice({ ...seen, processing: true, latestLastUpdated: fmt(now - 90 * DAY) }), null));

test("newestLastUpdated は時刻でいちばん新しいものを返す（文字列の長さや順では決めない）", () =>
  assert.equal(newestLastUpdated(["2025/12/31 23:59", "2026/01/02 00:05", "2026/01/02 00:04", "abc", null]), "2026/01/02 00:05"));
test("newestLastUpdated は読める値が無ければ null", () => {
  assert.equal(newestLastUpdated([]), null);
  assert.equal(newestLastUpdated(["abc", undefined, "--"]), null);
});
