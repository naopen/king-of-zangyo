const { test } = require("node:test");
const assert = require("node:assert/strict");
const { pickNotice, withDismissed, shouldOpenGuideOnInstalled, CHANGE_NOTICE_ID } = require("../notices.js");
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
