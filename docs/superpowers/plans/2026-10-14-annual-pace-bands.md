# 年間残業の色分けを年度の残りに応じて動かす 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 年間残業の黄と赤の境目を、保存データの最終更新日から見た年度の残りに応じて動かし、凡例と完了ダイアログで境目を知らせる。

**Architecture:** 判定と文言は新しいファイル `annual-pace.js` に集め、グローバルには `KingOfZangyoAnnualPace` という名前を1つだけ置く。DOMを触らない純粋な関数なので、Node標準の `node --test` で検証する。`content.js` は今の固定閾値の関数を消し、`annual-pace.js` の結果で背景色・凡例・ダイアログを組み立てる。

**Tech Stack:** Chrome拡張（Manifest V3、Vanilla JS、ビルドなし）、Node 22.18.0 の `node:test` と `node:assert/strict`（依存パッケージは入れない）

**Spec:** `docs/superpowers/specs/2026-10-14-annual-pace-bands-design.md`

## Global Constraints

- 緑と黄の境目は `360 − 20 ×（12 − t + 1）`、黄と赤の境目は `360 − 10 ×（12 − t）`。どちらも整数の時間に切り捨て、判定も切り捨てた値で行う
- t は保存データの `lastUpdated` の日付と `fiscalYearStart` から日割りで出す。今日の日付と今の設定は使わない
- 帯の色は `#c8e6c9` `#fff59d` `#ef9a9a` `#d32f2f` `#990099`（今の `ANNUAL_BANDS` と同じ）。540h以上は文字色 `#ffffff`
- 凡例の表示名は時間で書く（`～192h` `192～286h` `286～360h` `360～540h(超過)` `540h+(特別条項超過)`）
- グローバルに置く名前は `KingOfZangyoAnnualPace` の1つだけ
- 依存パッケージを入れない。`npm i -g` / `pip install` をしない
- コミットメッセージは1行の日本語で、既存の `feat:` `fix:` `docs:` `test:` `chore:` の接頭辞に合わせる。`Co-Authored-By` は付けない。`git add` はパスを指定する。`git push` はNaokiの許可が出るまでしない
- 画面の文言は設計書の逐語どおりにする（凡例のtitle、完了ダイアログの文）

## Review Focus

- `fiscalYearStart` が文字列（`"4"`）で保存されている場合。ポップアップから来た値（`request.month`）の型が保証されていないので、数値に直して同じ結果を返すこと（Task 1 にテストあり）
- `lastUpdated` が想定外の形（`"--"`、空文字、ISO形式）の場合。例外を投げず、古い形式として固定閾値で表示すること（Task 1 にテストあり）
- 月末ちょうど（t が整数）の境目。浮動小数の誤差で 180h が 179h にならないこと（Task 1 の 7/31 と 12/31 のテスト）
- 540h以上の年度から少ない年度へドロップダウンで切り替えたとき、白い文字色が残らないこと。今のコードは文字色が空のとき何もしないので残りうる（Task 3 で常に代入し、手動確認に入れる）
- 上書き前の保存データの読み込みに失敗した場合。完了処理を止めず、前のデータなしとして扱うこと（Task 3 の `.catch(() => null)`）

---

### Task 1: 境目と帯を計算する `computeAnnualBands`

**Files:**
- Create: `annual-pace.js`
- Create: `tests/annual-pace.test.js`
- Modify: `Makefile`

**Interfaces:**
- Consumes: なし
- Produces:
  - グローバル定数 `KingOfZangyoAnnualPace`（Nodeでは `module.exports`）
  - `KingOfZangyoAnnualPace.computeAnnualBands(annualData, fiscalYear)` が返すオブジェクト

    ```
    {
      kind: "pace" | "finished" | "legacy",
      thresholds: { yellow: number|null, red: number|null, crimson: 360, purple: 540 },
      bandIndex: -1〜4,        // 色の番号。annualData が null なら -1
      color: string,           // 背景色。bandIndex が -1 なら ""
      textColor: string,       // 540h以上なら "#ffffff"、それ以外は ""
      legend: [{ colorIndex, color, label, title }],  // 凡例に並べる帯
      legendIndex: number,     // legend の中で今いる帯の位置。無ければ -1
      asOf: string|null,       // "2026/08/20"。legacy なら null
    }
    ```

- [ ] **Step 1: 失敗するテストを書く**

`tests/annual-pace.test.js` を作る。

```js
const { test, describe } = require("node:test");
const assert = require("node:assert/strict");
const { computeAnnualBands } = require("../annual-pace.js");

// 保存データを作る。fiscalYearStart の既定は4月
function data(totalMinutes, date, fiscalYearStart = 4) {
  return { totalMinutes, lastUpdated: `${date} 18:05`, fiscalYearStart };
}
const h = (hours) => hours * 60;
const labels = (bands) => bands.legend.map((item) => item.label);

describe("境目（4月始まり、2026年度）", () => {
  const cases = [
    ["2026/07/31", 180, 280],
    ["2026/08/10", 186, 283],
    ["2026/08/20", 192, 286],
    ["2026/12/31", 280, 330],
    ["2027/02/28", 320, 350],
    ["2027/03/15", 329, 354],
    ["2027/03/31", 340, 360],
  ];
  for (const [date, yellow, red] of cases) {
    test(date, () => {
      const bands = computeAnnualBands(data(0, date), 2026);
      assert.equal(bands.kind, "pace");
      assert.equal(bands.thresholds.yellow, yellow);
      assert.equal(bands.thresholds.red, red);
      assert.equal(bands.asOf, date);
    });
  }
});

describe("凡例", () => {
  test("2026/08/20 は5つの帯を時間で出す", () => {
    const bands = computeAnnualBands(data(h(209), "2026/08/20"), 2026);
    assert.deepEqual(labels(bands), [
      "～192h", "192～286h", "286～360h", "360～540h(超過)", "540h+(特別条項超過)",
    ]);
    assert.equal(bands.legendIndex, 1);
    assert.equal(bands.legend[1].color, "#fff59d");
  });

  test("年度の最終日は幅0の赤の帯を出さない", () => {
    const bands = computeAnnualBands(data(h(345), "2027/03/31"), 2026);
    assert.deepEqual(labels(bands), [
      "～340h", "340～360h", "360～540h(超過)", "540h+(特別条項超過)",
    ]);
    assert.equal(bands.bandIndex, 1);
    assert.equal(bands.legendIndex, 1);
  });

  test("今年度の帯の title", () => {
    const titles = computeAnnualBands(data(0, "2026/08/20"), 2026).legend.map((i) => i.title);
    assert.deepEqual(titles, [
      "年360hまでの残り枠が、残りの月を月20hずつ使っても20h以上余る見込みです。",
      "年360hまでの残り枠が「残りの月数×20h＋20h」を切る見込みです。残りを月20h以内にすれば収まりますが、余裕は20h未満です。",
      "年360hまでの残り枠が「残りの月数×10h」を切る見込みです。残りの月を月10h未満にしないと、年360hを超えます。",
      "年360h（36協定の上限）を超えています。",
      "年540h（36協定の特別条項の上限）を超えています。",
    ]);
  });
});

describe("色が変わる日（毎日ボタンを押した場合）", () => {
  // [毎月の残業, 日付, その日の累計（分）, 期待する帯]
  const cases = [
    [45, "2026/07/30", 10713, 0],
    [45, "2026/07/31", 10800, 1],
    [45, "2026/08/01", 10887, 1], // 月替わりで緑に戻らない
    [45, "2026/10/26", 18465, 1],
    [45, "2026/10/27", 18552, 2],
    [30, "2027/01/28", 17826, 0],
    [30, "2027/01/29", 17884, 1],
    [30, "2027/03/29", 21484, 1],
    [30, "2027/03/30", 21542, 2],
    [25, "2027/03/30", 17952, 0],
  ];
  for (const [pace, date, minutes, expected] of cases) {
    test(`${pace}h/月 ${date}`, () => {
      assert.equal(computeAnnualBands(data(minutes, date), 2026).bandIndex, expected);
    });
  }
});

describe("境目ちょうど", () => {
  test("192:00 は黄、191:59 は緑（2026/08/20）", () => {
    assert.equal(computeAnnualBands(data(h(192), "2026/08/20"), 2026).bandIndex, 1);
    assert.equal(computeAnnualBands(data(h(192) - 1, "2026/08/20"), 2026).bandIndex, 0);
  });
  test("360:00 は濃赤、540:00 は紫で文字が白", () => {
    assert.equal(computeAnnualBands(data(h(360), "2026/08/20"), 2026).bandIndex, 3);
    const purple = computeAnnualBands(data(h(540), "2026/08/20"), 2026);
    assert.equal(purple.bandIndex, 4);
    assert.equal(purple.textColor, "#ffffff");
    assert.equal(purple.color, "#990099");
  });
});

describe("年度開始月", () => {
  test("10月始まりの2026年度、2027/01/31 は4か月経過", () => {
    const bands = computeAnnualBands(data(0, "2027/01/31", 10), 2026);
    assert.equal(bands.thresholds.yellow, 180);
    assert.equal(bands.thresholds.red, 280);
  });
  test("1月始まりの2026年度、2026/12/31 は最終日", () => {
    const bands = computeAnnualBands(data(0, "2026/12/31", 1), 2026);
    assert.equal(bands.kind, "pace");
    assert.equal(bands.thresholds.yellow, 340);
  });
  test("1月始まりの2026年度、2027/01/01 は終わった年度", () => {
    assert.equal(computeAnnualBands(data(0, "2027/01/01", 1), 2026).kind, "finished");
  });
  test("fiscalYearStart が文字列でも同じ結果", () => {
    const bands = computeAnnualBands(data(0, "2026/08/20", "4"), 2026);
    assert.equal(bands.kind, "pace");
    assert.equal(bands.thresholds.yellow, 192);
  });
  test("fiscalYear が文字列でも同じ結果", () => {
    assert.equal(computeAnnualBands(data(0, "2026/08/20"), "2026").thresholds.yellow, 192);
  });
});

describe("年度の外", () => {
  test("最終更新日が年度の始まりより前なら t=0", () => {
    const bands = computeAnnualBands(data(0, "2026/03/31"), 2026);
    assert.equal(bands.thresholds.yellow, 100);
    assert.equal(bands.thresholds.red, 240);
  });
  test("最終更新日が年度の終わり以降なら3つの帯", () => {
    const bands = computeAnnualBands(data(h(350), "2027/04/01"), 2026);
    assert.equal(bands.kind, "finished");
    assert.equal(bands.thresholds.yellow, null);
    assert.equal(bands.thresholds.red, null);
    assert.equal(bands.bandIndex, 0);
    assert.deepEqual(labels(bands), ["～360h", "360～540h(超過)", "540h+(特別条項超過)"]);
    assert.equal(bands.legend[0].title, "年360h（36協定の上限）以内に収まっています。");
  });
  test("過去の年度でも、最終更新日が年度の中なら5つの帯", () => {
    const bands = computeAnnualBands(data(0, "2026/03/20"), 2025);
    assert.equal(bands.kind, "pace");
    assert.equal(bands.legend.length, 5);
  });
});

describe("古い形式のデータ", () => {
  const legacyLabels = ["～300h", "300～330h", "330～360h", "360～540h(超過)", "540h+(特別条項超過)"];
  test("lastUpdated が無い", () => {
    const bands = computeAnnualBands({ totalMinutes: h(310), fiscalYearStart: 4 }, 2026);
    assert.equal(bands.kind, "legacy");
    assert.equal(bands.bandIndex, 1);
    assert.deepEqual(labels(bands), legacyLabels);
    assert.equal(bands.legend[0].title, "");
    assert.equal(bands.asOf, null);
  });
  test("fiscalYearStart が無い", () => {
    assert.equal(computeAnnualBands({ totalMinutes: 0, lastUpdated: "2026/08/20 18:05" }, 2026).kind, "legacy");
  });
  for (const bad of ["--", "", "2026-08-20T18:05:00Z", "13", 0]) {
    test(`fiscalYearStart や lastUpdated が想定外（${JSON.stringify(bad)}）`, () => {
      assert.equal(computeAnnualBands({ totalMinutes: 0, lastUpdated: bad, fiscalYearStart: 4 }, 2026).kind, "legacy");
      assert.equal(computeAnnualBands({ totalMinutes: 0, lastUpdated: "2026/08/20 18:05", fiscalYearStart: bad }, 2026).kind, "legacy");
    });
  }
  test("データが無い（null）なら bandIndex は -1", () => {
    const bands = computeAnnualBands(null, 2026);
    assert.equal(bands.kind, "legacy");
    assert.equal(bands.bandIndex, -1);
    assert.equal(bands.legendIndex, -1);
    assert.equal(bands.color, "");
  });
  test("totalMinutes が無ければ0として扱う", () => {
    assert.equal(computeAnnualBands({ lastUpdated: "2026/08/20 18:05", fiscalYearStart: 4 }, 2026).bandIndex, 0);
  });
});
```

`Makefile` に `test` ターゲットを足す。`.PHONY` にも加える。Node 22 の `node --test` にフォルダー（`tests/`）を渡すとファイルとして読もうとして失敗するので、ファイルのパターンを引用符つきで渡す（2026-10-14 に Node 22.18.0 で確認）。

```makefile
.PHONY: zip clean test

zip:
	zip -r king-of-zangyo.zip manifest.json popup.html popup.js popup.css content.js icons/

test:
	node --test "tests/*.test.js"

clean:
	rm -f king-of-zangyo.zip
```

- [ ] **Step 2: テストが落ちることを確かめる**

Run: `make test`
Expected: FAIL。`Cannot find module '../annual-pace.js'`

- [ ] **Step 3: `annual-pace.js` を書く**

```js
/**
 * 年間残業の色分けを、年度の残りに応じて決める
 * 設計: docs/superpowers/specs/2026-10-14-annual-pace-bands-design.md
 *
 * グローバルに置く名前は KingOfZangyoAnnualPace の1つだけ。
 * 入力は保存データだけで、今日の日付は使わない（同じ入力なら必ず同じ結果）。
 */
const KingOfZangyoAnnualPace = (() => {
  const COLORS = ["#c8e6c9", "#fff59d", "#ef9a9a", "#d32f2f", "#990099"];
  const LIMIT_HOURS = 360; // 36協定の年間上限
  const SPECIAL_LIMIT_HOURS = 540; // 特別条項の年間上限

  const OVER_TITLES = {
    3: "年360h（36協定の上限）を超えています。",
    4: "年540h（36協定の特別条項の上限）を超えています。",
  };
  const PACE_TITLES = {
    0: "年360hまでの残り枠が、残りの月を月20hずつ使っても20h以上余る見込みです。",
    1: "年360hまでの残り枠が「残りの月数×20h＋20h」を切る見込みです。残りを月20h以内にすれば収まりますが、余裕は20h未満です。",
    2: "年360hまでの残り枠が「残りの月数×10h」を切る見込みです。残りの月を月10h未満にしないと、年360hを超えます。",
    ...OVER_TITLES,
  };
  const FINISHED_TITLES = {
    0: "年360h（36協定の上限）以内に収まっています。",
    ...OVER_TITLES,
  };

  /** "2026/08/20 18:05" から年月日を取り出す。形が違えば null */
  function parseDate(lastUpdated) {
    const match = /^(\d{4})\/(\d{2})\/(\d{2})/.exec(
      typeof lastUpdated === "string" ? lastUpdated : "",
    );
    if (!match) return null;
    return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) };
  }

  /** 年度開始月を1〜12の数値にする。範囲外なら null */
  function parseStartMonth(value) {
    const month = Number(value);
    return Number.isInteger(month) && month >= 1 && month <= 12 ? month : null;
  }

  /** 浮動小数の誤差（179.9999…）で1時間ずれないように切り捨てる */
  function floorHours(value) {
    return Math.floor(value + 1e-9);
  }

  function bandIndexOf(totalMinutes, thresholds) {
    if (totalMinutes >= thresholds.purple * 60) return 4;
    if (totalMinutes >= thresholds.crimson * 60) return 3;
    if (thresholds.red !== null && totalMinutes >= thresholds.red * 60) return 2;
    if (thresholds.yellow !== null && totalMinutes >= thresholds.yellow * 60) return 1;
    return 0;
  }

  function legacyBands() {
    return {
      kind: "legacy",
      thresholds: { yellow: 300, red: 330, crimson: LIMIT_HOURS, purple: SPECIAL_LIMIT_HOURS },
      items: [
        { colorIndex: 0, label: "～300h" },
        { colorIndex: 1, label: "300～330h" },
        { colorIndex: 2, label: "330～360h" },
        { colorIndex: 3, label: "360～540h(超過)" },
        { colorIndex: 4, label: "540h+(特別条項超過)" },
      ],
      titles: {},
      asOf: null,
    };
  }

  function datedBands(date, fiscalYear, startMonth) {
    const asOf = `${date.year}/${String(date.month).padStart(2, "0")}/${String(date.day).padStart(2, "0")}`;
    const monthIndex = (date.year - fiscalYear) * 12 + (date.month - startMonth);

    if (monthIndex >= 12) {
      return {
        kind: "finished",
        thresholds: { yellow: null, red: null, crimson: LIMIT_HOURS, purple: SPECIAL_LIMIT_HOURS },
        items: [
          { colorIndex: 0, label: "～360h" },
          { colorIndex: 3, label: "360～540h(超過)" },
          { colorIndex: 4, label: "540h+(特別条項超過)" },
        ],
        titles: FINISHED_TITLES,
        asOf,
      };
    }

    // 経過月数（日割り）。最終更新日のその日の終わりまでを数える
    const daysInMonth = new Date(date.year, date.month, 0).getDate();
    const elapsed = monthIndex < 0 ? 0 : monthIndex + date.day / daysInMonth;
    const remaining = 12 - elapsed;
    const yellow = floorHours(LIMIT_HOURS - 20 * (remaining + 1));
    const red = floorHours(LIMIT_HOURS - 10 * remaining);

    const items = [
      { colorIndex: 0, label: `～${yellow}h` },
      { colorIndex: 1, label: `${yellow}～${red}h` },
      { colorIndex: 2, label: `${red}～360h` },
      { colorIndex: 3, label: "360～540h(超過)" },
      { colorIndex: 4, label: "540h+(特別条項超過)" },
    ].filter((item) => !(item.colorIndex === 2 && red >= LIMIT_HOURS)); // 年度の最終日は赤の幅が0

    return {
      kind: "pace",
      thresholds: { yellow, red, crimson: LIMIT_HOURS, purple: SPECIAL_LIMIT_HOURS },
      items,
      titles: PACE_TITLES,
      asOf,
    };
  }

  /**
   * 保存データから、年間の帯と凡例を作る
   * @param {Object|null} annualData - 保存データ（totalMinutes, lastUpdated, fiscalYearStart）
   * @param {number|string} fiscalYear - 保存キーの年度
   */
  function computeAnnualBands(annualData, fiscalYear) {
    const date = annualData ? parseDate(annualData.lastUpdated) : null;
    const startMonth = annualData ? parseStartMonth(annualData.fiscalYearStart) : null;
    const base =
      date && startMonth
        ? datedBands(date, Number(fiscalYear), startMonth)
        : legacyBands();

    const bandIndex = annualData
      ? bandIndexOf(annualData.totalMinutes || 0, base.thresholds)
      : -1;
    const legend = base.items.map((item) => ({
      colorIndex: item.colorIndex,
      color: COLORS[item.colorIndex],
      label: item.label,
      title: base.titles[item.colorIndex] || "",
    }));

    return {
      kind: base.kind,
      thresholds: base.thresholds,
      bandIndex,
      color: bandIndex >= 0 ? COLORS[bandIndex] : "",
      textColor: bandIndex === 4 ? "#ffffff" : "",
      legend,
      legendIndex: legend.findIndex((item) => item.colorIndex === bandIndex),
      asOf: base.asOf,
    };
  }

  return { computeAnnualBands };
})();

if (typeof module !== "undefined") {
  module.exports = KingOfZangyoAnnualPace;
}
```

- [ ] **Step 4: テストが通ることを確かめる**

Run: `make test`
Expected: PASS。失敗が0件

- [ ] **Step 5: コミット**

```bash
git add annual-pace.js tests/annual-pace.test.js Makefile
git commit -m "feat: 年間残業の境目を年度の残りから計算するannual-pace.jsを追加"
```

---

### Task 2: 完了ダイアログの文言を作る `buildBandsNotice`

**Files:**
- Modify: `annual-pace.js`
- Modify: `tests/annual-pace.test.js`

**Interfaces:**
- Consumes: `computeAnnualBands` の返り値（`kind`、`thresholds.yellow`、`thresholds.red`、`asOf`）
- Produces: `KingOfZangyoAnnualPace.buildBandsNotice(bands, previousAnnualData)` → `string`。足すものが無ければ `""`。行の区切りは `\n`

- [ ] **Step 1: 失敗するテストを書く**

`tests/annual-pace.test.js` の1行目の `require` を次に変え、末尾にテストを足す。

```js
const { computeAnnualBands, buildBandsNotice } = require("../annual-pace.js");
```

```js
describe("完了ダイアログに足す文言", () => {
  const current = computeAnnualBands(data(h(209), "2026/08/20"), 2026);
  const expected =
    "年度の残日数に合わせ、凡例を更新しました（2026/08/20 時点）\n" +
    "  緑↔黄  192h\n" +
    "  黄↔赤  286h";

  test("前のデータが新しい形式", () => {
    assert.equal(buildBandsNotice(current, data(h(194), "2026/08/10")), expected);
  });
  test("前のデータが無い（初めて取得）", () => {
    assert.equal(buildBandsNotice(current, null), expected);
  });
  test("前のデータが古い形式なら1行足す", () => {
    assert.equal(
      buildBandsNotice(current, { totalMinutes: h(194), yearRange: "2026/04-2027/03" }),
      expected + "\n年度の残りに応じて境目が動くようになりました。",
    );
  });
  test("過去の年度は何も足さない（前が古い形式でも）", () => {
    const finished = computeAnnualBands(data(h(350), "2027/04/01"), 2026);
    assert.equal(buildBandsNotice(finished, { totalMinutes: h(350) }), "");
  });
  test("今のデータが古い形式なら何も足さない", () => {
    const legacy = computeAnnualBands({ totalMinutes: h(100) }, 2026);
    assert.equal(buildBandsNotice(legacy, null), "");
  });
});
```

- [ ] **Step 2: テストが落ちることを確かめる**

Run: `make test`
Expected: FAIL。`buildBandsNotice is not a function`

- [ ] **Step 3: 実装する**

`annual-pace.js` の `computeAnnualBands` の後ろに足す。

```js
  /** 保存データが古い形式（最終更新日か年度開始月が無い）か */
  function isLegacyData(annualData) {
    return (
      !annualData ||
      !parseDate(annualData.lastUpdated) ||
      !parseStartMonth(annualData.fiscalYearStart)
    );
  }

  /**
   * 完了ダイアログに足す、凡例の境目の知らせ
   * @param {Object} bands - computeAnnualBands の返り値（取り直したデータ）
   * @param {Object|null} previousAnnualData - 上書きする前の保存データ
   * @returns {string} 足す文言。足すものが無ければ空文字
   */
  function buildBandsNotice(bands, previousAnnualData) {
    if (bands.kind !== "pace") return "";
    const lines = [
      `年度の残日数に合わせ、凡例を更新しました（${bands.asOf} 時点）`,
      `  緑↔黄  ${bands.thresholds.yellow}h`,
      `  黄↔赤  ${bands.thresholds.red}h`,
    ];
    if (previousAnnualData && isLegacyData(previousAnnualData)) {
      lines.push("年度の残りに応じて境目が動くようになりました。");
    }
    return lines.join("\n");
  }
```

公開するオブジェクトを変える。

```js
  return { computeAnnualBands, buildBandsNotice };
```

- [ ] **Step 4: テストが通ることを確かめる**

Run: `make test`
Expected: PASS。失敗が0件

- [ ] **Step 5: コミット**

```bash
git add annual-pace.js tests/annual-pace.test.js
git commit -m "feat: 完了ダイアログで凡例の境目を知らせる文言を追加"
```

---

### Task 3: `content.js` を新しい判定に差し替える

**Files:**
- Modify: `content.js`（`ANNUAL_BANDS` 736〜742行、`getAnnualBandIndex` 753〜759行、`buildLegendElement` 765〜812行、年別データの初期凡例 1185〜1193行、`updateAnnualDataDisplay` 1256〜1320行、`getAnnualOvertimeBackgroundColor` と `getAnnualOvertimeTextColor` 1324〜1359行、`completeFetchAnnualOvertime` 2004〜2049行、`showResultDialog` 2185〜2247行）
- Modify: `manifest.json`
- Modify: `Makefile`
- Modify: `test/extension-test.html`、`test/extension-test-時間休あり.html`、`test/extension-test-管理職.html`、`test/extension-test-早退型時間休.html`（サブモジュール）

**Interfaces:**
- Consumes: `KingOfZangyoAnnualPace.computeAnnualBands(annualData, fiscalYear)`、`KingOfZangyoAnnualPace.buildBandsNotice(bands, previousAnnualData)`
- Produces: `showResultDialog(annualData, fiscalYear, previousAnnualData)`（第3引数を追加）

行番号は作業前のもの。作業の途中でずれるので、関数名で探して置き換える。

- [ ] **Step 1: 読み込みを足す**

`manifest.json` の `content_scripts` の `js` を変える。

```json
      "js": ["annual-pace.js", "content.js"],
```

`Makefile` の zip の対象に `annual-pace.js` を足す。

```makefile
zip:
	zip -r king-of-zangyo.zip manifest.json popup.html popup.js popup.css annual-pace.js content.js icons/
```

- [ ] **Step 2: 固定閾値の定義と関数を消す**

`content.js` から次の3つを消す。`MONTHLY_BANDS` と `getMonthlyBandIndex` は残す。

- `const ANNUAL_BANDS = [ ... ];`
- `function getAnnualBandIndex(hours) { ... }`
- `getAnnualOvertimeBackgroundColor` と `getAnnualOvertimeTextColor`（JSDocコメントごと）

- [ ] **Step 3: 凡例に title を付けられるようにする**

`buildLegendElement` の `bands.forEach(({ color, label }, i) => {` を次に変える。月間の帯には `title` が無いので、月間の表示は変わらない。

```js
  bands.forEach(({ color, label, title }, i) => {
    const span = document.createElement("span");
    span.style.display = "inline-flex";
    span.style.alignItems = "center";
    span.style.marginLeft = "8px";
    if (title) {
      span.title = title;
    }
```

（`const span = ...` から `span.style.marginLeft = "8px";` までの既存の4行を、上の6行で置き換える）

- [ ] **Step 4: 年別データの初期凡例を差し替える**

`injectAnnualDataSection` の中の初期凡例を変える。

```js
    // 年間凡例（初期は現在地未確定なので等価表示、左寄せで目安年間残業セルに揃える）
    const annualLegendDiv = buildLegendElement(
      KingOfZangyoAnnualPace.computeAnnualBands(null, currentFiscalYear).legend,
      -1,
      ANNUAL_LEGEND_ID,
      "left",
    );
```

- [ ] **Step 5: `updateAnnualDataDisplay` を差し替える**

データが無いときの凡例（`if (!annualData) { ... }` の中）を変える。

```js
      const newLegend = buildLegendElement(
        KingOfZangyoAnnualPace.computeAnnualBands(null, fiscalYear).legend,
        -1,
        ANNUAL_LEGEND_ID,
        "left",
      );
```

「背景色と文字色を設定（時間単位で判定）」から「年間凡例を更新」の `replaceWith` までを次にする。文字色は空文字でも代入し、540h以上の年度から切り替えたときに白が残らないようにする。

```js
  // 背景色と文字色を設定（年度の残りに応じた境目で判定）
  const bands = KingOfZangyoAnnualPace.computeAnnualBands(annualData, fiscalYear);
  hoursCell.style.backgroundColor = bands.color;
  hoursCell.style.color = bands.textColor;

  // 最終更新を表示
  updatedCell.textContent = annualData.lastUpdated || "--";

  // 年間凡例を更新
  const existingAnnualLegend = document.getElementById(ANNUAL_LEGEND_ID);
  if (existingAnnualLegend) {
    const newLegend = buildLegendElement(
      bands.legend,
      bands.legendIndex,
      ANNUAL_LEGEND_ID,
      "left",
    );
    newLegend.style.display = overtimeEnabled && legendEnabled ? "" : "none";
    existingAnnualLegend.replaceWith(newLegend);
  }
```

あわせて、不要になった `const totalHours = totalMinutes / 60;` を消す。

- [ ] **Step 6: 完了処理で前の保存データを読む**

`completeFetchAnnualOvertime` の `await saveAnnualData(annualData, targetFiscalYear);` の直前に足す。読み込みに失敗しても完了処理は止めない。

```js
    // 上書き前のデータ（古い形式かどうかを完了ダイアログで使う）
    const previousAnnualData = await loadAnnualData(targetFiscalYear).catch(
      () => null,
    );
```

結果ダイアログの呼び出しを変える。

```js
    await showResultDialog(annualData, targetFiscalYear, previousAnnualData);
```

- [ ] **Step 7: `showResultDialog` を差し替える**

シグネチャとJSDocに第3引数を足す。

```js
/**
 * 結果ダイアログを表示する
 * @param {Object} annualData - 取り直した年間データ
 * @param {number} fiscalYear - 年度
 * @param {Object|null} previousAnnualData - 上書きする前の保存データ
 */
function showResultDialog(annualData, fiscalYear, previousAnnualData) {
```

`const totalHours = totalMinutes / 60;` を消し、「背景色と文字色を設定」のコメントから `body.style.borderRadius = "4px";` までを次にする。

```js
    // 背景色と文字色を設定
    const bands = KingOfZangyoAnnualPace.computeAnnualBands(annualData, fiscalYear);
    body.style.backgroundColor = bands.color;
    body.style.color = bands.textColor;
    body.style.padding = "20px";
    body.style.borderRadius = "4px";

    // 凡例の境目の知らせ（色の背景の外に出す。過去の年度などで空なら出さない）
    const noticeText = KingOfZangyoAnnualPace.buildBandsNotice(
      bands,
      previousAnnualData,
    );
    const notice = document.createElement("div");
    notice.textContent = noticeText;
    notice.style.padding = "12px 20px 4px";
    notice.style.fontSize = "14px";
    notice.style.lineHeight = "1.7";
    notice.style.color = "#333";
    notice.style.whiteSpace = "pre-line";
```

組み立ての部分を次にする。

```js
    dialog.appendChild(header);
    dialog.appendChild(body);
    if (noticeText) {
      dialog.appendChild(notice);
    }
    dialog.appendChild(footer);
```

- [ ] **Step 8: 消した関数の呼び出しが残っていないことを確かめる**

Run: `grep -n -E "ANNUAL_BANDS|getAnnualBandIndex|getAnnualOvertimeBackgroundColor|getAnnualOvertimeTextColor" content.js`
Expected: 何も出ない

Run: `node --check content.js && node --check annual-pace.js && make test`
Expected: 構文エラーなし、テストは PASS

- [ ] **Step 9: テスト用HTMLに読み込みを足す（サブモジュール）**

`test/` の4ファイルで、`<script src="../content.js"></script>` の直前に1行足す。

```html
  <script src="../annual-pace.js"></script>
  <script src="../content.js"></script>
```

Run: `grep -c 'src="../annual-pace.js"' test/*.html`
Expected: 4ファイルとも `1`

- [ ] **Step 10: 実機で確かめる（Naokiと一緒に）**

Naokiに `chrome://extensions` でこの拡張を再読み込みしてもらい、KING OF TIMEのタイムカード画面で次を見る。拡張のコンソール（DevTools の Console、コンテキストを拡張に切り替える）にエラーが出ていないことも見る。ここで `KingOfZangyoAnnualPace` が `content.js` から見えることを実測で確かめる（設計書で確かめていなかった前提）。見えずに `ReferenceError` が出たら、`annual-pace.js` の `const KingOfZangyoAnnualPace = ...` を `globalThis.KingOfZangyoAnnualPace = ...` に変えて、テストを回し直す。

- [ ] 年別データの凡例が、更新前は今の固定閾値（`～300h` など）で、title が出ない
- [ ] 「最新データで更新」を押すと、完了ダイアログに「年度の残日数に合わせ、凡例を更新しました（今日の日付 時点）」と緑↔黄・黄↔赤の時間が出る。初回は「年度の残りに応じて境目が動くようになりました。」の行も出る
- [ ] ダイアログを閉じたあと、凡例の時間がダイアログと同じ値になっている。帯にマウスを乗せると設計書の文が出る
- [ ] 年度のドロップダウンで過去の年度に切り替えると、その年度のデータの最終更新日に応じた帯になる。540h以上の年度から少ない年度へ切り替えて、文字が白のまま残らない
- [ ] 月間の凡例と色は変わっていない

- [ ] **Step 11: コミット（サブモジュール → 親リポジトリの順）**

サブモジュールは detached HEAD になっていることが多い。コミットの前に `git -C test status` でブランチを見て、detached なら `git -C test switch main` してから作業する（`main` が無ければ `git -C test branch -a` で既定のブランチを確かめる）。サブモジュール側の push はNaokiの許可が出るまでしない。push しないうちは、親リポジトリが指すサブモジュールのコミットはリモートに無い状態になる。

```bash
cd test
git add extension-test.html extension-test-時間休あり.html extension-test-管理職.html extension-test-早退型時間休.html
git commit -m "test: annual-pace.jsを読み込む行を追加"
cd ..
git add manifest.json Makefile content.js test
git commit -m "feat: 年間残業の色分けを年度の残りに応じた境目に切り替え、完了ダイアログで知らせる"
```

---

### Task 4: 文書とバージョン

**Files:**
- Modify: `README.md`（37〜48行の年間残業時間の表）
- Modify: `store-assets/STORE_LISTING.txt`（日本語版のバージョン履歴 18〜24行と年間残業時間 47〜52行、英語版のバージョン履歴 92〜98行と Annual Overtime 121〜125行）
- Modify: `manifest.json`（`version`）

**Interfaces:**
- Consumes: Task 3 までの挙動
- Produces: なし

- [ ] **Step 1: README の年間残業時間の表を書き直す**

`**年間残業時間（5 段階）**` から表の終わりまでを次にする。

```markdown
**年間残業時間（5 段階）**

黄と赤の境目は、年間データを最後に更新した日から見た年度の残りに応じて動きます。年360hまでの残り枠が「残りの月数×20h＋20h」を切ると黄、「残りの月数×10h」を切ると赤になります。境目は「最新データで更新」を押したときに計算し直し、完了ダイアログで知らせます。

| 色   | 範囲（2026/08/20 に更新した場合の例） | 意味         |
| ---- | ------------------------------------- | ------------ |
| 緑   | 0 ～ 192 時間                         | 安全範囲     |
| 黄   | 192 ～ 286 時間                       | 注意         |
| 赤   | 286 ～ 360 時間                       | 危険         |
| 濃赤 | 360 ～ 540 時間                       | 超過         |
| 紫   | 540 時間以上                          | 特別条項超過 |

年度が終わったあとに更新した年度は、緑（360 時間未満）、濃赤、紫の 3 段階で表示します。
```

- [ ] **Step 2: ストア掲載文の年間残業時間を書き直す**

日本語版の `年間残業時間:` から紫の行までを次にする。

```
年間残業時間:
  黄と赤の境目は、最後に更新した日から見た年度の残りに応じて動きます。
  （例: 2026/08/20 に更新した場合）
  - 緑: 0～192時間（安全範囲）
  - 黄: 192～286時間（注意）
  - 赤: 286～360時間（危険）
  - 濃赤: 360～540時間（超過）
  - 紫: 540時間以上（特別条項超過）
```

英語版の `Annual Overtime:` から Purple の行までを次にする。

```
Annual Overtime:
  The yellow and red thresholds move with the time left in the fiscal year, counted from the date you last updated.
  (Example: updated on 2026/08/20)
  - Green: 0-192 hours (safe range)
  - Yellow: 192-286 hours (caution)
  - Red: 286-360 hours (warning)
  - Dark Red: 360-540 hours (exceeded)
  - Purple: 540+ hours (special clause exceeded)
```

- [ ] **Step 3: バージョン履歴を足し、5行にそろえる**

日本語版の `【主なバージョン履歴】` の先頭に足し、末尾の `v2.2.0: 半日休暇（AM有休・PM有休）の残業時間を計算できるようになりました` を消す。

```
v2.5.0: 年間残業時間の黄と赤の境目が、年度の残りに応じて動くようになりました
```

英語版の `Key Version History` の先頭に足し、末尾の `v2.2.0: Added overtime calculation for half-day leave (AM/PM leave)` を消す。

```
v2.5.0: The yellow and red annual overtime thresholds now move with the time left in the fiscal year
```

- [ ] **Step 4: バージョン番号を上げる**

`manifest.json` の `"version": "2.4.8"` を `"version": "2.5.0"` にする。

- [ ] **Step 5: 確かめてコミット**

Run: `grep -n -E "300|330" README.md store-assets/STORE_LISTING.txt`
Expected: 年間残業時間の固定閾値（300、330）が出ない

Run: `grep -c "^v2\." store-assets/STORE_LISTING.txt`
Expected: `10`（日本語版5行と英語版5行）

```bash
git add README.md store-assets/STORE_LISTING.txt
git commit -m "docs: 年間残業の色分けの説明を年度の残りに応じた境目に書き直す"
git add manifest.json
git commit -m "chore: バージョン番号を2.5.0に更新"
```
