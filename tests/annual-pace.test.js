const { test, describe } = require("node:test");
const assert = require("node:assert/strict");
const { computeAnnualBands, buildBandsNotice } = require("../annual-pace.js");

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
  test("360:00 は濃赤、540:00 は紫で、どちらも文字がオフホワイト", () => {
    const crimson = computeAnnualBands(data(h(360), "2026/08/20"), 2026);
    assert.equal(crimson.bandIndex, 3);
    assert.equal(crimson.color, "#d32f2f");
    assert.equal(crimson.textColor, "#f5f5f5");
    const purple = computeAnnualBands(data(h(540), "2026/08/20"), 2026);
    assert.equal(purple.bandIndex, 4);
    assert.equal(purple.color, "#990099");
    assert.equal(purple.textColor, "#f5f5f5");
  });
  test("359:59 は赤で、文字色を指定しない", () => {
    assert.equal(computeAnnualBands(data(h(360) - 1, "2026/08/20"), 2026).textColor, "");
  });
  test("古い形式でも濃赤の文字はオフホワイト", () => {
    assert.equal(computeAnnualBands({ totalMinutes: h(400) }, 2026).textColor, "#f5f5f5");
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

// v1.3.0〜v2.4.8 が保存していた形。年間データは v1.3.0 で入り、最初から lastUpdated と fiscalYearStart を持つ
function savedByV248(totalMinutes, lastUpdated, fiscalYear) {
  return {
    totalHours: totalMinutes / 60,
    totalMinutes,
    lastUpdated,
    fiscalYearStart: 4,
    yearRange: `${fiscalYear}/04-${fiscalYear + 1}/03`,
    monthlyData: { [`${fiscalYear}-04`]: totalMinutes },
    fiscalYear,
  };
}

describe("v2.4.8 以前に保存したデータ", () => {
  test("アップデート直後から、最終更新日の時点の境目で描く", () => {
    const bands = computeAnnualBands(savedByV248(h(250), "2026/09/30 18:05", 2026), 2026);
    assert.equal(bands.kind, "pace");
    assert.equal(bands.thresholds.yellow, 220);
    assert.equal(bands.thresholds.red, 300);
    assert.equal(bands.bandIndex, 1); // 固定の境目では緑だった
  });
  test("年度が終わってから更新した年度は3段階", () => {
    const bands = computeAnnualBands(savedByV248(h(300), "2026/04/02 20:00", 2025), 2025);
    assert.equal(bands.kind, "finished");
    assert.equal(bands.bandIndex, 0); // 固定の境目では黄だった
  });
  test("取り直したとき、完了ダイアログに古い形式向けの1行を足さない", () => {
    const current = computeAnnualBands(data(h(260), "2026/10/26"), 2026);
    assert.equal(
      buildBandsNotice(current, savedByV248(h(250), "2026/09/30 18:05", 2026)),
      "年度の残日数に合わせ、凡例を更新しました（2026/10/26 時点）\n" +
        "  緑↔黄  236h\n" +
        "  黄↔赤  308h",
    );
  });
});

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
