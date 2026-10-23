/**
 * 年間残業の色分けを、年度の残りに応じて決める
 * 設計: docs/superpowers/specs/2026-10-14-annual-pace-bands-design.md
 *
 * グローバルに置く名前は KingOfZangyoAnnualPace の1つだけ。
 * 入力は保存データだけで、今日の日付は使わない（同じ入力なら必ず同じ結果）。
 */
const KingOfZangyoAnnualPace = (() => {
  const COLORS = ["#c8e6c9", "#fff59d", "#ef9a9a", "#d32f2f", "#990099"];
  // 濃赤と紫の上の文字色。白はギラつくので少し落とす（濃赤とのコントラスト比 4.57:1）
  const TEXT_ON_DARK = "#f5f5f5";
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
      textColor: bandIndex >= 3 ? TEXT_ON_DARK : "",
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
