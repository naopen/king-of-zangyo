/**
 * 使い方の吹き出しと、変更点の吹き出しの出し分けを決める
 * 設計: docs/superpowers/specs/2026-10-29-usage-guide-and-notices-design.md
 *
 * グローバルに置く名前は KingOfZangyoNotices の1つだけ。
 * 入力は渡された値だけで、今日の日付や保存領域は読まない（同じ入力なら必ず同じ結果）。
 */
const KingOfZangyoNotices = (() => {
  // 変更点の吹き出しを閉じたかどうかの記録に使うID
  const CHANGE_NOTICE_ID = "annual-pace-2.5.0";

  /** インストール直後（初回）のときだけ使い方ページを開く。更新では開かない */
  function shouldOpenGuideOnInstalled(reason) {
    return reason === "install";
  }

  /**
   * 出す吹き出しを決める。"usage" | "change" | null
   * 今の年度は見ない。前の年度だけ取得した人も「取得したことがある」に入る
   */
  function pickNotice({ enabled, processing, savedYears, dismissed }) {
    if (!enabled || processing) return null;
    if (savedYears.length === 0) return "usage";
    if (dismissed.includes(CHANGE_NOTICE_ID)) return null;
    return "change";
  }

  /** 閉じた記録にIDを足した新しい配列を返す（重複させず、元の配列は変えない） */
  function withDismissed(dismissed, id) {
    return dismissed.includes(id) ? [...dismissed] : [...dismissed, id];
  }

  return { CHANGE_NOTICE_ID, shouldOpenGuideOnInstalled, pickNotice, withDismissed };
})();

if (typeof module !== "undefined") {
  module.exports = KingOfZangyoNotices;
}
