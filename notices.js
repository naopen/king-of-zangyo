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

  // 更新を促す吹き出しを出すまでの日数。1か月は30日として数える
  const STALE_DAYS = 30;
  const STALE_MS = STALE_DAYS * 24 * 60 * 60 * 1000;

  /** インストール直後（初回）のときだけ使い方ページを開く。更新では開かない */
  function shouldOpenGuideOnInstalled(reason) {
    return reason === "install";
  }

  /**
   * 出す吹き出しを決める。"usage" | "change" | "stale" | null（優先順もこの並び）
   * now（ミリ秒）は呼び出し側が渡す。渡さなければ "stale" にしない
   * 今の年度は見ない。前の年度だけ取得した人も「取得したことがある」に入る
   */
  function pickNotice({
    enabled,
    processing,
    savedYears,
    dismissed,
    latestLastUpdated = null,
    staleDismissedAt = null,
    now = null,
  }) {
    if (!enabled || processing) return null;
    if (savedYears.length === 0) return "usage";
    if (!dismissed.includes(CHANGE_NOTICE_ID)) return "change";
    return isStale(latestLastUpdated, staleDismissedAt, now) ? "stale" : null;
  }

  /**
   * 更新ボタンに赤い点を出すか。吹き出しとは切り離し、最終更新から30日以上たっているあいだは出し続ける
   * 一度も取得していない人の点は、使い方の吹き出しと一緒に描くので、ここでは false
   */
  function shouldShowUpdateDot({ enabled, processing, savedYears, latestLastUpdated = null, now = null }) {
    if (!enabled || processing || savedYears.length === 0) return false;
    return isOutdated(latestLastUpdated, now);
  }

  /** 最終更新から30日以上たち、閉じてからも30日以上たっているか */
  function isStale(latestLastUpdated, staleDismissedAt, now) {
    if (!isOutdated(latestLastUpdated, now)) return false;
    return typeof staleDismissedAt !== "number" || now - staleDismissedAt >= STALE_MS;
  }

  /** 最終更新から30日以上たっているか。最終更新が読めなければ false */
  function isOutdated(latestLastUpdated, now) {
    const updatedAt = parseLastUpdated(latestLastUpdated);
    if (updatedAt === null || typeof now !== "number") return false;
    return now - updatedAt >= STALE_MS;
  }

  /** "YYYY/MM/DD HH:mm"（手元の時刻）をミリ秒にする。読めなければ null */
  function parseLastUpdated(value) {
    if (typeof value !== "string") return null;
    const m = /^(\d{4})\/(\d{1,2})\/(\d{1,2}) (\d{1,2}):(\d{2})$/.exec(value.trim());
    if (!m) return null;
    const [year, month, day, hour, minute] = m.slice(1).map(Number);
    const date = new Date(year, month - 1, day, hour, minute);
    return Number.isNaN(date.getTime()) ? null : date.getTime();
  }

  /** 最終更新の文字列の中から、時刻がいちばん新しいものを返す。読めるものが無ければ null */
  function newestLastUpdated(values) {
    let best = null;
    let bestTime = -Infinity;
    for (const value of values) {
      const time = parseLastUpdated(value);
      if (time !== null && time > bestTime) {
        best = value;
        bestTime = time;
      }
    }
    return best;
  }

  /** 閉じた記録にIDを足した新しい配列を返す（重複させず、元の配列は変えない） */
  function withDismissed(dismissed, id) {
    return dismissed.includes(id) ? [...dismissed] : [...dismissed, id];
  }

  return {
    CHANGE_NOTICE_ID,
    STALE_DAYS,
    shouldOpenGuideOnInstalled,
    pickNotice,
    shouldShowUpdateDot,
    withDismissed,
    newestLastUpdated,
  };
})();

if (typeof module !== "undefined") {
  module.exports = KingOfZangyoNotices;
}
