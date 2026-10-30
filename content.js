// King-of-Zangyo: 残業時間自動計算・表示スクリプト

// ストレージキー定数
const STORAGE_KEY = "kingOfZangyoEnabled";
const STANDARD_HOURS_KEY = "kingOfZangyoStandardHours";
const FISCAL_YEAR_START_KEY = "kingOfZangyoFiscalYearStartMonth";
const LEGEND_KEY = "kingOfZangyoLegendEnabled";
const ANNUAL_DATA_KEY_PREFIX = "kingOfZangyoAnnualData_"; // 年度別キーの接頭辞
const ANNUAL_DATA_YEARS_KEY = "kingOfZangyoAnnualDataYears"; // 保存済み年度リスト
const PROCESSING_STATE_KEY = "kingOfZangyoProcessingState";
const NEXT_DISPLAY_YEAR_KEY = "kingOfZangyoNextDisplayYear";

// DOM要素の一意なID
const OVERTIME_HEADER_ID = "king-of-zangyo-header";
const OVERTIME_CELL_ID = "king-of-zangyo-cell";
const ANNUAL_SECTION_ID = "king-of-zangyo-annual-section";
const MONTHLY_LEGEND_ID = "king-of-zangyo-monthly-legend";
const ANNUAL_LEGEND_ID = "king-of-zangyo-annual-legend";

// 所定労働時間（分）- 設定から読み込まれるまでのデフォルト値
let STANDARD_WORK_MINUTES = 450; // 7.5時間

// 年度開始月（1-12）- 設定から読み込まれるまでのデフォルト値
let FISCAL_YEAR_START_MONTH = 4; // 4月開始

// 残業時間表示状態 - 設定から読み込まれるまでのデフォルト値
let overtimeEnabled = true;

// 凡例表示状態 - 設定から読み込まれるまでのデフォルト値
let legendEnabled = true;

/**
 * 現在のページが対象ページかどうかをチェック
 * @returns {boolean} 対象ページの場合はtrue
 */
function isTargetPage() {
  // セッション切れページやエラーページを除外
  const hasErrorDialog = document.querySelector(".htBlock-dialog") !== null;
  const hasSessionError =
    document.body.textContent.includes("セッションが無効です") ||
    document.body.textContent.includes("セッションが切れました");

  if (hasErrorDialog || hasSessionError) {
    // セッション切れの場合は処理を中止
    clearProcessingState().catch((err) => {
      console.error("King-of-Zangyo: 処理状態のクリアに失敗しました", err);
    });

    // // セッション切れの場合はアラート表示
    // if (hasSessionError) {
    //   alert("セッションが切れました。再ログインして、もう一度お試しください。");
    // }

    return false;
  }

  const searchParams = new URLSearchParams(window.location.search);
  const pageId = searchParams.get("page_id");

  // 方法1: page_idパラメータをチェック
  if (
    pageId &&
    pageId.includes("/working/monthly_individual_working_list_custom")
  ) {
    return true;
  }

  // 方法2: DOM要素の存在をチェック（page_idがない場合）
  // タイムカード画面特有の要素が存在するかチェック
  const hasSummaryTable =
    document.querySelector(
      "div.htBlock-normalTable table.specific-table_800",
    ) !== null;
  const hasDailyDataTable =
    document.querySelector(".htBlock-adjastableTableF_inner > table") !== null;
  const hasTimecardTitle =
    document.querySelector("h1")?.textContent.includes("タイムカード") || false;

  const isDomMatch = hasSummaryTable && hasDailyDataTable;

  return isDomMatch;
}

/**
 * 設定をchrome.storage.syncから読み込み、グローバル変数に反映する
 *
 * 【重要】この関数は checkAndResumeProcessing() より前に await すること。
 * 再開処理は月ごとにページ遷移するため、設定の反映が後回しになると
 * 当月の残業計算（calculateTotalOvertime）がデフォルトの所定労働時間で
 * 行われ、年間合計に誤った値が混ざる。
 *
 * @returns {Promise<"ok"|"unavailable"|"error">} 読み込み結果
 *   - "ok": 設定を反映済み
 *   - "unavailable": chrome.storage.sync が使えない（デフォルト値のまま）
 *   - "error": 読み込みに失敗した（デフォルト値のまま）
 */
function loadSettings() {
  return new Promise((resolve) => {
    if (
      typeof chrome === "undefined" ||
      !chrome.storage ||
      !chrome.storage.sync
    ) {
      console.warn(
        "King-of-Zangyo: chrome.storage.sync が利用できません。デフォルトでオンにします。",
      );
      resolve("unavailable");
      return;
    }

    chrome.storage.sync.get(
      [STORAGE_KEY, STANDARD_HOURS_KEY, FISCAL_YEAR_START_KEY, LEGEND_KEY],
      (result) => {
        if (chrome.runtime.lastError) {
          console.error(
            "King-of-Zangyo: ストレージ取得エラー:",
            chrome.runtime.lastError,
          );
          resolve("error");
          return;
        }

        // 残業時間表示（デフォルトは「オン」）
        overtimeEnabled =
          result[STORAGE_KEY] !== undefined ? result[STORAGE_KEY] : true;

        // 凡例表示設定（デフォルトは「オン」）
        legendEnabled =
          result[LEGEND_KEY] !== undefined ? result[LEGEND_KEY] : true;

        // 所定労働時間を設定（デフォルトは7.5時間）
        const standardHours =
          result[STANDARD_HOURS_KEY] !== undefined
            ? result[STANDARD_HOURS_KEY]
            : 7.5;
        STANDARD_WORK_MINUTES = standardHours * 60;

        // 年度開始月を設定（デフォルトは4月）
        FISCAL_YEAR_START_MONTH =
          result[FISCAL_YEAR_START_KEY] !== undefined
            ? result[FISCAL_YEAR_START_KEY]
            : 4;

        resolve("ok");
      },
    );
  });
}

/**
 * アプリケーションの初期化
 */
async function initializeApp() {
  try {
    // 対象ページかチェック
    if (!isTargetPage()) {
      return;
    }

    // ダイアログスタイルを先に注入する
    // 【重要】checkAndResumeProcessing() は月ごとにページ遷移して戻ってこないため、
    // これより後ろでの注入は再開中の document では実行されない。
    // 進捗・完了ダイアログが素のまま描画されるのを防ぐため、必ずここで注入する。
    injectDialogStyles();

    // 設定を読み込む（再開処理より前に必ず反映させる）
    const settingsStatus = await loadSettings();

    // 処理状態をチェックして自動再開
    await checkAndResumeProcessing();

    if (settingsStatus === "unavailable") {
      injectOvertimeColumn();
      return;
    }

    if (settingsStatus === "error") {
      // エラーの場合はデフォルトでオンにする
      injectOvertimeColumn();
      await injectAnnualDataSection();
      return;
    }

    // 年別データセクションを注入（データ読み込みのため常に実行）
    await injectAnnualDataSection();

    if (overtimeEnabled) {
      injectOvertimeColumn();
    } else {
      // 非表示設定の場合、年別データセクションを非表示にする
      const annualSection = document.getElementById(ANNUAL_SECTION_ID);
      if (annualSection) {
        annualSection.style.display = "none";
      }
    }
  } catch (error) {
    console.error("King-of-Zangyo: 初期化エラー:", error);
    // エラーが発生してもデフォルトで機能を実行
    injectOvertimeColumn();
  }
}

/**
 * 残業時間列をテーブルに注入する
 */
function injectOvertimeColumn() {
  try {
    // 時間集計テーブルを取得（残業時間を表示する場所）
    const summaryTable = document.querySelector(
      "div.htBlock-normalTable table.specific-table_800",
    );

    if (!summaryTable) {
      console.log("King-of-Zangyo: 時間集計テーブルが見つかりません");
      return;
    }

    const thead = summaryTable.querySelector("thead");
    const tbody = summaryTable.querySelector("tbody");

    if (!thead || !tbody) {
      console.log("King-of-Zangyo: thead または tbody が見つかりません");
      return;
    }

    // 既存の列の幅を固定（初回のみ）
    if (!summaryTable.dataset.columnsFixed) {
      const headerRow = thead.querySelector("tr");
      if (headerRow) {
        const headers = headerRow.querySelectorAll("th");
        headers.forEach((th) => {
          const computedWidth = window.getComputedStyle(th).width;
          th.style.width = computedWidth;
          th.style.minWidth = computedWidth;
        });
      }

      const dataRow = tbody.querySelector("tr");
      if (dataRow) {
        const cells = dataRow.querySelectorAll("td");
        cells.forEach((td) => {
          const computedWidth = window.getComputedStyle(td).width;
          td.style.width = computedWidth;
          td.style.minWidth = computedWidth;
        });
      }

      summaryTable.dataset.columnsFixed = "true";
    }

    // 既に注入されているかチェック（冪等性の確保）
    const existingHeader = document.getElementById(OVERTIME_HEADER_ID);
    const existingCell = document.getElementById(OVERTIME_CELL_ID);

    // 残業時間を計算
    const overtimeMinutes = calculateTotalOvertime();
    const overtimeText = formatMinutesToTime(overtimeMinutes);

    if (existingHeader && existingCell) {
      // 既に存在する場合は内容を更新するのみ
      existingCell.textContent = overtimeText;
      existingCell.style.backgroundColor =
        getOvertimeBackgroundColor(overtimeMinutes);
      const textColor = getOvertimeTextColor(overtimeMinutes);
      existingCell.style.color = textColor || "";
      // 月間凡例も更新（テーブル幅を再計算して適用）
      const existingLegend = document.getElementById(MONTHLY_LEGEND_ID);
      if (existingLegend) {
        const updatedLegend = buildLegendElement(
          MONTHLY_BANDS,
          getMonthlyBandIndex(overtimeMinutes),
          MONTHLY_LEGEND_ID,
        );
        updatedLegend.style.width = summaryTable.offsetWidth + "px";
        updatedLegend.style.display =
          overtimeEnabled && legendEnabled ? "" : "none";
        existingLegend.replaceWith(updatedLegend);
      }
      return;
    }

    // ヘッダー行に新しい列を追加
    const headerRow = thead.querySelector("tr");
    if (headerRow) {
      const th = document.createElement("th");
      th.id = OVERTIME_HEADER_ID;
      th.appendChild(document.createTextNode("現時点の目安残業"));
      th.appendChild(document.createElement("br"));
      th.appendChild(document.createTextNode("(法定休日を除く)"));
      th.style.textAlign = "center";
      th.style.fontWeight = "bold";
      th.style.fontSize = "13px";
      th.style.width = "120px";
      th.style.minWidth = "120px";
      if (!overtimeEnabled) th.style.display = "none";
      headerRow.appendChild(th);
    }

    // データ行に新しいセルを追加
    const dataRow = tbody.querySelector("tr");
    if (dataRow) {
      const td = document.createElement("td");
      td.id = OVERTIME_CELL_ID;
      td.textContent = overtimeText;
      td.style.textAlign = "center";
      td.style.fontWeight = "bold";
      td.style.fontSize = "14px";
      td.style.width = "120px";
      td.style.minWidth = "120px";
      td.style.backgroundColor = getOvertimeBackgroundColor(overtimeMinutes);
      const textColor = getOvertimeTextColor(overtimeMinutes);
      if (textColor) {
        td.style.color = textColor;
      }
      if (!overtimeEnabled) td.style.display = "none";
      dataRow.appendChild(td);
    }

    // 月間凡例をテーブルの直後に挿入（テーブル幅に制限して右寄せ）
    const legendDiv = buildLegendElement(
      MONTHLY_BANDS,
      getMonthlyBandIndex(overtimeMinutes),
      MONTHLY_LEGEND_ID,
    );
    legendDiv.style.width = summaryTable.offsetWidth + "px";
    legendDiv.style.display = overtimeEnabled && legendEnabled ? "" : "none";
    summaryTable.insertAdjacentElement("afterend", legendDiv);
  } catch (error) {
    console.error("King-of-Zangyo: 列の注入中にエラーが発生しました:", error);
  }
}

/**
 * 時間文字列を分に変換する
 * @param {string} timeStr - "HH.MM"または"HH:MM"形式の時間文字列
 * @returns {number} 分、パース失敗時は0
 */
function parseHHMMToMinutes(timeStr) {
  const match = timeStr.match(/(\d+)[:.](\d{2})/);
  if (!match) return 0;
  return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
}

/**
 * スケジュールテキストから半日休暇の期待勤務時間（分）を取得する
 * @param {string} scheduleText - スケジュール欄のテキスト
 * @returns {{isHalfDayLeave: boolean, isAM: boolean, isPM: boolean, expectedMinutes: number}}
 */
function getHalfDayLeaveInfo(scheduleText) {
  const isAM = scheduleText.includes("AM有休");
  const isPM = scheduleText.includes("PM有休");
  const isHalfDayLeave = isAM || isPM;

  if (!isHalfDayLeave) {
    return {
      isHalfDayLeave: false,
      isAM: false,
      isPM: false,
      expectedMinutes: 0,
    };
  }

  // パターンを完全一致で判定
  let pattern = 2; // デフォルトはパターン２
  if (scheduleText.includes("フレックス（８始業１２休憩）")) {
    pattern = 1;
  } else if (scheduleText.includes("フレックス（９始業１２休憩）")) {
    pattern = 2;
  } else if (scheduleText.includes("フレックス（１０始業１３休憩）")) {
    pattern = 3;
  }

  // パターンと休暇種別に基づいて期待勤務時間を決定
  let expectedMinutes;
  if (pattern === 1) {
    expectedMinutes = isAM ? 210 : 240; // AM有休→3.5h, PM有休→4h
  } else {
    // パターン２、３は同じ
    expectedMinutes = isAM ? 270 : 180; // AM有休→4.5h, PM有休→3h
  }

  return { isHalfDayLeave, isAM, isPM, expectedMinutes };
}

/**
 * 日数集計の有休セルから月全体の時間休合計（分）を取得する
 * @returns {number} 時間休の合計時間（分）
 */
function getHourlyLeaveTotalFromSummary() {
  // 1. 有休セル（日数集計）から使用日数と時間休Hを取得
  const labels = document.querySelectorAll(
    ".specific-daysCount_1 label.holiday_count",
  );
  let paidLeaveDays = 0;
  let paidLeaveHours = 0;
  for (const label of labels) {
    if (label.textContent.trim() === "有休") {
      const valueDiv = label.nextElementSibling;
      if (!valueDiv) break;
      const text = valueDiv.textContent.trim();
      // "2.0 / 0H (残 10.5 / 2H)" → 日数と時間を抽出
      const daysMatch = text.match(/^([\d.]+)/);
      if (daysMatch) paidLeaveDays = parseFloat(daysMatch[1]);
      const hoursMatch = text.match(/\/\s*(\d+)H/);
      if (hoursMatch) paidLeaveHours = parseInt(hoursMatch[1], 10);
      break;
    }
  }
  if (paidLeaveDays === 0 && paidLeaveHours === 0) return 0;

  // 2. 日別データから全日有休・半日有休の日数をカウント
  const dailyDataTable = document.querySelector(
    ".htBlock-adjastableTableF_inner > table",
  );
  let fullDayLeaveCount = 0;
  if (dailyDataTable) {
    const rows = dailyDataTable.querySelectorAll("tbody tr");
    rows.forEach((row) => {
      const scheduleCell = row.querySelector("td.schedule");
      const scheduleCellText = scheduleCell?.textContent.trim() || "";
      if (!scheduleCellText.includes("有休")) return;
      const halfDayInfo = getHalfDayLeaveInfo(scheduleCellText);
      if (halfDayInfo.isHalfDayLeave) {
        fullDayLeaveCount += 0.5; // 半日有休は0.5日
      } else if (!scheduleCellText.includes("時間休")) {
        fullDayLeaveCount += 1; // 全日有休は1日（時間休を除く）
      }
    });
  }

  // 3. 時間休の合計を算出
  const hourlyLeaveDays = Math.max(0, paidLeaveDays - fullDayLeaveCount);
  // 換算レート: 所定労働時間を1時間単位に切り上げ（労基法施行規則24条の4）
  const hourlyLeavePerDay = Math.ceil(STANDARD_WORK_MINUTES / 60) * 60;

  return hourlyLeaveDays * hourlyLeavePerDay + paidLeaveHours * 60;
}

/**
 * 月内に管理職の勤務があるかどうかを判定
 * @returns {boolean} 管理職の場合true
 */
function isManagerMonth() {
  const dailyDataTable = document.querySelector(
    ".htBlock-adjastableTableF_inner > table",
  );
  if (!dailyDataTable) {
    return false;
  }
  const tbody = dailyDataTable.querySelector("tbody");
  if (!tbody) {
    return false;
  }
  const rows = tbody.querySelectorAll("tr");
  for (const row of rows) {
    const scheduleCell = row.querySelector("td.schedule");
    const scheduleCellText = scheduleCell?.textContent.trim() || "";
    if (scheduleCellText.includes("管理職")) {
      return true;
    }
  }
  return false;
}

/**
 * 累計残業時間を計算する
 * @returns {number} 累計残業時間（分）
 */
function calculateTotalOvertime() {
  // 月内に管理職の勤務があるかを事前に判定
  const isManager = isManagerMonth();
  // 日別データテーブルを取得
  const dailyDataTable = document.querySelector(
    ".htBlock-adjastableTableF_inner > table",
  );

  if (!dailyDataTable) {
    return 0;
  }

  const tbody = dailyDataTable.querySelector("tbody");
  if (!tbody) {
    return 0;
  }

  // 年月情報を取得（ページ上部のh2から）
  const currentYearMonth = getCurrentYearMonth();
  if (!currentYearMonth) {
    return 0;
  }

  const rows = tbody.querySelectorAll("tr");
  let totalOvertimeMinutes = 0;

  // 今日の日付を取得
  const today = new Date();
  today.setHours(0, 0, 0, 0); // 時間をリセット

  rows.forEach((row) => {
    // 日付を取得（2番目のtd - 1番目は申請ボタン）
    const dateCell = row.querySelector(
      "td.htBlock-scrollTable_day, td:nth-child(2)",
    );
    const dateCellText = dateCell?.textContent.trim();

    if (!dateCellText) {
      return; // 日付がない場合はスキップ
    }

    // 日付をパース（年月情報を使用）
    const rowDate = parseDate(
      dateCellText,
      currentYearMonth.year,
      currentYearMonth.month,
    );
    if (!rowDate || rowDate > today) {
      return; // 明日以降のデータはスキップ
    }

    const isToday = rowDate.getTime() === today.getTime();

    // スケジュール（5番目のtd）を取得
    const scheduleCell = row.querySelector("td.schedule, td:nth-child(5)");
    const scheduleCellText = scheduleCell?.textContent.trim() || "";

    // 半日休暇情報を取得
    const halfDayInfo = getHalfDayLeaveInfo(scheduleCellText);

    // 全日有休の場合のみスキップ（半日休暇・時間休は含めない）
    const isFullDayLeave =
      scheduleCellText.includes("有休") &&
      !halfDayInfo.isHalfDayLeave &&
      !scheduleCellText.includes("時間休");
    if (isFullDayLeave) {
      return;
    }

    // 勤務日種別（6番目のtd）を取得
    const workdayTypeCell = row.querySelector(
      "td.work_day_type, td:nth-child(6)",
    );
    const workdayTypeCellText = workdayTypeCell?.textContent.trim() || "";

    // 法定休日（多くは日曜日）は時間外労働の上限にカウントしない
    const isLegalHoliday = workdayTypeCellText === "法定休日";
    if (isLegalHoliday) {
      return; // この日はスキップ
    }

    // 平日かどうかを判定
    const isWeekday = workdayTypeCellText === "平日";
    // 法定外休日かどうかを判定（土曜日・祝日）
    const isNonLegalHoliday = workdayTypeCellText === "法定外休日";

    // 平日でも法定外休日でもない場合はスキップ
    if (!isWeekday && !isNonLegalHoliday) {
      return;
    }

    // 実働時間（class="custom12"を優先）を取得
    const actualWorkCell = row.querySelector("td.custom12");
    const actualWorkText = actualWorkCell?.textContent.trim() || "";

    if (!actualWorkText || actualWorkText === "") {
      return; // 実働時間が空の場合はスキップ
    }

    // 実働時間を分に変換
    const actualWorkMinutes = parseHHMMToMinutes(actualWorkText);
    if (actualWorkMinutes === 0) {
      return; // パース失敗時はスキップ
    }

    // 残業時間を計算
    // 平日（全日勤務）：実働時間 - 所定労働時間（450分 = 7.5時間）
    // 平日（半日休暇）：実働時間 - パターン別の期待勤務時間
    // 休日（土日祝）：実働時間 - 0分 = 実働時間（全て残業扱い）
    let standardMinutes;
    if (!isWeekday) {
      // 法定外休日（土曜・祝日）の場合
      standardMinutes = 0;
    } else if (halfDayInfo.isHalfDayLeave) {
      // 半日休暇の場合：パターン別の期待勤務時間を基準にする
      standardMinutes = halfDayInfo.expectedMinutes;
    } else {
      // 通常の平日勤務
      standardMinutes = STANDARD_WORK_MINUTES;
    }
    const dailyOvertimeMinutes = actualWorkMinutes - standardMinutes;

    // 当日かつ所定時間未達（勤務中）の場合はスキップ
    // 日種別ごとの standardMinutes を基準にするため、休日・半休でも正しく動作する
    if (isToday && dailyOvertimeMinutes < 0) {
      return;
    }

    // 累計に追加
    // 管理職の場合はマイナスを0に（所定時間の概念がないため）
    if (isManager && dailyOvertimeMinutes < 0) {
      totalOvertimeMinutes += 0;
    } else {
      totalOvertimeMinutes += dailyOvertimeMinutes;
    }
  });

  // 有休セルから時間休合計を取得し、一括加算
  totalOvertimeMinutes += getHourlyLeaveTotalFromSummary();

  return Math.round(totalOvertimeMinutes);
}

/**
 * 月別データテーブルから「①平日時間外＋法定外休勤務」の残業時間を取得する
 * 前月以前のデータ取得に使用する（公式値を使用）
 * @returns {number} 残業時間（分）
 */
function getMonthlyOvertimeFromSummary() {
  // table.specific-table_800 の tbody td.custom2 から値を取得
  const summaryTable = document.querySelector("table.specific-table_800");
  if (!summaryTable) {
    console.warn(
      "King-of-Zangyo: 月別データテーブル（table.specific-table_800）が見つかりません",
    );
    return 0;
  }

  const overtimeCell = summaryTable.querySelector("tbody td.custom2");
  if (!overtimeCell) {
    console.warn("King-of-Zangyo: 残業時間セル（td.custom2）が見つかりません");
    return 0;
  }

  const overtimeText = overtimeCell.textContent.trim();
  if (!overtimeText || overtimeText === "") {
    return 0;
  }

  // 値の形式: "37.23" = 37時間23分 → 分に変換
  const totalMinutes = parseHHMMToMinutes(overtimeText);
  if (totalMinutes === 0 && overtimeText !== "0.00") {
    console.warn(
      `King-of-Zangyo: 残業時間のパースに失敗しました: "${overtimeText}"`,
    );
  }

  return totalMinutes;
}

/**
 * 現在の年月を取得する
 * @returns {{year: number, month: number}|null} 年月オブジェクト
 */
function getCurrentYearMonth() {
  // ページ上部のh2から年月を取得（例: "2025/11/01(土) ～ 2025/11/30(日)"）
  const h2Element = document.querySelector(".htBlock-mainContents h2 span");
  if (!h2Element) {
    return null;
  }

  const h2Text = h2Element.textContent.trim();
  const match = h2Text.match(/(\d{4})\/(\d{1,2})/);

  if (!match) {
    return null;
  }

  return {
    year: parseInt(match[1], 10),
    month: parseInt(match[2], 10),
  };
}

/**
 * 日付文字列をDateオブジェクトに変換する
 * @param {string} dateStr - 日付文字列（例: "11/01（土）" または "2025/11/01(土)"）
 * @param {number} defaultYear - デフォルトの年
 * @param {number} defaultMonth - デフォルトの月
 * @returns {Date|null} Dateオブジェクト、パース失敗時はnull
 */
function parseDate(dateStr, defaultYear, defaultMonth) {
  // パターン1: 年月日形式（例: "2025/11/01(土)"）
  let match = dateStr.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);

  if (match) {
    const year = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1; // 月は0始まり
    const day = parseInt(match[3], 10);
    return new Date(year, month, day);
  }

  // パターン2: 月日のみ形式（例: "11/01（土）"）
  match = dateStr.match(/(\d{1,2})\/(\d{1,2})/);

  if (match) {
    const month = parseInt(match[1], 10);
    const day = parseInt(match[2], 10);

    // 表示されている月と異なる場合は年をまたいでいる可能性がある
    // 例: 12月のページで1月が表示される場合（年末年始）
    let year = defaultYear;
    if (month < defaultMonth) {
      year += 1; // 翌年
    }

    return new Date(year, month - 1, day); // 月は0始まり
  }

  return null;
}

/**
 * 分を HH:MM 形式の文字列に変換する
 * @param {number} minutes - 分
 * @returns {string} HH:MM 形式の文字列（例: "2:30", "-1:15"）
 */
function formatMinutesToTime(minutes) {
  const isNegative = minutes < 0;
  const absoluteMinutes = Math.abs(minutes);

  const hours = Math.floor(absoluteMinutes / 60);
  const mins = absoluteMinutes % 60;

  const formattedTime = `${hours}:${String(mins).padStart(2, "0")}`;

  return isNegative ? `-${formattedTime}` : formattedTime;
}

// 月間の閾値バンド定義（年間は annual-pace.js が年度の残りから作る）
const MONTHLY_BANDS = [
  { color: "#c8e6c9", label: "～30h" },
  { color: "#fff59d", label: "30～40h" },
  { color: "#ef9a9a", label: "40～45h" },
  { color: "#d32f2f", label: "45～75h(36協定超過)" },
  { color: "#990099", label: "75h+(36協定特別条項超過)" },
];

function getMonthlyBandIndex(overtimeMinutes) {
  const h = overtimeMinutes / 60;
  if (h < 30) return 0;
  if (h < 40) return 1;
  if (h < 45) return 2;
  if (h < 75) return 3;
  return 4;
}

/**
 * 年間凡例の置き場所だけを作る（中身は描かない）
 * 保存データを読み込んだら updateAnnualDataDisplay がこの要素を本物の凡例に差し替える
 */
function buildEmptyAnnualLegend() {
  const div = document.createElement("div");
  div.id = ANNUAL_LEGEND_ID;
  return div;
}

/**
 * 閾値凡例のdiv要素を構築する
 * currentBandIndex が -1 の場合はすべてのバンドを等価表示（未取得状態）
 * textAlign: "right"（月間・セルが右端）または "left"（年間・セルが左端）
 */
function buildLegendElement(
  bands,
  currentBandIndex,
  legendId,
  textAlign = "right",
) {
  const div = document.createElement("div");
  div.id = legendId;
  div.style.textAlign = textAlign;
  div.style.fontSize = "11px";
  div.style.color = "#555";
  div.style.marginTop = "3px";
  div.style.lineHeight = "1.4";

  bands.forEach(({ color, label, title }, i) => {
    const span = document.createElement("span");
    span.style.display = "inline-flex";
    span.style.alignItems = "center";
    span.style.marginLeft = "8px";
    if (title) {
      span.title = title;
    }
    const isCurrent = currentBandIndex >= 0 && i === currentBandIndex;
    const isUnresolved = currentBandIndex < 0;
    if (isCurrent) {
      span.style.fontWeight = "bold";
      span.style.borderBottom = "1.5px solid rgba(0,0,0,0.4)";
    }

    const swatch = document.createElement("span");
    swatch.style.display = "inline-block";
    swatch.style.width = "9px";
    swatch.style.height = "9px";
    swatch.style.backgroundColor = color;
    swatch.style.border = "1px solid rgba(0,0,0,0.15)";
    swatch.style.flexShrink = "0";
    swatch.style.marginRight = "2px";
    if (isCurrent) {
      swatch.style.boxShadow = "0 0 0 1.5px rgba(0,0,0,0.35)";
    }

    const labelSpan = document.createElement("span");
    labelSpan.textContent = label;
    labelSpan.style.color = isUnresolved || isCurrent ? "#555" : "#999";

    span.appendChild(swatch);
    span.appendChild(labelSpan);
    div.appendChild(span);
  });

  return div;
}

/**
 * 残業時間に応じた背景色を取得する
 * @param {number} overtimeMinutes - 残業時間（分）
 * @returns {string} 背景色（16進数カラーコード）
 */
function getOvertimeBackgroundColor(overtimeMinutes) {
  const overtimeHours = overtimeMinutes / 60;

  if (overtimeHours < 30) {
    // 0~29.9999時間：緑色
    return "#c8e6c9";
  } else if (overtimeHours < 40) {
    // 30~39.9999時間：黄色
    return "#fff59d";
  } else if (overtimeHours < 45) {
    // 40~44.9999時間：赤色
    return "#ef9a9a";
  } else if (overtimeHours < 75) {
    // 45~74.9999時間：クリムゾンレッド
    return "#d32f2f";
  } else {
    // 75時間以上：紫色（特別条項上限）
    return "#990099";
  }
}

/**
 * 月間残業時間に応じた文字色を取得する
 * @param {number} overtimeMinutes - 残業時間（分）
 * @returns {string} 文字色（16進数カラーコード）、変更不要の場合は空文字
 */
function getOvertimeTextColor(overtimeMinutes) {
  const overtimeHours = overtimeMinutes / 60;
  // 濃赤（45h以上）と紫（75h以上）の上は、白よりギラつかないオフホワイトにする（年間と同じ）
  if (overtimeHours >= 45) {
    return "#f5f5f5";
  }
  return "";
}

/**
 * 残業時間列の表示/非表示を切り替える
 * @param {boolean} enabled - 表示する場合はtrue、非表示にする場合はfalse
 */
function toggleOvertimeDisplay(enabled) {
  overtimeEnabled = enabled; // グローバル変数を更新
  const header = document.getElementById(OVERTIME_HEADER_ID);
  const cell = document.getElementById(OVERTIME_CELL_ID);
  const annualSection = document.getElementById(ANNUAL_SECTION_ID);
  const monthlyLegend = document.getElementById(MONTHLY_LEGEND_ID);
  const annualLegend = document.getElementById(ANNUAL_LEGEND_ID);

  if (enabled) {
    // 列が存在しない場合は新規作成
    if (!header || !cell) {
      injectOvertimeColumn();
    } else {
      // 既に存在する場合は表示
      header.style.display = "";
      cell.style.display = "";
    }
    // 年別データセクションを表示
    if (annualSection) {
      annualSection.style.display = "";
    }
    // 凡例は overtimeEnabled && legendEnabled の両方が満たされる場合のみ表示
    const showLegend = legendEnabled;
    if (monthlyLegend) monthlyLegend.style.display = showLegend ? "" : "none";
    if (annualLegend) annualLegend.style.display = showLegend ? "" : "none";
  } else {
    // 非表示
    if (header) header.style.display = "none";
    if (cell) cell.style.display = "none";
    if (annualSection) annualSection.style.display = "none";
    // 残業時間表示OFFなら凡例も強制非表示（legendEnabledは変えない）
    if (monthlyLegend) monthlyLegend.style.display = "none";
    if (annualLegend) annualLegend.style.display = "none";
  }
  // 吹き出しの出し分け（OFFなら消え、ONなら出し直す）
  renderNotices().catch((error) => {
    console.error("King-of-Zangyo: 吹き出し表示エラー:", error);
  });
}

/**
 * 凡例の表示/非表示を切り替える
 * @param {boolean} enabled - 表示する場合はtrue
 */
function toggleLegendDisplay(enabled) {
  legendEnabled = enabled; // グローバル変数を更新（シングルソースオブトゥルース）
  const monthlyLegend = document.getElementById(MONTHLY_LEGEND_ID);
  const annualLegend = document.getElementById(ANNUAL_LEGEND_ID);
  // 残業時間表示がOFFの場合は凡例設定に関わらず非表示のまま
  const show = overtimeEnabled && enabled;

  if (monthlyLegend) monthlyLegend.style.display = show ? "" : "none";
  if (annualLegend) annualLegend.style.display = show ? "" : "none";
  // 凡例の有無で吹き出しの位置が変わるので、置き直す
  renderNotices().catch((error) => {
    console.error("King-of-Zangyo: 吹き出し表示エラー:", error);
  });
}

// メッセージリスナーを設定（ポップアップからのメッセージを受信）
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "toggleDisplay") {
    toggleOvertimeDisplay(request.enabled);
    sendResponse({ success: true });
  } else if (request.action === "toggleLegendDisplay") {
    toggleLegendDisplay(request.enabled);
    sendResponse({ success: true });
  } else if (request.action === "updateStandardHours") {
    // 所定労働時間を更新
    STANDARD_WORK_MINUTES = request.hours * 60;
    // 残業時間を再計算して表示を更新
    injectOvertimeColumn();
    sendResponse({ success: true });
  } else if (request.action === "updateFiscalYearStart") {
    // 年度開始月を更新
    FISCAL_YEAR_START_MONTH = request.month;

    // 現在年度を再計算
    const currentFiscalYear = determineFiscalYear(
      new Date().getFullYear(),
      new Date().getMonth() + 1,
      FISCAL_YEAR_START_MONTH,
    );

    // 表示をリセット（全データ削除済みのため）
    updateAnnualDataDisplay(null, currentFiscalYear);

    // ドロップダウンをリセット（現在年度のみ表示）
    const fiscalYearSelect = document.getElementById(
      "annual-fiscal-year-select",
    );
    if (fiscalYearSelect) {
      fiscalYearSelect.replaceChildren();
      const option = document.createElement("option");
      option.value = currentFiscalYear;
      option.textContent = `${currentFiscalYear}年度`;
      option.selected = true;
      fiscalYearSelect.appendChild(option);
    }

    sendResponse({ success: true });
  }

  return true; // 非同期レスポンスを許可
});

/**
 * 年別データセクションをページに注入する
 * 「月別データ」セクション全体の右側に配置
 */
async function injectAnnualDataSection() {
  try {
    // 既にセクションが注入されている場合はスキップ
    if (document.getElementById(ANNUAL_SECTION_ID)) {
      return;
    }

    // 「月別データ」の見出しを探す
    const headings = document.querySelectorAll("h4.htBlock-box_subTitle");
    let monthlyDataHeading = null;

    for (const heading of headings) {
      if (heading.textContent && heading.textContent.includes("月別データ")) {
        monthlyDataHeading = heading;
        break;
      }
    }

    if (!monthlyDataHeading) {
      console.log("King-of-Zangyo: 月別データセクションが見つかりません");
      return;
    }

    // 月別データのh4の親要素を取得
    const parentElement = monthlyDataHeading.parentElement;

    if (!parentElement) {
      console.log(
        "King-of-Zangyo: 月別データセクションの親要素が見つかりません",
      );
      return;
    }

    // 「月別データ」と「日別データ」の両方を収集
    // 「月別データ」見出しから、「日別データ」とその後続要素を含めて収集
    const elementsToWrap = [];
    let currentElement = monthlyDataHeading;
    let h4Count = 0; // h4の出現回数をカウント

    while (currentElement) {
      // h4要素をカウント
      if (currentElement.tagName === "H4") {
        h4Count++;
        // 3つ目のh4が出現したら終了（月別データ、日別データの次）
        if (h4Count > 2) {
          break;
        }
      }

      elementsToWrap.push(currentElement);
      currentElement = currentElement.nextElementSibling;
    }

    // 横並びレイアウト用のフレックスコンテナを作成
    const flexContainer = document.createElement("div");
    flexContainer.style.display = "flex";
    flexContainer.style.gap = "0";
    flexContainer.style.alignItems = "flex-start";
    flexContainer.style.minWidth = "1200px";
    flexContainer.style.position = "relative";
    flexContainer.style.overflow = "visible";

    // 左側：月別データと日別データ用のコンテナ
    const leftContainer = document.createElement("div");
    leftContainer.style.flex = "0 1 auto";
    leftContainer.style.minWidth = "0";

    // 月別データのh4の位置にフレックスコンテナを挿入（要素を移動する前に）
    parentElement.insertBefore(flexContainer, monthlyDataHeading);

    // 左側コンテナをフレックスコンテナに追加
    flexContainer.appendChild(leftContainer);

    // 収集した要素を左側コンテナに移動
    elementsToWrap.forEach((element) => {
      leftContainer.appendChild(element);
    });

    // 年別データセクションを作成
    const sectionContainer = document.createElement("div");
    sectionContainer.id = ANNUAL_SECTION_ID;
    sectionContainer.style.position = "absolute";
    sectionContainer.style.left = "975px";
    sectionContainer.style.top = "0";

    // ★年度選択ドロップダウンを作成
    const currentFiscalYear = determineFiscalYear(
      new Date().getFullYear(),
      new Date().getMonth() + 1,
      FISCAL_YEAR_START_MONTH,
    );

    // 取得完了後に戻ってきた場合は、更新した年度を表示する
    const displayYear = (await takeNextDisplayYear()) ?? currentFiscalYear;

    // 保存済み年度リストを取得
    const savedYears = await getSavedFiscalYears();

    // 現在年度・表示年度・保存済み年度を結合（重複削除、降順）
    const allYears = [
      ...new Set([currentFiscalYear, displayYear, ...savedYears]),
    ].sort((a, b) => b - a);

    const fiscalYearSelect = document.createElement("select");
    fiscalYearSelect.id = "annual-fiscal-year-select";
    fiscalYearSelect.className = "htBlock-selectmenu"; // 既存UIに準拠

    // 年度オプションを生成
    allYears.forEach((year) => {
      const option = document.createElement("option");
      option.value = year;
      option.textContent = `${year}年度`;
      if (year === displayYear) {
        option.selected = true;
      }
      fiscalYearSelect.appendChild(option);
    });

    // 変更イベントリスナー
    fiscalYearSelect.addEventListener("change", handleFiscalYearChange);

    // 更新ボタン
    const updateButton = document.createElement("button");
    updateButton.type = "button";
    updateButton.className = "htBlock-buttonPrimary htBlock-buttonM";
    updateButton.style.position = "relative";
    // updateButton.style.top = "-1px";
    const buttonSpan = document.createElement("span");
    buttonSpan.textContent = "更新";
    updateButton.appendChild(buttonSpan);
    updateButton.addEventListener("click", handleAnnualUpdateButtonClick);

    // 使い方ページへの入口（更新ボタンの直後に置く）
    const guideLink = document.createElement("a");
    guideLink.className = "kot-zangyo-guide-link";
    guideLink.href = chrome.runtime.getURL("guide/index.html");
    guideLink.target = "_blank";
    guideLink.rel = "noopener";
    guideLink.title = "使い方";
    const guideMark = document.createElement("span");
    guideMark.className = "kot-zangyo-guide-link-mark";
    guideMark.textContent = "?";
    guideLink.appendChild(guideMark);
    guideLink.appendChild(document.createTextNode("使い方"));

    // セクション見出しとボタンを横並びにするコンテナ
    const titleContainer = document.createElement("div");
    titleContainer.style.display = "flex";
    titleContainer.style.alignItems = "baseline";
    titleContainer.style.gap = "4px";

    // セクション見出し
    const sectionTitle = document.createElement("h4");
    sectionTitle.className = "htBlock-box_subTitle";
    sectionTitle.textContent = "年別データ";
    sectionTitle.style.marginRight = "8px"; // 見出しの右側に自動マージンを追加

    // タイトルコンテナに見出し、ドロップダウン、ボタンを追加
    titleContainer.appendChild(sectionTitle);
    titleContainer.appendChild(fiscalYearSelect);
    titleContainer.appendChild(updateButton);
    titleContainer.appendChild(guideLink);

    // 年間集計テーブルの見出し
    const tableCaption = document.createElement("h5");
    tableCaption.className = "htBlock-box_caption";
    tableCaption.id = "annual-table-caption";
    tableCaption.textContent = "年間集計";
    tableCaption.style.marginTop = "8px";

    // テーブルコンテナ
    const tableContainer = document.createElement("div");
    tableContainer.className = "htBlock-normalTable specific-table";

    // テーブル作成
    const table = document.createElement("table");
    table.className = "specific-table_800";
    table.style.tableLayout = "fixed";
    table.style.width = "270px";

    // テーブルヘッダー
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    const headers = [
      { textLines: ["目安年間残業", "(法定休日を除く)"], width: "130px" },
      { textLines: ["最終更新"], width: "150px" },
    ];

    headers.forEach((header) => {
      const th = document.createElement("th");
      const p = document.createElement("p");
      header.textLines.forEach((line, index) => {
        if (index > 0) {
          p.appendChild(document.createElement("br"));
        }
        p.appendChild(document.createTextNode(line));
      });
      th.appendChild(p);
      th.style.textAlign = "center";
      th.style.width = header.width;
      th.style.minWidth = header.width;
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    // テーブルボディ
    const tbody = document.createElement("tbody");
    const dataRow = document.createElement("tr");

    // 年間残業時間セル
    const annualHoursCell = document.createElement("td");
    annualHoursCell.id = "annual-overtime-hours";
    annualHoursCell.style.textAlign = "center";
    annualHoursCell.style.fontWeight = "bold";
    annualHoursCell.style.fontSize = "14px";
    annualHoursCell.style.width = "120px";
    annualHoursCell.style.minWidth = "120px";
    annualHoursCell.textContent = "未取得";

    // 最終更新セル
    const lastUpdatedCell = document.createElement("td");
    lastUpdatedCell.id = "annual-last-updated";
    lastUpdatedCell.style.textAlign = "center";
    // lastUpdatedCell.style.fontWeight = "bold";
    lastUpdatedCell.style.fontSize = "14px";
    lastUpdatedCell.style.width = "120px";
    lastUpdatedCell.style.minWidth = "120px";
    lastUpdatedCell.textContent = "未取得";

    // セルを行に追加
    dataRow.appendChild(annualHoursCell);
    dataRow.appendChild(lastUpdatedCell);

    tbody.appendChild(dataRow);
    table.appendChild(tbody);

    tableContainer.appendChild(table);

    // 年間凡例は、境目が保存データの最終更新日で決まるため、読み込むまで中身を描かない
    const annualLegendDiv = buildEmptyAnnualLegend();

    // セクションを構築
    sectionContainer.appendChild(titleContainer);
    sectionContainer.appendChild(tableCaption);
    sectionContainer.appendChild(tableContainer);
    sectionContainer.appendChild(annualLegendDiv);

    // フレックスコンテナの右側に年別データセクションを配置
    flexContainer.appendChild(sectionContainer);

    // 保存されているデータを読み込んで表示
    loadAndDisplayAnnualData(displayYear);
  } catch (error) {
    console.error(
      "King-of-Zangyo: 年別データセクション注入中にエラーが発生しました:",
      error,
    );
  }
}

/**
 * 年度選択ドロップダウン変更時のハンドラ
 * @param {Event} event - 変更イベント
 */
async function handleFiscalYearChange(event) {
  const selectedYear = parseInt(event.target.value, 10);

  try {
    // 選択された年度のデータを読み込んで表示
    const annualData = await loadAnnualData(selectedYear);
    updateAnnualDataDisplay(annualData, selectedYear);
  } catch (error) {
    console.error("King-of-Zangyo: 年度データ読み込みエラー:", error);
  }
}

/**
 * 保存された年間データを読み込んで表示を更新する（初期表示用）
 * @param {number} fiscalYear - 表示する年度（ドロップダウンで選択した年度と同じ）
 */
async function loadAndDisplayAnnualData(fiscalYear) {
  try {
    const annualData = await loadAnnualData(fiscalYear);

    updateAnnualDataDisplay(annualData, fiscalYear);
  } catch (error) {
    console.error("King-of-Zangyo: 年間データ読み込みエラー:", error);
  }
}

/**
 * 年別データセクションの表示を更新する
 * @param {Object|null} annualData - 年間データオブジェクト（nullの場合は未取得表示）
 * @param {number} fiscalYear - 年度
 */
function updateAnnualDataDisplay(annualData, fiscalYear) {
  try {
    renderAnnualDataSection(annualData, fiscalYear);
  } finally {
    // 早期 return を含むどの経路でも、吹き出しの出し分けをやり直す
    renderNotices().catch((error) => {
      console.error("King-of-Zangyo: 吹き出し表示エラー:", error);
    });
  }
}

/**
 * 年別データセクションの表示を書き換える本体（updateAnnualDataDisplay から呼ぶ）
 */
function renderAnnualDataSection(annualData, fiscalYear) {
  const hoursCell = document.getElementById("annual-overtime-hours");
  const updatedCell = document.getElementById("annual-last-updated");
  const caption = document.getElementById("annual-table-caption");

  if (!hoursCell || !updatedCell || !caption) {
    return;
  }

  // データがない場合は未取得表示
  if (!annualData) {
    caption.textContent = `年間集計 (${fiscalYear}年度)`;
    hoursCell.textContent = "未取得";
    hoursCell.style.backgroundColor = "#F9F9F9"; // 初期値に戻す
    hoursCell.style.color = ""; // 文字色をリセット（360h超後のオフホワイト残留を防ぐ）
    updatedCell.textContent = "未取得";
    // 未取得のあいだは比べる値が無いので、凡例を出さない
    const existingAnnualLegend = document.getElementById(ANNUAL_LEGEND_ID);
    if (existingAnnualLegend) {
      existingAnnualLegend.replaceWith(buildEmptyAnnualLegend());
    }
    return;
  }

  // 年度範囲をタイトルに表示（2025/04 〜 2026/03 形式）
  if (annualData.yearRange) {
    const yearRangeFormatted = annualData.yearRange.replace("-", " 〜 ");
    caption.textContent = `年間集計 (${yearRangeFormatted})`;
  } else {
    caption.textContent = `年間集計 (${fiscalYear}年度)`;
  }

  // 年間残業時間を表示（HH:MM形式）
  const totalMinutes = annualData.totalMinutes || 0;
  hoursCell.textContent = formatMinutesToTime(totalMinutes);

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
}

const DISMISSED_NOTICES_KEY = "kingOfZangyoDismissedNotices";
const NOTICE_BUBBLE_CLASS = "kot-zangyo-notice-bubble";
const NOTICE_DOT_CLASS = "kot-zangyo-notice-dot";

/**
 * 前回描いた吹き出しと赤い点を消す
 */
function clearNoticeElements() {
  document
    .querySelectorAll(`.${NOTICE_BUBBLE_CLASS}, .${NOTICE_DOT_CLASS}`)
    .forEach((element) => element.remove());
}

/**
 * 保存データを読み、出すべきお知らせを1つ選んで描く
 * 呼ぶたびに前の表示を消してから描くので、年度の切り替えでも重複しない
 */
async function renderNotices() {
  const updateButton = document.querySelector(
    `#${ANNUAL_SECTION_ID} button.htBlock-buttonPrimary`,
  );
  if (!updateButton) {
    return;
  }

  const [stored, processingState, savedYears] = await Promise.all([
    new Promise((resolve) => {
      chrome.storage.sync.get([STORAGE_KEY, DISMISSED_NOTICES_KEY], resolve);
    }),
    loadProcessingState(),
    getSavedFiscalYears(),
  ]);

  // 読み込み中に別の描画が走っていても、描く直前に必ず消す
  clearNoticeElements();

  const notice = KingOfZangyoNotices.pickNotice({
    enabled: stored[STORAGE_KEY] !== undefined ? stored[STORAGE_KEY] : true,
    processing: processingState?.isProcessing === true,
    savedYears: savedYears || [],
    dismissed: stored[DISMISSED_NOTICES_KEY] || [],
  });

  if (notice === "usage") {
    renderUsageNotice(updateButton);
  }
  if (notice === "change") {
    renderChangeNotice();
  }
}

/**
 * 変更点の吹き出し（年の凡例の直下、凡例が無いときは年間集計の表の直下）を描く
 * 凡例が無いのは、閾値凡例表示がオフのときと、表示中の年度が未取得のとき
 */
function renderChangeNotice() {
  const section = document.getElementById(ANNUAL_SECTION_ID);
  if (!section) {
    return;
  }

  // 凡例が実際に見えているときだけ凡例を指す。空や非表示なら表を指す
  const legend = document.getElementById(ANNUAL_LEGEND_ID);
  const legendRect = legend ? legend.getBoundingClientRect() : null;
  const anchor =
    legendRect && legendRect.width > 0 && legendRect.height > 0
      ? legend
      : section.querySelector("table");
  if (!anchor) {
    return;
  }

  const bubble = document.createElement("div");
  bubble.className = NOTICE_BUBBLE_CLASS;
  bubble.setAttribute("role", "note");

  const arrow = document.createElement("span");
  arrow.className = "kot-zangyo-notice-arrow";
  bubble.appendChild(arrow);

  const text = document.createElement("p");
  text.className = "kot-zangyo-notice-text";
  text.textContent =
    "背景色の切り替わりが年度の残り日数に応じて変化するようになりました。";
  bubble.appendChild(text);

  const footer = document.createElement("div");
  footer.className = "kot-zangyo-notice-footer";

  const moreLink = document.createElement("a");
  moreLink.className = "kot-zangyo-notice-more";
  // 節は # ではなく ?section= で渡す。# だとブラウザが読み込みと同時にその節へ飛び、
  // 使い方ページの側で先頭からスクロールして見せられないため（guide/guide.js）
  moreLink.href = chrome.runtime.getURL("guide/index.html?section=annual-colors");
  moreLink.target = "_blank";
  moreLink.rel = "noopener";
  moreLink.textContent = "詳しく見る";

  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "kot-zangyo-notice-close";
  closeButton.textContent = "閉じる";
  closeButton.addEventListener("click", async () => {
    try {
      await markChangeNoticeDismissed();
    } catch (error) {
      console.error("King-of-Zangyo: 変更点の吹き出しの既読保存エラー:", error);
    }
    bubble.remove();
  });

  footer.appendChild(moreLink);
  footer.appendChild(closeButton);
  bubble.appendChild(footer);

  // 年別データセクション自身に付ける（絶対配置なので位置の基準になり、
  // セクションを隠すと吹き出しも一緒に隠れる）。実寸の差から座標を出す
  section.appendChild(bubble);
  const sectionRect = section.getBoundingClientRect();
  const anchorRect = anchor.getBoundingClientRect();
  bubble.style.left = `${anchorRect.left - sectionRect.left - section.clientLeft + section.scrollLeft}px`;
  bubble.style.top = `${anchorRect.bottom - sectionRect.top - section.clientTop + section.scrollTop + 8}px`;
}

/**
 * 変更点の吹き出しを「見た」扱いにして保存する
 * 保存済みの配列をその場で読み直してから足す（古い写しで上書きしない）
 */
async function markChangeNoticeDismissed() {
  const stored = await new Promise((resolve) => {
    chrome.storage.sync.get([DISMISSED_NOTICES_KEY], resolve);
  });
  await new Promise((resolve) => {
    chrome.storage.sync.set(
      {
        [DISMISSED_NOTICES_KEY]: KingOfZangyoNotices.withDismissed(
          stored[DISMISSED_NOTICES_KEY] || [],
          KingOfZangyoNotices.CHANGE_NOTICE_ID,
        ),
      },
      resolve,
    );
  });
}

/**
 * 使い方の吹き出し（更新ボタンの赤い点と、直下の説明）を描く
 * 閉じたことは保存しない（次に開けばまた出る）
 * @param {HTMLElement} updateButton - 更新ボタン
 */
function renderUsageNotice(updateButton) {
  const dot = document.createElement("span");
  dot.className = NOTICE_DOT_CLASS;
  updateButton.appendChild(dot);

  const bubble = document.createElement("div");
  bubble.className = NOTICE_BUBBLE_CLASS;
  bubble.setAttribute("role", "note");

  const arrow = document.createElement("span");
  arrow.className = "kot-zangyo-notice-arrow";
  bubble.appendChild(arrow);

  const text = document.createElement("p");
  text.className = "kot-zangyo-notice-text";
  text.textContent =
    "このボタンで年度の残業時間を集計できます。年360hに収まりそうかを色で表示します。";
  bubble.appendChild(text);

  const footer = document.createElement("div");
  footer.className = "kot-zangyo-notice-footer";

  const moreLink = document.createElement("a");
  moreLink.className = "kot-zangyo-notice-more";
  moreLink.href = chrome.runtime.getURL("guide/index.html?section=annual-fetch");
  moreLink.target = "_blank";
  moreLink.rel = "noopener";
  moreLink.textContent = "詳しく見る";

  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "kot-zangyo-notice-close";
  closeButton.textContent = "閉じる";
  closeButton.addEventListener("click", clearNoticeElements);

  footer.appendChild(moreLink);
  footer.appendChild(closeButton);
  bubble.appendChild(footer);

  // 更新ボタンの直下に、レイアウトを押さずに重ねる
  const host = updateButton.offsetParent || updateButton.parentElement;
  bubble.style.left = `${updateButton.offsetLeft}px`;
  bubble.style.top = `${updateButton.offsetTop + updateButton.offsetHeight + 8}px`;
  host.appendChild(bubble);

  // 更新ボタンを押したら案内は役目を終える
  updateButton.addEventListener("click", clearNoticeElements, { once: true });
}

/**
 * 更新ボタンクリックハンドラ
 * 確認ダイアログを表示し、OKなら年間残業時間の取得を開始する
 */
async function handleAnnualUpdateButtonClick() {
  try {
    // 処理中かチェック
    const existingState = await loadProcessingState();
    if (existingState && existingState.isProcessing) {
      const confirmed = await showConfirmDialog(
        "処理中の取得があります",
        "中断された処理を再開しますか？\n（キャンセルすると処理を中止します）",
      );
      if (confirmed) {
        await resumeFetchAnnualOvertime();
      } else {
        await clearProcessingState();
        alert("処理を中止しました。");
      }
      return;
    }

    // ★選択された年度を取得
    const fiscalYearSelect = document.getElementById(
      "annual-fiscal-year-select",
    );
    const selectedFiscalYear = fiscalYearSelect
      ? parseInt(fiscalYearSelect.value, 10)
      : null;

    if (!selectedFiscalYear) {
      alert("年度が選択されていません。");
      return;
    }

    // 所要時間の目安は取得する月数から出す。決め打ちの待機（開始0.5秒＋月の切り替え1秒×(月数−1)＋終了1秒）に
    // ページ読み込み1か月約0.9秒を足す（2026年10月の実測は7か月で11秒強、12か月で約23秒。短く出ないよう遅いほうに合わせた）
    const monthCount = calculateFiscalYearMonths(
      FISCAL_YEAR_START_MONTH,
      selectedFiscalYear,
    ).length;
    const durationNote =
      monthCount > 0
        ? `処理には約${Math.round(monthCount * 1.9 + 0.5)}秒かかります。\n`
        : "";

    // 確認ダイアログ（選択年度を表示）
    const confirmed = await showConfirmDialog(
      `データ更新`,
      `${selectedFiscalYear}年度の年間残業時間データを取得します。\n${durationNote}よろしいですか？`,
    );

    if (!confirmed) {
      return;
    }

    // ★年間残業時間を取得開始（年度を引数に渡す）
    await startFetchAnnualOvertime(selectedFiscalYear);
  } catch (error) {
    console.error("King-of-Zangyo: 更新中にエラーが発生しました", error);
    await clearProcessingState();
    alert(
      `エラーが発生しました: ${error.message}\nページをリロードしてもう一度お試しください。`,
    );
  }
}

/**
 * 指定ミリ秒待機する
 * @param {number} ms - 待機時間（ミリ秒）
 * @returns {Promise<void>}
 */
function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * ページ読み込み完了を待つ（Ajaxページ遷移対応）
 * @param {number} targetYear - 遷移先の年
 * @param {number} targetMonth - 遷移先の月
 * @returns {Promise<void>}
 */
function waitForPageLoad(targetYear, targetMonth) {
  return new Promise((resolve) => {
    const startTime = Date.now();
    const timeout = 10000; // 10秒タイムアウト

    const checkPageUpdated = () => {
      const yearInput = document.querySelector('input[name="year"]');
      const monthInput = document.querySelector('input[name="month"]');

      // 年月が目標値に更新されたかチェック
      if (
        yearInput &&
        monthInput &&
        parseInt(yearInput.value) === targetYear &&
        parseInt(monthInput.value) === targetMonth
      ) {
        // さらに、時間集計テーブルが存在するか確認
        const summaryTable = document.querySelector(
          "div.htBlock-normalTable table.specific-table_800",
        );
        if (summaryTable) {
          resolve();
          return;
        }
      }

      // タイムアウトチェック
      if (Date.now() - startTime > timeout) {
        console.warn(
          `King-of-Zangyo: ページ遷移タイムアウト (${targetYear}年${targetMonth}月)`,
        );
        resolve();
        return;
      }

      // 100msごとに再チェック
      setTimeout(checkPageUpdated, 100);
    };

    // 最初のチェック
    checkPageUpdated();
  });
}

/**
 * 指定月のページに遷移する
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 * @returns {Promise<void>}
 */
async function navigateToMonth(year, month) {
  const yearInput = document.querySelector('input[name="year"]');
  const monthInput = document.querySelector('input[name="month"]');
  const displayButton = document.querySelector("#display_button");

  if (!yearInput || !monthInput || !displayButton) {
    throw new Error("ページ遷移に必要な要素が見つかりません");
  }

  yearInput.value = year;
  monthInput.value = month;
  displayButton.click();

  // ページ読み込み完了を待つ（Ajaxページ遷移対応）
  await waitForPageLoad(year, month);
  // 要件: ページ読み込み完了後1秒待機
  await sleep(1000);
}

/**
 * 指定された年月から年度を判定する
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 * @param {number} fiscalYearStartMonth - 年度開始月（1-12）
 * @returns {number} 年度（年度開始年）
 *
 * 例: 2025年1月、年度開始4月 → 2024年度
 *     2025年4月、年度開始4月 → 2025年度
 */
function determineFiscalYear(year, month, fiscalYearStartMonth) {
  if (month >= fiscalYearStartMonth) {
    return year; // 年度開始月以降なら、その年が年度
  } else {
    return year - 1; // 年度開始月より前なら、前年が年度
  }
}

/**
 * 年度の月リストを計算する（現在月までのみ）
 * @param {number} fiscalYearStartMonth - 年度開始月（1-12）
 * @param {number|null} targetFiscalYear - 対象年度（nullなら現在年度）
 * @returns {Array<{year: number, month: number}>} 年度の月リスト（現在月まで）
 */
function calculateFiscalYearMonths(
  fiscalYearStartMonth,
  targetFiscalYear = null,
) {
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1; // 1-12

  // 対象年度の判定
  let fiscalYear;
  if (targetFiscalYear === null) {
    // 現在年度を自動判定
    fiscalYear = determineFiscalYear(
      currentYear,
      currentMonth,
      fiscalYearStartMonth,
    );
  } else {
    // 指定された年度を使用
    fiscalYear = targetFiscalYear;
  }

  // 年度の12ヶ月を生成
  const months = [];
  for (let i = 0; i < 12; i++) {
    const month = fiscalYearStartMonth + i;
    let targetYear, targetMonth;

    if (month <= 12) {
      targetYear = fiscalYear;
      targetMonth = month;
    } else {
      targetYear = fiscalYear + 1;
      targetMonth = month - 12;
    }

    // ★重要: 未来の月を除外
    const targetDate = new Date(targetYear, targetMonth - 1, 1);
    const currentDate = new Date(currentYear, currentMonth - 1, 1);

    if (targetDate > currentDate) {
      break; // 未来の月なのでループ終了
    }

    months.push({ year: targetYear, month: targetMonth });
  }

  return months;
}

/**
 * 日時を「YYYY/MM/DD HH:mm」形式でフォーマットする
 * @param {Date} date - フォーマットする日時
 * @returns {string} フォーマットされた日時文字列
 */
function formatDateTime(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${year}/${month}/${day} ${hours}:${minutes}`;
}

/**
 * 保存済みの年度リストを取得する
 * @returns {Promise<Array<number>>} 年度の配列（降順）
 */
function getSavedFiscalYears() {
  return new Promise((resolve, reject) => {
    chrome.storage.sync.get([ANNUAL_DATA_YEARS_KEY], (result) => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        const years = result[ANNUAL_DATA_YEARS_KEY] || [];
        resolve(years.sort((a, b) => b - a)); // 降順
      }
    });
  });
}

/**
 * 年度リストに年度を追加する（重複なし）
 * @param {number} fiscalYear - 追加する年度
 * @returns {Promise<void>}
 */
async function addFiscalYearToList(fiscalYear) {
  const years = await getSavedFiscalYears();
  if (!years.includes(fiscalYear)) {
    years.push(fiscalYear);
    years.sort((a, b) => b - a); // 降順
    await saveFiscalYearsList(years);
  }
}

/**
 * 年度リストを保存する
 * @param {Array<number>} years - 年度の配列
 * @returns {Promise<void>}
 */
function saveFiscalYearsList(years) {
  return new Promise((resolve, reject) => {
    chrome.storage.sync.set({ [ANNUAL_DATA_YEARS_KEY]: years }, () => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        resolve();
      }
    });
  });
}

/**
 * 年間データをchrome.storage.syncに保存する
 * @param {Object} annualData - 年間データ
 * @param {number} fiscalYear - 年度（年度開始年）
 * @returns {Promise<void>}
 */
async function saveAnnualData(annualData, fiscalYear) {
  const key = `${ANNUAL_DATA_KEY_PREFIX}${fiscalYear}`;

  // データに年度を追加
  const dataToSave = {
    ...annualData,
    fiscalYear: fiscalYear,
  };

  return new Promise((resolve, reject) => {
    chrome.storage.sync.set({ [key]: dataToSave }, async () => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        // 年度リストに追加
        try {
          await addFiscalYearToList(fiscalYear);
          resolve();
        } catch (error) {
          reject(error);
        }
      }
    });
  });
}

/**
 * 指定年度の年間データをchrome.storage.syncから読み込む
 * @param {number} fiscalYear - 年度（年度開始年）
 * @returns {Promise<Object|null>} 年間データまたはnull
 */
function loadAnnualData(fiscalYear) {
  const key = `${ANNUAL_DATA_KEY_PREFIX}${fiscalYear}`;

  return new Promise((resolve, reject) => {
    chrome.storage.sync.get([key], (result) => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        const annualData = result[key] || null;
        resolve(annualData);
      }
    });
  });
}

/**
 * 指定年度のデータを削除する
 * @param {number} fiscalYear - 年度
 * @returns {Promise<void>}
 */
async function deleteAnnualData(fiscalYear) {
  const key = `${ANNUAL_DATA_KEY_PREFIX}${fiscalYear}`;

  return new Promise(async (resolve, reject) => {
    chrome.storage.sync.remove([key], async () => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        console.log(`King-of-Zangyo: ${fiscalYear}年度のデータを削除しました`);

        // 年度リストから削除
        try {
          const years = await getSavedFiscalYears();
          const updatedYears = years.filter((y) => y !== fiscalYear);
          await saveFiscalYearsList(updatedYears);
          resolve();
        } catch (error) {
          reject(error);
        }
      }
    });
  });
}

/**
 * 処理状態をchrome.storage.localに保存する
 * @param {Object} state - 処理状態オブジェクト
 * @returns {Promise<void>}
 */
function saveProcessingState(state) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.set({ [PROCESSING_STATE_KEY]: state }, () => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        resolve();
      }
    });
  });
}

/**
 * 処理状態をchrome.storage.localから読み込む
 * @returns {Promise<Object|null>} 処理状態オブジェクト（存在しない場合はnull）
 */
function loadProcessingState() {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get([PROCESSING_STATE_KEY], (result) => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        const state = result[PROCESSING_STATE_KEY] || null;
        resolve(state);
      }
    });
  });
}

/**
 * 処理状態をchrome.storage.localから削除する
 * @returns {Promise<void>}
 */
function clearProcessingState() {
  return new Promise((resolve, reject) => {
    chrome.storage.local.remove([PROCESSING_STATE_KEY], () => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        resolve();
      }
    });
  });
}

/**
 * 次のページ読み込みで表示する年度を保存する
 * 取得完了後に元のページへ戻ると拡張が読み込み直されるため、更新した年度をそこへ引き継ぐ
 * @param {number} fiscalYear - 表示する年度
 * @returns {Promise<void>}
 */
function saveNextDisplayYear(fiscalYear) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.set(
      { [NEXT_DISPLAY_YEAR_KEY]: { fiscalYear, savedAt: Date.now() } },
      () => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
        } else {
          resolve();
        }
      },
    );
  });
}

/**
 * 保存した表示年度を1回だけ取り出して消す
 * 保存から1分を過ぎた値は使わない（ページ遷移に失敗して残った値で、後日過去の年度が出るのを防ぐ）
 * @returns {Promise<number|null>} 表示する年度（無い場合はnull）
 */
function takeNextDisplayYear() {
  return new Promise((resolve) => {
    chrome.storage.local.get([NEXT_DISPLAY_YEAR_KEY], (result) => {
      const saved = chrome.runtime.lastError
        ? null
        : result[NEXT_DISPLAY_YEAR_KEY];
      chrome.storage.local.remove([NEXT_DISPLAY_YEAR_KEY], () => {
        const isFresh = saved && Date.now() - saved.savedAt < 60 * 1000;
        resolve(isFresh ? saved.fiscalYear : null);
      });
    });
  });
}

/**
 * ページロード時に処理状態をチェックして自動再開
 */
async function checkAndResumeProcessing() {
  try {
    const state = await loadProcessingState();

    if (!state || !state.isProcessing) {
      // 処理中でない場合は何もしない
      return;
    }

    // タイムアウトチェック（30分）
    const elapsed = Date.now() - state.lastUpdatedAt;
    const TIMEOUT_MS = 30 * 60 * 1000; // 30分

    if (elapsed > TIMEOUT_MS) {
      console.error("King-of-Zangyo: 処理がタイムアウトしました");
      await clearProcessingState();
      alert(
        "年間残業時間の取得がタイムアウトしました（30分経過）。\nページをリロードして再度お試しください。",
      );
      return;
    }

    // 処理を再開
    await resumeFetchAnnualOvertime();
  } catch (error) {
    console.error("King-of-Zangyo: 処理再開中にエラーが発生しました", error);
    await clearProcessingState();
  }
}

/**
 * 年間残業時間の取得を開始する
 * @returns {Promise<void>}
 */
async function startFetchAnnualOvertime(targetFiscalYear) {
  try {
    // 現在のページ情報を保存
    const currentYearInput = document.querySelector('input[name="year"]');
    const currentMonthInput = document.querySelector('input[name="month"]');
    const originalYear = currentYearInput
      ? parseInt(currentYearInput.value)
      : null;
    const originalMonth = currentMonthInput
      ? parseInt(currentMonthInput.value)
      : null;

    // 指定年度の月リストを計算（未来の月は除外される）
    const fiscalYearMonths = calculateFiscalYearMonths(
      FISCAL_YEAR_START_MONTH,
      targetFiscalYear,
    );

    // 未来の月のみの場合（データがない）
    if (fiscalYearMonths.length === 0) {
      alert(`${targetFiscalYear}年度のデータは未来のため取得できません。`);
      return;
    }

    // 初期状態を作成
    const initialState = {
      isProcessing: true,
      startTime: Date.now(),
      currentMonthIndex: 0,
      fiscalYearMonths: fiscalYearMonths,
      monthlyData: {},
      totalMinutes: 0,
      originalYear: originalYear,
      originalMonth: originalMonth,
      fiscalYearStart: FISCAL_YEAR_START_MONTH,
      targetFiscalYear: targetFiscalYear, // ★年度を保存
      lastUpdatedAt: Date.now(),
    };

    // chrome.storage.localに保存
    await saveProcessingState(initialState);

    // 進捗ダイアログを作成・表示（年度情報を渡す）
    const progressDialog = createProgressDialog(targetFiscalYear);
    document.body.appendChild(progressDialog);
    progressDialog.showModal();

    // 最初の月へ遷移
    const firstMonth = fiscalYearMonths[0];

    await sleep(500); // ダイアログ表示を確実にするため少し待機
    await navigateToMonth(firstMonth.year, firstMonth.month);

    // ここでページリロードが発生し、処理が中断される
    // resumeFetchAnnualOvertime() が次のロードで呼ばれる
  } catch (error) {
    console.error(
      "King-of-Zangyo: 年間残業時間の取得開始中にエラーが発生しました",
      error,
    );
    await clearProcessingState();
    throw error;
  }
}

/**
 * 年間残業時間の取得を再開する（ページリロード後）
 * @returns {Promise<void>}
 */
async function resumeFetchAnnualOvertime() {
  try {
    // 処理状態を読み込み
    const state = await loadProcessingState();

    if (!state || !state.isProcessing) {
      console.log("King-of-Zangyo: 再開する処理がありません");
      return;
    }

    const currentMonth = state.fiscalYearMonths[state.currentMonthIndex];

    // 進捗ダイアログを再表示（年度情報を渡す）
    const progressDialog = createProgressDialog(state.targetFiscalYear);
    document.body.appendChild(progressDialog);
    progressDialog.showModal();
    updateProgress(
      state.currentMonthIndex,
      state.fiscalYearMonths.length,
      progressDialog,
    );

    // 現在の月のデータを取得
    // 【設計仕様】当月と前月以前で計算方式が異なる:
    //   当月: calculateTotalOvertime() → 日別データから算出（時間休・半休・管理職補正を含む独自計算）
    //   前月以前: getMonthlyOvertimeFromSummary() → KoT公式値 td.custom2「①平日時間外＋法定外休勤務」を使用
    // 年間合計は計算方式の異なる値の合算になるが、これは意図的な設計。
    const now = new Date();
    const isCurrentMonth =
      currentMonth.year === now.getFullYear() &&
      currentMonth.month === now.getMonth() + 1;

    const overtimeMinutes = isCurrentMonth
      ? calculateTotalOvertime()
      : getMonthlyOvertimeFromSummary();

    // データを更新
    const monthKey = `${currentMonth.year}-${String(
      currentMonth.month,
    ).padStart(2, "0")}`;
    state.monthlyData[monthKey] = overtimeMinutes;
    state.totalMinutes += overtimeMinutes;
    state.currentMonthIndex += 1;
    state.lastUpdatedAt = Date.now();

    // 進捗を保存
    await saveProcessingState(state);

    // 次の月があるかチェック
    if (state.currentMonthIndex < state.fiscalYearMonths.length) {
      // 次の月へ遷移
      const nextMonth = state.fiscalYearMonths[state.currentMonthIndex];
      updateProgress(
        state.currentMonthIndex,
        state.fiscalYearMonths.length,
        progressDialog,
      );

      // 1秒待機してから次の月へ
      await sleep(1000);
      await navigateToMonth(nextMonth.year, nextMonth.month);

      // ここでページリロードが発生し、再度このフローが呼ばれる
    } else {
      // 全ての月が完了
      await completeFetchAnnualOvertime(state, progressDialog);
    }
  } catch (error) {
    console.error("King-of-Zangyo: 処理再開中にエラーが発生しました", error);
    await clearProcessingState();

    // エラーダイアログを表示
    const progressDialog = document.getElementById(
      "kot-zangyo-progress-dialog",
    );
    if (progressDialog) {
      progressDialog.close();
      progressDialog.remove();
    }

    alert(
      `エラーが発生しました: ${error.message}\n処理を中止しました。ページをリロードして再度お試しください。`,
    );
  }
}

/**
 * 年間残業時間の取得を完了する
 * @param {Object} state - 処理状態
 * @param {HTMLDialogElement} progressDialog - 進捗ダイアログ
 * @returns {Promise<void>}
 */
async function completeFetchAnnualOvertime(state, progressDialog) {
  try {
    // プログレスバーを100%に更新
    updateProgress(
      state.fiscalYearMonths.length,
      state.fiscalYearMonths.length,
      progressDialog,
    );

    // 100%表示を1秒間見せてから進捗ダイアログを閉じる
    await sleep(1000);

    if (progressDialog) {
      progressDialog.close();
      progressDialog.remove();
    }

    // 年度を取得
    const targetFiscalYear = state.targetFiscalYear;

    // 年間データを作成
    const totalHours = state.totalMinutes / 60;
    const lastUpdated = formatDateTime(new Date());

    const firstMonth = state.fiscalYearMonths[0];

    // 年度の最終月を計算（開始月の11ヶ月後）
    const endMonthIndex = state.fiscalYearStart + 11;
    const fiscalYearEndMonth =
      endMonthIndex > 12 ? endMonthIndex - 12 : endMonthIndex;
    const fiscalYearEndYear =
      endMonthIndex > 12 ? targetFiscalYear + 1 : targetFiscalYear;

    const yearRange = `${firstMonth.year}/${String(firstMonth.month).padStart(
      2,
      "0",
    )}-${fiscalYearEndYear}/${String(fiscalYearEndMonth).padStart(2, "0")}`;

    const annualData = {
      totalHours: totalHours,
      totalMinutes: state.totalMinutes,
      lastUpdated: lastUpdated,
      fiscalYearStart: state.fiscalYearStart,
      yearRange: yearRange,
      monthlyData: state.monthlyData,
    };

    // 上書き前のデータ（古い形式かどうかを完了ダイアログで使う）
    const previousAnnualData = await loadAnnualData(targetFiscalYear).catch(
      () => null,
    );

    // 指定年度のデータとして保存
    await saveAnnualData(annualData, targetFiscalYear);

    // 取得を完了した人には変更点の吹き出しを出さない（失敗しても完了処理は続ける）
    try {
      await markChangeNoticeDismissed();
    } catch (error) {
      console.error("King-of-Zangyo: 変更点の既読保存エラー:", error);
    }

    // 処理状態をクリア
    await clearProcessingState();

    // ★年度選択ドロップダウンを更新（新しく保存した年度を追加）
    const fiscalYearSelect = document.getElementById(
      "annual-fiscal-year-select",
    );
    if (fiscalYearSelect) {
      const savedYears = await getSavedFiscalYears();
      const currentFiscalYear = determineFiscalYear(
        new Date().getFullYear(),
        new Date().getMonth() + 1,
        FISCAL_YEAR_START_MONTH,
      );
      const allYears = [...new Set([currentFiscalYear, ...savedYears])].sort(
        (a, b) => b - a,
      );

      fiscalYearSelect.replaceChildren();
      allYears.forEach((year) => {
        const option = document.createElement("option");
        option.value = year;
        option.textContent = `${year}年度`;
        if (year === targetFiscalYear) {
          option.selected = true; // ★更新した年度を選択
        }
        fiscalYearSelect.appendChild(option);
      });
    }

    // 年別データセクションの表示を更新（年度を指定）
    updateAnnualDataDisplay(annualData, targetFiscalYear);

    // 結果ダイアログを表示（年度を指定）
    await showResultDialog(annualData, targetFiscalYear, previousAnnualData);

    // 元のページに戻る（読み込み直したあとも、更新した年度を表示する）
    if (state.originalYear && state.originalMonth) {
      await saveNextDisplayYear(targetFiscalYear);
      await navigateToMonth(state.originalYear, state.originalMonth);
    }
  } catch (error) {
    console.error("King-of-Zangyo: 完了処理中にエラーが発生しました", error);
    await clearProcessingState();
    throw error;
  }
}

/**
 * 確認ダイアログを表示する
 * @param {string} title - ダイアログのタイトル
 * @param {string} message - 確認メッセージ
 * @returns {Promise<boolean>} OKならtrue、キャンセルならfalse
 */
function showConfirmDialog(title, message) {
  return new Promise((resolve) => {
    const dialog = document.createElement("dialog");
    dialog.className = "kot-zangyo-dialog";

    const header = document.createElement("div");
    header.className = "kot-zangyo-dialog-header";
    header.textContent = title;

    const body = document.createElement("div");
    body.className = "kot-zangyo-dialog-body";
    body.style.whiteSpace = "pre-line"; // 改行を有効にする
    body.textContent = message;

    const footer = document.createElement("div");
    footer.className = "kot-zangyo-dialog-footer";

    const okButton = document.createElement("button");
    okButton.className = "kot-zangyo-btn kot-zangyo-btn-primary";
    okButton.textContent = "OK";
    okButton.onclick = () => {
      dialog.close();
      dialog.remove();
      resolve(true);
    };

    const cancelButton = document.createElement("button");
    cancelButton.className = "kot-zangyo-btn kot-zangyo-btn-cancel";
    cancelButton.textContent = "キャンセル";
    cancelButton.onclick = () => {
      dialog.close();
      dialog.remove();
      resolve(false);
    };

    footer.appendChild(cancelButton);
    footer.appendChild(okButton);

    dialog.appendChild(header);
    dialog.appendChild(body);
    dialog.appendChild(footer);

    document.body.appendChild(dialog);
    dialog.showModal();
  });
}

/**
 * 進捗ダイアログを作成する
 * @returns {HTMLDialogElement} 作成された進捗ダイアログ要素
 */
function createProgressDialog(fiscalYear) {
  const dialog = document.createElement("dialog");
  dialog.className = "kot-zangyo-dialog";
  dialog.id = "kot-zangyo-progress-dialog";

  const header = document.createElement("div");
  header.className = "kot-zangyo-dialog-header";
  header.textContent = `取得中`;

  const body = document.createElement("div");
  body.className = "kot-zangyo-dialog-body";

  const progressBarContainer = document.createElement("div");
  progressBarContainer.className = "progress-bar-container";

  const progressBar = document.createElement("div");
  progressBar.className = "progress-bar";
  progressBar.id = "kot-zangyo-progress-bar";
  progressBar.style.width = "0%";
  progressBar.textContent = "0%";

  progressBarContainer.appendChild(progressBar);

  const progressText = document.createElement("div");
  progressText.className = "progress-text";
  progressText.id = "kot-zangyo-progress-text";
  progressText.textContent = "処理を開始しています…";

  body.appendChild(progressBarContainer);
  body.appendChild(progressText);

  dialog.appendChild(header);
  dialog.appendChild(body);

  return dialog;
}

/**
 * 進捗を更新する
 * @param {number} currentMonth - 現在の月番号（1-12）
 * @param {number} totalMonths - 総月数（通常12）
 * @param {HTMLDialogElement} progressDialog - 進捗ダイアログ要素
 */
function updateProgress(currentMonth, totalMonths, progressDialog) {
  const percentage = Math.round((currentMonth / totalMonths) * 100);

  const progressBar = progressDialog.querySelector("#kot-zangyo-progress-bar");
  const progressText = progressDialog.querySelector(
    "#kot-zangyo-progress-text",
  );

  if (progressBar) {
    progressBar.style.width = `${percentage}%`;
    progressBar.textContent = `${percentage}%`;
  }

  if (progressText) {
    progressText.textContent = `${currentMonth}/${totalMonths}ヶ月取得中…`;
  }
}

/**
 * 結果ダイアログを表示する
 * @param {Object} annualData - 取り直した年間データ
 * @param {number} fiscalYear - 年度
 * @param {Object|null} previousAnnualData - 上書きする前の保存データ
 * @returns {Promise<void>}
 */
function showResultDialog(annualData, fiscalYear, previousAnnualData) {
  return new Promise((resolve) => {
    const dialog = document.createElement("dialog");
    dialog.className = "kot-zangyo-dialog";

    const header = document.createElement("div");
    header.className = "kot-zangyo-dialog-header";
    header.textContent = `完了`;

    const body = document.createElement("div");
    body.className = "kot-zangyo-dialog-body";

    const totalMinutes = annualData.totalMinutes || 0;

    // 結果メッセージを作成（HH:MM形式）
    const yearRangeFormatted = annualData.yearRange
      ? annualData.yearRange.replace("-", " 〜 ")
      : `${fiscalYear}年度`;

    const introText = document.createTextNode(`正常に取得が完了しました。\n\n`);
    body.appendChild(introText);

    const resultText = document.createElement("div");
    resultText.textContent = `年間残業時間 (${yearRangeFormatted}): \n${formatMinutesToTime(
      totalMinutes,
    )}`;
    resultText.style.fontWeight = "bold";
    body.appendChild(resultText);

    body.style.whiteSpace = "pre-line";

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

    const footer = document.createElement("div");
    footer.className = "kot-zangyo-dialog-footer";

    const closeButton = document.createElement("button");
    closeButton.className = "kot-zangyo-btn kot-zangyo-btn-primary";
    closeButton.textContent = "閉じる";
    closeButton.onclick = () => {
      dialog.close();
      dialog.remove();
      resolve();
    };

    footer.appendChild(closeButton);

    dialog.appendChild(header);
    dialog.appendChild(body);
    if (noticeText) {
      dialog.appendChild(notice);
    }
    dialog.appendChild(footer);

    document.body.appendChild(dialog);
    dialog.showModal();
  });
}

/**
 * ダイアログ用のCSSスタイルをページに注入する（初回のみ）
 * King of Time公式のデザインに準拠
 */
function injectDialogStyles() {
  // 既にスタイルが注入されている場合はスキップ
  if (document.getElementById("king-of-zangyo-dialog-styles")) {
    return;
  }

  const styleElement = document.createElement("style");
  styleElement.id = "king-of-zangyo-dialog-styles";
  styleElement.textContent = `
    /* 使い方ページへの入口（更新ボタンの右） */
    .kot-zangyo-guide-link {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      height: 28px;
      padding: 0 10px 0 5px;
      margin-left: 4px;
      border: 1px solid #999;
      border-radius: 14px;
      background: #fff;
      color: #555;
      font-size: 13px;
      text-decoration: none;
      position: relative;
      top: -1px;
      box-sizing: border-box;
      white-space: nowrap;
    }
    .kot-zangyo-guide-link:hover {
      background: #f2f2f2;
      text-decoration: none;
    }
    /* KING OF TIME の a:visited（青）、a:hover、a:active が .kot-zangyo-guide-link より強く効くので、
       擬似クラスを付けて同じ灰色に固定する。押したあとに色が変わらないようにするため */
    .kot-zangyo-guide-link:link,
    .kot-zangyo-guide-link:visited,
    .kot-zangyo-guide-link:hover,
    .kot-zangyo-guide-link:active,
    .kot-zangyo-guide-link:focus {
      color: #555;
    }
    .kot-zangyo-guide-link-mark {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 18px;
      height: 18px;
      border-radius: 50%;
      background: #777;
      color: #fff;
      font-weight: bold;
      font-size: 12px;
    }

    /* 使い方の吹き出し（更新ボタンの直下）と赤い点 */
    .kot-zangyo-notice-dot {
      position: absolute;
      top: -4px;
      right: -4px;
      width: 10px;
      height: 10px;
      border-radius: 50%;
      background: #d32f2f;
      border: 2px solid #fff;
      box-sizing: content-box;
      pointer-events: none;
    }
    .kot-zangyo-notice-bubble {
      position: absolute;
      z-index: 100;
      width: 250px;
      box-sizing: border-box;
      padding: 10px 12px 8px;
      border: 1px solid #b7c7d6;
      border-radius: 4px;
      background: #f4f8fb;
      box-shadow: 0 2px 6px rgba(0, 0, 0, 0.12);
      color: #333;
      font-size: 13px;
      line-height: 1.6;
      text-align: left;
    }
    .kot-zangyo-notice-arrow {
      position: absolute;
      top: -6px;
      left: 18px;
      width: 10px;
      height: 10px;
      background: #f4f8fb;
      border-top: 1px solid #b7c7d6;
      border-left: 1px solid #b7c7d6;
      transform: rotate(45deg);
    }
    .kot-zangyo-notice-text {
      margin: 0;
      color: #333;
      font-size: 13px;
      line-height: 1.6;
    }
    .kot-zangyo-notice-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 6px;
      font-size: 13px;
    }
    .kot-zangyo-notice-more {
      color: #1565c0;
      text-decoration: underline;
    }
    .kot-zangyo-notice-close {
      border: none;
      background: none;
      padding: 0;
      color: #555;
      font-size: 13px;
      cursor: pointer;
    }

    /* ダイアログ基本スタイル（King of Time公式に準拠） */
    .kot-zangyo-dialog {
      border: none;
      border-radius: 4px;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1), 0 2px 4px rgba(0, 0, 0, 0.06);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Helvetica Neue", Arial, sans-serif;
      font-size: 16px;
      padding: 0;
      max-width: 600px;
      width: 90%;
      outline: none;
    }

    .kot-zangyo-dialog::backdrop {
      background-color: rgba(0, 0, 0, 0.5);
    }

    /* ダイアログヘッダー */
    .kot-zangyo-dialog-header {
      background-color: #f5f5f5;
      border-bottom: 1px solid #e0e0e0;
      padding: 16px 20px;
      font-size: 18px;
      font-weight: bold;
      color: #333;
    }

    /* ダイアログボディ */
    .kot-zangyo-dialog-body {
      background-color: #fff;
      padding: 20px;
      color: #333;
      line-height: 1.6;
    }

    /* ダイアログフッター */
    .kot-zangyo-dialog-footer {
      background-color: #fafafa;
      border-top: 1px solid #e0e0e0;
      padding: 16px 20px;
      text-align: right;
      display: flex;
      justify-content: flex-end;
      gap: 12px;
    }

    /* ボタン基本スタイル */
    .kot-zangyo-btn {
      border: none;
      border-radius: 4px;
      padding: 8px 16px;
      font-size: 16px;
      font-weight: 500;
      cursor: pointer;
      transition: background-color 0.2s, box-shadow 0.2s;
      outline: none;
    }

    .kot-zangyo-btn:hover {
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.15);
    }

    .kot-zangyo-btn:active {
      transform: translateY(1px);
    }

    /* プライマリボタン（King of Time公式の緑色） */
    .kot-zangyo-btn-primary {
      background-color: #1d9e48;
      color: #fff;
    }

    .kot-zangyo-btn-primary:hover {
      background-color: #008736;
    }

    /* キャンセルボタン（グレー） */
    .kot-zangyo-btn-cancel {
      background-color: #e0e0e0;
      color: #333;
    }

    .kot-zangyo-btn-cancel:hover {
      background-color: #d0d0d0;
    }

    /* プログレスバーコンテナ */
    .progress-bar-container {
      width: 100%;
      height: 36px;
      background-color: #f0f0f0;
      border-radius: 16px;
      overflow: hidden;
      margin-top: 16px;
      margin-bottom: 8px;
      position: relative;
    }

    /* プログレスバー */
    .progress-bar {
      height: 100%;
      background-color: #1d9e48;
      border-radius: 16px;
      display: flex;
      align-items: center;
      justify-content: center;
      color: #fff;
      font-size: 20px;
      font-weight: bold;
    }

    /* 進捗テキスト */
    .progress-text {
      margin-top: 8px;
      text-align: center;
      color: #666;
      font-size: 16px;
    }
  `;

  document.head.appendChild(styleElement);
}

// アプリケーションの初期化を実行
initializeApp();
