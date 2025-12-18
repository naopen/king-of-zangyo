// King-of-Zangyo: 残業時間自動計算・表示スクリプト

// ストレージキー定数
const STORAGE_KEY = "kingOfZangyoEnabled";
const STANDARD_HOURS_KEY = "kingOfZangyoStandardHours";

// DOM要素の一意なID
const OVERTIME_HEADER_ID = "king-of-zangyo-header";
const OVERTIME_CELL_ID = "king-of-zangyo-cell";

// 所定労働時間（分）- 設定から読み込まれるまでのデフォルト値
let STANDARD_WORK_MINUTES = 450; // 7.5時間

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
    // セッション切れやエラーページの場合は対象外
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
      "div.htBlock-normalTable table.specific-table_800"
    ) !== null;
  const hasDailyDataTable =
    document.querySelector(".htBlock-adjastableTableF_inner > table") !== null;
  const hasTimecardTitle =
    document.querySelector("h1")?.textContent.includes("タイムカード") || false;

  const isDomMatch = hasSummaryTable && hasDailyDataTable;

  return isDomMatch;
}

/**
 * アプリケーションの初期化
 */
function initializeApp() {
  try {
    // 対象ページかチェック
    if (!isTargetPage()) {
      return;
    }

    // chrome.storage APIが利用可能かチェック
    if (
      typeof chrome === "undefined" ||
      !chrome.storage ||
      !chrome.storage.sync
    ) {
      console.warn(
        "King-of-Zangyo: chrome.storage.sync が利用できません。デフォルトでオンにします。"
      );
      injectOvertimeColumn();
      return;
    }

    chrome.storage.sync.get([STORAGE_KEY, STANDARD_HOURS_KEY], (result) => {
      if (chrome.runtime.lastError) {
        console.error(
          "King-of-Zangyo: ストレージ取得エラー:",
          chrome.runtime.lastError
        );
        // エラーの場合はデフォルトでオンにする
        injectOvertimeColumn();
        return;
      }

      // デフォルトは「オン」(true)
      const isEnabled =
        result[STORAGE_KEY] !== undefined ? result[STORAGE_KEY] : true;

      // 所定労働時間を設定（デフォルトは7.5時間）
      const standardHours =
        result[STANDARD_HOURS_KEY] !== undefined
          ? result[STANDARD_HOURS_KEY]
          : 7.5;
      STANDARD_WORK_MINUTES = standardHours * 60;

      if (isEnabled) {
        injectOvertimeColumn();
      }
    });
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
      "div.htBlock-normalTable table.specific-table_800"
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
      return;
    }

    // ヘッダー行に新しい列を追加
    const headerRow = thead.querySelector("tr");
    if (headerRow) {
      const th = document.createElement("th");
      th.id = OVERTIME_HEADER_ID;
      th.textContent = "現時点の目安残業";
      th.style.textAlign = "center";
      th.style.fontWeight = "bold";
      th.style.fontSize = "13px";
      th.style.width = "120px";
      th.style.minWidth = "120px";
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
      dataRow.appendChild(td);
    }
  } catch (error) {
    console.error("King-of-Zangyo: 列の注入中にエラーが発生しました:", error);
  }
}

/**
 * 累計残業時間を計算する
 * @returns {number} 累計残業時間（分）
 */
function calculateTotalOvertime() {
  // 日別データテーブルを取得
  const dailyDataTable = document.querySelector(
    ".htBlock-adjastableTableF_inner > table"
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
      "td.htBlock-scrollTable_day, td:nth-child(2)"
    );
    const dateCellText = dateCell?.textContent.trim();

    if (!dateCellText) {
      return; // 日付がない場合はスキップ
    }

    // 日付をパース（年月情報を使用）
    const rowDate = parseDate(
      dateCellText,
      currentYearMonth.year,
      currentYearMonth.month
    );
    if (!rowDate || rowDate > today) {
      return; // 明日以降のデータはスキップ
    }

    // 当日の場合、実働時間が所定時間（7.5h = 450分）以上の場合のみ含める
    const isToday = rowDate.getTime() === today.getTime();
    if (isToday) {
      // 実働時間を先に取得して判定
      const actualWorkCell = row.querySelector("td.custom12");
      const actualWorkText = actualWorkCell?.textContent.trim() || "";

      if (!actualWorkText || actualWorkText === "") {
        return;
      }

      const match = actualWorkText.match(/^(\d+)\.(\d+)$/);
      if (!match) {
        return;
      }

      const hours = parseInt(match[1], 10);
      const minutes = parseInt(match[2], 10);
      const actualWorkMinutes = hours * 60 + minutes;

      // 所定時間未満の場合はスキップ
      if (actualWorkMinutes < STANDARD_WORK_MINUTES) {
        return;
      }
    }

    // スケジュール（5番目のtd）を取得
    const scheduleCell = row.querySelector("td.schedule, td:nth-child(5)");
    const scheduleCellText = scheduleCell?.textContent.trim() || "";

    // 「有休」が含まれている場合はスキップ
    if (scheduleCellText.includes("有休")) {
      return;
    }

    // 勤務日種別（6番目のtd）を取得
    const workdayTypeCell = row.querySelector(
      "td.work_day_type, td:nth-child(6)"
    );
    const workdayTypeCellText = workdayTypeCell?.textContent.trim() || "";

    // 平日かどうかを判定
    const isWeekday = workdayTypeCellText === "平日";
    // 休日かどうかを判定（法定休日 = 日曜日、法定外休日 = 土曜日・祝日）
    const isHoliday = workdayTypeCellText === "法定休日" || workdayTypeCellText === "法定外休日";

    // 平日でも休日でもない場合はスキップ（念のため）
    if (!isWeekday && !isHoliday) {
      return;
    }

    // 実働時間（class="custom12"を優先）を取得
    const actualWorkCell = row.querySelector("td.custom12");
    const actualWorkText = actualWorkCell?.textContent.trim() || "";

    if (!actualWorkText || actualWorkText === "") {
      return; // 実働時間が空の場合はスキップ
    }

    // 実働時間を60進法からパース（例: "8.48" → 8時間48分 = 528分）
    const match = actualWorkText.match(/^(\d+)\.(\d+)$/);

    if (!match) {
      // マッチしない場合はスキップ
      return;
    }

    const hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);

    // 実働時間を分に変換（60進法: 8.48 = 8時間48分）
    const actualWorkMinutes = hours * 60 + minutes;

    // 残業時間を計算
    // 平日：実働時間 - 所定労働時間（450分 = 7.5時間）
    // 休日（土日祝）：実働時間 - 0分 = 実働時間（全て残業扱い）
    const standardMinutes = isWeekday ? STANDARD_WORK_MINUTES : 0;
    const dailyOvertimeMinutes = actualWorkMinutes - standardMinutes;

    // 累計に追加
    totalOvertimeMinutes += dailyOvertimeMinutes;
  });

  return Math.round(totalOvertimeMinutes);
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
  } else {
    // 45時間以上：クリムゾンレッド
    return "#d32f2f";
  }
}

/**
 * 残業時間列の表示/非表示を切り替える
 * @param {boolean} enabled - 表示する場合はtrue、非表示にする場合はfalse
 */
function toggleOvertimeDisplay(enabled) {
  const header = document.getElementById(OVERTIME_HEADER_ID);
  const cell = document.getElementById(OVERTIME_CELL_ID);

  if (enabled) {
    // 列が存在しない場合は新規作成
    if (!header || !cell) {
      injectOvertimeColumn();
    } else {
      // 既に存在する場合は表示
      header.style.display = "";
      cell.style.display = "";
    }
  } else {
    // 非表示
    if (header) {
      header.style.display = "none";
    }
    if (cell) {
      cell.style.display = "none";
    }
  }
}

// メッセージリスナーを設定（ポップアップからのメッセージを受信）
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "toggleDisplay") {
    toggleOvertimeDisplay(request.enabled);
    sendResponse({ success: true });
  } else if (request.action === "updateStandardHours") {
    // 所定労働時間を更新
    STANDARD_WORK_MINUTES = request.hours * 60;
    // 残業時間を再計算して表示を更新
    injectOvertimeColumn();
    sendResponse({ success: true });
  }

  return true; // 非同期レスポンスを許可
});

// アプリケーションの初期化を実行
initializeApp();
