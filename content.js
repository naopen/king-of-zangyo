// King-of-Zangyo: 残業時間自動計算・表示スクリプト

// ストレージキー定数
const STORAGE_KEY = "kingOfZangyoEnabled";
const STANDARD_HOURS_KEY = "kingOfZangyoStandardHours";
const FISCAL_YEAR_START_KEY = "kingOfZangyoFiscalYearStartMonth";
const ANNUAL_DATA_KEY_PREFIX = "kingOfZangyoAnnualData_"; // 年度別キーの接頭辞
const ANNUAL_DATA_YEARS_KEY = "kingOfZangyoAnnualDataYears"; // 保存済み年度リスト
const PROCESSING_STATE_KEY = "kingOfZangyoProcessingState";

// DOM要素の一意なID
const OVERTIME_HEADER_ID = "king-of-zangyo-header";
const OVERTIME_CELL_ID = "king-of-zangyo-cell";
const ANNUAL_SECTION_ID = "king-of-zangyo-annual-section";

// 所定労働時間（分）- 設定から読み込まれるまでのデフォルト値
let STANDARD_WORK_MINUTES = 450; // 7.5時間

// 年度開始月（1-12）- 設定から読み込まれるまでのデフォルト値
let FISCAL_YEAR_START_MONTH = 4; // 4月開始

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
 * アプリケーションの初期化
 */
async function initializeApp() {
  try {
    // 対象ページかチェック
    if (!isTargetPage()) {
      return;
    }

    // 処理状態をチェックして自動再開
    checkAndResumeProcessing();

    // chrome.storage APIが利用可能かチェック
    if (
      typeof chrome === "undefined" ||
      !chrome.storage ||
      !chrome.storage.sync
    ) {
      console.warn(
        "King-of-Zangyo: chrome.storage.sync が利用できません。デフォルトでオンにします。",
      );
      injectOvertimeColumn();
      return;
    }

    chrome.storage.sync.get(
      [STORAGE_KEY, STANDARD_HOURS_KEY, FISCAL_YEAR_START_KEY],
      async (result) => {
        if (chrome.runtime.lastError) {
          console.error(
            "King-of-Zangyo: ストレージ取得エラー:",
            chrome.runtime.lastError,
          );
          // エラーの場合はデフォルトでオンにする
          injectOvertimeColumn();
          await injectAnnualDataSection();
          injectDialogStyles();
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

        // 年度開始月を設定（デフォルトは4月）
        const fiscalYearStart =
          result[FISCAL_YEAR_START_KEY] !== undefined
            ? result[FISCAL_YEAR_START_KEY]
            : 4;
        FISCAL_YEAR_START_MONTH = fiscalYearStart;

        // 年別データセクションを注入（データ読み込みのため常に実行）
        await injectAnnualDataSection();

        // ダイアログスタイルを注入
        injectDialogStyles();

        if (isEnabled) {
          injectOvertimeColumn();
        } else {
          // 非表示設定の場合、年別データセクションを非表示にする
          const annualSection = document.getElementById(ANNUAL_SECTION_ID);
          if (annualSection) {
            annualSection.style.display = "none";
          }
        }
      },
    );
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
      return;
    }

    // ヘッダー行に新しい列を追加
    const headerRow = thead.querySelector("tr");
    if (headerRow) {
      const th = document.createElement("th");
      th.id = OVERTIME_HEADER_ID;
      th.appendChild(document.createTextNode("現時点の目安残業"));
      th.appendChild(document.createElement("br"));
      th.appendChild(document.createTextNode("(日曜を除く)"));
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
      "td.work_day_type, td:nth-child(6)",
    );
    const workdayTypeCellText = workdayTypeCell?.textContent.trim() || "";

    // 法定休日（日曜日）は時間外労働の上限にカウントしない
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
  const annualSection = document.getElementById(ANNUAL_SECTION_ID);

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
  } else {
    // 非表示
    if (header) {
      header.style.display = "none";
    }
    if (cell) {
      cell.style.display = "none";
    }
    // 年別データセクションを非表示
    if (annualSection) {
      annualSection.style.display = "none";
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

    // 保存済み年度リストを取得
    const savedYears = await getSavedFiscalYears();

    // 現在年度と保存済み年度を結合（重複削除、降順）
    const allYears = [...new Set([currentFiscalYear, ...savedYears])].sort(
      (a, b) => b - a,
    );

    const fiscalYearSelect = document.createElement("select");
    fiscalYearSelect.id = "annual-fiscal-year-select";
    fiscalYearSelect.className = "htBlock-selectmenu"; // 既存UIに準拠

    // 年度オプションを生成
    allYears.forEach((year) => {
      const option = document.createElement("option");
      option.value = year;
      option.textContent = `${year}年度`;
      if (year === currentFiscalYear) {
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
    buttonSpan.textContent = "最新データで更新";
    updateButton.appendChild(buttonSpan);
    updateButton.addEventListener("click", handleAnnualUpdateButtonClick);

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
      { textLines: ["年間残業時間", "(日曜を除く)"], width: "120px" },
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

    // セクションを構築
    sectionContainer.appendChild(titleContainer);
    sectionContainer.appendChild(tableCaption);
    sectionContainer.appendChild(tableContainer);

    // フレックスコンテナの右側に年別データセクションを配置
    flexContainer.appendChild(sectionContainer);

    // 保存されているデータを読み込んで表示
    loadAndDisplayAnnualData();
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
 */
async function loadAndDisplayAnnualData() {
  try {
    // 現在年度を判定
    const currentFiscalYear = determineFiscalYear(
      new Date().getFullYear(),
      new Date().getMonth() + 1,
      FISCAL_YEAR_START_MONTH,
    );

    // 現在年度のデータを読み込み
    const annualData = await loadAnnualData(currentFiscalYear);

    updateAnnualDataDisplay(annualData, currentFiscalYear);
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
    updatedCell.textContent = "未取得";
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

  // 背景色と文字色を設定（時間単位で判定）
  const totalHours = totalMinutes / 60;
  hoursCell.style.backgroundColor =
    getAnnualOvertimeBackgroundColor(totalHours);

  // 最終更新を表示
  updatedCell.textContent = annualData.lastUpdated || "--";
}

/**
 * 年間残業時間の背景色を取得する
 * @param {number} hours - 年間残業時間（時間単位）
 * @returns {string} 背景色（16進数カラーコード）
 */
function getAnnualOvertimeBackgroundColor(hours) {
  if (hours < 300) {
    return "#c8e6c9"; // 緑
  } else if (hours < 330) {
    return "#fff59d"; // 黄
  } else if (hours < 360) {
    return "#ffcc80"; // 橙
  } else {
    return "#ef9a9a"; // 赤
  }
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

    // 確認ダイアログ（選択年度を表示）
    const confirmed = await showConfirmDialog(
      `データ更新`,
      `${selectedFiscalYear}年度の年間残業時間データを取得します。\n処理には10秒〜20秒かかる場合があります。\nよろしいですか？`,
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
    const overtimeMinutes = calculateTotalOvertime();

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

    // 指定年度のデータとして保存
    await saveAnnualData(annualData, targetFiscalYear);

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
    await showResultDialog(annualData, targetFiscalYear);

    // 元のページに戻る
    if (state.originalYear && state.originalMonth) {
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
 * @param {Object} annualData - 年間データ
 * @returns {Promise<void>}
 */
function showResultDialog(annualData, fiscalYear) {
  return new Promise((resolve) => {
    const dialog = document.createElement("dialog");
    dialog.className = "kot-zangyo-dialog";

    const header = document.createElement("div");
    header.className = "kot-zangyo-dialog-header";
    header.textContent = `完了`;

    const body = document.createElement("div");
    body.className = "kot-zangyo-dialog-body";

    const totalMinutes = annualData.totalMinutes || 0;
    const totalHours = totalMinutes / 60;

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

    // 背景色を設定
    body.style.backgroundColor = getAnnualOvertimeBackgroundColor(totalHours);
    body.style.padding = "20px";
    body.style.borderRadius = "4px";

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
