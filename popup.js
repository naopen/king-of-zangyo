// ポップアップ画面の初期化と設定管理

// ストレージキー定数
const STORAGE_KEY = "kingOfZangyoEnabled";
const STANDARD_HOURS_KEY = "kingOfZangyoStandardHours";
const FISCAL_YEAR_START_KEY = "kingOfZangyoFiscalYearStartMonth";
const ANNUAL_DATA_KEY_PREFIX = "kingOfZangyoAnnualData_"; // 年度別キーの接頭辞
const ANNUAL_DATA_YEARS_KEY = "kingOfZangyoAnnualDataYears"; // 保存済み年度リスト

// ページ読み込み時に現在の設定を読み込む
document.addEventListener("DOMContentLoaded", () => {
  loadSettings();
  setupEventListeners();
});

/**
 * 保存されている設定を読み込んでUIに反映する
 */
function loadSettings() {
  chrome.storage.sync.get(
    [STORAGE_KEY, STANDARD_HOURS_KEY, FISCAL_YEAR_START_KEY],
    (result) => {
      // 残業時間表示のON/OFF設定（デフォルトは「オン」）
      const isEnabled =
        result[STORAGE_KEY] !== undefined ? result[STORAGE_KEY] : true;
      const toggleSwitch = document.getElementById("toggle-switch");

      if (toggleSwitch) {
        toggleSwitch.checked = isEnabled;
      }

      // 所定労働時間の設定（デフォルトは7.5時間）
      const standardHours =
        result[STANDARD_HOURS_KEY] !== undefined
          ? result[STANDARD_HOURS_KEY]
          : 7.5;
      const standardHoursInput = document.getElementById("standard-hours");

      if (standardHoursInput) {
        standardHoursInput.value = standardHours;
      }

      // 年度開始月の設定（デフォルトは4月）
      const fiscalYearStart =
        result[FISCAL_YEAR_START_KEY] !== undefined
          ? result[FISCAL_YEAR_START_KEY]
          : 4;
      const fiscalYearStartSelect =
        document.getElementById("fiscal-year-start");

      if (fiscalYearStartSelect) {
        fiscalYearStartSelect.value = fiscalYearStart.toString();
      }

      // デフォルト値が設定されていない場合は保存（初期化時のみ）
      // 各保存関数を使わず直接設定値のみを保存することで、
      // 年度データの削除や不要なタブへのメッセージ送信を回避する
      const defaults = {};
      if (result[STORAGE_KEY] === undefined) {
        defaults[STORAGE_KEY] = true;
      }
      if (result[STANDARD_HOURS_KEY] === undefined) {
        defaults[STANDARD_HOURS_KEY] = 7.5;
      }
      if (result[FISCAL_YEAR_START_KEY] === undefined) {
        defaults[FISCAL_YEAR_START_KEY] = 4;
      }
      if (Object.keys(defaults).length > 0) {
        chrome.storage.sync.set(defaults);
      }
    },
  );
}

/**
 * 設定をストレージに保存する
 * @param {boolean} isEnabled - 機能の有効/無効状態
 */
function saveSettings(isEnabled) {
  chrome.storage.sync.set({ [STORAGE_KEY]: isEnabled }, () => {
    console.log(
      `King-of-Zangyo: 設定を保存しました (${isEnabled ? "有効" : "無効"})`,
    );

    // アクティブなタブにメッセージを送信して、表示を更新
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      if (tabs[0]) {
        chrome.tabs.sendMessage(
          tabs[0].id,
          {
            action: "toggleDisplay",
            enabled: isEnabled,
          },
          (response) => {
            if (chrome.runtime.lastError) {
              console.log(
                "King-of-Zangyo: メッセージ送信エラー（ページをリロードしてください）",
              );
            } else {
              console.log("King-of-Zangyo: 表示を更新しました");
            }
          },
        );
      }
    });
  });
}

/**
 * 所定労働時間をストレージに保存する
 * @param {number} hours - 所定労働時間（時間単位）
 */
function saveStandardHours(hours) {
  chrome.storage.sync.set({ [STANDARD_HOURS_KEY]: hours }, () => {
    console.log(`King-of-Zangyo: 所定労働時間を保存しました (${hours}時間)`);

    // アクティブなタブにメッセージを送信して、計算を更新
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      if (tabs[0]) {
        chrome.tabs.sendMessage(
          tabs[0].id,
          {
            action: "updateStandardHours",
            hours: hours,
          },
          (response) => {
            if (chrome.runtime.lastError) {
              console.log(
                "King-of-Zangyo: メッセージ送信エラー（ページをリロードしてください）",
              );
            } else {
              console.log("King-of-Zangyo: 所定労働時間を更新しました");
            }
          },
        );
      }
    });
  });
}

/**
 * 年度開始月をストレージに保存する
 * @param {number} month - 年度開始月（1-12）
 */
async function saveFiscalYearStart(month) {
  // 全年度データを削除
  const result = await chrome.storage.sync.get([ANNUAL_DATA_YEARS_KEY]);
  const years = result[ANNUAL_DATA_YEARS_KEY] || [];

  // 各年度のデータを削除
  for (const year of years) {
    const key = `${ANNUAL_DATA_KEY_PREFIX}${year}`;
    await chrome.storage.sync.remove([key]);
    console.log(`King-of-Zangyo: ${year}年度のデータを削除しました`);
  }

  // 年度リストを削除
  await chrome.storage.sync.remove([ANNUAL_DATA_YEARS_KEY]);
  console.log("King-of-Zangyo: 年度リストを削除しました");

  // 年度開始月を保存
  await chrome.storage.sync.set({ [FISCAL_YEAR_START_KEY]: month });
  console.log(`King-of-Zangyo: 年度開始月を保存しました (${month}月)`);

  // アクティブなタブにメッセージを送信して、年度開始月を更新
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    if (tabs[0]) {
      chrome.tabs.sendMessage(
        tabs[0].id,
        {
          action: "updateFiscalYearStart",
          month: month,
        },
        (response) => {
          if (chrome.runtime.lastError) {
            console.log(
              "King-of-Zangyo: メッセージ送信エラー（ページをリロードしてください）",
            );
          } else {
            console.log("King-of-Zangyo: 年度開始月を更新しました");
          }
        },
      );
    }
  });
}

/**
 * イベントリスナーをセットアップする
 */
function setupEventListeners() {
  const toggleSwitch = document.getElementById("toggle-switch");

  if (toggleSwitch) {
    toggleSwitch.addEventListener("change", (event) => {
      const isEnabled = event.target.checked;
      saveSettings(isEnabled);
    });
  }

  const standardHoursInput = document.getElementById("standard-hours");

  if (standardHoursInput) {
    // 入力値が変更されたときに保存
    standardHoursInput.addEventListener("change", (event) => {
      let hours = parseFloat(event.target.value);

      // バリデーション
      if (isNaN(hours) || hours <= 0) {
        // 無効な値の場合はデフォルトの7.5時間に戻す
        hours = 7.5;
        event.target.value = hours;
      } else if (hours > 24) {
        // 24時間を超える場合は24時間に制限
        hours = 24;
        event.target.value = hours;
      }

      saveStandardHours(hours);
    });
  }

  const fiscalYearStartSelect = document.getElementById("fiscal-year-start");

  if (fiscalYearStartSelect) {
    // 変更前の値を保持
    let previousFiscalYearStart = fiscalYearStartSelect.value;

    // 年度開始月が変更されたときに保存
    fiscalYearStartSelect.addEventListener("change", async (event) => {
      let month = parseInt(event.target.value);

      // バリデーション
      if (isNaN(month) || month < 1 || month > 12) {
        // 無効な値の場合はデフォルトの4月に戻す
        month = 4;
        event.target.value = month;
      }

      // 確認ダイアログを表示
      const confirmed = confirm(
        "年度開始月を変更すると、保存されている全ての年度データが削除されます。\n" +
          "（年度サイクル変更により、データの整合性を保つため）\n" +
          "よろしいですか？",
      );

      if (!confirmed) {
        // キャンセルされた場合、前の値に戻す
        fiscalYearStartSelect.value = previousFiscalYearStart;
        return;
      }

      // データを削除して年度開始月を保存
      await saveFiscalYearStart(month);

      // 保存成功後、現在の値を保持
      previousFiscalYearStart = month;
    });
  }
}
