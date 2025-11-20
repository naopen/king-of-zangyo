// ポップアップ画面の初期化と設定管理

// ストレージキー定数
const STORAGE_KEY = "kingOfZangyoEnabled";
const STANDARD_HOURS_KEY = "kingOfZangyoStandardHours";

// ページ読み込み時に現在の設定を読み込む
document.addEventListener("DOMContentLoaded", () => {
  loadSettings();
  setupEventListeners();
});

/**
 * 保存されている設定を読み込んでUIに反映する
 */
function loadSettings() {
  chrome.storage.sync.get([STORAGE_KEY, STANDARD_HOURS_KEY], (result) => {
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

    // デフォルト値が設定されていない場合は保存
    if (result[STORAGE_KEY] === undefined) {
      saveSettings(true);
    }
    if (result[STANDARD_HOURS_KEY] === undefined) {
      saveStandardHours(7.5);
    }
  });
}

/**
 * 設定をストレージに保存する
 * @param {boolean} isEnabled - 機能の有効/無効状態
 */
function saveSettings(isEnabled) {
  chrome.storage.sync.set({ [STORAGE_KEY]: isEnabled }, () => {
    console.log(
      `King-of-Zangyo: 設定を保存しました (${isEnabled ? "有効" : "無効"})`
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
                "King-of-Zangyo: メッセージ送信エラー（ページをリロードしてください）"
              );
            } else {
              console.log("King-of-Zangyo: 表示を更新しました");
            }
          }
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
                "King-of-Zangyo: メッセージ送信エラー（ページをリロードしてください）"
              );
            } else {
              console.log("King-of-Zangyo: 所定労働時間を更新しました");
            }
          }
        );
      }
    });
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
}
