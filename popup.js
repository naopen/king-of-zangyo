// ポップアップ画面の初期化と設定管理

// ストレージキー定数
const STORAGE_KEY = "kingOfZangyoEnabled";

// ページ読み込み時に現在の設定を読み込む
document.addEventListener("DOMContentLoaded", () => {
  loadSettings();
  setupEventListeners();
});

/**
 * 保存されている設定を読み込んでUIに反映する
 */
function loadSettings() {
  chrome.storage.sync.get([STORAGE_KEY], (result) => {
    // デフォルトは「オン」(true)
    const isEnabled =
      result[STORAGE_KEY] !== undefined ? result[STORAGE_KEY] : true;
    const toggleSwitch = document.getElementById("toggle-switch");

    if (toggleSwitch) {
      toggleSwitch.checked = isEnabled;
    }

    // デフォルト値が設定されていない場合は保存
    if (result[STORAGE_KEY] === undefined) {
      saveSettings(true);
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
}
