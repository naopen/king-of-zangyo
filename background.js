// インストール直後だけ使い方ページを開く（判定は notices.js 側）
importScripts("notices.js");

chrome.runtime.onInstalled.addListener(({ reason }) => {
  if (KingOfZangyoNotices.shouldOpenGuideOnInstalled(reason)) {
    chrome.tabs.create({ url: chrome.runtime.getURL("guide/index.html") });
  }
});
