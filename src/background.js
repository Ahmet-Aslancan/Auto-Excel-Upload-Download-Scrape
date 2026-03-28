const POPUP_WIDTH = 480;
const POPUP_HEIGHT = 760;

function isValidUrl(url) {
  return url && url.startsWith("http");
}

chrome.action.onClicked.addListener(async (tab) => {
  if (!tab?.id || !isValidUrl(tab.url)) {
    console.warn("Cannot run extension on this page.");
    return;
  }

  try {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ["build/content.js"]
    });
  } catch (err) {
    console.error("Script injection failed:", err);
    return;
  }

  const url = chrome.runtime.getURL(`popup/popup.html?tabId=${tab.id}`);

  await chrome.windows.create({
    url,
    type: "popup",
    width: POPUP_WIDTH,
    height: POPUP_HEIGHT
  });
});