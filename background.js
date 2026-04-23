const LOG_PREFIX = "[SN MEETING]";

console.info(`${LOG_PREFIX} Background service worker loaded`);

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || message.type !== "OPEN_OUTLOOK_DEEPLINK" || !message.url) {
    console.log(`${LOG_PREFIX} BG ignored message`, {
      messageType: message?.type,
      hasUrl: Boolean(message?.url)
    });
    return;
  }

  console.log(`${LOG_PREFIX} BG opening Outlook tab`, {
    fromTabId: sender?.tab?.id,
    urlPreview: String(message.url).slice(0, 250)
  });

  chrome.tabs.create({
    url: message.url,
    active: true
  });

  sendResponse({ ok: true });
});
