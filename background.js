chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || message.type !== "OPEN_OUTLOOK_DEEPLINK" || !message.url) {
    return;
  }

  chrome.tabs.create({
    url: message.url,
    active: true
  });

  sendResponse({ ok: true });
});
