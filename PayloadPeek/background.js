// background.js — toolbar click toggles the in-page panel; also grants the content
// script access to chrome.storage.session (which is trusted-only by default).

async function allowSessionFromContentScripts() {
  try { await chrome.storage.session.setAccessLevel({ accessLevel: 'TRUSTED_AND_UNTRUSTED_CONTEXTS' }); } catch (_) {}
}
allowSessionFromContentScripts();                                   // on every service-worker start
chrome.runtime.onInstalled.addListener(allowSessionFromContentScripts);
chrome.runtime.onStartup.addListener(allowSessionFromContentScripts);

chrome.action.onClicked.addListener(async (tab) => {
  if (!tab || !/^https:\/\/[^/]+\.hana\.ondemand\.com\//.test(tab.url || '')) return;
  try {
    await chrome.tabs.sendMessage(tab.id, { type: 'toggle' });
  } catch (_) {
    try {
      await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ['content.js'] });
      await chrome.tabs.sendMessage(tab.id, { type: 'toggle' });
    } catch (e) { /* not a reachable page */ }
  }
});
