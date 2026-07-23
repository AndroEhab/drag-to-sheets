// Open side panel when extension icon is clicked
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch(() => {});

chrome.commands.onCommand.addListener(async (command) => {
  if (command !== 'open-panel') return;

  try {
    const window = await chrome.windows.getCurrent();
    await chrome.sidePanel.open({ windowId: window.id });
  } catch (_) {
    // Side panel may not be available in this context
  }
});

