const fs = require('fs');
const path = require('path');

describe('Background service worker', () => {
  let commandListener;

  beforeAll(() => {
    jest.clearAllMocks();
    global.self = global;
    const code = fs.readFileSync(path.resolve(__dirname, '../background.js'), 'utf-8');
    eval(code);
    commandListener = chrome.commands.onCommand.addListener.mock.calls[0][0];
  });

  beforeEach(() => {
    chrome.runtime.sendMessage = jest.fn().mockResolvedValue(undefined);
    chrome.storage.local.get = jest.fn().mockResolvedValue({});
    chrome.storage.local.set = jest.fn().mockResolvedValue(undefined);
    chrome.identity.getAuthToken = jest.fn().mockResolvedValue({ token: 'mock-token' });
    chrome.sidePanel.open = jest.fn().mockResolvedValue(undefined);
  });

  test('sets panel behavior to open on action click', () => {
    expect(chrome.sidePanel.setPanelBehavior).toHaveBeenCalledWith({
      openPanelOnActionClick: true,
    });
  });

  test('registers a command listener', () => {
    expect(chrome.commands.onCommand.addListener).toHaveBeenCalledTimes(1);
    expect(typeof commandListener).toBe('function');
  });

  test('opens side panel on "open-panel" command', async () => {
    await commandListener('open-panel');

    expect(chrome.windows.getCurrent).toHaveBeenCalled();
    expect(chrome.sidePanel.open).toHaveBeenCalledWith({ windowId: 1 });
  });

  test('does not open side panel for other commands', async () => {
    chrome.sidePanel.open.mockClear();
    await commandListener('some-other-command');

    expect(chrome.sidePanel.open).not.toHaveBeenCalled();
  });
});
