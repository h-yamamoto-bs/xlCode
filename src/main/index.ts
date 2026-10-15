import path from 'node:path';
import { app, BrowserWindow, dialog, Menu, shell } from 'electron';
import { readAppearance, resolveAppearance, WINDOW_COLORS } from './appearance';
import { registerIpc } from './ipc';

/**
 * アプリのメニュー。
 * Windows / Linux では既定のメニューを外す（既定のメニューは Ctrl+W で閉じる・Ctrl+R で再読み込みなどの
 * ショートカットを持ち、画面側のタブ操作（Ctrl+W = タブを閉じる、Ctrl+R = 最近使ったプロジェクト）とぶつかるため）。
 * macOS ではコピー・貼り付けなどがメニューのロール経由で動くので、最小限のメニューを残す。
 */
function setupMenu(): void {
  if (process.platform !== 'darwin') {
    Menu.setApplicationMenu(null);
    return;
  }
  Menu.setApplicationMenu(Menu.buildFromTemplate([{ role: 'appMenu' }, { role: 'editMenu' }, { role: 'windowMenu' }]));
}

function createWindow(): void {
  // 前回の表示モードで開く（画面が読み込まれたら、画面の設定で改めて決まる）
  const colors = WINDOW_COLORS[resolveAppearance(readAppearance())];
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 900,
    minHeight: 560,
    show: false,
    backgroundColor: colors.background,
    title: 'xlCode',
    titleBarStyle: 'hidden',
    titleBarOverlay:
      process.platform === 'darwin' ? undefined : { color: colors.bar, symbolColor: colors.symbol, height: 35 },
    webPreferences: {
      preload: path.join(import.meta.dirname, '../preload/index.cjs'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
    },
  });
  win.once('ready-to-show', () => win.show());
  // 開発用: F12 / Ctrl+Shift+I で開発者ツール（メニューを外したので自前で扱う）
  win.webContents.on('before-input-event', (_event, input) => {
    if (app.isPackaged || input.type !== 'keyDown') return;
    if (input.key === 'F12' || (input.control && input.shift && input.key.toUpperCase() === 'I')) {
      win.webContents.toggleDevTools();
    }
  });
  // 画面側が beforeunload で閉じるのを止めた（保存していない Markdown がある）ときに確認する
  win.webContents.on('will-prevent-unload', (event) => {
    const choice = dialog.showMessageBoxSync(win, {
      type: 'warning',
      title: 'xlCode',
      message: '保存していない Markdown があります',
      detail: '閉じると編集内容は失われます。',
      buttons: ['保存せずに閉じる', 'キャンセル'],
      defaultId: 1,
      cancelId: 1,
    });
    if (choice === 0) event.preventDefault();
  });
  // 外部リンクは既定のブラウザで開く
  win.webContents.setWindowOpenHandler(({ url }) => {
    void shell.openExternal(url);
    return { action: 'deny' };
  });
  if (process.env.ELECTRON_RENDERER_URL) void win.loadURL(process.env.ELECTRON_RENDERER_URL);
  else void win.loadFile(path.join(import.meta.dirname, '../renderer/index.html'));
}

// 自動テストでは設定（localStorage など）の保存先を分ける
if (process.env.XLCODE_USER_DATA) app.setPath('userData', process.env.XLCODE_USER_DATA);

// 同じプロジェクトを2つのウィンドウから操作すると処理がぶつかるため、1つだけ起動する
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', () => {
    const win = BrowserWindow.getAllWindows()[0];
    if (!win) return;
    if (win.isMinimized()) win.restore();
    win.focus();
  });
}

app.whenReady().then(() => {
  setupMenu();
  registerIpc();
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
