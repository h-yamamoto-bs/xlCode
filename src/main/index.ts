import path from 'node:path';
import { app, BrowserWindow, dialog, nativeTheme, shell } from 'electron';
import { registerIpc } from './ipc';

function createWindow(): void {
  nativeTheme.themeSource = 'dark';
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 900,
    minHeight: 560,
    show: false,
    backgroundColor: '#1f1f1f',
    title: 'xlCode',
    titleBarStyle: 'hidden',
    titleBarOverlay:
      process.platform === 'darwin' ? undefined : { color: '#181818', symbolColor: '#cccccc', height: 35 },
    webPreferences: {
      preload: path.join(import.meta.dirname, '../preload/index.cjs'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
    },
  });
  win.once('ready-to-show', () => win.show());
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

app.whenReady().then(() => {
  registerIpc();
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
