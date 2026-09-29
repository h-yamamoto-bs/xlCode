/**
 * electron-builder の設定
 *
 * Windows: 普通のデスクトップアプリとしてインストールする（NSIS インストーラー）。
 * - インストール先はフォルダ（onedir）。起動のたびに展開する portable 形式より速い
 * - スタートメニュー・「アプリと機能」に登録される
 * - 「自分だけ（管理者権限不要）」か「すべてのユーザー」かをインストール時に選べる
 */
module.exports = {
  appId: 'local.xlcode',
  productName: 'xlCode',
  directories: { output: 'dist', buildResources: 'build' },
  files: ['out/**', 'package.json'],
  asar: true,
  win: {
    target: [{ target: 'nsis', arch: ['x64'] }],
    icon: 'build/icon.png',
    // exe へのアイコン・バージョン情報の書き込み（rcedit）。Linux からビルドする場合は wine が必要
  },
  nsis: {
    oneClick: false,
    perMachine: false,
    allowToChangeInstallationDirectory: true,
    createDesktopShortcut: true,
    createStartMenuShortcut: true,
    shortcutName: 'xlCode',
    uninstallDisplayName: 'xlCode',
    artifactName: 'xlCode-Setup-${version}.exe',
    // アンインストールしても設定（最近使ったプロジェクト等）は残す
    deleteAppDataOnUninstall: false,
  },
  mac: { target: 'dmg', icon: 'build/icon.png' },
  linux: { target: 'AppImage', icon: 'build/icon.png' },
};
