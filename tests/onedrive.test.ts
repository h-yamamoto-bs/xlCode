import { describe, expect, it } from 'vitest';
import { joinUrl, parseSyncRoots, toWebUrl } from '../src/main/onedrive';

const REG = `
HKEY_CURRENT_USER\\Software\\SyncEngines\\Providers\\OneDrive\\Business1
    MountPoint    REG_SZ    C:\\Users\\me\\OneDrive - Contoso
    UrlNamespace    REG_SZ    https://contoso-my.sharepoint.com/personal/me_contoso_com/Documents/
    LibraryType    REG_SZ    mysite

HKEY_CURRENT_USER\\Software\\SyncEngines\\Providers\\OneDrive\\abc123
    MountPoint    REG_SZ    C:\\Users\\me\\Contoso\\開発チーム - ドキュメント
    UrlNamespace    REG_SZ    https://contoso.sharepoint.com/sites/dev/Shared Documents/
`;

describe('OneDrive の Web URL', () => {
  const roots = parseSyncRoots(REG);
  it('reg query の出力から同期フォルダを取り出す', () => {
    expect(roots).toHaveLength(2);
    expect(roots[0].mountPoint).toBe('C:\\Users\\me\\OneDrive - Contoso');
  });
  it('ローカルパスを Excel for the web の URL に変換する', () => {
    expect(toWebUrl('C:\\Users\\me\\OneDrive - Contoso\\shop\\app\\app.xlcode.xlsx', roots, '\\')).toBe(
      'https://contoso-my.sharepoint.com/personal/me_contoso_com/Documents/shop/app/app.xlcode.xlsx?web=1',
    );
    expect(toWebUrl('c:\\users\\me\\contoso\\開発チーム - ドキュメント\\x\\x.xlcode.xlsx', roots, '\\')).toMatch(
      /^https:\/\/contoso\.sharepoint\.com\/sites\/dev\/Shared Documents\/x\/x\.xlcode\.xlsx\?web=1$/,
    );
  });
  it('同期フォルダ外なら null', () => {
    expect(toWebUrl('D:\\work\\a.xlsx', roots, '\\')).toBeNull();
  });
  it('日本語・空白を URL エンコードする', () => {
    expect(joinUrl('https://x/Documents', ['開発', 'a b.xlsx'])).toBe(
      'https://x/Documents/%E9%96%8B%E7%99%BA/a%20b.xlsx?web=1',
    );
  });
});
