import { describe, expect, it } from 'vitest';
import { isPersonalUrl, joinUrl, parseSyncRoots, rootsFromRaw, toWebUrl } from '../src/main/onedrive';
import { buildDirTree } from '../src/renderer/src/dirTree';

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

describe('OneDrive の同期フォルダ（PowerShell 経由のレジストリ）', () => {
  it('SyncEngines を優先し、無いアカウントは Accounts から求める', () => {
    const roots = rootsFromRaw([
      {
        mountPoint: 'C:\\Users\\me\\OneDrive - 株式会社コントソ',
        urlNamespace: 'https://contoso-my.sharepoint.com/personal/me/Documents/',
      },
      {
        account: 'Business1',
        userFolder: 'C:\\Users\\me\\OneDrive - 株式会社コントソ',
        serviceEndpointUri: 'https://contoso-my.sharepoint.com/personal/me/_api',
      },
      {
        account: 'Business2',
        userFolder: 'C:\\Users\\me\\OneDrive - Fabrikam',
        serviceEndpointUri: 'https://fabrikam-my.sharepoint.com/personal/me_fabrikam_com/_api',
      },
      { account: 'Personal', userFolder: 'C:\\Users\\me\\OneDrive', cid: 'abc123' },
    ]);
    expect(roots).toEqual([
      {
        mountPoint: 'C:\\Users\\me\\OneDrive - 株式会社コントソ',
        urlNamespace: 'https://contoso-my.sharepoint.com/personal/me/Documents/',
      },
      {
        mountPoint: 'C:\\Users\\me\\OneDrive - Fabrikam',
        urlNamespace: 'https://fabrikam-my.sharepoint.com/personal/me_fabrikam_com/Documents/',
      },
      { mountPoint: 'C:\\Users\\me\\OneDrive', urlNamespace: 'https://d.docs.live.net/abc123/' },
    ]);
    expect(toWebUrl('C:\\Users\\me\\OneDrive - 株式会社コントソ\\a\\a.xlcode.xlsx', roots, '\\')).toBe(
      'https://contoso-my.sharepoint.com/personal/me/Documents/a/a.xlcode.xlsx?web=1',
    );
    expect(isPersonalUrl(toWebUrl('C:\\Users\\me\\OneDrive\\a.xlsx', roots, '\\')!)).toBe(true);
  });
  it('ブラウザからコピーした URL のクエリは除く', () => {
    expect(joinUrl('https://x/Documents/shop?csf=1&web=1', ['a.xlsx'])).toBe('https://x/Documents/shop/a.xlsx?web=1');
  });
});

describe('サイドバーの階層', () => {
  it('ディレクトリごとにまとめ、子が 1 つだけの階層は a/b にまとめる', () => {
    const t = buildDirTree(
      ['', 'src/app', 'src/app/x', 'src/lib', 'tools/deep/one'],
      (d) => d,
      (d) => d,
    );
    expect(t.items).toEqual(['']);
    expect(t.dirs.map((d) => d.label)).toEqual(['src', 'tools/deep/one']);
    expect(t.dirs[0].dirs.map((d) => [d.label, d.items])).toEqual([
      ['app', ['src/app']],
      ['lib', ['src/lib']],
    ]);
    expect(t.dirs[0].dirs[0].dirs[0].path).toBe('src/app/x');
  });
});
