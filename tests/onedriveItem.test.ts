import { DatabaseSync } from 'node:sqlite';
import { describe, expect, it } from 'vitest';
import { isPersonalUrl, personalCid } from '../src/main/onedrive';
import { findItemId, lookupInDb, personalWebUrl, type ItemRow } from '../src/main/onedriveItem';

const folders: ItemRow[] = [
  { id: 'ABC!101', parent: null, name: 'root' },
  { id: 'ABC!200', parent: 'ABC!101', name: 'ドキュメント' },
  { id: 'ABC!201', parent: 'ABC!200', name: 'shop' },
  { id: 'ABC!202', parent: 'ABC!101', name: 'shop' },
];
const files: ItemRow[] = [
  { id: 'ABC!300', parent: 'ABC!201', name: 'app.xlcode.xlsx' },
  { id: 'ABC!301', parent: 'ABC!202', name: 'app.xlcode.xlsx' },
  { id: 'ABC!302', parent: 'ABC!101', name: 'top.xlsx' },
];

describe('個人用 OneDrive のファイル ID', () => {
  it('パスをたどって同じ名前のファイルを区別する', () => {
    expect(findItemId(files, folders, ['ドキュメント', 'shop', 'app.xlcode.xlsx'])).toBe('ABC!300');
    expect(findItemId(files, folders, ['shop', 'APP.xlcode.xlsx'])).toBe('ABC!301');
    expect(findItemId(files, folders, ['top.xlsx'])).toBe('ABC!302');
    expect(findItemId(files, folders, ['other', 'app.xlcode.xlsx'])).toBeNull();
    expect(findItemId(files, folders, ['app.xlcode.xlsx'])).toBeNull();
  });
  it('同期データベース（SQLite）から探す', () => {
    const db = new DatabaseSync(':memory:');
    db.exec(`CREATE TABLE od_ClientFile_Records (resourceID TEXT, parentResourceID TEXT, fileName TEXT, eTag TEXT);
             CREATE TABLE od_ClientFolder_Records (resourceID TEXT, parentResourceID TEXT, folderName TEXT);`);
    for (const f of folders)
      db.prepare('INSERT INTO od_ClientFolder_Records VALUES (?, ?, ?)').run(f.id, f.parent, f.name);
    for (const f of files)
      db.prepare('INSERT INTO od_ClientFile_Records VALUES (?, ?, ?, ?)').run(f.id, f.parent, f.name, 'x');
    expect(lookupInDb(db, ['ドキュメント', 'shop', 'app.xlcode.xlsx'])).toBe('ABC!300');
    db.close();
  });
  it('Web で開く URL を作る', () => {
    expect(personalWebUrl('abc', 'ABC!300')).toBe(
      'https://onedrive.live.com/edit.aspx?cid=abc&resid=ABC!300&app=Excel',
    );
    expect(personalWebUrl('abc', '300')).toBe('https://onedrive.live.com/edit.aspx?cid=abc&resid=ABC!300&app=Excel');
  });
  it('d.docs.live.net の URL から cid を取り出す', () => {
    expect(personalCid('https://d.docs.live.net/0123abcd/')).toBe('0123abcd');
    expect(personalCid('https://contoso-my.sharepoint.com/personal/me/Documents/')).toBeNull();
    expect(isPersonalUrl('https://d.docs.live.net/0123abcd')).toBe(true);
  });
});
