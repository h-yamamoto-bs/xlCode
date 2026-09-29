import { describe, expect, it } from 'vitest';
import { buildScript, interpret, isUnder, type RawSyncInfo } from '../src/main/syncStatus';

const raw = (p: Partial<RawSyncInfo>): RawSyncInfo => ({ path: 'x', attrs: 0x20, sts: 0, sps: null, err: null, ...p });

describe('OneDrive の同期状態の解釈', () => {
  it.each([
    [{ sts: 0 }, 'synced'],
    [{ sts: 0x1 }, 'uploading'],
    [{ sts: 0x2 }, 'downloading'],
    [{ sts: 0x4 }, 'syncing'],
    [{ sts: 0x8 | 0x1 }, 'paused'],
    [{ sts: 0x10 | 0x4 }, 'error'],
    [{ sts: 0, attrs: 0x400000 }, 'online-only'],
    [{ sts: null, attrs: 0x20 }, 'unknown'],
    [{ sts: null, attrs: 0x1000 }, 'online-only'],
    [{ err: 'Access denied' }, 'unknown'],
  ] as const)('%o → %s', (p, expected) => {
    expect(interpret(raw(p))).toBe(expected);
  });
});

describe('PowerShell スクリプト', () => {
  it('日本語パスを Base64 の UTF-8 で埋め込む', () => {
    const paths = ['C:\\Users\\me\\OneDrive - 会社\\開発\\app.xlcode.xlsx'];
    const script = buildScript(paths);
    const b64 = /FromBase64String\('([^']+)'\)/.exec(script)![1];
    expect(JSON.parse(Buffer.from(b64, 'base64').toString('utf8'))).toEqual(paths);
    expect(script).not.toContain('会社');
    expect(script).toContain('System.SyncTransferStatus');
  });
});

describe('OneDrive のフォルダ内か', () => {
  const roots = ['C:\\Users\\me\\OneDrive - Contoso\\'];
  it.each([
    ['C:\\Users\\me\\OneDrive - Contoso\\shop\\a.xlsx', true],
    ['c:\\users\\ME\\onedrive - contoso\\a.xlsx', true],
    ['C:\\Users\\me\\OneDrive - Contoso2\\a.xlsx', false],
    ['D:\\work\\a.xlsx', false],
  ] as const)('%s', (file, expected) => {
    expect(isUnder(file, roots)).toBe(expected);
  });
});
