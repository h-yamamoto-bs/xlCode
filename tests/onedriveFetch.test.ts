import { describe, expect, it } from 'vitest';
import {
  buildFetchScript,
  cfMessage,
  compactScript,
  fetchLatest,
  outcomeOf,
  type RawFetchInfo,
} from '../src/main/onedriveFetch';
import { describePowerShellError } from '../src/main/powershell';
import { findMatches, highlight } from '../src/renderer/src/find';

const raw = (over: Partial<RawFetchInfo>): RawFetchInfo => ({
  path: 'C:\\OneDrive\\shop\\app.xlcode.xlsx',
  before: '1:100',
  after: '2:120',
  step: 'done',
  wasOnline: false,
  how: 'cf',
  cfErr: null,
  err: null,
  sts: 0,
  attrs: 32,
  ...over,
});

describe('OneDrive から最新を取り直す', () => {
  it('取り寄せたあと更新日時かサイズが変わっていれば updated、同じなら unchanged', () => {
    expect(outcomeOf(raw({})).result).toBe('updated');
    expect(outcomeOf(raw({ after: '1:100' })).result).toBe('unchanged');
    // オンラインのみだったファイルは、捨てる手順を飛ばして取り寄せただけ
    const o = outcomeOf(raw({ wasOnline: true, after: '1:100' }));
    expect(o.result).toBe('unchanged');
    expect(o.reason).toContain('オンラインのみ');
  });

  it('同期の途中・捨てられなかった・失敗は、理由つきで skipped / failed', () => {
    expect(outcomeOf(raw({ step: 'busy' }))).toMatchObject({
      result: 'skipped',
      reason: expect.stringContaining('同期の途中'),
    });
    expect(outcomeOf(raw({ step: 'no-dehydrate', how: null }))).toMatchObject({
      result: 'skipped',
      reason: expect.stringContaining('ファイル オンデマンド'),
    });
    // Cloud Files API のエラーは理由を日本語にして付ける
    expect(outcomeOf(raw({ step: 'no-dehydrate', how: null, cfErr: '0x8007016B' })).reason).toContain(
      '同期し終えていません（しばらくしてからもう一度）（0x8007016B）',
    );
    expect(outcomeOf(raw({ step: 'no-dehydrate', how: null, cfErr: '0x8007016A' })).reason).toContain(
      'プレースホルダーではない',
    );
    expect(cfMessage('0x8007017A')).toContain('常に保持する');
    expect(cfMessage('0x80070005')).toBe('Cloud Files API のエラー 0x80070005');
    expect(cfMessage('load: Add-Type failed')).toContain('読み込めませんでした（Add-Type failed）');
    expect(cfMessage(null)).toBeNull();
    expect(outcomeOf(raw({ err: 'Access denied' }))).toMatchObject({
      result: 'failed',
      reason: expect.stringContaining('Access denied'),
    });
    expect(outcomeOf(raw({ step: '' })).result).toBe('failed');
    // 診断用の生の値を付ける
    expect(JSON.parse(outcomeOf(raw({})).detail!)).toMatchObject({ step: 'done', before: '1:100', after: '2:120' });
  });

  it('スクリプトはパスの一覧をファイルで受け取り、「捨てる」→ 読む → 待つ → 元に戻す の順になっている', () => {
    const script = buildFetchScript("C:\\Users\\me\\AppData\\Local\\Temp\\xlcode-fetch-1\\it's.json");
    // ファイル名の ' はエスケープする。コメント行は送らない（コマンドラインの長さの上限のため）
    expect(script).toContain(
      "Get-Content -LiteralPath 'C:\\Users\\me\\AppData\\Local\\Temp\\xlcode-fetch-1\\it''s.json'",
    );
    expect(script.split('\n').some((l) => /^\s*#/.test(l))).toBe(false);
    expect(script).not.toContain('attrib.exe');
    const dehydrate = script.indexOf('SetPin $p $false $true');
    const cf = script.indexOf('CfDehydrate $p');
    const read = script.indexOf('ReadAllBytes');
    const restore = script.lastIndexOf('SetPin $p $pinned $unpinned');
    expect(dehydrate).toBeGreaterThan(0);
    expect(cf).toBeGreaterThan(dehydrate);
    expect(script).toContain('CfDehydratePlaceholder');
    expect(read).toBeGreaterThan(cf);
    expect(script.indexOf('Start-Sleep', read)).toBeGreaterThan(read);
    expect(restore).toBeGreaterThan(read);
    // 同期の途中（アップロード待ち・転送中・一時停止・エラー）のファイルには触らない
    expect(script).toContain('$BUSY = 0x1D');
    expect(script).toContain("step = 'busy'");
    // 途中で落ちても JSON で理由を返す
    expect(script).toContain('fatal = $_.Exception.Message');
    // -EncodedCommand（UTF-16 → Base64）にしても、Windows のコマンドラインの上限に十分収まる
    expect(Buffer.from(script, 'utf16le').toString('base64').length).toBeLessThan(16_000);
  });

  it('compactScript はコメント行と行末のコメントだけを取り除く', () => {
    expect(compactScript("# 説明\n$a = 1   # 値\n\n$b = '# これは文字列'\n")).toBe("$a = 1\n$b = '# これは文字列'");
  });

  it('PowerShell の失敗を、原因が分かる文にする', () => {
    expect(describePowerShellError({ killed: true }, 90_000)).toBe('90 秒たっても終わりませんでした');
    expect(describePowerShellError({ code: 'ENOENT' }, 1)).toBe('powershell.exe が見つかりません');
    expect(describePowerShellError({ code: 'E2BIG' }, 1, 40_000)).toBe('コマンドラインが長すぎます（40000 文字）');
    expect(describePowerShellError({ code: 1, stderr: 'At line:3 char:5\r\n+ oops\r\n' }, 1)).toBe(
      '終了コード 1: At line:3 char:5 + oops',
    );
    expect(describePowerShellError({ message: 'boom' }, 1)).toBe('boom');
  });

  it('Windows 以外では何もせず、すべて skipped', async () => {
    if (process.platform === 'win32') return;
    const r = await fetchLatest({ 'app/app.xlcode.xlsx': '/tmp/app.xlcode.xlsx' });
    expect(r).toEqual({
      oneDriveRunning: null,
      books: { 'app/app.xlcode.xlsx': { result: 'skipped', reason: 'Windows でだけ行います' } },
    });
  });
});

describe('ファイル内の検索', () => {
  const lines = ['const a = 1;', 'const b = A + a;', '', 'return a;'];

  it('大文字小文字を区別しない検索は左から順に、重ならずに見つける', () => {
    expect(findMatches(lines, 'a', false)).toEqual([
      { line: 0, start: 6, end: 7 },
      { line: 1, start: 10, end: 11 },
      { line: 1, start: 14, end: 15 },
      { line: 3, start: 7, end: 8 },
    ]);
    expect(findMatches(lines, 'A', true)).toEqual([{ line: 1, start: 10, end: 11 }]);
    expect(findMatches(lines, '', false)).toEqual([]);
    expect(findMatches(['aaaa'], 'aa', false)).toEqual([
      { line: 0, start: 0, end: 2 },
      { line: 0, start: 2, end: 4 },
    ]);
  });

  it('行を該当箇所とそれ以外に分ける', () => {
    expect(
      highlight('const b = A + a;', [
        { start: 10, end: 11, index: 1 },
        { start: 14, end: 15, index: 2 },
      ]),
    ).toEqual([
      { text: 'const b = ', match: null },
      { text: 'A', match: 1 },
      { text: ' + ', match: null },
      { text: 'a', match: 2 },
      { text: ';', match: null },
    ]);
    expect(highlight('abc', [{ start: 0, end: 3, index: 0 }])).toEqual([{ text: 'abc', match: 0 }]);
  });
});
