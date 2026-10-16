import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { FetchOutcome, FetchReport } from '../shared/api';
import { runPowerShell } from './powershell';
import { isUnder, oneDriveRoots, oneDriveRunning } from './syncStatus';

/**
 * OneDrive から、ブックの最新の内容を取り直す（Windows のみ）。
 *
 * OneDrive は、Excel（デスクトップ版）で開いていないブックの変更を、すぐにはこの PC に届けないことがある。
 * Web 版で編集したあとに xlCode を見ても「Excel 側の変更あり」にならないのはそのため。
 * 同期を指示する公開の API は無いので、ファイル オンデマンドの仕組みを使う:
 *   1. Cloud Files API の CfDehydratePlaceholder（エクスプローラーの「空き領域を増やす」と同じ）で
 *      ローカルの中身をその場で捨てる。API が使えなければ attrib +U で OneDrive に捨てさせる（遅い）
 *   2. ファイルを読んで（「このデバイス上で常に保持する」と同じ要求）クラウドの今の内容を取り寄せ
 *   3. 同期状態が落ち着くまで待つ
 * アップロード待ち・転送中・一時停止・エラーのブックと、Excel で開いているブックには触らない。
 */

/** PowerShell から受け取る、ブックごとの結果 */
export interface RawFetchInfo {
  path: string;
  /** 取得前後の更新日時（ticks）とサイズ */
  before: string | null;
  after: string | null;
  /** どこまで進んだか: busy（同期中などで触らない）/ no-dehydrate（捨てられなかった）/ done */
  step: string;
  /** 取得前は「オンラインのみ」だった（捨てる手順を飛ばした） */
  wasOnline: boolean;
  /** 捨てた方法: cf（Cloud Files API）/ attrib（OneDrive に任せた）/ null（捨てていない） */
  how: string | null;
  /** Cloud Files API のエラー（HRESULT の 16 進）。API が読み込めなかったときは load: 〜 */
  cfErr: string | null;
  err: string | null;
  sts: number | null;
  attrs: number | null;
}

/** 1 冊あたりに待つ上限（ミリ秒）。捨てるのに 15 秒、取り寄せるのに 60 秒 */
export const DEHYDRATE_WAIT_MS = 15_000;
export const HYDRATE_WAIT_MS = 60_000;

/**
 * PowerShell に渡すスクリプト。
 * パスの一覧は UTF-8 の JSON ファイル（pathsFile）で渡す。コマンドライン（-EncodedCommand）には
 * 長さの上限（約 32,000 文字）があり、ブックが多いと収まらないため。
 * コメント行は送る前に取り除く（compactScript）
 */
export function buildFetchScript(pathsFile: string): string {
  const file = pathsFile.replace(/'/g, "''");
  return compactScript(`
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
try {
$paths = ConvertFrom-Json (Get-Content -LiteralPath '${file}' -Raw -Encoding UTF8)
$sh = New-Object -ComObject Shell.Application
# Cloud Files API（cldapi.dll）。エクスプローラーの「空き領域を増やす」が使うもの
$cfLoadErr = $null
try {
  Add-Type -Namespace XlCode -Name Cf -MemberDefinition @'
[DllImport("cldapi.dll", CharSet = CharSet.Unicode)]
public static extern int CfOpenFileWithOplock(string FilePath, int Flags, out IntPtr ProtectedHandle);
[DllImport("cldapi.dll")]
public static extern void CfCloseHandle(IntPtr FileHandle);
[DllImport("cldapi.dll")]
public static extern int CfDehydratePlaceholder(IntPtr FileHandle, long StartingOffset, long Length, int DehydrateFlags, IntPtr Overlapped);
'@
} catch { $cfLoadErr = $_.Exception.Message }
function CfDehydrate($p) {
  $h = [IntPtr]::Zero
  $hr = [XlCode.Cf]::CfOpenFileWithOplock($p, 0, [ref]$h)
  if ($hr -ne 0) { return ('0x{0:X8}' -f $hr) }
  try { $hr = [XlCode.Cf]::CfDehydratePlaceholder($h, 0, -1, 0, [IntPtr]::Zero) } finally { [XlCode.Cf]::CfCloseHandle($h) }
  if ($hr -ne 0) { return ('0x{0:X8}' -f $hr) }
  return $null
}
$ONLINE = 0x401000   # FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS | FILE_ATTRIBUTE_OFFLINE
$PINNED = 0x80000
$UNPINNED = 0x100000
$BUSY = 0x1D         # STS_NEEDSUPLOAD | STS_TRANSFERRING | STS_PAUSED | STS_HASERROR
$MOVING = 0x6        # STS_NEEDSDOWNLOAD | STS_TRANSFERRING
function Stamp($p) { $f = Get-Item -LiteralPath $p -Force; "$($f.LastWriteTimeUtc.Ticks):$($f.Length)" }
function Attrs($p) { [int64][int](Get-Item -LiteralPath $p -Force).Attributes }
# 「常に保持」「空き領域を増やす」の印（FILE_ATTRIBUTE_PINNED / UNPINNED）を付け外しする。attrib +P / +U と同じ
function SetPin($p, $pin, $unpin) {
  $a = Attrs $p
  $a = $a -band (-bnot ($PINNED -bor $UNPINNED))
  if ($pin) { $a = $a -bor $PINNED }
  if ($unpin) { $a = $a -bor $UNPINNED }
  # SetFileAttributes で付けられる属性だけ渡す（ReparsePoint・RecallOnDataAccess などは渡せない）
  $a = $a -band 0x1831A7
  if ($a -eq 0) { $a = 0x80 }
  [IO.File]::SetAttributes($p, [IO.FileAttributes]$a)
}
function Sts($p) {
  $i = $sh.Namespace([IO.Path]::GetDirectoryName($p)).ParseName([IO.Path]::GetFileName($p))
  if ($i) { $v = $i.ExtendedProperty('System.SyncTransferStatus'); if ($null -ne $v) { return [int64]$v } }
  return $null
}
$out = @()
foreach ($p in @($paths)) {
  $r = [ordered]@{ path = $p; before = $null; after = $null; step = ''; wasOnline = $false; how = $null; cfErr = $null; err = $null; sts = $null; attrs = $null }
  try {
    $r.before = Stamp $p
    $attrs = Attrs $p
    $sts = Sts $p
    $r.sts = $sts; $r.attrs = $attrs
    if (($null -ne $sts) -and ($sts -band $BUSY)) { $r.step = 'busy'; $out += [pscustomobject]$r; continue }
    $pinned = ($attrs -band $PINNED) -ne 0
    $unpinned = ($attrs -band $UNPINNED) -ne 0
    $online = ($attrs -band $ONLINE) -ne 0
    $r.wasOnline = $online
    if (-not $online) {
      # 1. 空き領域を増やす（ローカルの中身を捨てる）。「常に保持」のままでは捨てられないので先に外す
      SetPin $p $false $true
      if ($cfLoadErr) { $r.cfErr = "load: $cfLoadErr" } else { $r.cfErr = CfDehydrate $p }
      if (((Attrs $p) -band $ONLINE) -ne 0) { $online = $true; $r.how = 'cf' }
      if (-not $online) {
        # API で捨てられなければ、attrib +U を見た OneDrive が捨てるのを待つ（時間がかかることがある）
        $deadline = (Get-Date).AddMilliseconds(${DEHYDRATE_WAIT_MS})
        while ((Get-Date) -lt $deadline) {
          if (((Attrs $p) -band $ONLINE) -ne 0) { $online = $true; $r.how = 'attrib'; break }
          Start-Sleep -Milliseconds 500
        }
      }
      if (-not $online) {
        # ファイル オンデマンドが無効か、OneDrive が応じない。元に戻して諦める
        SetPin $p $pinned $unpinned
        $r.step = 'no-dehydrate'; $r.attrs = Attrs $p; $r.sts = Sts $p
        $out += [pscustomobject]$r; continue
      }
    }
    # 2. 読んで取り寄せる（クラウドの今の内容が降りてくる）
    [void][IO.File]::ReadAllBytes($p)
    # 3. 落ち着くまで待つ
    $deadline = (Get-Date).AddMilliseconds(${HYDRATE_WAIT_MS})
    while ((Get-Date) -lt $deadline) {
      $a = Attrs $p; $s = Sts $p
      if ((($a -band $ONLINE) -eq 0) -and (($null -eq $s) -or (($s -band $MOVING) -eq 0))) { break }
      Start-Sleep -Milliseconds 500
    }
    # 元の「常に保持」「空き領域を増やす」の設定に戻す
    SetPin $p $pinned $unpinned
    $r.after = Stamp $p
    $r.attrs = Attrs $p; $r.sts = Sts $p
    $r.step = 'done'
  } catch {
    $r.err = $_.Exception.Message
    # 取り寄せに失敗してオンラインのみのままなら、OneDrive が動き出したときに取り寄せるよう「常に保持」にしておく
    try { if (((Attrs $p) -band $ONLINE) -ne 0) { SetPin $p $true $false } } catch {}
  }
  $out += [pscustomobject]$r
}
ConvertTo-Json -InputObject @($out) -Compress
} catch {
  ConvertTo-Json -InputObject @{ fatal = $_.Exception.Message; at = "$($_.InvocationInfo.ScriptLineNumber)" } -Compress
}
`);
}

/** コメント行と行末のコメントを取り除く（コマンドラインを短くする。文字列の中の # は無い前提） */
export function compactScript(script: string): string {
  return script
    .split('\n')
    .filter((l) => !/^\s*#/.test(l) && l.trim() !== '')
    .map((l) => l.replace(/\s+#\s[^'"]*$/, ''))
    .join('\n');
}

/** Cloud Files API のエラー（winerror.h の ERROR_CLOUD_FILE_*）を日本語に */
export function cfMessage(hresult: string | null): string | null {
  if (!hresult) return null;
  if (hresult.startsWith('load: ')) return `Cloud Files API を読み込めませんでした（${hresult.slice(6)}）`;
  const code = Number.parseInt(hresult.slice(-4), 16);
  const known: Record<number, string> = {
    0x16a: 'OneDrive のファイル オンデマンドの対象ではありません（プレースホルダーではない）',
    0x16b: 'OneDrive がまだこのファイルをクラウドと同期し終えていません（しばらくしてからもう一度）',
    0x174: 'OneDrive の認証が切れています',
    0x176: 'ネットワークに接続できません',
    0x178: 'OneDrive の同期フォルダの中ではありません',
    0x179: '他のアプリが開いています',
    0x17a: '「このデバイス上で常に保持する」になっています',
    0x17d: 'アクセスが拒否されました',
    0x186: 'OneDrive が動いていません',
    0x187: 'OneDrive の同期フォルダではありません',
  };
  const win32 = /^0x8007[0-9a-f]{4}$/i.test(hresult) ? known[code] : undefined;
  return win32 ? `${win32}（${hresult}）` : `Cloud Files API のエラー ${hresult}`;
}

/** 生の結果から、画面に出す結果を決める */
export function outcomeOf(raw: RawFetchInfo): FetchOutcome {
  const detail = JSON.stringify({
    step: raw.step,
    how: raw.how,
    cf: raw.cfErr,
    before: raw.before,
    after: raw.after,
    attrs: raw.attrs,
    sts: raw.sts,
  });
  if (raw.err) {
    return { result: 'failed', reason: `OneDrive から取り寄せられませんでした: ${raw.err}`, detail };
  }
  if (raw.step === 'busy') {
    return {
      result: 'skipped',
      reason: '同期の途中（アップロード待ち・転送中・一時停止・エラー）なので触りませんでした',
      detail,
    };
  }
  if (raw.step === 'no-dehydrate') {
    const cf = cfMessage(raw.cfErr);
    return {
      result: 'skipped',
      reason: cf
        ? `ローカルの中身を捨てられませんでした: ${cf}`
        : 'OneDrive がローカルの中身を捨ててくれませんでした（ファイル オンデマンドが無効か、同期が止まっています）',
      detail,
    };
  }
  if (raw.step !== 'done') return { result: 'failed', reason: `結果が分かりません（${raw.step || '不明'}）`, detail };
  const result = raw.before === raw.after ? 'unchanged' : 'updated';
  if (raw.wasOnline) return { result, reason: 'オンラインのみだったので取り寄せました', detail };
  return { result, detail };
}

/**
 * 各ブックの最新の内容を OneDrive から取り直し、落ち着くまで待つ。
 * `books` は rel → 絶対パス。Excel で開いているブックは呼ぶ側で除いておく
 */
export async function fetchLatest(books: Record<string, string>): Promise<FetchReport> {
  const entries = Object.entries(books);
  const out: Record<string, FetchOutcome> = {};
  if (process.platform !== 'win32') {
    for (const [rel] of entries) out[rel] = { result: 'skipped', reason: 'Windows でだけ行います' };
    return { oneDriveRunning: null, books: out };
  }
  const roots = await oneDriveRoots();
  const inside = entries.filter(([, abs]) => isUnder(abs, roots));
  for (const [rel] of entries)
    if (!inside.some(([r]) => r === rel)) out[rel] = { result: 'skipped', reason: 'OneDrive 外' };
  if (inside.length === 0) return { oneDriveRunning: null, books: out };
  const running = await oneDriveRunning();
  if (running === false) {
    for (const [rel] of inside) out[rel] = { result: 'skipped', reason: 'OneDrive が起動していません' };
    return { oneDriveRunning: running, books: out };
  }
  let raws: RawFetchInfo[];
  const dir = await mkdtemp(path.join(os.tmpdir(), 'xlcode-fetch-'));
  try {
    const pathsFile = path.join(dir, 'paths.json');
    await writeFile(pathsFile, '\uFEFF' + JSON.stringify(inside.map(([, abs]) => abs)), 'utf8');
    const timeout = 30_000 + inside.length * (DEHYDRATE_WAIT_MS + HYDRATE_WAIT_MS + 15_000);
    const stdout = await runPowerShell(buildFetchScript(pathsFile), timeout);
    const parsed = JSON.parse(stdout) as RawFetchInfo[] | { fatal: string; at: string };
    if (!Array.isArray(parsed)) throw new Error(`スクリプトの ${parsed.at} 行目: ${parsed.fatal}`);
    raws = parsed;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    for (const [rel] of inside) out[rel] = { result: 'failed', reason: `PowerShell を実行できません: ${msg}` };
    return { oneDriveRunning: running, books: out };
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
  for (const [rel, abs] of inside) {
    const raw = raws.find((r) => r.path === abs);
    out[rel] = raw ? outcomeOf(raw) : { result: 'failed', reason: '結果が返りませんでした' };
  }
  return { oneDriveRunning: running, books: out };
}
