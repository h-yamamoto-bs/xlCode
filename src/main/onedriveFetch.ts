import type { FetchOutcome, FetchReport } from '../shared/api';
import { runPowerShell } from './powershell';
import { isUnder, oneDriveRoots, oneDriveRunning } from './syncStatus';

/**
 * OneDrive から、ブックの最新の内容を取り直す（Windows のみ）。
 *
 * OneDrive は、Excel（デスクトップ版）で開いていないブックの変更を、すぐにはこの PC に届けないことがある。
 * Web 版で編集したあとに xlCode を見ても「Excel 側の変更あり」にならないのはそのため。
 * 同期を指示する公開の API は無いので、ファイル オンデマンドの仕組みを使う:
 *   1. attrib +U（「空き領域を増やす」）でローカルの中身を捨てさせ
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
  err: string | null;
  sts: number | null;
  attrs: number | null;
}

/** 1 冊あたりに待つ上限（ミリ秒）。捨てるのに 15 秒、取り寄せるのに 60 秒 */
export const DEHYDRATE_WAIT_MS = 15_000;
export const HYDRATE_WAIT_MS = 60_000;

/** PowerShell に渡すスクリプト。パスは Base64 の UTF-8 JSON で渡す */
export function buildFetchScript(paths: string[]): string {
  const b64 = Buffer.from(JSON.stringify(paths), 'utf8').toString('base64');
  return `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$paths = ConvertFrom-Json ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${b64}')))
$sh = New-Object -ComObject Shell.Application
$ONLINE = 0x401000   # FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS | FILE_ATTRIBUTE_OFFLINE
$PINNED = 0x80000
$UNPINNED = 0x100000
$BUSY = 0x1D         # STS_NEEDSUPLOAD | STS_TRANSFERRING | STS_PAUSED | STS_HASERROR
$MOVING = 0x6        # STS_NEEDSDOWNLOAD | STS_TRANSFERRING
function Stamp($p) { $f = Get-Item -LiteralPath $p -Force; "$($f.LastWriteTimeUtc.Ticks):$($f.Length)" }
function Attrs($p) { [int64][int](Get-Item -LiteralPath $p -Force).Attributes }
function Sts($p) {
  $i = $sh.Namespace([IO.Path]::GetDirectoryName($p)).ParseName([IO.Path]::GetFileName($p))
  if ($i) { $v = $i.ExtendedProperty('System.SyncTransferStatus'); if ($null -ne $v) { return [int64]$v } }
  return $null
}
$out = @()
foreach ($p in @($paths)) {
  $r = [ordered]@{ path = $p; before = $null; after = $null; step = ''; wasOnline = $false; err = $null; sts = $null; attrs = $null }
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
      # 1. 空き領域を増やす（ローカルの中身を捨てる）。同期済みのファイルなら数秒で消える
      & attrib.exe +U -P $p | Out-Null
      $deadline = (Get-Date).AddMilliseconds(${DEHYDRATE_WAIT_MS})
      while ((Get-Date) -lt $deadline) {
        if (((Attrs $p) -band $ONLINE) -ne 0) { $online = $true; break }
        Start-Sleep -Milliseconds 500
      }
      if (-not $online) {
        # ファイル オンデマンドが無効か、OneDrive が応じない。元に戻して諦める
        if ($pinned) { & attrib.exe +P -U $p | Out-Null } elseif (-not $unpinned) { & attrib.exe -U $p | Out-Null }
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
    if ($pinned) { & attrib.exe +P -U $p | Out-Null } elseif (-not $unpinned) { & attrib.exe -U $p | Out-Null }
    $r.after = Stamp $p
    $r.attrs = Attrs $p; $r.sts = Sts $p
    $r.step = 'done'
  } catch {
    $r.err = $_.Exception.Message
    # 取り寄せに失敗してオンラインのみのままなら、OneDrive が動き出したときに取り寄せるよう「常に保持」にしておく
    try { if (((Attrs $p) -band $ONLINE) -ne 0) { & attrib.exe +P -U $p | Out-Null } } catch {}
  }
  $out += [pscustomobject]$r
}
ConvertTo-Json -InputObject @($out) -Compress
`;
}

/** 生の結果から、画面に出す結果を決める */
export function outcomeOf(raw: RawFetchInfo): FetchOutcome {
  const detail = JSON.stringify({
    step: raw.step,
    before: raw.before,
    after: raw.after,
    attrs: raw.attrs,
    sts: raw.sts,
  });
  if (raw.err) {
    return {
      result: 'failed',
      reason: `OneDrive から取り寄せられませんでした: ${raw.err}`,
      detail,
    };
  }
  if (raw.step === 'busy') {
    return {
      result: 'skipped',
      reason: '同期の途中（アップロード待ち・転送中・一時停止・エラー）なので触りませんでした',
      detail,
    };
  }
  if (raw.step === 'no-dehydrate') {
    return {
      result: 'skipped',
      reason:
        'OneDrive がローカルの中身を捨ててくれませんでした（ファイル オンデマンドが無効か、同期が止まっています）',
      detail,
    };
  }
  if (raw.step !== 'done') return { result: 'failed', reason: `結果が分かりません（${raw.step || '不明'}）`, detail };
  if (raw.wasOnline)
    return {
      result: raw.before === raw.after ? 'unchanged' : 'updated',
      reason: 'オンラインのみだったので取り寄せました',
      detail,
    };
  return { result: raw.before === raw.after ? 'unchanged' : 'updated', detail };
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
  try {
    const timeout = 30_000 + inside.length * (DEHYDRATE_WAIT_MS + HYDRATE_WAIT_MS + 15_000);
    raws = JSON.parse(await runPowerShell(buildFetchScript(inside.map(([, abs]) => abs)), timeout)) as RawFetchInfo[];
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    for (const [rel] of inside) out[rel] = { result: 'failed', reason: `PowerShell を実行できません: ${msg}` };
    return { oneDriveRunning: running, books: out };
  }
  for (const [rel, abs] of inside) {
    const raw = raws.find((r) => r.path === abs);
    out[rel] = raw ? outcomeOf(raw) : { result: 'failed', reason: '結果が返りませんでした' };
  }
  return { oneDriveRunning: running, books: out };
}
