import { execFile } from 'node:child_process';
import { stat } from 'node:fs/promises';
import path from 'node:path';
import { promisify } from 'node:util';
import type { BookSync, SyncReport, SyncState } from '../shared/api';
import { readSyncRoots } from './onedrive';
import { runPowerShell } from './powershell';

const exec = promisify(execFile);

/** PowerShell から受け取る生の値 */
export interface RawSyncInfo {
  path: string;
  /** ファイル属性（FILE_ATTRIBUTE_*） */
  attrs: number | null;
  /** シェルプロパティ System.SyncTransferStatus（STS_* のビットの組み合わせ） */
  sts: number | null;
  /** シェルプロパティ System.StorageProviderState（参考値。解釈には使わない） */
  sps: number | null;
  err: string | null;
}

// FILE_ATTRIBUTE_*
const RECALL_ON_DATA_ACCESS = 0x400000;
const OFFLINE = 0x1000;
// STS_*（propkey.h の SYNC_TRANSFER_STATUS）
const STS_NEEDSUPLOAD = 0x1;
const STS_NEEDSDOWNLOAD = 0x2;
const STS_TRANSFERRING = 0x4;
const STS_PAUSED = 0x8;
const STS_HASERROR = 0x10;

/** 生の値から同期状態を決める（優先度: エラー > 一時停止 > 転送中 > 未ダウンロード > 同期済み） */
export function interpret(raw: RawSyncInfo): SyncState {
  if (raw.err) return 'unknown';
  const sts = raw.sts;
  const attrs = raw.attrs ?? 0;
  if (typeof sts === 'number') {
    if (sts & STS_HASERROR) return 'error';
    if (sts & STS_PAUSED) return 'paused';
    if (sts & STS_NEEDSUPLOAD) return 'uploading';
    if (sts & STS_NEEDSDOWNLOAD) return 'downloading';
    if (sts & STS_TRANSFERRING) return 'syncing';
  }
  if (attrs & (RECALL_ON_DATA_ACCESS | OFFLINE)) return 'online-only';
  if (typeof sts === 'number') return 'synced';
  return 'unknown';
}

/** PowerShell に渡すスクリプト。パスは Base64 の UTF-8 JSON で渡し、日本語パスの文字化けを避ける */
export function buildScript(paths: string[]): string {
  const b64 = Buffer.from(JSON.stringify(paths), 'utf8').toString('base64');
  return `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$paths = ConvertFrom-Json ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${b64}')))
$sh = New-Object -ComObject Shell.Application
$out = @()
foreach ($p in @($paths)) {
  $r = [ordered]@{ path = $p; attrs = $null; sts = $null; sps = $null; err = $null }
  try {
    $r.attrs = [int64][int](Get-Item -LiteralPath $p -Force).Attributes
    $item = $sh.Namespace([IO.Path]::GetDirectoryName($p)).ParseName([IO.Path]::GetFileName($p))
    if ($item) {
      $v = $item.ExtendedProperty('System.SyncTransferStatus'); if ($null -ne $v) { $r.sts = [int64]$v }
      $v = $item.ExtendedProperty('System.StorageProviderState'); if ($null -ne $v) { $r.sps = [int64]$v }
    }
  } catch { $r.err = $_.Exception.Message }
  $out += [pscustomobject]$r
}
ConvertTo-Json -InputObject @($out) -Compress
`;
}

async function oneDriveRunning(): Promise<boolean | null> {
  try {
    const { stdout } = await exec('tasklist.exe', ['/FI', 'IMAGENAME eq OneDrive.exe', '/FO', 'CSV', '/NH'], {
      windowsHide: true,
      timeout: 10_000,
    });
    return /onedrive\.exe/i.test(stdout);
  } catch {
    return null;
  }
}

/** OneDrive の同期フォルダの一覧（レジストリと環境変数から） */
async function oneDriveRoots(): Promise<string[]> {
  const roots = (await readSyncRoots()).map((r) => r.mountPoint);
  for (const k of ['OneDrive', 'OneDriveCommercial', 'OneDriveConsumer']) {
    const v = process.env[k];
    if (v) roots.push(v);
  }
  return roots;
}

export function isUnder(file: string, roots: string[]): boolean {
  const f = path.win32.normalize(file).toLowerCase();
  return roots.some((r) => {
    const root = path.win32
      .normalize(r)
      .replace(/[\\/]+$/, '')
      .toLowerCase();
    return f.startsWith(root + '\\');
  });
}

/** ブックの OneDrive 同期状態を調べる（Windows 以外では unsupported） */
export async function querySyncStatus(books: Record<string, string>): Promise<SyncReport> {
  const entries = Object.entries(books);
  if (process.platform !== 'win32') {
    return { oneDriveRunning: null, books: Object.fromEntries(entries.map(([k]) => [k, { state: 'unsupported' }])) };
  }
  const roots = await oneDriveRoots();
  const out: Record<string, BookSync> = {};
  const inside = entries.filter(([, abs]) => isUnder(abs, roots));
  for (const [rel] of entries) if (!inside.some(([r]) => r === rel)) out[rel] = { state: 'outside' };
  if (inside.length > 0) {
    let raws: RawSyncInfo[] = [];
    try {
      raws = JSON.parse(await runPowerShell(buildScript(inside.map(([, abs]) => abs)))) as RawSyncInfo[];
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      for (const [rel] of inside) out[rel] = { state: 'unknown', detail: `PowerShell を実行できません: ${msg}` };
    }
    for (const [rel, abs] of inside) {
      const raw = raws.find((r) => r.path === abs);
      if (raw)
        out[rel] = {
          state: interpret(raw),
          detail: JSON.stringify({ attrs: raw.attrs, sts: raw.sts, sps: raw.sps, err: raw.err }),
        };
    }
  }
  return { oneDriveRunning: inside.length > 0 ? await oneDriveRunning() : null, books: out };
}

/** ブックの更新日時とサイズ（Web 版での編集が届いたかの判定に使う） */
export async function bookStamp(abs: string): Promise<string> {
  const s = await stat(abs);
  return `${s.mtimeMs}:${s.size}`;
}
