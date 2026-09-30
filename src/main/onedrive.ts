import { execFile } from 'node:child_process';
import path from 'node:path';
import { promisify } from 'node:util';
import { runPowerShell } from './powershell';

const exec = promisify(execFile);

export interface SyncRoot {
  /** ローカルの同期フォルダ（例: C:\Users\me\OneDrive - Contoso） */
  mountPoint: string;
  /** 対応する Web 上の URL（例: https://contoso.sharepoint.com/personal/me/Documents/） */
  urlNamespace: string;
}

/**
 * `reg query HKCU\Software\SyncEngines\Providers\OneDrive /s` の出力から同期フォルダを取り出す。
 * キーごとに MountPoint と UrlNamespace の組がある。
 */
export function parseSyncRoots(regOutput: string): SyncRoot[] {
  const roots: SyncRoot[] = [];
  let cur: Partial<SyncRoot> = {};
  const flush = () => {
    if (cur.mountPoint && cur.urlNamespace) roots.push(cur as SyncRoot);
    cur = {};
  };
  for (const line of regOutput.split(/\r?\n/)) {
    if (/^HKEY_/.test(line)) {
      flush();
      continue;
    }
    const m = /^\s+(\S+)\s+REG_\w+\s+(.*)$/.exec(line);
    if (!m) continue;
    if (m[1] === 'MountPoint') cur.mountPoint = m[2].trim();
    if (m[1] === 'UrlNamespace') cur.urlNamespace = m[2].trim();
  }
  flush();
  return roots;
}

/** ローカルパスを含む同期フォルダ（最も深いもの） */
export function findRoot(fileAbs: string, roots: SyncRoot[], sep = path.sep): SyncRoot | null {
  const norm = (p: string) => p.replace(/[\\/]+$/, '').toLowerCase();
  const file = norm(fileAbs);
  return (
    roots
      .filter(
        (r) => file.startsWith(norm(r.mountPoint) + sep.toLowerCase()) || file.startsWith(norm(r.mountPoint) + '/'),
      )
      .sort((a, b) => b.mountPoint.length - a.mountPoint.length)[0] ?? null
  );
}

/** ローカルパスに対応する Web URL を作る（最も深い同期フォルダを採用） */
export function toWebUrl(fileAbs: string, roots: SyncRoot[], sep = path.sep): string | null {
  const hit = findRoot(fileAbs, roots, sep);
  if (!hit) return null;
  const rel = fileAbs.slice(hit.mountPoint.replace(/[\\/]+$/, '').length + 1).split(/[\\/]/);
  return joinUrl(hit.urlNamespace, rel);
}

/** ベース URL と相対パスから、ブラウザ（Excel for the web）で開く URL を作る */
export function joinUrl(base: string, segments: string[]): string {
  // ブラウザからコピーした URL の ?web=1 などは除く
  const bare = base.replace(/[?#].*$/, '');
  const b = bare.endsWith('/') ? bare : `${bare}/`;
  return `${b}${segments.map(encodeURIComponent).join('/')}?web=1`;
}

/** PowerShell から受け取るレジストリの値（同期フォルダ、または OneDrive のアカウント） */
export interface RawRoot {
  mountPoint?: string;
  urlNamespace?: string;
  /** HKCU\Software\Microsoft\OneDrive\Accounts\<account> */
  account?: string;
  userFolder?: string;
  cid?: string;
  serviceEndpointUri?: string;
}

/**
 * レジストリの値から同期フォルダ一覧を作る。
 * SyncEngines（同期フォルダごとの URL）を優先し、無ければ OneDrive のアカウント設定から個人用フォルダの URL を求める
 */
export function rootsFromRaw(raws: RawRoot[]): SyncRoot[] {
  const roots: SyncRoot[] = [];
  for (const r of raws) {
    if (r.mountPoint && r.urlNamespace) roots.push({ mountPoint: r.mountPoint, urlNamespace: r.urlNamespace });
  }
  for (const r of raws) {
    if (!r.userFolder || roots.some((x) => x.mountPoint.toLowerCase() === r.userFolder!.toLowerCase())) continue;
    // 職場・学校: https://contoso-my.sharepoint.com/personal/me_contoso_com/_api → …/Documents/
    const api = r.serviceEndpointUri?.match(/^(https:\/\/[^/]+\/personal\/[^/]+)\/_api\/?$/i);
    if (api) roots.push({ mountPoint: r.userFolder, urlNamespace: `${api[1]}/Documents/` });
    else if (r.account?.toLowerCase() === 'personal' && r.cid)
      roots.push({ mountPoint: r.userFolder, urlNamespace: `https://d.docs.live.net/${r.cid}/` });
  }
  return roots;
}

/** 個人用 OneDrive（d.docs.live.net）の URL か。ブラウザでは開けない（404 になる） */
export function isPersonalUrl(url: string): boolean {
  return /^https:\/\/d\.docs\.live\.net\//i.test(url);
}

/** 個人用 OneDrive の URL から cid を取り出す */
export function personalCid(url: string): string | null {
  return /^https:\/\/d\.docs\.live\.net\/([^/?#]+)/i.exec(url)?.[1] ?? null;
}

const REGISTRY_SCRIPT = `
$ErrorActionPreference = 'SilentlyContinue'
$out = @()
foreach ($k in @(Get-ChildItem -LiteralPath 'HKCU:\\Software\\SyncEngines\\Providers\\OneDrive')) {
  $p = Get-ItemProperty -LiteralPath $k.PSPath
  if ($p.MountPoint -and $p.UrlNamespace) {
    $out += [pscustomobject]@{ mountPoint = [string]$p.MountPoint; urlNamespace = [string]$p.UrlNamespace }
  }
}
foreach ($k in @(Get-ChildItem -LiteralPath 'HKCU:\\Software\\Microsoft\\OneDrive\\Accounts')) {
  $p = Get-ItemProperty -LiteralPath $k.PSPath
  if ($p.UserFolder) {
    $out += [pscustomobject]@{ account = [string]$k.PSChildName; userFolder = [string]$p.UserFolder; cid = [string]$p.cid; serviceEndpointUri = [string]$p.ServiceEndpointUri }
  }
}
ConvertTo-Json -InputObject @($out) -Compress
`;

/**
 * Windows の OneDrive 設定から同期フォルダ一覧を取得する（他 OS では空）。
 * reg.exe の出力はコードページ（日本語環境では CP932）で届き、「OneDrive - 株式会社…」などが文字化けするため、PowerShell で UTF-8 にして読む
 */
export async function readSyncRoots(): Promise<SyncRoot[]> {
  if (process.platform !== 'win32') return [];
  try {
    const out = (await runPowerShell(REGISTRY_SCRIPT)).trim();
    if (out) return rootsFromRaw(JSON.parse(out) as RawRoot[]);
  } catch {
    // 下の reg.exe に任せる
  }
  try {
    const { stdout } = await exec('reg', ['query', 'HKCU\\Software\\SyncEngines\\Providers\\OneDrive', '/s'], {
      windowsHide: true,
    });
    return parseSyncRoots(stdout);
  } catch {
    return [];
  }
}

/**
 * エクスプローラーの右クリックメニューにある OneDrive の「オンラインで表示」を実行する。
 * 実行できたら ok: true。できなければ、見つかったメニュー項目を返す（原因を調べる用）
 */
export async function openOnlineViaOneDrive(fileAbs: string): Promise<{ ok: boolean; verbs: string[] }> {
  if (process.platform !== 'win32') return { ok: false, verbs: [] };
  const b64 = Buffer.from(fileAbs, 'utf8').toString('base64');
  try {
    const out = await runPowerShell(`
$ErrorActionPreference = 'Stop'
$p = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${b64}'))
$sh = New-Object -ComObject Shell.Application
$item = $sh.Namespace([IO.Path]::GetDirectoryName($p)).ParseName([IO.Path]::GetFileName($p))
$verbs = @($item.Verbs())
$v = $verbs | Where-Object { ($_.Name -replace '&', '') -match '(オンラインで表示|オンラインで開く|View online|Open online)' } | Select-Object -First 1
if ($v) { $v.DoIt(); 'ok' } else { ($verbs | ForEach-Object { ($_.Name -replace '&', '') } | Where-Object { $_ }) -join "\n" }
`);
    const lines = out.trim().split(/\r?\n/).filter(Boolean);
    return lines[0] === 'ok' ? { ok: true, verbs: [] } : { ok: false, verbs: lines };
  } catch (e) {
    return { ok: false, verbs: [`（メニューを調べられませんでした: ${e instanceof Error ? e.message : String(e)}）`] };
  }
}
