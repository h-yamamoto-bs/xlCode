import { execFile } from 'node:child_process';
import path from 'node:path';
import { promisify } from 'node:util';

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

/** ローカルパスに対応する Web URL を作る（最も深い同期フォルダを採用） */
export function toWebUrl(fileAbs: string, roots: SyncRoot[], sep = path.sep): string | null {
  const norm = (p: string) => p.replace(/[\\/]+$/, '').toLowerCase();
  const file = norm(fileAbs);
  const hit = roots
    .filter((r) => file.startsWith(norm(r.mountPoint) + sep.toLowerCase()) || file.startsWith(norm(r.mountPoint) + '/'))
    .sort((a, b) => b.mountPoint.length - a.mountPoint.length)[0];
  if (!hit) return null;
  const rel = fileAbs.slice(hit.mountPoint.replace(/[\\/]+$/, '').length + 1).split(/[\\/]/);
  return joinUrl(hit.urlNamespace, rel);
}

/** ベース URL と相対パスから、ブラウザ（Excel for the web）で開く URL を作る */
export function joinUrl(base: string, segments: string[]): string {
  const b = base.endsWith('/') ? base : `${base}/`;
  return `${b}${segments.map(encodeURIComponent).join('/')}?web=1`;
}

/** Windows の OneDrive 設定から同期フォルダ一覧を取得する（他 OS では空） */
export async function readSyncRoots(): Promise<SyncRoot[]> {
  if (process.platform !== 'win32') return [];
  try {
    const { stdout } = await exec('reg', ['query', 'HKCU\\Software\\SyncEngines\\Providers\\OneDrive', '/s'], {
      windowsHide: true,
    });
    return parseSyncRoots(stdout);
  } catch {
    return [];
  }
}
