import { copyFile, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

/**
 * 個人用 OneDrive のファイルを Excel for the web で開く URL を求める。
 *
 * 個人用 OneDrive の同期設定にある URL（https://d.docs.live.net/<cid>/…）は Office アプリ用で、ブラウザでは 404 になる。
 * ブラウザで開くにはファイルの ID（resid）が要るため、OneDrive がこの PC に持つ同期データベース
 * （%LOCALAPPDATA%\Microsoft\OneDrive\settings\Personal\SyncEngineDatabase.db）から、パスをたどって ID を探す。
 */

/** データベースの 1 行（ファイルまたはフォルダ） */
export interface ItemRow {
  id: string;
  parent: string | null;
  name: string;
}

/**
 * 同期フォルダからの相対パス（segments）に当たるファイルの ID を探す。
 * ファイル名で候補を絞り、親フォルダの名前をたどってパスと一致するものを選ぶ
 */
export function findItemId(files: ItemRow[], folders: ItemRow[], segments: string[]): string | null {
  if (segments.length === 0) return null;
  const eq = (a: string, b: string) => a.localeCompare(b, undefined, { sensitivity: 'accent' }) === 0;
  const byId = new Map(folders.map((f) => [f.id, f]));
  const dirs = segments.slice(0, -1).reverse();
  const hits = files.filter((f) => eq(f.name, segments[segments.length - 1]));
  for (const f of hits) {
    let parent = f.parent ? byId.get(f.parent) : undefined;
    let ok = true;
    for (const d of dirs) {
      if (!parent || !eq(parent.name, d)) {
        ok = false;
        break;
      }
      parent = parent.parent ? byId.get(parent.parent) : undefined;
    }
    // たどり終えたら同期フォルダの直下（残りはルートのフォルダ 1 つまで）
    const rest = parent ? (parent.parent && byId.has(parent.parent) ? 2 : 1) : 0;
    if (ok && rest <= 1) return f.id;
  }
  return null;
}

/** ファイルの ID から Excel for the web で開く URL を作る */
export function personalWebUrl(cid: string, resid: string): string {
  const id = resid.includes('!') ? resid : `${cid.toUpperCase()}!${resid}`;
  const owner = id.split('!')[0].toLowerCase() || cid.toLowerCase();
  return `https://onedrive.live.com/edit.aspx?cid=${encodeURIComponent(owner)}&resid=${encodeURIComponent(id)}&app=Excel`;
}

interface Db {
  prepare(sql: string): { all(...params: unknown[]): unknown[] };
  close(): void;
}

/** テーブルの列名から、ID・親 ID・名前の列を選ぶ */
function readRows(db: Db, table: string, nameCol: RegExp): ItemRow[] {
  const cols = (db.prepare(`PRAGMA table_info("${table}")`).all() as { name: string }[]).map((c) => c.name);
  const id = cols.find((c) => /^resourceid$/i.test(c));
  const parent = cols.find((c) => /^parentresourceid$/i.test(c));
  const name = cols.find((c) => nameCol.test(c));
  if (!id || !parent || !name) return [];
  return (
    db.prepare(`SELECT "${id}" AS id, "${parent}" AS parent, "${name}" AS name FROM "${table}"`).all() as {
      id: unknown;
      parent: unknown;
      name: unknown;
    }[]
  )
    .filter((r) => typeof r.id === 'string' && typeof r.name === 'string')
    .map((r) => ({
      id: r.id as string,
      parent: typeof r.parent === 'string' ? r.parent : null,
      name: r.name as string,
    }));
}

/** データベースから、相対パスに当たるファイルの ID を探す */
export function lookupInDb(db: Db, segments: string[]): string | null {
  const tables = (db.prepare(`SELECT name FROM sqlite_master WHERE type = 'table'`).all() as { name: string }[]).map(
    (t) => t.name,
  );
  const fileTable = tables.find((t) => /clientfile_records$/i.test(t));
  const folderTable = tables.find((t) => /clientfolder_records$/i.test(t));
  if (!fileTable || !folderTable) return null;
  return findItemId(readRows(db, fileTable, /^filename$/i), readRows(db, folderTable, /^foldername$/i), segments);
}

/**
 * 個人用 OneDrive のファイルの Web URL を求める（求められなければ null）。
 * OneDrive が使用中のデータベースは直接開かず、一時フォルダへ写してから読む
 */
export async function resolvePersonalUrl(fileAbs: string, mountPoint: string, cid: string): Promise<string | null> {
  const local = process.env.LOCALAPPDATA;
  if (!local) return null;
  const src = path.join(local, 'Microsoft', 'OneDrive', 'settings', 'Personal', 'SyncEngineDatabase.db');
  const rel = path.win32.relative(mountPoint, fileAbs);
  if (!rel || rel.startsWith('..')) return null;
  const tmp = await mkdtemp(path.join(tmpdir(), 'xlcode-od-'));
  try {
    const dst = path.join(tmp, 'db.sqlite');
    await copyFile(src, dst);
    for (const ext of ['-wal', '-shm']) await copyFile(src + ext, dst + ext).catch(() => {});
    const { DatabaseSync } = await import('node:sqlite');
    // 写しなので読み書きで開く（WAL の取り込みに書き込みが要る）
    const db = new DatabaseSync(dst);
    try {
      const id = lookupInDb(db, rel.split(/[\\/]/));
      return id ? personalWebUrl(cid, id) : null;
    } finally {
      db.close();
    }
  } catch {
    return null;
  } finally {
    await rm(tmp, { recursive: true, force: true }).catch(() => {});
  }
}
