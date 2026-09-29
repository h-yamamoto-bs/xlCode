import { access, open } from 'node:fs/promises';
import path from 'node:path';

export interface OpenCheck {
  open: boolean;
  reason?: string;
}

async function exists(p: string): Promise<boolean> {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

/**
 * ブックがデスクトップ版 Excel で開かれているかを判定する（11.2）。
 * - Excel が作る所有者ファイル（~$ で始まる）の有無
 * - Windows では開いているファイルを書き込みモードで開けない（EBUSY 等）
 * 注意: Web 版 Excel で開いている状態はローカルからは検知できない。
 */
export async function checkBookOpen(bookAbs: string, opts: { quick?: boolean } = {}): Promise<OpenCheck> {
  const dir = path.dirname(bookAbs);
  const name = path.basename(bookAbs);
  // 長いファイル名では先頭2文字が ~$ に置き換わる
  for (const owner of [`~$${name}`, `~$${name.slice(2)}`]) {
    if (await exists(path.join(dir, owner)))
      return { open: true, reason: `Excel のロックファイル ${owner} があります` };
  }
  // quick: 定期確認用。ファイルを掴むと、同じ瞬間に Excel が開こうとしたとき「使用中」になり得るため、
  // 所有者ファイルだけを見る。書き込み用に開けるかの確認は、書き込み直前だけ行う。
  if (opts.quick) return { open: false };
  try {
    const fh = await open(bookAbs, 'r+');
    await fh.close();
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    if (code === 'EBUSY' || code === 'EPERM' || code === 'EACCES') {
      return { open: true, reason: `ブックを書き込み用に開けません（${code}）` };
    }
    if (code !== 'ENOENT') throw e;
  }
  return { open: false };
}
