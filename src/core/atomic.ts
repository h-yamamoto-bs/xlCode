import { randomBytes } from 'node:crypto';
import { chmod, copyFile, rename, rm, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';

/**
 * 一時ファイルに書いてから置き換える。書き込み途中で失敗しても元のファイルは壊れない。
 * 一時ファイル名は「~$」で始め、xlCode の走査対象や Git の対象から外れるようにする。
 */
export async function atomicWrite(file: string, data: string | Uint8Array): Promise<void> {
  const tmp = path.join(path.dirname(file), `~$xlcode-${randomBytes(4).toString('hex')}.tmp`);
  try {
    await writeFile(tmp, data, { flush: true });
    // 実行権限など元のファイルのモードを引き継ぐ
    const mode = await stat(file).then(
      (s) => s.mode,
      () => null,
    );
    if (mode !== null) await chmod(tmp, mode);
    await renameWithRetry(tmp, file);
  } catch (e) {
    await rm(tmp, { force: true });
    throw e;
  }
}

/** Windows ではウイルス対策・OneDrive が一瞬ファイルを掴むことがあるため、短く再試行する */
async function renameWithRetry(from: string, to: string): Promise<void> {
  for (let i = 0; ; i++) {
    try {
      await rename(from, to);
      return;
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code;
      if (i >= 4 || (code !== 'EPERM' && code !== 'EBUSY' && code !== 'EACCES')) throw e;
      await new Promise((r) => setTimeout(r, 100 * 2 ** i));
    }
  }
}

/** 上書き前のバックアップ（file.bak）を取る。元が無ければ何もしない */
export async function backup(file: string): Promise<void> {
  await copyFile(file, `${file}.bak`).catch((e: NodeJS.ErrnoException) => {
    if (e.code !== 'ENOENT') throw e;
  });
}
