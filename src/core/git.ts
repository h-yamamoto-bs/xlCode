import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const exec = promisify(execFile);

async function git(root: string, args: string[]): Promise<string> {
  const { stdout } = await exec('git', args, { cwd: root, maxBuffer: 16 * 1024 * 1024 });
  return stdout;
}

export async function isGitRepo(root: string): Promise<boolean> {
  try {
    return (await git(root, ['rev-parse', '--is-inside-work-tree'])).trim() === 'true';
  } catch {
    return false;
  }
}

/** ブックが担当するディレクトリ直下のファイルだけを指す pathspec（ブック自体は除く） */
function dirPathspec(dirRel: string): string[] {
  const prefix = dirRel === '' ? '' : `${dirRel}/`;
  return [`:(top,glob)${prefix}*`, ':(top,exclude,glob)**/*.xlcode.xlsx', ':(top,exclude,glob)**/~$*'];
}

/** 対象ディレクトリの未コミット変更（git status --porcelain の行） */
export async function uncommittedChanges(root: string, dirRel: string): Promise<string[]> {
  const out = await git(root, ['status', '--porcelain=v1', '--untracked-files=all', '--', ...dirPathspec(dirRel)]);
  return out.split('\n').filter((l) => l.trim() !== '');
}

/**
 * Build / Sync 直前の自動コミット（8.3）。対象ディレクトリ直下のみをコミットする。
 * 変更が無ければ何もしない。
 */
export async function autoCommit(root: string, dirRel: string, message: string): Promise<boolean> {
  const spec = dirPathspec(dirRel);
  if ((await uncommittedChanges(root, dirRel)).length === 0) return false;
  await git(root, ['add', '-A', '--', ...spec]);
  const staged = await git(root, ['diff', '--cached', '--name-only', '--', ...spec]);
  if (staged.trim() === '') return false;
  await git(root, ['commit', '-q', '-m', message, '--', ...spec]);
  return true;
}
