import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const exec = promisify(execFile);

async function git(root: string, args: string[]): Promise<string> {
  // core.quotepath=false: 日本語のファイル名を \346\227… のようにエスケープしない
  const { stdout } = await exec('git', ['-c', 'core.quotepath=false', ...args], {
    cwd: root,
    maxBuffer: 16 * 1024 * 1024,
    windowsHide: true,
  });
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

/** git status --porcelain=v1 -z の出力からパスを取り出す（名前変更は新しい方のパス） */
export function parseStatusZ(out: string): string[] {
  const parts = out.split('\0');
  const paths: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (p.length < 4) continue;
    paths.push(p.slice(3));
    // R / C は次の要素が元のパス
    if (p[0] === 'R' || p[0] === 'C') i++;
  }
  return paths;
}

/** 対象ディレクトリの未コミット変更（パスの一覧） */
export async function uncommittedChanges(root: string, dirRel: string): Promise<string[]> {
  const out = await git(root, [
    'status',
    '--porcelain=v1',
    '-z',
    '--untracked-files=all',
    '--',
    ...dirPathspec(dirRel),
  ]);
  return parseStatusZ(out);
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

export interface GitSummary {
  repo: boolean;
  branch?: string;
  changes: string[];
}

/** GUI のステータスバー用 */
export async function gitSummary(root: string): Promise<GitSummary> {
  if (!(await isGitRepo(root))) return { repo: false, changes: [] };
  const branch = (await git(root, ['branch', '--show-current'])).trim() || '(detached)';
  const out = await git(root, ['status', '--porcelain=v1', '-z', '--untracked-files=all']);
  return { repo: true, branch, changes: parseStatusZ(out) };
}
