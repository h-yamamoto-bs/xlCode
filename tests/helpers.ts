import { execFileSync } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { Book } from '../src/core/workbook';

export interface Fixture {
  root: string;
  file: (rel: string) => string;
  write: (rel: string, text: string) => Promise<void>;
  read: (rel: string) => Promise<string>;
  git: (...args: string[]) => string;
  /** Copilot の編集を模してブックを書き換える */
  editBook: (rel: string, fn: (b: Book) => void) => Promise<void>;
  sheet: (rel: string, name: string) => Promise<string[]>;
}

export async function fixture(files: Record<string, string> = {}, opts: { git?: boolean } = {}): Promise<Fixture> {
  const root = await mkdtemp(path.join(tmpdir(), 'xlcode-'));
  const file = (rel: string) => path.join(root, rel);
  const write = async (rel: string, text: string) => {
    await mkdir(path.dirname(file(rel)), { recursive: true });
    await writeFile(file(rel), text);
  };
  for (const [rel, text] of Object.entries(files)) await write(rel, text);
  const git = (...args: string[]) => execFileSync('git', args, { cwd: root, encoding: 'utf8' });
  if (opts.git) {
    git('init', '-q', '-b', 'main');
    git('config', 'user.name', 'test');
    git('config', 'user.email', 'test@example.com');
    git('add', '-A');
    git('commit', '-q', '-m', 'init', '--allow-empty');
  }
  return {
    root,
    file,
    write,
    read: (rel) => readFile(file(rel), 'utf8'),
    git,
    editBook: async (rel, fn) => {
      const b = await Book.load(file(rel));
      fn(b);
      await b.save(file(rel));
    },
    sheet: async (rel, name) => {
      const b = await Book.load(file(rel));
      return b.readSheet(name).lines;
    },
  };
}
