/**
 * GUI 完成までの動作確認用コマンドライン
 *   npx tsx src/cli.ts <command> [project] [book] [--yes] [--discard]
 */
import path from 'node:path';
import { createInterface } from 'node:readline/promises';
import { bookStatus, build, createBook, initProject, refreshTree, sync, type OpResult } from './core';

const [cmd, projectArg = '.', target, ...rest] = process.argv.slice(2);
const flags = new Set([target, ...rest].filter((a) => a?.startsWith('--')));
const root = path.resolve(projectArg);
const book = target && !target.startsWith('--') ? path.resolve(root, target) : undefined;

async function ask(q: string): Promise<boolean> {
  if (flags.has('--yes')) return true;
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  const a = await rl.question(`${q} [y/N] `);
  rl.close();
  return a.trim().toLowerCase() === 'y';
}

function print(r: OpResult): void {
  for (const c of r.changes) console.log(`  ${c.action.padEnd(15)} ${c.target}`);
  for (const w of r.warnings) console.log(`  警告: ${w}`);
  for (const e of r.errors) console.log(`  エラー: ${e}`);
  for (const c of r.conflicts) console.log(`  衝突: ${c.file} → ${c.sheet}`);
  console.log(`結果: ${r.status}`);
}

async function withConfirm(run: (confirmed: boolean) => Promise<OpResult>): Promise<OpResult> {
  const r = await run(false);
  if (r.status !== 'confirm') return r;
  for (const c of r.confirmations) {
    console.log(`確認: ${c.message}`);
    for (const f of c.files) console.log(`  - ${f}`);
  }
  return (await ask('続行しますか？')) ? run(true) : r;
}

function requireBook(): string {
  if (!book) throw new Error('ブックのパスを指定してください');
  return book;
}

async function main(): Promise<void> {
  switch (cmd) {
    case 'init':
      for (const l of await initProject(root)) console.log(l);
      break;
    case 'create': {
      const r = await createBook(root, path.resolve(root, target ?? '.'));
      console.log(`${r.book} を作成: ${r.sheets.join(', ')}`);
      for (const s of r.skipped) console.log(`  スキップ: ${s}`);
      break;
    }
    case 'status': {
      const s = await bookStatus(root, requireBook());
      for (const f of s.files) console.log(`  ${f.status.padEnd(15)} ${f.name}`);
      for (const e of s.errors) console.log(`  エラー: ${e}`);
      for (const w of s.warnings) console.log(`  警告: ${w}`);
      break;
    }
    case 'build':
      print(await withConfirm((confirmed) => build(root, requireBook(), { confirmed })));
      break;
    case 'sync': {
      let r = await withConfirm((confirmed) => sync(root, requireBook(), { confirmed }));
      if (r.status === 'needs-decision') {
        console.log(`未 Build のシート変更があります: ${r.unbuilt.join(', ')}`);
        if (
          flags.has('--discard') ||
          (await ask('シートの変更を破棄して Sync しますか？（先に Build する場合は N）'))
        ) {
          r = await withConfirm((confirmed) => sync(root, requireBook(), { confirmed, discardExcelChanges: true }));
        }
      }
      print(r);
      break;
    }
    case 'tree': {
      const r = await refreshTree(root);
      console.log(`tree-version: ${r.version}`);
      for (const b of r.books) console.log(`  ${b.ok ? 'OK ' : 'NG '} ${b.book}${b.error ? ` (${b.error})` : ''}`);
      if (r.partial) console.log('警告: 一部のブックだけ更新されました。Refresh Tree を再実行してください');
      break;
    }
    default:
      console.log(`使い方: npx tsx src/cli.ts <command> [project] [book|dir]
  init   <project>                 .gitignore と Agents.md を用意
  create <project> <dir>           ディレクトリにブックを作成
  status <project> <book>          各ファイルの状態
  build  <project> <book> [--yes]  Excel → ソース
  sync   <project> <book> [--yes] [--discard]  ソース → Excel
  tree   <project>                 全ブックの #tree を更新`);
  }
}

main().catch((e) => {
  console.error(e instanceof Error ? e.message : e);
  process.exit(1);
});
