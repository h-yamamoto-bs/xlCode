import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const exec = promisify(execFile);

/**
 * PowerShell スクリプトを実行して標準出力を返す。
 * スクリプトは -EncodedCommand（UTF-16LE）で渡し、出力は UTF-8 にさせる（日本語のパスが文字化けしないように）
 */
export async function runPowerShell(script: string, timeout = 20_000): Promise<string> {
  const full = `[Console]::OutputEncoding = [Text.Encoding]::UTF8\n${script}`;
  const encoded = Buffer.from(full, 'utf16le').toString('base64');
  try {
    const { stdout } = await exec(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded],
      { windowsHide: true, timeout, maxBuffer: 4 * 1024 * 1024, encoding: 'utf8' },
    );
    return stdout;
  } catch (e) {
    throw new Error(describePowerShellError(e, timeout, encoded.length), { cause: e });
  }
}

/** 実行に失敗した理由を、原因が分かる形にする（終了コード・stderr の末尾・時間切れ・コマンドラインの長さ） */
export function describePowerShellError(e: unknown, timeout: number, commandLength = 0): string {
  const err = e as NodeJS.ErrnoException & { killed?: boolean; signal?: string; stderr?: string; code?: unknown };
  if (err.killed || err.signal === 'SIGTERM') return `${Math.round(timeout / 1000)} 秒たっても終わりませんでした`;
  if (err.code === 'ENOENT') return 'powershell.exe が見つかりません';
  if (err.code === 'E2BIG' || err.code === 'ENAMETOOLONG') return `コマンドラインが長すぎます（${commandLength} 文字）`;
  const stderr = (err.stderr ?? '').replace(/\s+/g, ' ').trim();
  const tail = stderr.length > 400 ? `…${stderr.slice(-400)}` : stderr;
  const code = typeof err.code === 'number' ? `終了コード ${err.code}` : (err.message ?? String(e));
  return tail ? `${code}: ${tail}` : code;
}
