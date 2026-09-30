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
  const { stdout } = await exec(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded],
    { windowsHide: true, timeout, maxBuffer: 4 * 1024 * 1024, encoding: 'utf8' },
  );
  return stdout;
}
