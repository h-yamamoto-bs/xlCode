import iconv from 'iconv-lite';

/**
 * テキストファイルの形式（文字コード・改行コード）。
 * Excel のセルには文字しか入らないため、ファイルごとの形式は xlCode が覚えておいて書き戻す。
 */
export type Encoding = 'utf8' | 'utf8bom' | 'utf16le' | 'sjis';
export type Eol = 'lf' | 'crlf';

export interface TextFormat {
  encoding: Encoding;
  eol: Eol;
}

const UTF8_BOM = Buffer.from([0xef, 0xbb, 0xbf]);
const UTF16LE_BOM = Buffer.from([0xff, 0xfe]);

/**
 * 新しく作るファイルの形式（Windows 11 日本語環境を想定）
 * - .bat / .cmd / .vbs: Shift_JIS・CRLF（コマンドプロンプト・WSH の標準）
 * - .bas / .cls / .frm / .refs: Shift_JIS・CRLF（VBA モード）
 * - .ps1 など: UTF-8（BOM あり）・CRLF（Windows PowerShell 5.1 は BOM なしを Shift_JIS として読む）
 * - .reg: UTF-16 LE・CRLF（レジストリエディタの形式）
 * - それ以外: UTF-8・LF
 */
export function defaultFormat(fileName: string): TextFormat {
  const ext = fileName.toLowerCase().split('.').pop() ?? '';
  // VBA（.bas / .cls / .frm）と参照設定（.refs）も、VBE と同じ Shift_JIS・CRLF にする
  if (['bat', 'cmd', 'vbs', 'bas', 'cls', 'frm', 'refs'].includes(ext)) return { encoding: 'sjis', eol: 'crlf' };
  if (['ps1', 'psm1', 'psd1'].includes(ext)) return { encoding: 'utf8bom', eol: 'crlf' };
  if (ext === 'reg') return { encoding: 'utf16le', eol: 'crlf' };
  return { encoding: 'utf8', eol: 'lf' };
}

export function formatLabel(f: TextFormat): string {
  const enc = { utf8: 'UTF-8', utf8bom: 'UTF-8 BOM', utf16le: 'UTF-16 LE', sjis: 'Shift_JIS' }[f.encoding];
  return `${enc} / ${f.eol === 'crlf' ? 'CRLF' : 'LF'}`;
}

export type Decoded = { kind: 'text'; text: string; format: TextFormat } | { kind: 'binary'; reason: string };

function detectEol(text: string, fallback: Eol): Eol {
  const crlf = (text.match(/\r\n/g) ?? []).length;
  const lf = (text.match(/\n/g) ?? []).length - crlf;
  if (crlf === 0 && lf === 0) return fallback;
  return crlf >= lf ? 'crlf' : 'lf';
}

function isAscii(buf: Buffer): boolean {
  for (const b of buf) if (b >= 0x80) return false;
  return true;
}

/** ファイルの内容を文字列にし、形式を判定する */
export function decodeFile(buf: Buffer, fileName: string): Decoded {
  const def = defaultFormat(fileName);
  const text = (enc: Encoding, s: string): Decoded => ({
    kind: 'text',
    text: s,
    format: { encoding: enc, eol: detectEol(s, def.eol) },
  });
  try {
    if (buf.subarray(0, 3).equals(UTF8_BOM)) {
      return text('utf8bom', new TextDecoder('utf-8', { fatal: true }).decode(buf.subarray(3)));
    }
    if (buf.subarray(0, 2).equals(UTF16LE_BOM)) {
      return text('utf16le', new TextDecoder('utf-16le', { fatal: true }).decode(buf.subarray(2)));
    }
  } catch {
    return { kind: 'binary', reason: 'BOM はあるが内容が壊れています' };
  }
  if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff)
    return { kind: 'binary', reason: 'UTF-16 BE は未対応です' };
  if (buf.includes(0)) return { kind: 'binary', reason: 'バイナリファイルです' };
  if (isAscii(buf)) {
    // ASCII だけのファイルはどの文字コードとも読める。日本語が加わったときに備え、拡張子の標準に合わせる
    // （UTF-16 だけは ASCII のファイルを別物に変えてしまうため UTF-8 にする）
    return text(def.encoding === 'utf16le' ? 'utf8' : def.encoding, buf.toString('latin1'));
  }
  try {
    return text('utf8', new TextDecoder('utf-8', { fatal: true }).decode(buf));
  } catch {
    // UTF-8 として読めなければ Shift_JIS を試す
  }
  const sjis = iconv.decode(buf, 'cp932');
  if (!sjis.includes('�')) return text('sjis', sjis);
  return { kind: 'binary', reason: 'UTF-8 でも Shift_JIS でもありません' };
}

export class EncodeError extends Error {}

/**
 * LF に正規化済みのテキストを、指定の形式でファイルの内容にする。
 * Shift_JIS で表せない文字があれば、化けさせずに例外にする。
 */
export function encodeFile(text: string, format: TextFormat, fileName: string): Buffer {
  const s = format.eol === 'crlf' ? text.replace(/\r?\n/g, '\r\n') : text;
  switch (format.encoding) {
    case 'utf8':
      return Buffer.from(s, 'utf8');
    case 'utf8bom':
      return Buffer.concat([UTF8_BOM, Buffer.from(s, 'utf8')]);
    case 'utf16le':
      return Buffer.concat([UTF16LE_BOM, Buffer.from(s, 'utf16le')]);
    case 'sjis': {
      const out = iconv.encode(s, 'cp932');
      if (iconv.decode(out, 'cp932') !== s) {
        const bad = [...new Set([...s].filter((c) => iconv.decode(iconv.encode(c, 'cp932'), 'cp932') !== c))];
        throw new EncodeError(
          `「${fileName}」は Shift_JIS のファイルですが、Shift_JIS で表せない文字があります: ${bad.slice(0, 10).join(' ')}`,
        );
      }
      return out;
    }
  }
}
