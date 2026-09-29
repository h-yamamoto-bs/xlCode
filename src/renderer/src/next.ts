import type { BookSummary } from '../../shared/api';
import { EXCEL_SIDE, SOURCE_SIDE } from './status';

export type NextKind = 'build' | 'sync' | 'refresh' | 'open' | 'fix' | 'close' | 'resolve' | 'wait';

/** 画面に出す「次の操作」。ボタンの強調（primary）もこれで決める */
export interface Next {
  kind: NextKind;
  text: string;
  detail?: string;
}

/**
 * ブックの状態から、次に何をすればよいかを 1 つに決める。
 * 迷わないことを優先し、必ず 1 つだけ提案する（優先順: 止まっている理由 → 反映が必要な変更 → その他）。
 */
export function nextAction(book: BookSummary, treeVersion: string, busy: boolean): Next {
  const files = book.files ?? [];
  const excel = files.filter((f) => EXCEL_SIDE.includes(f.status)).length;
  const source = files.filter((f) => SOURCE_SIDE.includes(f.status)).length;
  const conflict = files.filter((f) => f.status === 'conflict').length;
  if (busy) return { kind: 'wait', text: '処理中です。終わるまでお待ちください' };
  if (book.loadError || (book.errors?.length ?? 0) > 0) {
    return { kind: 'fix', text: '上のエラーを解決すると Build / Sync できます' };
  }
  if (book.open) {
    return {
      kind: 'close',
      text: 'Excel を閉じると Build / Sync できます',
      detail: '閉じたことは自動で検知します。閉じてもこの表示が消えないときは「再読み込み」を押してください',
    };
  }
  if ((book.conflictSheets?.length ?? 0) > 0) {
    return {
      kind: 'resolve',
      text: 'Excel で衝突シートを統合してから Build',
      detail: 'Copilot に「A列とB列を統合して元のシートに書き、衝突シートを削除して」と指示してください',
    };
  }
  if (conflict > 0) {
    return {
      kind: 'build',
      text: `Build で衝突シートを作成（${conflict} ファイルが両側で変更）`,
      detail: '作成した衝突シートを Excel で Copilot に統合させ、もう一度 Build します',
    };
  }
  if (excel > 0) {
    return {
      kind: 'build',
      text: `Build で Excel 側の変更 ${excel} ファイルをソースコードへ出力`,
      detail: source > 0 ? `エディタ側の変更 ${source} ファイルも、あわせてシートへ取り込みます` : undefined,
    };
  }
  if (source > 0) {
    return {
      kind: 'sync',
      text: `Sync でエディタ側の変更 ${source} ファイルをシートへ反映`,
      detail: '「開く」でも Sync してから Excel を開きます',
    };
  }
  if (book.treeVersion !== treeVersion) {
    return { kind: 'refresh', text: 'Refresh Tree で #tree を最新にする', detail: 'ファイル構成が変わりました' };
  }
  return {
    kind: 'open',
    text: 'すべて同期済み。Excel で開いて Copilot に編集させるか、エディタで編集してください',
  };
}
