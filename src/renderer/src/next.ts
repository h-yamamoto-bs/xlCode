import type { ProjectMode } from '../../core';
import type { BookSummary } from '../../shared/api';
import { EXCEL_SIDE, SOURCE_SIDE } from './status';

export type NextKind = 'build' | 'sync' | 'refresh' | 'open' | 'fix' | 'close' | 'resolve' | 'wait';

/** 画面に出す「次の操作」。ボタンの強調（primary）もこれで決める */
export interface Next {
  kind: NextKind;
  text: string;
  detail?: string;
}

export interface NextContext {
  treeVersion: string;
  busy: boolean;
  /** プロジェクトの種類（VBA なら Build でビルド結果も作る） */
  projectMode?: ProjectMode;
  /** OS（VBA の Build は Windows が必要） */
  platform?: string;
}

/**
 * ブックの状態から、次に何をすればよいかを 1 つに決める。
 * 迷わないことを優先し、必ず 1 つだけ提案する（優先順: 止まっている理由 → 反映が必要な変更 → その他）。
 */
export function nextAction(book: BookSummary, ctx: NextContext): Next {
  const { treeVersion, busy } = ctx;
  const vba = ctx.projectMode === 'vba';
  const files = book.files ?? [];
  const excel = files.filter((f) => EXCEL_SIDE.includes(f.status)).length;
  const source = files.filter((f) => SOURCE_SIDE.includes(f.status)).length;
  const conflict = files.filter((f) => f.status === 'conflict').length;
  const outName = book.vba?.output.split(/[\\/]/).pop() ?? '.xlsm';
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
  // VBA の Build は Windows のデスクトップ版 Excel が要る。押しても失敗するだけなので、先に伝える
  const noExcel = vba && ctx.platform !== undefined && ctx.platform !== 'win32';
  const vbaBuild = (text: string, detail?: string): Next =>
    noExcel
      ? {
          kind: 'fix',
          text: 'VBA の Build は Windows のデスクトップ版 Excel が必要です',
          detail: 'この PC では Build できません（Sync とエディタでの編集はできます）',
        }
      : { kind: 'build', text, detail };
  if (conflict > 0) {
    return vbaBuild(
      `Build で衝突シートを作成（${conflict} ファイルが両側で変更）`,
      '作成した衝突シートを Excel で Copilot に統合させ、もう一度 Build します',
    );
  }
  if (excel > 0) {
    const also = source > 0 ? `エディタ側の変更 ${source} ファイルも、あわせてシートへ取り込みます` : undefined;
    return vbaBuild(
      vba
        ? `Build で Excel 側の変更 ${excel} ファイルをソースコードへ出力し、${outName} を作り直す`
        : `Build で Excel 側の変更 ${excel} ファイルをソースコードへ出力`,
      also,
    );
  }
  if (source > 0) {
    return {
      kind: 'sync',
      text: `Sync でエディタ側の変更 ${source} ファイルをシートへ反映`,
      detail: vba
        ? `${outName} に入れるには、Sync のあと Build します（「開く」でも Sync してから Excel を開きます）`
        : '「開く」でも Sync してから Excel を開きます',
    };
  }
  if (book.treeVersion !== treeVersion) {
    return { kind: 'refresh', text: 'Refresh Tree で #tree を最新にする', detail: 'ファイル構成が変わりました' };
  }
  if (vba && book.vba && !book.vba.exists) {
    return vbaBuild(
      `Build で ${outName} を作る`,
      'シートをソースコードに書き出してから、画面のシートと VBA を入れた .xlsm を作ります',
    );
  }
  return {
    kind: 'open',
    text: vba
      ? `すべて同期済み。${outName} を開いて動作を確かめるか、Excel で開いて Copilot に編集させてください`
      : 'すべて同期済み。Excel で開いて Copilot に編集させるか、エディタで編集してください',
  };
}
