import clsx from 'clsx';
import type { ReactNode } from 'react';
import type { ExcelMode, ProjectInfo } from '../../../shared/api';
import { ModeLabel } from './SettingsView';
import { EXCEL_SIDE } from '../status';
import { Icon } from './Icons';

function Item({
  children,
  onClick,
  title,
  className,
}: {
  children: ReactNode;
  onClick?: () => void;
  title?: string;
  className?: string;
}) {
  const Tag = onClick ? 'button' : 'div';
  return (
    <Tag
      title={title}
      onClick={onClick}
      className={clsx('flex h-full items-center gap-1 px-2', onClick && 'hover:bg-[#ffffff1f]', className)}
    >
      {children}
    </Tag>
  );
}

export function StatusBar({
  project,
  busy,
  errors,
  warnings,
  mode,
  onMode,
  onProblems,
}: {
  project: ProjectInfo | null;
  busy: string | null;
  errors: number;
  warnings: number;
  mode: ExcelMode;
  onMode: () => void;
  onProblems: () => void;
}) {
  const editing =
    project?.books.flatMap((b) => (b.files ?? []).filter((f) => EXCEL_SIDE.includes(f.status)).map((f) => f.name)) ??
    [];
  const stale = project?.books.filter((b) => b.treeVersion !== project.treeVersion).length ?? 0;
  return (
    <div className="flex h-[22px] shrink-0 items-center border-t border-line bg-side text-[12px] text-fg">
      <div className="flex h-full items-center bg-accent px-2.5 text-white">
        <Icon.Excel size={14} />
      </div>
      {project?.git.repo ? (
        <Item title={`${project.git.changes.length} 件の未コミット変更`}>
          <Icon.Branch size={14} />
          {project.git.branch}
          {project.git.changes.length > 0 && '*'}
        </Item>
      ) : (
        project && (
          <Item className="text-warn" title="Git リポジトリではありません。Build / Sync 前のバックアップが作られません">
            <Icon.Warning size={13} /> Git なし
          </Item>
        )
      )}
      {project && (
        <Item onClick={onProblems} title="問題">
          <Icon.Error size={13} /> {errors}
          <Icon.Warning size={13} className="ml-1" /> {warnings}
        </Item>
      )}
      {busy && (
        <Item>
          <Icon.Spinner size={13} /> {busy}
        </Item>
      )}
      <div className="flex-1" />
      {project && editing.length > 0 && (
        <Item title={editing.join('\n')} className="text-modified">
          Excel側で編集中: {editing.length}
        </Item>
      )}
      {project && (
        <Item title="全ブックの #tree の状態">
          <Icon.Tree size={13} className={stale > 0 ? 'text-warn' : undefined} />
          {stale > 0 ? `#tree 古い (${stale})` : '#tree 最新'}
        </Item>
      )}
      <Item onClick={onMode} title="使う Excel（クリックで設定を開く）">
        {mode === 'web' ? <Icon.Cloud size={14} /> : <Icon.Desktop size={14} />}
        Excel: {ModeLabel(mode)}
      </Item>
    </div>
  );
}
