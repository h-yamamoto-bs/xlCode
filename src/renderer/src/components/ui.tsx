import clsx from 'clsx';
import type { ButtonHTMLAttributes, ComponentProps, ReactNode } from 'react';

export function Button({
  variant = 'secondary',
  className,
  ...rest
}: ComponentProps<'button'> & { variant?: 'primary' | 'secondary' | 'danger' }) {
  return (
    <button
      className={clsx(
        'inline-flex h-[26px] items-center gap-1.5 rounded-[2px] px-3 text-[13px] whitespace-nowrap disabled:cursor-default disabled:opacity-50',
        variant === 'primary' && 'bg-accent text-white enabled:hover:bg-accent-hover',
        variant === 'secondary' && 'bg-[#313131] text-fg enabled:hover:bg-[#3c3c3c]',
        variant === 'danger' && 'bg-[#8b2020] text-white enabled:hover:bg-[#a52a2a]',
        className,
      )}
      {...rest}
    />
  );
}

/** ツールバーのアイコンボタン */
export function IconButton({ title, className, ...rest }: ButtonHTMLAttributes<HTMLButtonElement> & { title: string }) {
  return (
    <button
      title={title}
      aria-label={title}
      className={clsx(
        'inline-flex h-[22px] w-[22px] items-center justify-center rounded-[5px] text-fg enabled:hover:bg-[#ffffff1a] disabled:opacity-40',
        className,
      )}
      {...rest}
    />
  );
}

/** サイドバーの折りたたみセクション */
export function Section({
  title,
  open,
  onToggle,
  actions,
  children,
  grow,
}: {
  title: string;
  open: boolean;
  onToggle: () => void;
  actions?: ReactNode;
  children: ReactNode;
  grow?: boolean;
}) {
  return (
    <div className={clsx('flex min-h-0 flex-col border-t border-line first:border-t-0', open && grow && 'flex-1')}>
      <div className="group flex h-[22px] shrink-0 cursor-pointer items-center pr-2 pl-0.5" onClick={onToggle}>
        <span className="flex w-5 justify-center text-fg">{open ? '▾' : '▸'}</span>
        <span className="flex-1 text-[11px] font-bold tracking-wide text-fg uppercase">{title}</span>
        <div className="flex gap-0.5 opacity-0 group-hover:opacity-100" onClick={(e) => e.stopPropagation()}>
          {actions}
        </div>
      </div>
      {open && <div className="min-h-0 flex-1 overflow-auto">{children}</div>}
    </div>
  );
}

export function Kbd({ children }: { children: ReactNode }) {
  return (
    <kbd className="rounded-[3px] border border-[#ffffff26] bg-[#ffffff0f] px-1 font-mono text-[11px] text-muted">
      {children}
    </kbd>
  );
}
