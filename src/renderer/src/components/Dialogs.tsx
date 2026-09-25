import { useCallback, useEffect, useRef, useState, type ReactNode } from 'react';
import { Button } from './ui';
import { Icon } from './Icons';

export interface DialogButton<T> {
  label: string;
  value: T;
  variant?: 'primary' | 'secondary' | 'danger';
}

export interface DialogSpec<T> {
  title: string;
  icon?: 'warning' | 'info' | 'error';
  body: ReactNode;
  buttons: DialogButton<T>[];
  /** Esc・外側クリック時の値 */
  cancelValue: T;
  checkbox?: string;
}

interface Pending {
  spec: DialogSpec<unknown>;
  resolve: (v: { value: unknown; checked: boolean }) => void;
}

/** Promise で結果を返すモーダル。await ask({...}) で使う */
export function useDialog() {
  const [pending, setPending] = useState<Pending | null>(null);
  const ask = useCallback(
    <T,>(spec: DialogSpec<T>) =>
      new Promise<{ value: T; checked: boolean }>((resolve) => {
        setPending({ spec: spec as DialogSpec<unknown>, resolve: resolve as Pending['resolve'] });
      }),
    [],
  );
  const close = (value: unknown, checked: boolean) => {
    pending?.resolve({ value, checked });
    setPending(null);
  };
  const node = pending ? <Modal spec={pending.spec} onClose={close} /> : null;
  return { ask, node };
}

function Modal({ spec, onClose }: { spec: DialogSpec<unknown>; onClose: (v: unknown, checked: boolean) => void }) {
  const [checked, setChecked] = useState(false);
  const primary = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    primary.current?.focus();
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose(spec.cancelValue, false);
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [spec, onClose]);
  const icon =
    spec.icon === 'error' ? (
      <Icon.Error size={28} className="text-deleted" />
    ) : spec.icon === 'info' ? (
      <Icon.Info size={28} className="text-info" />
    ) : (
      <Icon.Warning size={28} className="text-warn" />
    );
  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center bg-black/40 pt-[12vh]"
      onMouseDown={() => onClose(spec.cancelValue, false)}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label={spec.title}
        className="w-[520px] max-w-[92vw] rounded-md border border-[#454545] bg-widget shadow-[0_0_16px_rgba(0,0,0,0.6)]"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="flex gap-4 px-5 pt-5 pb-3">
          <div className="shrink-0 pt-0.5">{icon}</div>
          <div className="min-w-0 flex-1">
            <div className="mb-2 text-[14px] font-semibold text-fg-strong">{spec.title}</div>
            <div className="selectable max-h-[45vh] overflow-auto text-[13px] leading-relaxed text-fg">{spec.body}</div>
            {spec.checkbox && (
              <label className="mt-3 flex items-center gap-2 text-[12px] text-muted">
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={(e) => setChecked(e.target.checked)}
                  className="accent-accent"
                />
                {spec.checkbox}
              </label>
            )}
          </div>
        </div>
        <div className="flex justify-end gap-2 px-5 pt-1 pb-4">
          {spec.buttons.map((b, i) => (
            <Button
              key={b.label}
              ref={i === 0 ? primary : undefined}
              variant={b.variant ?? 'secondary'}
              onClick={() => onClose(b.value, checked)}
            >
              {b.label}
            </Button>
          ))}
        </div>
      </div>
    </div>
  );
}
