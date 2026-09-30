import clsx from 'clsx';
import { useEffect, useState, type ReactNode } from 'react';
import type { XlcodeConfig } from '../../../core';
import type { ExcelMode } from '../../../shared/api';
import { api, unwrap } from '../api';
import { Icon } from './Icons';
import { Button } from './ui';

export const EXCEL_MODES: { id: ExcelMode; label: string; desc: string }[] = [
  {
    id: 'both',
    label: 'デスクトップ版と Web 版の両方',
    desc: 'デスクトップ版で開いているかは自動で検知します。Web 版で開いた可能性があるときだけ、Build / Sync 前に閉じたかを確認します。',
  },
  {
    id: 'desktop',
    label: 'デスクトップ版のみ',
    desc: '開いているかを自動で検知し、開いていれば Build / Sync を止めます。確認ダイアログは出しません。',
  },
  {
    id: 'web',
    label: 'Web 版のみ',
    desc: 'Web 版で開いている状態は検知できないため、Build / Sync 前に毎回閉じたかを確認します。',
  },
];

export function ModeLabel(mode: ExcelMode): string {
  return { both: '両方', desktop: 'デスクトップ版', web: 'Web版' }[mode];
}

function Row({ title, desc, children }: { title: string; desc?: ReactNode; children: ReactNode }) {
  return (
    <div className="border-b border-line py-4">
      <div className="text-[13px] font-semibold text-fg-strong">{title}</div>
      {desc && <div className="mt-0.5 mb-2 text-[12px] leading-relaxed text-muted">{desc}</div>}
      {children}
    </div>
  );
}

const input =
  'h-[26px] rounded-[2px] border border-[#3c3c3c] bg-input px-2 text-[13px] text-fg outline-none focus:border-accent';

export function SettingsView({
  root,
  bookRoot,
  bookCount,
  onRelocate,
  mode,
  onMode,
  onSaved,
  onError,
}: {
  root: string;
  bookRoot: string | null;
  bookCount: number;
  onRelocate: (newRoot: string | null) => void;
  mode: ExcelMode;
  onMode: (m: ExcelMode) => void;
  onSaved: () => void;
  onError: (msg: string) => void;
}) {
  const [config, setConfig] = useState<XlcodeConfig | null>(null);
  const [saved, setSaved] = useState('');
  useEffect(() => {
    unwrap(api.readConfig(root))
      .then((c) => {
        setConfig(c);
        setSaved(JSON.stringify(c));
      })
      .catch((e: Error) => onError(e.message));
  }, [root, onError]);

  const dirty = config !== null && JSON.stringify(config) !== saved;
  const save = async () => {
    if (!config) return;
    const c = { ...config, webUrlBase: config.webUrlBase?.trim() || undefined };
    try {
      await unwrap(api.writeConfig(root, c));
      setConfig(c);
      setSaved(JSON.stringify(c));
      onSaved();
    } catch (e) {
      onError((e as Error).message);
    }
  };
  const set = <K extends keyof XlcodeConfig>(k: K, v: XlcodeConfig[K]) => config && setConfig({ ...config, [k]: v });

  return (
    <div className="h-full overflow-auto">
      <div className="mx-auto max-w-[760px] px-6 py-5">
        <div className="mb-1 text-[20px] font-light text-fg-strong">設定</div>

        <div className="mt-4 mb-1 text-[11px] tracking-wide text-muted uppercase">このパソコン</div>
        <Row
          title="使う Excel"
          desc="ブックをどの Excel で編集するか。「両方」では、xlCode からデスクトップ版で開いたブックは閉じたことを自動で検知できるため確認を省きます。"
        >
          <div className="flex flex-col gap-1.5" role="radiogroup" aria-label="使う Excel">
            {EXCEL_MODES.map((m) => (
              <label
                key={m.id}
                className={clsx(
                  'flex cursor-pointer gap-2.5 rounded-[3px] border px-3 py-2',
                  mode === m.id ? 'border-accent bg-focus/40' : 'border-line hover:bg-hover',
                )}
              >
                <input
                  type="radio"
                  name="mode"
                  checked={mode === m.id}
                  onChange={() => onMode(m.id)}
                  className="mt-1 accent-accent"
                />
                <span>
                  <span className="text-fg">{m.label}</span>
                  <span className="block text-[12px] leading-relaxed text-muted">{m.desc}</span>
                </span>
              </label>
            ))}
          </div>
        </Row>

        <div className="mt-6 mb-1 flex items-center text-[11px] tracking-wide text-muted uppercase">
          このプロジェクト（.xlcode/config.json）
          <Button variant="primary" className="ml-auto normal-case" onClick={save} disabled={!dirty}>
            <Icon.Save size={14} /> 保存
          </Button>
        </div>
        <Row
          title="ブックの置き場所"
          desc={
            <>
              OneDrive 内の専用フォルダを指定すると、ソースと同じフォルダ構成でブックだけをそこに置きます。
              ソース・Git・node_modules は OneDrive の外に置けます。
              {bookCount > 0 && ' 変更すると、既存のブックも移動します。'}
            </>
          }
        >
          <div className="flex flex-wrap items-center gap-2">
            <span
              className="min-w-0 flex-1 truncate rounded-[2px] border border-line bg-side px-2 py-1 font-mono text-[12.5px]"
              title={bookRoot ?? ''}
            >
              {bookRoot ?? <span className="text-muted">ソースの各フォルダの中（{root}）</span>}
            </span>
            <Button
              onClick={async () => {
                const dir = await api.pickFolder('ブックの置き場所（OneDrive 内の専用フォルダ）');
                if (dir) onRelocate(dir);
              }}
            >
              <Icon.Folder size={14} /> フォルダを選ぶ…
            </Button>
            {bookRoot && <Button onClick={() => onRelocate(null)}>ソースの中に戻す</Button>}
          </div>
        </Row>
        {config && (
          <>
            <Row title="プロジェクトの種類" desc="最初のブックを作るときに決めます。あとから変えることはできません。">
              <div className="text-[13px] text-fg">
                {config.mode === 'vba'
                  ? 'VBA モード（.bas / .cls / .frm のシートを VBA として書き込んだ .xlsm を生成）'
                  : config.mode === 'source'
                    ? 'ソースコードモード（シートをソースコードのファイルとして書き出す）'
                    : '未設定（最初のブックを作るときに選びます）'}
              </div>
            </Row>
            <Row
              title="Web 版の URL"
              desc={
                <>
                  Web 版で開くときの、ブックの置き場所（未設定ならプロジェクトルート）の URL。空欄なら Windows の
                  OneDrive 同期設定から自動で求めます。
                  <br />
                  例: https://contoso-my.sharepoint.com/personal/me_contoso_com/Documents/shop
                </>
              }
            >
              <input
                aria-label="Web 版の URL"
                className={clsx(input, 'w-full font-mono')}
                placeholder="自動"
                value={config.webUrlBase ?? ''}
                onChange={(e) => set('webUrlBase', e.target.value)}
              />
            </Row>
            <Row title="省略検知" desc="Copilot がコードを要約・省略していないか、行数・文字数の急減で検知します。">
              <div className="flex flex-wrap items-center gap-2 text-[13px]">
                <input
                  aria-label="閾値"
                  type="number"
                  min={5}
                  max={95}
                  className={clsx(input, 'w-20')}
                  value={Math.round(config.shrinkThreshold * 100)}
                  onChange={(e) => set('shrinkThreshold', Number(e.target.value) / 100)}
                />
                % 以上減ったら確認（前回
                <input
                  aria-label="最小行数"
                  type="number"
                  min={1}
                  className={clsx(input, 'w-20')}
                  value={config.shrinkMinLines}
                  onChange={(e) => set('shrinkMinLines', Number(e.target.value))}
                />
                行以上のファイルが対象）
              </div>
            </Row>
            <Row title="整形・正規化">
              <div className="flex flex-col gap-1.5">
                {(
                  [
                    ...(config.mode === 'vba' ? [] : ([['format', 'Prettier で整形する']] as const)),
                    ['trimTrailingWhitespace', '行末の空白を削除する（Markdown は対象外）'],
                    ['autoCommit', 'Build / Sync の前に Git へ自動コミットする'],
                  ] as const
                ).map(([k, label]) => (
                  <label key={k} className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      className="accent-accent"
                      checked={config[k]}
                      onChange={(e) => set(k, e.target.checked)}
                    />
                    {label}
                  </label>
                ))}
              </div>
            </Row>
          </>
        )}
      </div>
    </div>
  );
}
