import type { SVGProps } from 'react';
import logo from '../assets/logo.png';

/** Codicon 風の最小アイコンセット（16px グリッド） */
type P = SVGProps<SVGSVGElement> & { size?: number };

function Svg({ size = 16, children, ...rest }: P & { children: React.ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.2}
      strokeLinecap="round"
      strokeLinejoin="round"
      {...rest}
    >
      {children}
    </svg>
  );
}

export const Icon = {
  Files: (p: P) => (
    <Svg {...p}>
      <path d="M9.5 1.5h-5a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h7a1 1 0 0 0 1-1v-7z" />
      <path d="M9.5 1.5v3h3" />
      <path d="M1.5 5.5v8a1 1 0 0 0 1 1h6" />
    </Svg>
  ),
  Rules: (p: P) => (
    <Svg {...p}>
      <path d="M3 2.5h10v11H3z" />
      <path d="M5.5 5.5h5M5.5 8h5M5.5 10.5h3" />
    </Svg>
  ),
  Book: (p: P) => (
    <Svg {...p}>
      <rect x="2" y="2.5" width="12" height="11" rx="1" />
      <path d="M2 6h12M2 9.5h12M6 6v7.5" />
    </Svg>
  ),
  Folder: (p: P) => (
    <Svg {...p}>
      <path d="M1.5 3.5a1 1 0 0 1 1-1h3.3l1.5 1.5h6.2a1 1 0 0 1 1 1v7.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1z" />
    </Svg>
  ),
  File: (p: P) => (
    <Svg {...p}>
      <path d="M9.5 1.5h-5a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h7a1 1 0 0 0 1-1v-9z" />
      <path d="M9.5 1.5v3h3" />
    </Svg>
  ),
  Sync: (p: P) => (
    <Svg {...p}>
      <path d="M13.5 6.5A5.5 5.5 0 0 0 3.2 4.6M2.5 9.5a5.5 5.5 0 0 0 10.3 1.9" />
      <path d="M3 1.8v3h3M13 14.2v-3h-3" />
    </Svg>
  ),
  Build: (p: P) => (
    <Svg {...p}>
      <path d="M4.5 2.5v11l9-5.5z" />
    </Svg>
  ),
  Refresh: (p: P) => (
    <Svg {...p}>
      <path d="M13.5 8a5.5 5.5 0 1 1-1.6-3.9" />
      <path d="M12.5 1.5v3h-3" />
    </Svg>
  ),
  Terminal: (p: P) => (
    <Svg {...p}>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M4 6l2 2-2 2M8 10.5h4" />
    </Svg>
  ),
  Excel: (p: P) => (
    <Svg {...p}>
      <rect x="1.5" y="3" width="8" height="10" rx="1" />
      <path d="M3.7 5.8l3.6 4.4M7.3 5.8l-3.6 4.4" />
      <path d="M9.5 4.5h4a1 1 0 0 1 1 1v5a1 1 0 0 1-1 1h-4" />
    </Svg>
  ),
  Reveal: (p: P) => (
    <Svg {...p}>
      <path d="M1.5 3.5a1 1 0 0 1 1-1h3.3l1.5 1.5h6.2a1 1 0 0 1 1 1v7.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1z" />
      <path d="M8 7v4M6 9l2 2 2-2" />
    </Svg>
  ),
  Plus: (p: P) => (
    <Svg {...p}>
      <path d="M8 3v10M3 8h10" />
    </Svg>
  ),
  ChevronRight: (p: P) => (
    <Svg {...p}>
      <path d="M6 4l4 4-4 4" />
    </Svg>
  ),
  ChevronDown: (p: P) => (
    <Svg {...p}>
      <path d="M4 6l4 4 4-4" />
    </Svg>
  ),
  Lock: (p: P) => (
    <Svg {...p}>
      <rect x="3.5" y="7" width="9" height="7" rx="1" />
      <path d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2" />
    </Svg>
  ),
  Error: (p: P) => (
    <Svg {...p}>
      <circle cx="8" cy="8" r="6.5" />
      <path d="M5.8 5.8l4.4 4.4M10.2 5.8l-4.4 4.4" />
    </Svg>
  ),
  Warning: (p: P) => (
    <Svg {...p}>
      <path d="M8 1.8l6.5 12H1.5z" />
      <path d="M8 6.5v3.2M8 11.8v.1" />
    </Svg>
  ),
  Info: (p: P) => (
    <Svg {...p}>
      <circle cx="8" cy="8" r="6.5" />
      <path d="M8 7.2v4M8 4.8v.1" />
    </Svg>
  ),
  Check: (p: P) => (
    <Svg {...p}>
      <path d="M3 8.5l3 3 7-7" />
    </Svg>
  ),
  Branch: (p: P) => (
    <Svg {...p}>
      <circle cx="4.5" cy="3.5" r="1.5" />
      <circle cx="4.5" cy="12.5" r="1.5" />
      <circle cx="11.5" cy="5.5" r="1.5" />
      <path d="M4.5 5v6M11.5 7c0 3-7 2-7 4" />
    </Svg>
  ),
  Close: (p: P) => (
    <Svg {...p}>
      <path d="M4 4l8 8M12 4l-8 8" />
    </Svg>
  ),
  Spinner: (p: P) => (
    <Svg {...p} className={`spin ${p.className ?? ''}`}>
      <path d="M8 1.5a6.5 6.5 0 1 1-6.5 6.5" />
    </Svg>
  ),
  Tree: (p: P) => (
    <Svg {...p}>
      <path d="M3 2v11.5h3M3 6.5h3" />
      <rect x="6.5" y="4.5" width="7" height="4" rx=".5" />
      <rect x="6.5" y="11.5" width="7" height="3" rx=".5" />
    </Svg>
  ),
  Save: (p: P) => (
    <Svg {...p}>
      <path d="M2.5 2.5h9l2 2v9h-11z" />
      <path d="M5 2.5v3.5h5V2.5M5 13.5v-4h6v4" />
    </Svg>
  ),
  Gear: (p: P) => (
    <Svg {...p}>
      <circle cx="8" cy="8" r="2" />
      <path d="M8 1.5v2M8 12.5v2M1.5 8h2M12.5 8h2M3.4 3.4l1.4 1.4M11.2 11.2l1.4 1.4M3.4 12.6l1.4-1.4M11.2 4.8l1.4-1.4" />
    </Svg>
  ),
  Desktop: (p: P) => (
    <Svg {...p}>
      <rect x="1.5" y="2.5" width="13" height="8.5" rx="1" />
      <path d="M5.5 13.5h5M8 11v2.5" />
    </Svg>
  ),
  Cloud: (p: P) => (
    <Svg {...p}>
      <path d="M4.5 12.5a3 3 0 0 1-.3-6A4 4 0 0 1 12 6.2a3.2 3.2 0 0 1 .3 6.3z" />
    </Svg>
  ),
  Undo: (p: P) => (
    <Svg {...p}>
      <path d="M5.5 4.5 2.5 7.5l3 3" />
      <path d="M2.5 7.5h7a3 3 0 0 1 0 6H6" />
    </Svg>
  ),
  ArrowRight: (p: P) => (
    <Svg {...p}>
      <path d="M2.5 8h11M9.5 4l4 4-4 4" />
    </Svg>
  ),
  ArrowLeft: (p: P) => (
    <Svg {...p}>
      <path d="M13.5 8h-11M6.5 4l-4 4 4 4" />
    </Svg>
  ),
  Diff: (p: P) => (
    <Svg {...p}>
      <path d="M5 2.5v5M2.5 5h5M2.5 11.5h5M8.5 2.5v11" />
    </Svg>
  ),
  Lightbulb: (p: P) => (
    <Svg {...p}>
      <path d="M6 13.5h4M6.5 11.5h3M8 2a4 4 0 0 0-2.5 7.1c.4.4.7 1 .7 1.6v.8h3.6v-.8c0-.6.3-1.2.7-1.6A4 4 0 0 0 8 2z" />
    </Svg>
  ),
  /**
   * アプリのロゴ（build/icon.png と同じ画像）。画像は周囲に余白が多いので、
   * 小さく出すときは中央を切り出して図柄が大きく見えるようにする
   */
  Logo: ({ size = 16, className, crop = 1.6 }: P & { crop?: number }) => (
    <span
      className={`inline-block shrink-0 overflow-hidden rounded-[3px] align-middle ${className ?? ''}`}
      style={{ width: size, height: size }}
    >
      <img
        src={logo}
        alt=""
        draggable={false}
        style={{
          width: size * crop,
          height: size * crop,
          maxWidth: 'none',
          margin: -((size * crop - size) / 2),
          display: 'block',
        }}
      />
    </span>
  ),
};
