import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import { applyDocumentTheme, PREFERS_DARK, readSavedAppTheme, resolveAppTheme } from './appTheme';
import './index.css';

// 最初の描画の前に、保存してある表示モードを当てる（一瞬だけ別の色で光らないように）
applyDocumentTheme(resolveAppTheme(readSavedAppTheme(), window.matchMedia(PREFERS_DARK).matches));

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
