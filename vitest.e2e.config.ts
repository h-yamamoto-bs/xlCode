import { defineConfig } from 'vitest/config';

// GUI の E2E テスト（事前に electron-vite build が必要。Linux では xvfb-run 経由で実行）
export default defineConfig({
  test: {
    include: ['e2e/**/*.e2e.ts'],
    testTimeout: 60_000,
    hookTimeout: 60_000,
    fileParallelism: false,
  },
});
