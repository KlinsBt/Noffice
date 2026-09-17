import { defineConfig } from 'vitest/config';
import { sveltekit } from '@sveltejs/kit/vite';
export default defineConfig({
  plugins: [sveltekit()],
  test: {
    environment: 'jsdom',
    include: ['src/**/*.test.ts'],
    setupFiles: ['src/test-setup.ts'],
    // Cold Office conversions contend for memory/CPU across jsdom workers.
    // Keep the per-test timeout intact while bounding concurrent conversions.
    maxWorkers: 2,
  },
  build: { chunkSizeWarningLimit: 1800 },
});
