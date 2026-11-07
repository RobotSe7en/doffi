import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests/browser",
  fullyParallel: false,
  workers: 1,
  forbidOnly: Boolean(process.env.CI),
  // Passing tests need no trace capture. A retry records the complete failure
  // context; even a passing retry remains a failed gate.
  retries: 1,
  failOnFlakyTests: true,
  snapshotPathTemplate: "{testDir}/{testFilePath}-snapshots/{arg}-{platform}{ext}",
  projects: [
    { name: "chromium", use: { browserName: "chromium", ...(process.env.DOFFI_BROWSER_EXECUTABLE ? { launchOptions: { executablePath: process.env.DOFFI_BROWSER_EXECUTABLE } } : {}) } },
    { name: "firefox", testMatch: /(?:workspace|workbook-(?:storage|grid|preview|input|read-worker|reading-performance|playground|drawing(?:-window)?)|pagination|compatibility|remote-access|word-closure|presentation(?:-(?:editor|workbench|closure))?)\.spec\.ts/, use: { browserName: "firefox" } },
    { name: "webkit", testMatch: /(?:workspace|workbook-(?:storage|grid|preview|input|read-worker|reading-performance|playground|drawing(?:-window)?)|pagination|compatibility|remote-access|word-closure|presentation(?:-(?:editor|workbench|closure))?)\.spec\.ts/, use: { browserName: "webkit" } },
  ],
  use: {
    browserName: "chromium",
    ...(process.env.DOFFI_BROWSER_CHANNEL
      ? { channel: process.env.DOFFI_BROWSER_CHANNEL }
      : {}),
    baseURL: "http://127.0.0.1:4173",
    viewport: { width: 1280, height: 800 },
    trace: "on-first-retry",
    screenshot: "only-on-failure",
  },
  webServer: {
    command: "pnpm --filter @doffi/playground preview --port 4173 --strictPort",
    url: "http://127.0.0.1:4173",
    reuseExistingServer: false,
    timeout: 30_000,
  },
});
