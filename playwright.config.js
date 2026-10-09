import { defineConfig } from '@playwright/test';

export default defineConfig({
    testDir: './tests/browser',
    fullyParallel: true,
    workers: process.env.CI ? 2 : 3,
    reporter: process.env.CI ? 'github' : 'list',
    use: { baseURL: 'http://127.0.0.1:19483', headless: true },
    projects: [
        { name: 'chrome', use: { browserName: 'chromium', channel: 'chrome' } },
        { name: 'edge', use: { browserName: 'chromium', channel: 'msedge' } },
        { name: 'firefox', use: { browserName: 'firefox' } }
    ],
    webServer: {
        command: 'node scripts/serve.mjs', url: 'http://127.0.0.1:19483/tests/fixtures/quotation.html',
        reuseExistingServer: !process.env.CI
    }
});
