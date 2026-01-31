const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests",
  timeout: 90000,
  expect: { timeout: 25000 },
  use: {
    headless: true,
    navigationTimeout: 60000,
    actionTimeout: 60000,
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
    video: "retain-on-failure",
  },
  reporter: [["html"], ["list"]],
});