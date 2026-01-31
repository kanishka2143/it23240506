const { test, expect } = require("@playwright/test");
const fs = require("fs");
const path = require("path");
const { loadTestCases } = require("../scripts/readExcel");

const BASE_URL = "https://www.swifttranslator.com/";

const CONFIG = {
  typingDelay: 20,
  outputTimeout: 25000,
};

function normalizeText(s) {
  return (s || "")
    .replace(/\r\n/g, "\n")
    .trim()
    .normalize("NFC");
}

// 1) Scope to the translator container using "Swap Languages" (unique to translator area). [1](https://oss.csbodima.lk/npm-singlish/)[2](https://swifttranslator.com/)
async function getTranslatorContainer(page) {
  const swapBtn = page.getByRole("button", { name: /Swap Languages/i });
  await expect(swapBtn).toBeVisible({ timeout: 20000 });

  // Go up to a container that includes both panels
  const container = swapBtn.locator("..").locator("..");
  await expect(container).toBeVisible({ timeout: 20000 });
  return container;
}

// 2) Input textbox (as shown in snapshot). [1](https://oss.csbodima.lk/npm-singlish/)
async function getInput(page) {
  const input = page.getByRole("textbox", { name: /Input Your Singlish Text Here/i });
  await expect(input).toBeVisible({ timeout: 20000 });
  return input;
}

// 3) Read Sinhala output ONLY from Sinhala panel
async function readSinhalaOutput(container) {
  // Sinhala header inside translator area (not the top switch)
  const sinhalaHeader = container.getByText("Sinhala", { exact: true }).first();

  // Sinhala panel is the parent of the "Sinhala" heading
  const sinhalaPanel = sinhalaHeader.locator("..");
  await expect(sinhalaPanel).toBeVisible({ timeout: 20000 });

  // First: most UIs show output as the first non-button text right under Sinhala header
  const sibling = sinhalaHeader.locator("xpath=following-sibling::*[1]");

  // Wait for sibling to contain Sinhala characters
  try {
    await expect(sibling).toContainText(/[\u0D80-\u0DFF]/, { timeout: 8000 });
    return normalizeText(await sibling.innerText());
  } catch {
    // Fallback: search inside Sinhala panel only (exclude buttons)
    const candidates = sinhalaPanel
      .locator("div, span, p, pre, textarea, [role='textbox']")
      .filter({ hasNot: sinhalaPanel.locator("button") });

    // Wait until we find some Sinhala output inside Sinhala panel
    await expect
      .poll(async () => {
        const count = await candidates.count();
        let best = "";
        for (let i = 0; i < Math.min(count, 25); i++) {
          const txt = normalizeText(await candidates.nth(i).innerText().catch(() => ""));
          if (/[\u0D80-\u0DFF]/.test(txt) && txt.length > best.length) best = txt;
        }
        return best;
      }, { timeout: CONFIG.outputTimeout })
      .not.toBe("");

    // Return best Sinhala-looking text
    const count = await candidates.count();
    let best = "";
    for (let i = 0; i < Math.min(count, 25); i++) {
      const txt = normalizeText(await candidates.nth(i).innerText().catch(() => ""));
      if (/[\u0D80-\u0DFF]/.test(txt) && txt.length > best.length) best = txt;
    }
    return best;
  }
}

// Load Excel test cases (your file has Notes above headers; loader already handles that). [3](https://mysliit-my.sharepoint.com/personal/it23355446_my_sliit_lk/_layouts/15/Doc.aspx?sourcedoc=%7B5630C3D4-38E3-4FB4-B4EE-5F516FAF39DA%7D&file=IT23355446-Test_cases.xlsx&action=default&mobileredirect=true)
const testCases = loadTestCases()
  .filter((tc) => tc.tcId && tc.name)
  .map((tc, i) => ({ ...tc, __index: i }));

const results = [];

test.describe("SwiftTranslator Singlish → Sinhala (Excel-driven)", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto(BASE_URL, { waitUntil: "domcontentloaded" });
  });

  for (const tc of testCases) {
    const title = `${tc.tcId} - ${tc.name} - ${tc.__index}`;

    test(title, async ({ page }) => {
      if (!tc.input) test.skip(true, "Skipped: empty input in Excel");

      const container = await getTranslatorContainer(page);
      const input = await getInput(page);

      // Clear and input
      await input.fill("");

      const isUiTest =
        String(tc.tcId).startsWith("Pos_UI") || String(tc.tcId).startsWith("Neg_UI");

      if (isUiTest) {
        // UI test: typing for real-time behavior
        await input.click();
        await input.type(tc.input, { delay: CONFIG.typingDelay });
      } else {
        // Functional test: fill instantly for stability
        await input.fill(tc.input);
      }

      // Read Sinhala output
      const actual = await readSinhalaOutput(container);
      const expected = normalizeText(tc.expected);

      // ✅ Pass/Fail rule: Actual == Expected (your requirement). 
      const pass = normalizeText(actual) === expected;

      results.push({
        tcId: tc.tcId,
        name: tc.name,
        input: tc.input,
        expected: tc.expected,
        actual,
        status: pass ? "Pass" : "Fail",
      });

      // Assertion
      expect(normalizeText(actual)).toBe(expected);
    });
  }

  test.afterAll(async () => {
    const outDir = path.join(__dirname, "..", "results");
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);

    fs.writeFileSync(
      path.join(outDir, "test-results.json"),
      JSON.stringify({ executedAt: new Date().toISOString(), results }, null, 2),
      "utf-8"
    );

    console.log("Saved results/test-results.json");
  });
});