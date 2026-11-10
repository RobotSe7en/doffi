import { expect, test } from "@playwright/test";

test("HTTP remote origins can open DOCX, annotate and create a correlated request without randomUUID", async ({ page, context }) => {
  await context.route("http://office.test:4173/**", async (route) => {
    const response = await route.fetch({ url: route.request().url().replace("http://office.test:4173", "http://127.0.0.1:4173"), headers: { ...route.request().headers(), host: "127.0.0.1:4173" } });
    await route.fulfill({ response });
  });
  await page.addInitScript(() => localStorage.setItem("doffi:editing", "preview"));
  await page.goto("http://office.test:4173/");
  expect(await page.evaluate(() => isSecureContext)).toBe(false);
  expect(await page.evaluate(() => typeof crypto.randomUUID)).toBe("undefined");
  await page.getByRole("button", { name: "打开 DOCX 样例", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await page.frameLocator("iframe").getByRole("heading", { name: "实施报告" }).evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event("selectionchange"));
  });
  await page.getByRole("tab", { name: "审阅", exact: true }).click(); await page.getByRole("button", { name: "新建批注", exact: true }).click();
  await page.getByLabel("新批注内容").fill("远程 HTTP 批注"); await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-annotation-count", "1");
  await page.getByText("选择提交的批注", { exact: true }).click(); await page.getByLabel("提交批注：远程 HTTP 批注").check(); await page.getByRole("button", { name: "统一提交批注", exact: true }).click();
  await expect(page.getByText("Playground 已接收批注请求；使用固定提案验证返回与应用。", { exact: true })).toBeVisible();
  await expect(page.getByRole("alert")).toHaveCount(0);
});
