// Đo tốc độ lần mở đầu trên 4G giả lập (9 Mbps, trễ 170 ms, không bộ nhớ đệm, khổ 390) — cùng thông số
// biên bản GĐ2 lần 1. Chụp màn hình ở giây 3/5/8 và ghi FCP, DOMContentLoaded, load vào perf4g.json.
//   node perf4g.js <thư_mục_lưu_ảnh>      (cần Playwright + Chromium)
const { chromium } = require('playwright');
(async () => {
  const out = process.argv[2];
  const browser = await chromium.launch();
  const res = [];
  for (const p of ['/login', '/', '/register']) {
    for (let run = 0; run < 2; run++) {
      const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 3, isMobile: true, hasTouch: true,
        userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1' });
      const page = await ctx.newPage();
      const cdp = await ctx.newCDPSession(page);
      await cdp.send('Network.enable');
      await cdp.send('Network.setCacheDisabled', { cacheDisabled: true });
      await cdp.send('Network.emulateNetworkConditions', { offline: false, latency: 170, downloadThroughput: 9 * 1024 * 1024 / 8, uploadThroughput: 9 * 1024 * 1024 / 8 });
      let bytes = 0; page.on('response', async r => { try { const h = await r.headerValue('content-length'); bytes += Number(h || 0); } catch {} });
      const t0 = Date.now();
      await page.goto('https://timviec123.vn' + p, { waitUntil: 'commit', timeout: 90000 });
      const shots = {};
      for (const s of [3, 5, 8]) { await page.waitForTimeout(Math.max(0, s * 1000 - (Date.now() - t0))); if (run === 0) { const f = `${out}/4g${p.replace(/\//g,'_')||'_home'}_${s}s.png`; await page.screenshot({ path: f }); shots[s] = f; } }
      await page.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
      const m = await page.evaluate(() => { const fcp = performance.getEntriesByName('first-contentful-paint')[0]; const nav = performance.getEntriesByType('navigation')[0];
        return { fcp: fcp ? Math.round(fcp.startTime) : null, dcl: Math.round(nav.domContentLoadedEventEnd), load: Math.round(nav.loadEventEnd) }; });
      res.push({ page: p, run, ...m, idle_ms: Date.now() - t0, kb: Math.round(bytes / 1024) });
      await ctx.close();
    }
  }
  console.log(JSON.stringify(res, null, 1));
  require('fs').writeFileSync(out + '/perf4g.json', JSON.stringify(res, null, 1));
  await browser.close();
})();
