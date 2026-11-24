/**
 * Placeholder pictures of the deck template (assets/template/*.jpg) until the real ones are given: soft,
 * neutral textures in the template's colors (ivory, gold, navy), drawn by a headless browser.
 * Replace the files with the template's photos (same names) and logo.png to use the real brand.
 */
import { chromium } from "playwright-core";

const scenes: Record<string, { w: number; h: number; css: string }> = {
  "cover.jpg": { w: 1200, h: 1350, css: "radial-gradient(120% 80% at 70% 20%, #f7f1e6 0%, #e9dcc4 38%, #cdb48a 70%, #8c7347 100%)" },
  "section.jpg": { w: 2400, h: 1350, css: "radial-gradient(90% 120% at 85% 30%, #55607a 0%, #2f374a 45%, #1b2030 100%)" },
  "agenda.jpg": { w: 800, h: 1350, css: "linear-gradient(170deg, #efe6d6 0%, #d9c6a3 50%, #b39251 100%)" },
  "end.jpg": { w: 2400, h: 1350, css: "radial-gradient(70% 90% at 50% 40%, #3b3a37 0%, #1d1c1a 60%, #0d0d0c 100%)" },
};

const b = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" }).catch(() => chromium.launch());
for (const [name, s] of Object.entries(scenes)) {
  const p = await b.newPage({ viewport: { width: s.w, height: s.h } });
  // Silk-like folds (soft light curves) and a little grain, so the flat gradient reads as a photograph.
  await p.setContent(`<body style="margin:0"><canvas id="c" width="${s.w}" height="${s.h}" style="display:block"></canvas>
  <div style="position:fixed;inset:0;background:${s.css};mix-blend-mode:multiply"></div>
  <script>
    const c = document.getElementById('c'), x = c.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
    for (let i = 0; i < 14; i++) {
      const g = x.createLinearGradient(0, 0, c.width, c.height);
      g.addColorStop(0, 'rgba(255,255,255,0)'); g.addColorStop(0.5, 'rgba(0,0,0,' + (0.04 + Math.random() * 0.05) + ')'); g.addColorStop(1, 'rgba(255,255,255,0)');
      x.strokeStyle = g; x.lineWidth = 60 + Math.random() * 160; x.beginPath();
      const y0 = Math.random() * c.height; x.moveTo(-100, y0);
      x.bezierCurveTo(c.width * 0.3, y0 - 300 + Math.random() * 600, c.width * 0.7, y0 - 300 + Math.random() * 600, c.width + 100, y0 + Math.random() * 200 - 100);
      x.stroke();
    }
    const d = x.getImageData(0, 0, c.width, c.height);
    for (let i = 0; i < d.data.length; i += 4) { const n = (Math.random() - 0.5) * 14; d.data[i] += n; d.data[i+1] += n; d.data[i+2] += n; }
    x.putImageData(d, 0, 0);
  </script></body>`);
  await p.screenshot({ path: `assets/template/${name}`, type: "jpeg", quality: 85 });
  await p.close();
  console.log(name);
}
await b.close();
