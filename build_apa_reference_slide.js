const fs = require("fs");
const path = require("path");
process.env.NODE_PATH = [
  process.env.NODE_PATH,
  "C:\\Users\\閻星澄\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules",
  "C:\\Users\\閻星澄\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules\\.pnpm\\pptxgenjs@4.0.1\\node_modules",
  "C:\\Users\\閻星澄\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules\\.pnpm\\sharp@0.34.5\\node_modules",
].filter(Boolean).join(path.delimiter);
require("module").Module._initPaths();
const pptxgen = require("C:\\Users\\閻星澄\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules\\.pnpm\\pptxgenjs@4.0.1\\node_modules\\pptxgenjs");
const sharp = require("C:\\Users\\閻星澄\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules\\.pnpm\\sharp@0.34.5\\node_modules\\sharp");

const outDir = path.join(__dirname, "簡報輸出");
fs.mkdirSync(outDir, { recursive: true });

const pptxPath = path.join(outDir, "核心文獻_APA格式補充頁.pptx");
const svgPath = path.join(outDir, "核心文獻_APA格式補充頁.svg");
const pngPath = path.join(outDir, "核心文獻_APA格式補充頁.png");

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";
pptx.author = "Codex";
pptx.subject = "核心文獻 APA 格式";
pptx.title = "核心文獻 APA 格式補充頁";
pptx.lang = "zh-TW";
pptx.theme = {
  headFontFace: "Microsoft JhengHei",
  bodyFontFace: "Microsoft JhengHei",
  lang: "zh-TW",
};

const C = {
  bg: "F7F8FA",
  ink: "172033",
  muted: "5B6475",
  line: "D8DEE8",
  soft: "EDF2F7",
  white: "FFFFFF",
  accent: "2F6F73",
  warm: "C56A2C",
};

const usageRows = [
  ["文獻", "本文參考用途"],
  ["《華語流行語的使用分析與教學建議》", "流行語來源、結構、語意、語用分類"],
  ["User Lifecycle and Linguistic Change in Online Communities", "線上社群語言規範與使用者生命週期"],
  ["Labov, Principles of Linguistic Change: Social Factors", "社會因素、世代、身份認同對語言變化的影響"],
];

const apaEntries = [
  [
    "吳思平（2019）。華語流行語的使用分析與教學建議：以卡提諾狂新聞為語料範圍〔碩士論文，國立政治大學〕。",
    "政大學術集成。http://nccur.lib.nccu.edu.tw/handle/140.119/125631",
  ],
  [
    "Danescu-Niculescu-Mizil, C., West, R., Jurafsky, D., Leskovec, J., & Potts, C. (2013).",
    "No country for old members: User lifecycle and linguistic change in online communities.",
    "In Proceedings of the 22nd International Conference on World Wide Web (pp. 307–318). ACM.",
    "https://doi.org/10.1145/2488388.2488416",
  ],
  [
    "Labov, W. (2001). Principles of linguistic change: Social factors (Vol. 2). Blackwell Publishers.",
  ],
];

function addText(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: "Microsoft JhengHei",
    color: opts.color || C.ink,
    fontSize: opts.fontSize || 16,
    bold: opts.bold || false,
    italic: opts.italic || false,
    margin: opts.margin ?? 0,
    fit: "shrink",
    valign: opts.valign || "top",
    align: opts.align || "left",
    breakLine: false,
  });
}

function addLine(slide, x, y, w, color = C.line) {
  slide.addShape(pptx.ShapeType.line, {
    x, y, w, h: 0,
    line: { color, width: 1.05 },
  });
}

function addTable(slide, x, y, w, rows, colWidths, opts = {}) {
  const rowH = opts.rowH || 0.42;
  const headH = opts.headH || 0.34;
  const fontSize = opts.fontSize || 10.5;
  const headSize = opts.headSize || 10.5;
  const h = headH + rowH * (rows.length - 1);
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: C.white },
    line: { color: C.line, width: 1 },
  });
  let yy = y;
  rows.forEach((row, r) => {
    const rh = r === 0 ? headH : rowH;
    slide.addShape(pptx.ShapeType.rect, {
      x, y: yy, w, h: rh,
      fill: { color: r === 0 ? C.soft : C.white },
      line: { color: C.line, width: 0.55 },
    });
    let xx = x;
    row.forEach((cell, c) => {
      addText(slide, cell, xx + 0.14, yy + (r === 0 ? 0.09 : 0.1), colWidths[c] - 0.22, rh - 0.1, {
        fontSize: r === 0 ? headSize : fontSize,
        bold: r === 0 || (opts.boldFirstCol && c === 0 && r > 0),
        color: r === 0 ? C.muted : C.ink,
      });
      if (c > 0) {
        slide.addShape(pptx.ShapeType.line, {
          x: xx, y, w: 0, h,
          line: { color: C.line, width: 0.7 },
        });
      }
      xx += colWidths[c];
    });
    yy += rh;
  });
  return h;
}

const slide = pptx.addSlide();
slide.background = { color: C.bg };

addText(slide, "核心文獻", 0.7, 0.45, 2.0, 0.3, { fontSize: 14, bold: true, color: C.accent });
addText(slide, "核心文獻與 APA 格式", 0.7, 0.84, 6.0, 0.5, { fontSize: 29, bold: true });
addText(slide, "上方保留原本的參考用途；下方補上可放在簡報最後的 APA 格式。", 0.72, 1.48, 9.8, 0.28, { fontSize: 13.5, color: C.muted });
addLine(slide, 0.7, 1.95, 11.95);

addText(slide, "文獻用途", 0.72, 2.22, 1.4, 0.22, { fontSize: 12, bold: true, color: C.warm });
addTable(slide, 0.72, 2.56, 11.55, usageRows, [5.0, 6.55], {
  rowH: 0.43,
  headH: 0.32,
  fontSize: 10.6,
  headSize: 10.2,
  boldFirstCol: true,
});

addText(slide, "APA 格式參考文獻", 0.72, 4.28, 2.4, 0.22, { fontSize: 12, bold: true, color: C.warm });
const apaY = [4.62, 5.25, 6.32];
const apaH = [0.5, 0.9, 0.35];
apaEntries.forEach((entry, i) => {
  const y = apaY[i];
  addText(slide, `${i + 1}.`, 0.78, y + 0.02, 0.22, 0.18, { fontSize: 9.8, bold: true, color: C.accent });
  addText(slide, entry.join("\n"), 1.05, y, 11.05, apaH[i], { fontSize: 8.8, color: C.ink });
});

addLine(slide, 0.7, 6.92, 11.95);
addText(slide, "註：此頁可直接接在原簡報「核心文獻」頁後，或取代原本的文獻頁。", 0.72, 7.1, 10.5, 0.22, { fontSize: 9.5, color: C.muted });

function esc(s) {
  return String(s).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function svgTable(x, y, widths, rows, opts = {}) {
  const rowH = opts.rowH || 52;
  const headH = opts.headH || 40;
  const font = opts.font || 20;
  const totalW = widths.reduce((a, b) => a + b, 0);
  const totalH = headH + rowH * (rows.length - 1);
  let out = `<rect x="${x}" y="${y}" width="${totalW}" height="${totalH}" fill="#fff" stroke="#${C.line}" stroke-width="2"/>`;
  let yy = y;
  rows.forEach((row, r) => {
    const rh = r === 0 ? headH : rowH;
    out += `<rect x="${x}" y="${yy}" width="${totalW}" height="${rh}" fill="${r === 0 ? "#" + C.soft : "#fff"}" stroke="#${C.line}"/>`;
    let xx = x;
    row.forEach((cell, c) => {
      if (c > 0) out += `<line x1="${xx}" y1="${y}" x2="${xx}" y2="${y + totalH}" stroke="#${C.line}"/>`;
      out += `<text x="${xx + 20}" y="${yy + (r === 0 ? 27 : 32)}" class="font ${r === 0 ? "muted" : "ink"}" font-size="${r === 0 ? font - 2 : font}" font-weight="${r === 0 || (opts.boldFirstCol && c === 0 && r > 0) ? 700 : 400}">${esc(cell)}</text>`;
      xx += widths[c];
    });
    yy += rh;
  });
  return out;
}

const svgApaY = [555, 625, 755];
const svgApa = apaEntries.map((entry, i) => {
  const y = svgApaY[i];
  return `<text x="94" y="${y}" class="font accent" font-size="20" font-weight="700">${i + 1}.</text>
    ${entry.map((line, j) => `<text x="126" y="${y + j * 25}" class="font ink" font-size="19">${esc(line)}</text>`).join("")}`;
}).join("");

const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900">
  <rect width="1600" height="900" fill="#${C.bg}"/>
  <style>
    .font{font-family:"Microsoft JhengHei","Noto Sans TC",Arial,sans-serif}
    .ink{fill:#${C.ink}}.muted{fill:#${C.muted}}.accent{fill:#${C.accent}}.warm{fill:#${C.warm}}
  </style>
  <text x="84" y="82" class="font accent" font-size="27" font-weight="700">核心文獻</text>
  <text x="84" y="150" class="font ink" font-size="50" font-weight="700">核心文獻與 APA 格式</text>
  <text x="86" y="210" class="font muted" font-size="24">上方保留原本的參考用途；下方補上可放在簡報最後的 APA 格式。</text>
  <line x1="84" y1="234" x2="1518" y2="234" stroke="#${C.line}" stroke-width="2"/>
  <text x="86" y="286" class="font warm" font-size="24" font-weight="700">文獻用途</text>
  ${svgTable(86, 307, [600, 786], usageRows, {font:20, rowH:52, headH:38, boldFirstCol:true})}
  <text x="86" y="535" class="font warm" font-size="24" font-weight="700">APA 格式參考文獻</text>
  ${svgApa}
  <line x1="84" y1="834" x2="1518" y2="834" stroke="#${C.line}" stroke-width="2"/>
  <text x="86" y="870" class="font muted" font-size="18">註：此頁可直接接在原簡報「核心文獻」頁後，或取代原本的文獻頁。</text>
</svg>`;

fs.writeFileSync(svgPath, svg, "utf8");

(async () => {
  await pptx.writeFile({ fileName: pptxPath });
  await sharp(Buffer.from(svg)).png().toFile(pngPath);
  console.log(JSON.stringify({ pptxPath, svgPath, pngPath }, null, 2));
})();
