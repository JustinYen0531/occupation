const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");

const outDir = path.join(__dirname, "簡報輸出");
fs.mkdirSync(outDir, { recursive: true });

const pptxPath = path.join(outDir, "流行語死語化_第一頁研究核心.pptx");
const svgPath = path.join(outDir, "流行語死語化_第一頁研究核心.svg");
const pngPath = path.join(outDir, "流行語死語化_第一頁研究核心.png");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Codex";
pptx.subject = "流行語死語化研究架構";
pptx.title = "研究核心";
pptx.company = "核通";
pptx.lang = "zh-TW";
pptx.theme = {
  headFontFace: "Microsoft JhengHei",
  bodyFontFace: "Microsoft JhengHei",
  lang: "zh-TW",
};
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";

const slide = pptx.addSlide();
slide.background = { color: "F7F8FA" };

const C = {
  ink: "172033",
  muted: "5B6475",
  line: "D8DEE8",
  soft: "EDF2F7",
  accent: "2F6F73",
  accent2: "C56A2C",
  white: "FFFFFF",
};

function addText(text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontFace: "Microsoft JhengHei",
    color: opts.color || C.ink,
    fontSize: opts.fontSize || 18,
    bold: opts.bold || false,
    margin: opts.margin ?? 0,
    breakLine: false,
    fit: "shrink",
    valign: opts.valign || "top",
    align: opts.align || "left",
    paraSpaceAfterPt: opts.paraSpaceAfterPt || 0,
  });
}

function addRule(x, y, w, color = C.line, transparency = 0) {
  slide.addShape(pptx.ShapeType.line, {
    x,
    y,
    w,
    h: 0,
    line: { color, width: 1.1, transparency },
  });
}

// Header
addText("研究核心", 0.7, 0.48, 2.3, 0.34, {
  fontSize: 14,
  bold: true,
  color: C.accent,
});
addText("流行語為什麼會紅？又為什麼會退燒、變成死語？", 0.7, 0.88, 8.65, 0.86, {
  fontSize: 30,
  bold: true,
});
addText("流行語的生命週期不只和詞語本身有關，也和來源、語境、二創能力、社群規範與社會環境有關。", 0.72, 1.83, 10.9, 0.48, {
  fontSize: 16,
  color: C.muted,
});
addRule(0.7, 2.48, 11.95);

// Left thesis block
addText("主題", 0.72, 2.85, 1.8, 0.3, { fontSize: 13, bold: true, color: C.accent2 });
addText("本報告不是做流行語字典，而是看流行語如何經歷：", 0.72, 3.25, 5.0, 0.4, {
  fontSize: 17,
  color: C.ink,
});

const stages = [
  ["上升", "被看見"],
  ["成熟", "被共用"],
  ["衰退", "被說過氣"],
  ["死語化", "需要解釋"],
];
stages.forEach((s, i) => {
  const x = 0.72 + i * 1.29;
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y: 3.86,
    w: 1.06,
    h: 0.72,
    rectRadius: 0.08,
    fill: { color: i < 2 ? "E7F2F1" : "F7EFEA" },
    line: { color: i < 2 ? "BFD9D7" : "E5CFC2", width: 1 },
  });
  addText(s[0], x + 0.08, 3.98, 0.9, 0.24, {
    fontSize: 15,
    bold: true,
    align: "center",
    color: i < 2 ? C.accent : C.accent2,
  });
  addText(s[1], x + 0.08, 4.25, 0.9, 0.18, {
    fontSize: 8.5,
    align: "center",
    color: C.muted,
  });
  if (i < stages.length - 1) {
    addText("→", x + 1.1, 4.08, 0.28, 0.24, { fontSize: 15, color: C.muted, align: "center" });
  }
});

addText("核心觀點", 0.72, 5.02, 1.8, 0.28, { fontSize: 13, bold: true, color: C.accent2 });
addText("流行語的生命週期，是來源、語境與社群共同推動的變化過程。", 0.72, 5.37, 5.25, 0.74, {
  fontSize: 20,
  bold: true,
});

// Right model table
addText("本報告模型", 7.05, 2.85, 2.0, 0.3, { fontSize: 13, bold: true, color: C.accent2 });
slide.addShape(pptx.ShapeType.rect, {
  x: 7.04,
  y: 3.25,
  w: 5.45,
  h: 2.86,
  fill: { color: C.white },
  line: { color: C.line, width: 1 },
});

const rows = [
  ["關鍵因素", "作用"],
  ["來源", "決定爆紅起點"],
  ["語境依賴", "決定是否容易過時"],
  ["二創能力", "決定是否進入成熟期"],
  ["社群與環境", "決定是否被接續使用"],
];
const rowH = [0.44, 0.56, 0.56, 0.56, 0.56];
let y = 3.25;
rows.forEach((r, idx) => {
  const isHead = idx === 0;
  slide.addShape(pptx.ShapeType.rect, {
    x: 7.04,
    y,
    w: 5.45,
    h: rowH[idx],
    fill: { color: isHead ? C.soft : C.white },
    line: { color: C.line, width: 0.6 },
  });
  addText(r[0], 7.25, y + 0.13, 1.8, rowH[idx] - 0.14, {
    fontSize: isHead ? 12 : 15,
    bold: isHead || idx === 1,
    color: isHead ? C.muted : C.ink,
  });
  addText(r[1], 9.05, y + 0.13, 3.15, rowH[idx] - 0.14, {
    fontSize: isHead ? 12 : 15,
    bold: isHead,
    color: isHead ? C.muted : C.ink,
  });
  y += rowH[idx];
});
addRule(8.86, 3.25, 0, C.line);
slide.addShape(pptx.ShapeType.line, {
  x: 8.86,
  y: 3.25,
  w: 0,
  h: 2.86,
  line: { color: C.line, width: 0.8 },
});

// Footer note
addRule(0.7, 6.55, 11.95, C.line);
addText("口頭補充：本報告不是做流行語字典，而是看流行語如何經歷「上升、成熟、衰退、死語化」。", 0.72, 6.76, 11.5, 0.25, {
  fontSize: 11,
  color: C.muted,
});

// Matching SVG/PNG preview.
const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900">
  <rect width="1600" height="900" fill="#F7F8FA"/>
  <style>
    .font{font-family:"Microsoft JhengHei","Noto Sans TC",Arial,sans-serif}
    .muted{fill:#5B6475}.ink{fill:#172033}.accent{fill:#2F6F73}.warm{fill:#C56A2C}
  </style>
  <text x="84" y="91" class="font accent" font-size="28" font-weight="700">研究核心</text>
  <text x="84" y="164" class="font ink" font-size="51" font-weight="700">流行語為什麼會紅？又為什麼會退燒、變成死語？</text>
  <text x="86" y="255" class="font muted" font-size="27">流行語的生命週期不只和詞語本身有關，也和來源、語境、二創能力、社群規範與社會環境有關。</text>
  <line x1="84" y1="298" x2="1518" y2="298" stroke="#D8DEE8" stroke-width="2"/>

  <text x="86" y="375" class="font warm" font-size="24" font-weight="700">主題</text>
  <text x="86" y="438" class="font ink" font-size="31">本報告不是做流行語字典，而是看流行語如何經歷：</text>
  ${stages.map((s, i) => {
    const x = 86 + i * 155;
    const fill = i < 2 ? "#E7F2F1" : "#F7EFEA";
    const stroke = i < 2 ? "#BFD9D7" : "#E5CFC2";
    const color = i < 2 ? "#2F6F73" : "#C56A2C";
    const arrow = i < stages.length - 1 ? `<text x="${x + 130}" y="505" class="font muted" font-size="34">→</text>` : "";
    return `<rect x="${x}" y="463" width="128" height="86" rx="12" fill="${fill}" stroke="${stroke}" stroke-width="2"/>
      <text x="${x + 64}" y="505" class="font" text-anchor="middle" font-size="28" font-weight="700" fill="${color}">${s[0]}</text>
      <text x="${x + 64}" y="532" class="font muted" text-anchor="middle" font-size="16">${s[1]}</text>${arrow}`;
  }).join("")}
  <text x="86" y="635" class="font warm" font-size="24" font-weight="700">核心觀點</text>
  <text x="86" y="705" class="font ink" font-size="36" font-weight="700">流行語的生命週期，是來源、語境與社群</text>
  <text x="86" y="748" class="font ink" font-size="36" font-weight="700">共同推動的變化過程。</text>

  <text x="846" y="375" class="font warm" font-size="24" font-weight="700">本報告模型</text>
  <rect x="845" y="390" width="654" height="343" fill="#FFFFFF" stroke="#D8DEE8" stroke-width="2"/>
  <rect x="845" y="390" width="654" height="53" fill="#EDF2F7" stroke="#D8DEE8"/>
  <line x1="1064" y1="390" x2="1064" y2="733" stroke="#D8DEE8" stroke-width="2"/>
  ${rows.map((r, idx) => {
    const heights = [53, 67, 67, 67, 67];
    const top = 390 + heights.slice(0, idx).reduce((a,b)=>a+b,0);
    const head = idx === 0;
    return `<line x1="845" y1="${top}" x2="1499" y2="${top}" stroke="#D8DEE8"/>
      <text x="871" y="${top + (head ? 34 : 43)}" class="font ${head ? "muted" : "ink"}" font-size="${head ? 22 : 28}" font-weight="${head || idx === 1 ? 700 : 400}">${r[0]}</text>
      <text x="1088" y="${top + (head ? 34 : 43)}" class="font ${head ? "muted" : "ink"}" font-size="${head ? 22 : 28}" font-weight="${head ? 700 : 400}">${r[1]}</text>`;
  }).join("")}
  <line x1="84" y1="786" x2="1518" y2="786" stroke="#D8DEE8" stroke-width="2"/>
  <text x="86" y="830" class="font muted" font-size="22">口頭補充：本報告不是做流行語字典，而是看流行語如何經歷「上升、成熟、衰退、死語化」。</text>
</svg>`;

fs.writeFileSync(svgPath, svg, "utf8");

(async () => {
  await pptx.writeFile({ fileName: pptxPath });
  await sharp(Buffer.from(svg)).png().toFile(pngPath);
  console.log(JSON.stringify({ pptxPath, svgPath, pngPath }, null, 2));
})();
