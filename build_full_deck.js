const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");

const outDir = path.join(__dirname, "簡報輸出");
const previewDir = path.join(outDir, "八頁預覽");
fs.mkdirSync(previewDir, { recursive: true });

const pptxPath = path.join(outDir, "流行語死語化_八頁簡報.pptx");
const montagePath = path.join(outDir, "流行語死語化_八頁簡報_總覽.png");

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";
pptx.author = "Codex";
pptx.subject = "流行語死語化研究架構";
pptx.title = "流行語死語化八頁簡報";
pptx.company = "核通";
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
  greenSoft: "E7F2F1",
  warmSoft: "F7EFEA",
};

function addText(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: "Microsoft JhengHei",
    color: opts.color || C.ink,
    fontSize: opts.fontSize || 18,
    bold: opts.bold || false,
    margin: opts.margin ?? 0,
    fit: "shrink",
    valign: opts.valign || "top",
    align: opts.align || "left",
    breakLine: false,
    paraSpaceAfterPt: opts.paraSpaceAfterPt || 0,
  });
}

function addLine(slide, x, y, w, color = C.line) {
  slide.addShape(pptx.ShapeType.line, {
    x, y, w, h: 0,
    line: { color, width: 1.05 },
  });
}

function addHeader(slide, page, title, subtitle = "") {
  slide.background = { color: C.bg };
  addText(slide, `第 ${page} 頁`, 0.7, 0.42, 1.0, 0.28, { fontSize: 12, bold: true, color: C.accent });
  addText(slide, title, 0.7, 0.78, 9.6, 0.58, { fontSize: 28, bold: true });
  if (subtitle) addText(slide, subtitle, 0.72, 1.5, 11.4, 0.36, { fontSize: 14, color: C.muted });
  addLine(slide, 0.7, 2.06, 11.95);
}

function sectionLabel(slide, text, x, y) {
  addText(slide, text, x, y, 2.8, 0.24, { fontSize: 12, bold: true, color: C.warm });
}

function addTable(slide, x, y, w, rows, colWidths, opts = {}) {
  const rowH = opts.rowH || 0.52;
  const headH = opts.headH || 0.42;
  const fontSize = opts.fontSize || 13.5;
  const headSize = opts.headSize || 11.5;
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
      addText(slide, cell, xx + 0.15, yy + (r === 0 ? 0.11 : 0.13), colWidths[c] - 0.24, rh - 0.12, {
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

function addBullets(slide, items, x, y, w, opts = {}) {
  const gap = opts.gap || 0.48;
  items.forEach((item, i) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x, y: y + i * gap + 0.07, w: 0.09, h: 0.09,
      fill: { color: opts.color || C.accent },
      line: { color: opts.color || C.accent },
    });
    addText(slide, item, x + 0.18, y + i * gap, w - 0.2, 0.28, {
      fontSize: opts.fontSize || 14,
      color: opts.textColor || C.ink,
    });
  });
}

function addFooter(slide, text) {
  addLine(slide, 0.7, 6.58, 11.95);
  addText(slide, `口頭補充：${text}`, 0.72, 6.78, 11.5, 0.25, { fontSize: 10.5, color: C.muted });
}

function slide1() {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addText(slide, "研究核心", 0.7, 0.48, 2.3, 0.34, { fontSize: 14, bold: true, color: C.accent });
  addText(slide, "流行語為什麼會紅？又為什麼會退燒、變成死語？", 0.7, 0.88, 8.65, 0.86, { fontSize: 30, bold: true });
  addText(slide, "流行語的生命週期不只和詞語本身有關，也和來源、語境、二創能力、社群規範與社會環境有關。", 0.72, 1.83, 10.9, 0.48, { fontSize: 16, color: C.muted });
  addLine(slide, 0.7, 2.48, 11.95);
  sectionLabel(slide, "主題", 0.72, 2.85);
  addText(slide, "本報告不是做流行語字典，而是看流行語如何經歷：", 0.72, 3.25, 5.0, 0.4, { fontSize: 17 });
  [["上升", "被看見"], ["成熟", "被共用"], ["衰退", "被說過氣"], ["死語化", "需要解釋"]].forEach((s, i) => {
    const x = 0.72 + i * 1.29;
    const green = i < 2;
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: 3.86, w: 1.06, h: 0.72, rectRadius: 0.08,
      fill: { color: green ? C.greenSoft : C.warmSoft },
      line: { color: green ? "BFD9D7" : "E5CFC2", width: 1 },
    });
    addText(slide, s[0], x + 0.08, 3.98, 0.9, 0.24, { fontSize: 15, bold: true, align: "center", color: green ? C.accent : C.warm });
    addText(slide, s[1], x + 0.08, 4.25, 0.9, 0.18, { fontSize: 8.5, align: "center", color: C.muted });
    if (i < 3) addText(slide, "→", x + 1.1, 4.08, 0.28, 0.24, { fontSize: 15, color: C.muted, align: "center" });
  });
  sectionLabel(slide, "核心觀點", 0.72, 5.02);
  addText(slide, "流行語的生命週期，是來源、語境與社群\n共同推動的變化過程。", 0.72, 5.37, 5.25, 0.74, { fontSize: 20, bold: true });
  sectionLabel(slide, "本報告模型", 7.05, 2.85);
  addTable(slide, 7.04, 3.25, 5.45, [
    ["關鍵因素", "作用"],
    ["來源", "決定爆紅起點"],
    ["語境依賴", "決定是否容易過時"],
    ["二創能力", "決定是否進入成熟期"],
    ["社群與環境", "決定是否被接續使用"],
  ], [1.8, 3.65], { boldFirstCol: true, rowH: 0.56, fontSize: 15 });
  addFooter(slide, "本報告不是做流行語字典，而是看流行語如何經歷「上升、成熟、衰退、死語化」。");
}

function slide2() {
  const slide = pptx.addSlide();
  addHeader(slide, 2, "來源分類：流行語從哪裡來，會影響它活多久", "事件型與情緒表達型，是本報告最重要的兩種來源。");
  sectionLabel(slide, "最重要的兩種來源", 0.72, 2.42);
  addTable(slide, 0.72, 2.82, 11.55, [
    ["類型", "特徵", "生命週期"],
    ["事件型", "來自新聞、影片名場面、網路事件、人物", "爆紅快，也退燒快"],
    ["情緒表達型", "表達尷尬、崩潰、嘲諷、無奈、驚訝", "較容易進入日常使用"],
    ["其他來源", "平台型、圈層型、語音／諧音型、二創型", "影響傳播速度與使用範圍"],
  ], [1.55, 5.45, 4.55], { rowH: 0.62, fontSize: 14.2, boldFirstCol: true });
  sectionLabel(slide, "關鍵判斷", 0.72, 5.42);
  addTable(slide, 0.72, 5.82, 6.25, [
    ["問題", "意義"],
    ["是否依賴特定事件？", "越依賴，越容易死語化"],
    ["是否能表達普遍情緒？", "越能表達日常情緒，越容易留下"],
  ], [2.65, 3.6], { rowH: 0.45, fontSize: 12.8, headH: 0.34 });
  addText(slide, "小結：流行語的來源會決定它的起點，語境依賴程度會影響它的壽命。", 7.35, 5.88, 4.8, 0.78, { fontSize: 17, bold: true });
  addFooter(slide, "事件型靠共同記憶活著，情緒型靠日常情緒續命。");
}

function slide3() {
  const slide = pptx.addSlide();
  addHeader(slide, 3, "四個輔助分析層次", "四個層次用來解釋流行語如何誕生、傳播、被理解與被日常使用。");
  addTable(slide, 0.72, 2.48, 11.88, [
    ["層次", "決定什麼", "例子方向"],
    ["構詞來源", "初始爆紅方式", "事件、平台、輸入習慣"],
    ["詞語結構", "傳播效率", "擬音、英數混合、短語模板"],
    ["語意指涉", "語境依賴程度", "事件、情緒、評價、圈層"],
    ["語用功能", "是否能日常使用", "吐槽、接梗、表態、共鳴"],
  ], [2.0, 3.4, 6.48], { rowH: 0.62, fontSize: 15, boldFirstCol: true });
  sectionLabel(slide, "小結", 0.78, 5.95);
  addBullets(slide, [
    "它怎麼紅？",
    "它為什麼好傳？",
    "它離開語境後還能不能懂？",
    "它能不能變成日常互動工具？",
  ], 0.82, 6.28, 7.4, { gap: 0.27, fontSize: 11.5 });
  addFooter(slide, "最重要的是語意指涉和語用功能，因為它們直接關係到流行語能不能續命。");
}

function slide4() {
  const slide = pptx.addSlide();
  addHeader(slide, 4, "生命週期模型", "本文將流行語的盛衰分成四個階段。");
  addTable(slide, 0.72, 2.45, 11.88, [
    ["階段", "特徵", "可觀察指標"],
    ["上升期", "被事件或語境帶起", "搜尋量、貼文數增加"],
    ["成熟期", "大量使用、改寫、二創", "變體、梗圖、短影音增加"],
    ["衰退期", "使用下降、被說過氣", "出現老梗、過氣討論"],
    ["死語期", "只剩回顧或解釋", "新使用者需要被解釋"],
  ], [2.0, 4.0, 5.88], { rowH: 0.62, fontSize: 15, boldFirstCol: true });
  sectionLabel(slide, "成熟期判準", 0.76, 5.82);
  addText(slide, "二創能力 = 是否成熟的重要指標", 0.76, 6.16, 4.6, 0.34, { fontSize: 18, bold: true, color: C.accent });
  addText(slide, "流行語不是被看見就成熟，而是要被群體拿來共同使用、改寫與模仿。", 5.25, 5.95, 6.9, 0.55, { fontSize: 18, bold: true });
  addFooter(slide, "成熟期的重點不是單純很紅，而是大家能不能把它拿去玩、拿去改、拿去套用。");
}

function slide5() {
  const slide = pptx.addSlide();
  addHeader(slide, 5, "線上社群理論輔證", "文獻：“User Lifecycle and Linguistic Change in Online Communities”。");
  addTable(slide, 0.72, 2.35, 11.88, [
    ["本文階段", "論文概念", "解釋"],
    ["上升期", "Linguistic innovation", "小群體先帶入新用法"],
    ["成熟期", "Community norm / Conforming", "成員跟著用，變成社群規範"],
    ["衰退期", "Linguistic adolescence 後的距離", "社群語言繼續變，舊詞停住"],
    ["死語期", "Conservative phase", "舊使用者保存，新使用者不自然採用"],
  ], [1.7, 4.35, 5.83], { rowH: 0.64, fontSize: 14.5, boldFirstCol: true });
  sectionLabel(slide, "關鍵句", 0.78, 5.9);
  addText(slide, "線上社群的語言會隨使用者加入、離開與互動而改變。", 0.78, 6.22, 10.8, 0.4, { fontSize: 20, bold: true });
  addFooter(slide, "這篇文獻支撐的是「流行語成熟不是個人使用，而是群體接受」。");
}

function slide6() {
  const slide = pptx.addSlide();
  addHeader(slide, 6, "社會與平台環境因素", "理論來源：William Labov, Principles of Linguistic Change: Social Factors。");
  addText(slide, "流行語的盛衰不只是詞語本身的問題，也受到使用者、平台與身份認同影響。", 0.75, 2.36, 11.4, 0.34, { fontSize: 17, bold: true });
  sectionLabel(slide, "影響因素", 0.75, 3.08);
  addTable(slide, 0.72, 3.45, 5.65, [
    ["因素", "影響"],
    ["使用者世代", "小圈層推動流行；新世代不接續就衰退"],
    ["平台環境", "短影音爆紅快、退燒快；論壇較容易累積文字梗"],
    ["身分認同", "懂梗代表圈內身份；過度大眾化會失去酷感"],
  ], [1.55, 4.1], { rowH: 0.66, fontSize: 12.8, boldFirstCol: true });
  sectionLabel(slide, "轉折解釋", 7.0, 3.08);
  addTable(slide, 6.98, 3.45, 5.25, [
    ["轉折", "可能原因"],
    ["上升到成熟", "使用者擴大、群體模仿"],
    ["成熟到衰退", "過度大眾化、平台風向轉移"],
    ["衰退到死語", "新世代不接續、語境消失"],
  ], [1.7, 3.55], { rowH: 0.66, fontSize: 13.2, boldFirstCol: true });
  addFooter(slide, "一個詞變過氣，有時不是詞不好，而是社群換了一套語言。");
}

function slide7() {
  const slide = pptx.addSlide();
  addHeader(slide, 7, "模型小結", "把來源、語境、二創與退燒放在同一個生命週期中理解。");
  sectionLabel(slide, "來源-語境-二創-退燒模型", 0.75, 2.42);
  addTable(slide, 0.72, 2.82, 5.65, [
    ["階段", "關鍵問題"],
    ["來源", "它從哪裡來？"],
    ["語境", "它是否依賴特定事件？"],
    ["二創", "它是否能被改寫、模仿、套用？"],
    ["退燒", "新使用者是否還自然使用？"],
  ], [1.45, 4.2], { rowH: 0.58, fontSize: 14, boldFirstCol: true });
  sectionLabel(slide, "本文主張", 7.0, 2.42);
  addText(slide, "流行語的生命週期可以看成社群規範變化的過程：", 7.0, 2.82, 5.1, 0.5, { fontSize: 15.5 });
  addTable(slide, 6.98, 3.55, 5.25, [
    ["狀態", "意義"],
    ["被群體吸收", "進入成熟"],
    ["不再被新成員自然使用", "開始死語化"],
  ], [2.55, 2.7], { rowH: 0.68, fontSize: 14, boldFirstCol: true });
  addText(slide, "社群往前走了，但某個詞停在原地，它就開始變老。", 1.0, 5.95, 11.0, 0.42, { fontSize: 21, bold: true, align: "center", color: C.accent });
  addFooter(slide, "這頁可以把分類、理論、環境因素全部接回生命週期。");
}

function slide8() {
  const slide = pptx.addSlide();
  addHeader(slide, 8, "結論與核心文獻", "用兩種來源差異，收束整個生命週期模型。");
  sectionLabel(slide, "結論", 0.75, 2.35);
  addTable(slide, 0.72, 2.75, 5.35, [
    ["類型", "結果"],
    ["事件型流行語", "依賴共同記憶，容易快速死亡"],
    ["情緒表達型流行語", "可進入日常情緒表達，較容易留下"],
  ], [2.1, 3.25], { rowH: 0.66, fontSize: 13.4, boldFirstCol: true });
  sectionLabel(slide, "核心觀點", 6.75, 2.35);
  addText(slide, "流行語的死亡，不只是因為大家不再使用它，而是支撐它的語境、情緒共鳴、社群規範與二創活動逐漸消失。", 6.75, 2.75, 5.45, 1.1, { fontSize: 16.5, bold: true });
  sectionLabel(slide, "核心文獻", 0.75, 4.62);
  addTable(slide, 0.72, 5.0, 11.88, [
    ["文獻", "本文參考用途"],
    ["《華語流行語的使用分析與教學建議》", "流行語來源、結構、語意、語用分類"],
    ["User Lifecycle and Linguistic Change in Online Communities", "線上社群語言規範與使用者生命週期"],
    ["Labov, Principles of Linguistic Change: Social Factors", "社會因素、世代、身份認同對語言變化的影響"],
  ], [5.2, 6.68], { rowH: 0.48, fontSize: 11.8, headH: 0.34, boldFirstCol: true });
  addFooter(slide, "本報告把流行語視為社群互動與社會環境共同推動的語言生命週期。");
}

slide1();
slide2();
slide3();
slide4();
slide5();
slide6();
slide7();
slide8();

function esc(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function svgBase(title, page, body) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900">
<rect width="1600" height="900" fill="#${C.bg}"/>
<style>.font{font-family:"Microsoft JhengHei","Noto Sans TC",Arial,sans-serif}.ink{fill:#${C.ink}}.muted{fill:#${C.muted}}.accent{fill:#${C.accent}}.warm{fill:#${C.warm}}</style>
<text x="84" y="82" class="font accent" font-size="24" font-weight="700">第 ${page} 頁</text>
<text x="84" y="152" class="font ink" font-size="44" font-weight="700">${esc(title)}</text>
<line x1="84" y1="247" x2="1518" y2="247" stroke="#${C.line}" stroke-width="2"/>
${body}
</svg>`;
}

function svgTable(x, y, widths, rows, opts = {}) {
  const rowH = opts.rowH || 72;
  const headH = opts.headH || 52;
  const font = opts.font || 25;
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
      out += `<text x="${xx + 24}" y="${yy + (r === 0 ? 34 : Math.min(44, rh - 18))}" class="font ${r === 0 ? "muted" : "ink"}" font-size="${r === 0 ? Math.max(18, font - 5) : font}" font-weight="${r === 0 || (opts.boldFirstCol && c === 0 && r > 0) ? 700 : 400}">${esc(cell)}</text>`;
      xx += widths[c];
    });
    yy += rh;
  });
  return out;
}

function writeSvg(name, svg) {
  const svgPath = path.join(previewDir, `${name}.svg`);
  const pngPath = path.join(previewDir, `${name}.png`);
  fs.writeFileSync(svgPath, svg, "utf8");
  return sharp(Buffer.from(svg)).png().toFile(pngPath).then(() => pngPath);
}

async function previews() {
  const paths = [];
  // For visual QA previews, use simplified vector renderings that match the PPT style and text.
  paths.push(await writeSvg("01_研究核心", fs.readFileSync(path.join(outDir, "流行語死語化_第一頁研究核心.svg"), "utf8")));
  paths.push(await writeSvg("02_來源分類", svgBase("來源分類：流行語從哪裡來，會影響它活多久", 2,
    `<text x="88" y="310" class="font warm" font-size="26" font-weight="700">最重要的兩種來源</text>
    ${svgTable(86, 345, [190, 610, 560], [["類型","特徵","生命週期"],["事件型","來自新聞、影片名場面、網路事件、人物","爆紅快，也退燒快"],["情緒表達型","表達尷尬、崩潰、嘲諷、無奈、驚訝","較容易進入日常使用"],["其他來源","平台型、圈層型、語音／諧音型、二創型","影響傳播速度與使用範圍"]], {font:24,boldFirstCol:true})}
    <text x="88" y="665" class="font warm" font-size="24" font-weight="700">關鍵判斷</text>
    ${svgTable(86, 695, [330, 430], [["問題","意義"],["是否依賴特定事件？","越依賴，越容易死語化"],["是否能表達普遍情緒？","越能表達日常情緒，越容易留下"]], {font:21,rowH:52,headH:40})}
    <text x="890" y="736" class="font ink" font-size="30" font-weight="700">小結：來源決定起點，</text><text x="890" y="775" class="font ink" font-size="30" font-weight="700">語境依賴影響壽命。</text>`)));
  paths.push(await writeSvg("03_四個輔助分析層次", svgBase("四個輔助分析層次", 3,
    `${svgTable(86, 300, [240, 400, 720], [["層次","決定什麼","例子方向"],["構詞來源","初始爆紅方式","事件、平台、輸入習慣"],["詞語結構","傳播效率","擬音、英數混合、短語模板"],["語意指涉","語境依賴程度","事件、情緒、評價、圈層"],["語用功能","是否能日常使用","吐槽、接梗、表態、共鳴"]], {font:27,boldFirstCol:true})}
    <text x="92" y="742" class="font warm" font-size="24" font-weight="700">小結</text>
    <text x="92" y="790" class="font ink" font-size="24">它怎麼紅？為什麼好傳？離開語境後還能不能懂？能不能變成日常互動工具？</text>`)));
  paths.push(await writeSvg("04_生命週期模型", svgBase("生命週期模型", 4,
    `${svgTable(86, 295, [240, 460, 660], [["階段","特徵","可觀察指標"],["上升期","被事件或語境帶起","搜尋量、貼文數增加"],["成熟期","大量使用、改寫、二創","變體、梗圖、短影音增加"],["衰退期","使用下降、被說過氣","出現老梗、過氣討論"],["死語期","只剩回顧或解釋","新使用者需要被解釋"]], {font:27,boldFirstCol:true})}
    <text x="92" y="710" class="font warm" font-size="24" font-weight="700">成熟期判準</text>
    <text x="92" y="762" class="font accent" font-size="31" font-weight="700">二創能力 = 是否成熟的重要指標</text>
    <text x="640" y="762" class="font ink" font-size="29" font-weight="700">不是被看見就成熟，而是被共同使用、改寫與模仿。</text>`)));
  paths.push(await writeSvg("05_線上社群理論輔證", svgBase("線上社群理論輔證", 5,
    `<text x="88" y="305" class="font muted" font-size="27">文獻：“User Lifecycle and Linguistic Change in Online Communities”</text>
    ${svgTable(86, 350, [190, 500, 670], [["本文階段","論文概念","解釋"],["上升期","Linguistic innovation","小群體先帶入新用法"],["成熟期","Community norm / Conforming","成員跟著用，變成社群規範"],["衰退期","Linguistic adolescence 後的距離","社群語言繼續變，舊詞停住"],["死語期","Conservative phase","舊使用者保存，新使用者不自然採用"]], {font:24,boldFirstCol:true})}
    <text x="92" y="765" class="font warm" font-size="24" font-weight="700">關鍵句</text>
    <text x="92" y="815" class="font ink" font-size="32" font-weight="700">線上社群的語言會隨使用者加入、離開與互動而改變。</text>`)));
  paths.push(await writeSvg("06_社會與平台環境因素", svgBase("社會與平台環境因素", 6,
    `<text x="88" y="300" class="font ink" font-size="28" font-weight="700">流行語的盛衰不只是詞語本身的問題，也受到使用者、平台與身份認同影響。</text>
    <text x="88" y="375" class="font warm" font-size="24" font-weight="700">影響因素</text>
    ${svgTable(86, 410, [190, 500], [["因素","影響"],["使用者世代","小圈層推動流行；新世代不接續就衰退"],["平台環境","短影音爆紅快、退燒快；論壇較容易累積文字梗"],["身分認同","懂梗代表圈內身份；過度大眾化會失去酷感"]], {font:20,rowH:78,boldFirstCol:true})}
    <text x="850" y="375" class="font warm" font-size="24" font-weight="700">轉折解釋</text>
    ${svgTable(840, 410, [220, 410], [["轉折","可能原因"],["上升到成熟","使用者擴大、群體模仿"],["成熟到衰退","過度大眾化、平台風向轉移"],["衰退到死語","新世代不接續、語境消失"]], {font:22,rowH:78,boldFirstCol:true})}`)));
  paths.push(await writeSvg("07_模型小結", svgBase("模型小結", 7,
    `<text x="88" y="310" class="font warm" font-size="24" font-weight="700">來源-語境-二創-退燒模型</text>
    ${svgTable(86, 350, [180, 500], [["階段","關鍵問題"],["來源","它從哪裡來？"],["語境","它是否依賴特定事件？"],["二創","它是否能被改寫、模仿、套用？"],["退燒","新使用者是否還自然使用？"]], {font:25,boldFirstCol:true,rowH:64})}
    <text x="850" y="310" class="font warm" font-size="24" font-weight="700">本文主張</text>
    <text x="850" y="365" class="font ink" font-size="26">流行語的生命週期可以看成社群規範變化的過程：</text>
    ${svgTable(840, 440, [310, 320], [["狀態","意義"],["被群體吸收","進入成熟"],["不再被新成員自然使用","開始死語化"]], {font:23,boldFirstCol:true,rowH:78})}
    <text x="800" y="765" text-anchor="middle" class="font accent" font-size="38" font-weight="700">社群往前走了，但某個詞停在原地，它就開始變老。</text>`)));
  paths.push(await writeSvg("08_結論與核心文獻", svgBase("結論與核心文獻", 8,
    `<text x="88" y="305" class="font warm" font-size="24" font-weight="700">結論</text>
    ${svgTable(86, 340, [260, 420], [["類型","結果"],["事件型流行語","依賴共同記憶，容易快速死亡"],["情緒表達型流行語","可進入日常情緒表達，較容易留下"]], {font:21,rowH:72,boldFirstCol:true})}
    <text x="825" y="305" class="font warm" font-size="24" font-weight="700">核心觀點</text>
    <text x="825" y="360" class="font ink" font-size="28" font-weight="700">流行語的死亡，不只是因為大家不再使用它，</text>
    <text x="825" y="400" class="font ink" font-size="28" font-weight="700">而是支撐它的語境、情緒共鳴、社群規範</text>
    <text x="825" y="440" class="font ink" font-size="28" font-weight="700">與二創活動逐漸消失。</text>
    <text x="88" y="560" class="font warm" font-size="24" font-weight="700">核心文獻</text>
    ${svgTable(86, 595, [590, 770], [["文獻","本文參考用途"],["《華語流行語的使用分析與教學建議》","流行語來源、結構、語意、語用分類"],["User Lifecycle and Linguistic Change in Online Communities","線上社群語言規範與使用者生命週期"],["Labov, Principles of Linguistic Change: Social Factors","社會因素、世代、身份認同對語言變化的影響"]], {font:19,rowH:56,headH:40,boldFirstCol:true})}`)));

  const thumbs = await Promise.all(paths.map((p) => sharp(p).resize(400, 225).toBuffer()));
  const canvas = sharp({
    create: {
      width: 1600,
      height: 450,
      channels: 4,
      background: "#F7F8FA",
    },
  });
  await canvas.composite(thumbs.map((input, i) => ({
    input,
    left: (i % 4) * 400,
    top: Math.floor(i / 4) * 225,
  }))).png().toFile(montagePath);
  return paths;
}

(async () => {
  await pptx.writeFile({ fileName: pptxPath });
  const previewPaths = await previews();
  console.log(JSON.stringify({ pptxPath, montagePath, previewPaths }, null, 2));
})();
