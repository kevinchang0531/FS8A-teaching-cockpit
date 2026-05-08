const pptxgen = require("pptxgenjs");

const BG = "1C3A1C";       // dark blackboard green
const WHITE = "FFFFFF";
const CHALK_YELLOW = "F5E642";
const CHALK_BLUE = "A8D8EA";
const CHALK_ORANGE = "FFB347";
const DIM = "C8C8C8";      // dimmed white for secondary text
const LINE_COLOR = "FFFFFF";

const TITLE_FONT = "Impact";
const BODY_FONT = "Calibri";

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "FS8A 福衛八號飛行操作訓練";

// ─────────────────────────────────────────
// Helper: chalk underline (white line under text)
function chalkLine(slide, x, y, w) {
  slide.addShape(pres.shapes.LINE, {
    x, y, w, h: 0,
    line: { color: LINE_COLOR, width: 1.5, dashType: "dash" }
  });
}

// Helper: section header bar
function sectionBar(slide, text, y) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y, w: 10, h: 0.55,
    fill: { color: "2E5E2E" }, line: { color: "2E5E2E" }
  });
  slide.addText(text, {
    x: 0.3, y: y + 0.04, w: 9.4, h: 0.47,
    fontSize: 18, fontFace: TITLE_FONT, color: CHALK_YELLOW,
    bold: true, valign: "middle", margin: 0
  });
}

// Helper: bullet list
function addBullets(slide, items, x, y, w, h, opts = {}) {
  const textArr = items.map((item, i) => ({
    text: item.text || item,
    options: {
      bullet: true,
      breakLine: i < items.length - 1,
      color: item.color || opts.color || WHITE,
      fontSize: item.size || opts.size || 14,
      bold: item.bold || false,
      fontFace: BODY_FONT,
    }
  }));
  slide.addText(textArr, { x, y, w, h, valign: "top", margin: 0.1 });
}

// ─────────────────────────────────────────
// SLIDE 1 – Cover
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };

  // Chalk frame border
  slide_border(s);

  // Big title
  s.addText("FS8A 福衛八號", {
    x: 0.5, y: 0.8, w: 9, h: 1.2,
    fontSize: 54, fontFace: TITLE_FONT, color: WHITE, bold: true, align: "center"
  });
  s.addText("飛行操作訓練", {
    x: 0.5, y: 1.9, w: 9, h: 1.0,
    fontSize: 44, fontFace: TITLE_FONT, color: CHALK_YELLOW, bold: true, align: "center"
  });

  // Chalk divider
  s.addShape(pres.shapes.LINE, {
    x: 1.5, y: 3.05, w: 7, h: 0,
    line: { color: WHITE, width: 2, dashType: "lgDash" }
  });

  // Subtitle
  s.addText("計畫管理 及 衛星電機系統", {
    x: 0.5, y: 3.2, w: 9, h: 0.55,
    fontSize: 22, fontFace: BODY_FONT, color: DIM, align: "center"
  });
  s.addText("系工組 | 林信嘉　2025/02/07", {
    x: 0.5, y: 3.85, w: 9, h: 0.45,
    fontSize: 16, fontFace: BODY_FONT, color: DIM, align: "center"
  });

  // Small chalk drawing: satellite icon using shapes
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.4, y: 4.55, w: 1.2, h: 0.6,
    fill: { color: "3A6A3A" }, line: { color: WHITE, width: 1.5 }
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.2, y: 4.72, w: 1.1, h: 0.25, fill: { color: "3A6A3A" }, line: { color: WHITE, width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.7, y: 4.72, w: 1.1, h: 0.25, fill: { color: "3A6A3A" }, line: { color: WHITE, width: 1 } });
  s.addShape(pres.shapes.LINE, { x: 4.95, y: 4.1, w: 0, h: 0.45, line: { color: WHITE, width: 1 } });
  s.addShape(pres.shapes.OVAL, { x: 4.75, y: 4.0, w: 0.4, h: 0.25, fill: { color: "3A6A3A" }, line: { color: WHITE, width: 1 } });
}

// ─────────────────────────────────────────
// SLIDE 2 – Outline
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "課程大綱  Course Outline", 0);

  const topics = [
    { num: "01", title: "計畫實施", sub: "大事紀 · 先導計畫 → 福衛八號", color: CHALK_YELLOW },
    { num: "02", title: "系統需求", sub: "軌道 · 解析度 · 質量 · SNR規格", color: CHALK_BLUE },
    { num: "03", title: "衛星構型", sub: "Bus系統 · RSI酬載 · 國內自製", color: CHALK_ORANGE },
    { num: "04", title: "衛星電機介面", sub: "電氣架構 · CAN Bus · TT&C · X-band · RCS", color: "98FB98" },
  ];

  topics.forEach((t, i) => {
    const x = (i % 2) * 4.8 + 0.4;
    const y = Math.floor(i / 2) * 2.1 + 0.85;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 1.8,
      fill: { color: "254525" }, line: { color: t.color, width: 2 }
    });
    s.addText(t.num, {
      x: x + 0.15, y: y + 0.1, w: 0.7, h: 0.6,
      fontSize: 28, fontFace: TITLE_FONT, color: t.color, bold: true, margin: 0
    });
    s.addText(t.title, {
      x: x + 0.9, y: y + 0.1, w: 3.2, h: 0.55,
      fontSize: 20, fontFace: TITLE_FONT, color: WHITE, bold: true, margin: 0
    });
    s.addShape(pres.shapes.LINE, {
      x: x + 0.15, y: y + 0.8, w: 3.9, h: 0,
      line: { color: t.color, width: 1, dashType: "dash" }
    });
    s.addText(t.sub, {
      x: x + 0.15, y: y + 0.9, w: 3.95, h: 0.8,
      fontSize: 12, fontFace: BODY_FONT, color: DIM, valign: "top", margin: 0
    });
  });
}

// ─────────────────────────────────────────
// SLIDE 3 – 計畫大事紀
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "計畫大事紀  Project Milestones", 0);

  // Phase 1 box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 0.75, w: 4.3, h: 2.15,
    fill: { color: "254525" }, line: { color: CHALK_YELLOW, width: 2 }
  });
  s.addText("Phase 1：先導型衛星任務", {
    x: 0.45, y: 0.8, w: 4.0, h: 0.45,
    fontSize: 15, fontFace: TITLE_FONT, color: CHALK_YELLOW, bold: true, margin: 0
  });
  s.addText("106~109年（MDR / SDR / PDR）", {
    x: 0.45, y: 1.25, w: 4.0, h: 0.3,
    fontSize: 12, fontFace: BODY_FONT, color: DIM, margin: 0
  });
  addBullets(s, [
    { text: "250kg 光學遙測衛星", size: 12 },
    { text: "561km SSO 軌道，6顆星座", size: 12 },
    { text: "目標解析度 PAN 0.7m (super-res)", size: 12 },
  ], 0.45, 1.55, 4.0, 1.3, { color: WHITE, size: 12 });

  // Phase 2 box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 0.75, w: 4.3, h: 2.15,
    fill: { color: "254525" }, line: { color: CHALK_BLUE, width: 2 }
  });
  s.addText("Phase 2：福衛八號計畫", {
    x: 5.45, y: 0.8, w: 4.0, h: 0.45,
    fontSize: 15, fontFace: TITLE_FONT, color: CHALK_BLUE, bold: true, margin: 0
  });
  s.addText("108~117年（先導型高解析度衛星星系）", {
    x: 5.45, y: 1.25, w: 4.0, h: 0.3,
    fontSize: 11, fontFace: BODY_FONT, color: DIM, margin: 0
  });
  addBullets(s, [
    { text: "400kg 光學遙測衛星", size: 12 },
    { text: "8A~8F 星座佈建", size: 12 },
    { text: "1m黑白 / 2m彩色解析度", size: 12 },
  ], 5.45, 1.55, 4.0, 1.3, { color: WHITE, size: 12 });

  // Timeline
  sectionBar(s, "FS-8A 重要里程碑", 3.05);

  const milestones = [
    { year: "112/5", label: "Delta ITR", color: CHALK_YELLOW },
    { year: "114/4", label: "ITR", color: CHALK_ORANGE },
    { year: "114/7", label: "PSR", color: CHALK_BLUE },
    { year: "114/10", label: "🚀 發射\nFalcon-9", color: "98FB98" },
  ];
  s.addShape(pres.shapes.LINE, {
    x: 0.5, y: 4.4, w: 9, h: 0,
    line: { color: WHITE, width: 2 }
  });
  milestones.forEach((m, i) => {
    const x = 0.5 + i * 2.3;
    s.addShape(pres.shapes.OVAL, {
      x: x - 0.12, y: 4.28, w: 0.24, h: 0.24,
      fill: { color: m.color }, line: { color: m.color }
    });
    s.addText(m.year, {
      x: x - 0.5, y: 3.65, w: 1.3, h: 0.35,
      fontSize: 11, fontFace: BODY_FONT, color: m.color, bold: true, align: "center", margin: 0
    });
    s.addText(m.label, {
      x: x - 0.5, y: 4.55, w: 1.5, h: 0.6,
      fontSize: 11, fontFace: BODY_FONT, color: WHITE, align: "center", margin: 0
    });
  });
}

// ─────────────────────────────────────────
// SLIDE 4 – 衛星系統規格
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "衛星系統規格  Satellite System Specifications", 0);

  const specs = [
    { param: "任務軌道", value: "561 km 太陽同步軌道 (SSO)", color: CHALK_YELLOW },
    { param: "任務壽命", value: "3 年（目標 5 年）", color: WHITE },
    { param: "過赤道時間 (LTDN)", value: "FS-8A：上午 10:00~11:00", color: WHITE },
    { param: "衛星再訪週期", value: "1 天", color: WHITE },
    { param: "GSD", value: "1m (PAN) ／ 2m (MS)", color: CHALK_BLUE },
    { param: "Swath", value: "≥ 10 km　[PAN: 12288 px, MS: 6144 px]", color: WHITE },
    { param: "Bands", value: "2 PAN + 6 MS", color: WHITE },
    { param: "SNR", value: "≥ 80 (PAN, TDI 8)　／　≥ 90 (MS, TDI 16)", color: CHALK_ORANGE },
    { param: "System CTF (PAN)", value: "≥ 0.07", color: WHITE },
    { param: "質量", value: "≤ 410 kg", color: WHITE },
  ];

  specs.forEach((sp, i) => {
    const y = 0.75 + i * 0.47;
    if (i % 2 === 1) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.2, y: y - 0.04, w: 9.6, h: 0.44,
        fill: { color: "254525" }, line: { color: "254525" }
      });
    }
    s.addText(sp.param, {
      x: 0.3, y, w: 3.0, h: 0.38,
      fontSize: 13, fontFace: BODY_FONT, color: DIM, bold: true, valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.LINE, {
      x: 3.4, y: y + 0.19, w: 0, h: 0.05,
      line: { color: WHITE, width: 1 }
    });
    s.addText(sp.value, {
      x: 3.5, y, w: 6.2, h: 0.38,
      fontSize: 13, fontFace: BODY_FONT, color: sp.color, valign: "middle", margin: 0
    });
  });
}

// ─────────────────────────────────────────
// SLIDE 5 – 衛星構型
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "衛星構型  Satellite Configuration", 0);

  // Three columns
  const cols = [
    {
      title: "Bus 衛星匯流排", color: CHALK_YELLOW, x: 0.25,
      items: ["EPS — 電力系統\n（電池、太陽能板、PCU）",
              "C&DH — 星上電腦\n（OBC、IOC）",
              "TT&C — 通訊\n（S-band 收發器 A/B）",
              "AOCS — 姿態控制\n（Star Tracker、GPSR、陀螺、磁力計、RW）",
              "RCS — 推進系統\n（H₂O₂推力器 x4）"]
    },
    {
      title: "RSI 遙測酬載", color: CHALK_BLUE, x: 3.55,
      items: ["OSA — 光學鏡組\n（望遠鏡、M1/M2）",
              "FPA — 焦面組件\n（線列感測器、CMOS）",
              "SSR — 固態記錄器\n（SSR-A / SSR-B）",
              "EU — 電子單元\n（EU-A / EU-B）"]
    },
    {
      title: "國內自製元件", color: CHALK_ORANGE, x: 6.85,
      items: ["TASA OBC / PCU",
              "GPSR-B（衛星導航）",
              "X-band 發射器",
              "S-band Helix / Patch 天線",
              "RCS 推進模組",
              "RSI 光機 (OSA / FPA)"]
    }
  ];

  cols.forEach(col => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: col.x, y: 0.7, w: 3.1, h: 4.65,
      fill: { color: "254525" }, line: { color: col.color, width: 2 }
    });
    s.addText(col.title, {
      x: col.x + 0.1, y: 0.75, w: 2.9, h: 0.45,
      fontSize: 14, fontFace: TITLE_FONT, color: col.color, bold: true, margin: 0, align: "center"
    });
    s.addShape(pres.shapes.LINE, {
      x: col.x + 0.1, y: 1.22, w: 2.9, h: 0,
      line: { color: col.color, width: 1 }
    });
    col.items.forEach((item, j) => {
      s.addText(item, {
        x: col.x + 0.15, y: 1.32 + j * 0.75, w: 2.8, h: 0.7,
        fontSize: 11, fontFace: BODY_FONT, color: WHITE, valign: "top", margin: 0
      });
      if (j < col.items.length - 1) {
        s.addShape(pres.shapes.LINE, {
          x: col.x + 0.15, y: 1.32 + (j + 1) * 0.75, w: 2.8, h: 0,
          line: { color: "3A6A3A", width: 0.5 }
        });
      }
    });
  });
}

// ─────────────────────────────────────────
// SLIDE 6 – 電氣架構
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "電氣架構  Electrical Architecture", 0);

  // Power flow (left column)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.25, y: 0.72, w: 3.0, h: 4.6,
    fill: { color: "254525" }, line: { color: CHALK_YELLOW, width: 2 }
  });
  s.addText("電力流程", {
    x: 0.35, y: 0.77, w: 2.8, h: 0.38,
    fontSize: 14, fontFace: TITLE_FONT, color: CHALK_YELLOW, bold: true, margin: 0, align: "center"
  });
  const pwrFlow = ["☀ 太陽能板 (Solar Array)", "↓", "⚡ BCR (Battery Charge Reg.)", "↓", "🔋 電池 (Battery)", "↓", "🔌 PDM (Power Dist. Module)", "↓", "各分系統供電"];
  pwrFlow.forEach((item, i) => {
    s.addText(item, {
      x: 0.35, y: 1.22 + i * 0.44, w: 2.8, h: 0.4,
      fontSize: item === "↓" ? 16 : 12, fontFace: BODY_FONT,
      color: item === "↓" ? CHALK_YELLOW : WHITE,
      align: "center", margin: 0
    });
  });

  // CAN Bus (right)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.55, y: 0.72, w: 6.2, h: 2.15,
    fill: { color: "254525" }, line: { color: CHALK_BLUE, width: 2 }
  });
  s.addText("CAN Bus 雙備援架構", {
    x: 3.65, y: 0.77, w: 6.0, h: 0.38,
    fontSize: 14, fontFace: TITLE_FONT, color: CHALK_BLUE, bold: true, margin: 0
  });
  addBullets(s, [
    { text: "CAN Bus-A（主用）/ CAN Bus-B（備援）— 500 Kbps" },
    { text: "OBC PM-A / PM-B 備援；預設 A-side，可指令切換" },
    { text: "連接節點：PCU PDM-A/B、PCU BCR-A/B、GPSR-B、MEMS IRU、XTx-B" },
  ], 3.65, 1.22, 6.0, 1.6, { color: WHITE, size: 12 });

  // FDIR
  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.55, y: 3.05, w: 6.2, h: 2.27,
    fill: { color: "254525" }, line: { color: CHALK_ORANGE, width: 2 }
  });
  s.addText("FDIR 故障偵測 / 隔離 / 復原", {
    x: 3.65, y: 3.1, w: 6.0, h: 0.38,
    fontSize: 14, fontFace: TITLE_FONT, color: CHALK_ORANGE, bold: true, margin: 0
  });
  addBullets(s, [
    { text: "PM-A 故障 → TMTC-B 通知切換至 PM-B" },
    { text: "WDT 觸發 ARO 前須先 disable WDT" },
    { text: "PCU PC-A/B 可透過 TMTC 送 PC 指令" },
    { text: "Star Tracker 雙機：Port-0/1 分散至 UART13/14" },
  ], 3.65, 3.55, 6.0, 1.7, { color: WHITE, size: 12 });
}

// ─────────────────────────────────────────
// SLIDE 7 – TT&C
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "S-band TT&C 遙控遙測介面", 0);

  // Left: Telecommand
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.25, y: 0.72, w: 4.6, h: 4.6,
    fill: { color: "254525" }, line: { color: CHALK_YELLOW, width: 2 }
  });
  s.addText("📡 遙控 (TC — Telecommand)", {
    x: 0.35, y: 0.77, w: 4.4, h: 0.45,
    fontSize: 15, fontFace: TITLE_FONT, color: CHALK_YELLOW, bold: true, margin: 0
  });
  chalkLine(s, 0.35, 1.28, 4.4);
  addBullets(s, [
    { text: "收發器：SRx-A（收發器-A）/ SRx-B（收發器-B）", size: 12 },
    { text: "OBC TMTC-A / TMTC-B 雙路接收", size: 12 },
    { text: "格式：NRZ-L，RS-422", size: 12 },
    { text: "速率：32 Kbps / 250 Kbps", size: 12 },
    { text: "Carrier Lock、TC Valid 狀態回報", size: 12 },
  ], 0.35, 1.35, 4.4, 3.85, { color: WHITE, size: 12 });

  // Right: Telemetry
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.15, y: 0.72, w: 4.6, h: 4.6,
    fill: { color: "254525" }, line: { color: CHALK_BLUE, width: 2 }
  });
  s.addText("📶 遙測 (TM — Telemetry)", {
    x: 5.25, y: 0.77, w: 4.4, h: 0.45,
    fontSize: 15, fontFace: TITLE_FONT, color: CHALK_BLUE, bold: true, margin: 0
  });
  chalkLine(s, 5.25, 1.28, 4.4);
  addBullets(s, [
    { text: "發射器：STx-A / STx-B 雙路輸出", size: 12 },
    { text: "格式：NRZ-L，RS-422", size: 12 },
    { text: "速率：2 Mbps / 250 Kbps / 50 Kbps", size: 12 },
    { text: "TM Data A/B + TM CLK A/B", size: 12 },
    { text: "由 PM/FSW 控制 Enable/Disable", size: 12 },
  ], 5.25, 1.35, 4.4, 3.85, { color: WHITE, size: 12 });

  // Bottom: Antenna summary
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.25, y: 5.1, w: 9.5, h: 0.42, // adjusted since 5.625 total height
    fill: { color: "2E5E2E" }, line: { color: "2E5E2E" }
  });
  s.addText("天線配置：S-band Patch x4（2A/2B/1A/1B）+ Helix x1 (1A)　—　Diplexer A/B + Power Divider", {
    x: 0.35, y: 5.13, w: 9.3, h: 0.35,
    fontSize: 11, fontFace: BODY_FONT, color: DIM, margin: 0, align: "center"
  });
}

// ─────────────────────────────────────────
// SLIDE 8 – X-band
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "X-band 下傳系統  XDS Architecture", 0);

  // Key specs
  const xspecs = [
    { label: "傳輸速率", value: "600 Mbps", color: CHALK_YELLOW },
    { label: "調變方式", value: "8PSK 4D-TCM", color: CHALK_BLUE },
    { label: "頻率範圍", value: "8200 ± 162 MHz", color: CHALK_ORANGE },
    { label: "發射器", value: "XTx-A (Honeywell) + XTx-B (TASA)", color: "98FB98" },
  ];
  xspecs.forEach((sp, i) => {
    const x = (i % 2) * 4.8 + 0.25;
    const y = Math.floor(i / 2) * 1.2 + 0.75;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.4, h: 1.05,
      fill: { color: "254525" }, line: { color: sp.color, width: 2 }
    });
    s.addText(sp.label, {
      x: x + 0.15, y: y + 0.08, w: 2.0, h: 0.4,
      fontSize: 12, fontFace: BODY_FONT, color: DIM, bold: true, margin: 0
    });
    s.addText(sp.value, {
      x: x + 0.15, y: y + 0.5, w: 4.0, h: 0.45,
      fontSize: 15, fontFace: TITLE_FONT, color: sp.color, bold: true, margin: 0
    });
  });

  // Data flow diagram (text-based)
  sectionBar(s, "資料流程", 3.22);

  const flow = ["RSI SSR-A", "→", "XTx-A\n(Honeywell)", "→", "BPF\n(帶通濾波)", "→", "X-band\n天線 A/B", "→", "地面站"];
  let fx = 0.15;
  flow.forEach((f, i) => {
    const isArrow = f === "→";
    s.addText(f, {
      x: fx, y: 3.75, w: isArrow ? 0.4 : 1.1, h: 0.95,
      fontSize: isArrow ? 20 : 11, fontFace: isArrow ? "Arial" : BODY_FONT,
      color: isArrow ? CHALK_YELLOW : WHITE, align: "center", valign: "middle", margin: 0,
      ...(isArrow ? {} : { fill: { color: "254525" } })
    });
    if (!isArrow) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: fx, y: 3.77, w: 1.1, h: 0.9,
        fill: { color: "254525" }, line: { color: CHALK_BLUE, width: 1 }
      });
      s.addText(f, {
        x: fx + 0.05, y: 3.8, w: 1.0, h: 0.84,
        fontSize: 10, fontFace: BODY_FONT, color: WHITE, align: "center", valign: "middle", margin: 0
      });
    }
    fx += isArrow ? 0.4 : 1.1;
  });
  s.addText("SSR-B → XTx-B (TASA) 路徑為備援", {
    x: 0.3, y: 4.85, w: 9.4, h: 0.4,
    fontSize: 12, fontFace: BODY_FONT, color: DIM, align: "center", margin: 0
  });
}

// ─────────────────────────────────────────
// SLIDE 9 – RCS
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "RCS 推進系統  Reaction Control System", 0);

  // 2x3 grid of component boxes
  const comps = [
    { title: "推進劑", icon: "⚗", detail: "H₂O₂（過氧化氫）\n壓力槽 x2\n（Tank-1 / Tank-2）", color: CHALK_YELLOW },
    { title: "推力器", icon: "🔥", detail: "Thruster x4\nNX / NY / PX / PY\n（±X / ±Y 方向）", color: CHALK_ORANGE },
    { title: "電磁閥", icon: "🔧", detail: "Ball Latching Valve x2\nFill & Drain Valve\nFill & Vent Valve", color: CHALK_BLUE },
    { title: "加熱器", icon: "🌡", detail: "Cat Bed Heater x4\n（燃燒床預熱）", color: "98FB98" },
    { title: "壓力監測", icon: "📊", detail: "Pressure Transducer x2\n（即時壓力監測）", color: "FFD700" },
    { title: "注意事項", icon: "⚠", detail: "NCR-0203：1路\nHeater短路燒毀\n→ 切斷該路電源", color: "FF6B6B" },
  ];

  comps.forEach((c, i) => {
    const x = (i % 3) * 3.25 + 0.2;
    const y = Math.floor(i / 3) * 2.15 + 0.72;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 3.0, h: 2.0,
      fill: { color: "254525" }, line: { color: c.color, width: 2 }
    });
    s.addText(c.icon + " " + c.title, {
      x: x + 0.1, y: y + 0.08, w: 2.8, h: 0.42,
      fontSize: 14, fontFace: TITLE_FONT, color: c.color, bold: true, margin: 0
    });
    s.addShape(pres.shapes.LINE, {
      x: x + 0.1, y: y + 0.55, w: 2.8, h: 0,
      line: { color: c.color, width: 1, dashType: "dash" }
    });
    s.addText(c.detail, {
      x: x + 0.1, y: y + 0.63, w: 2.8, h: 1.3,
      fontSize: 11, fontFace: BODY_FONT, color: WHITE, valign: "top", margin: 0
    });
  });
}

// ─────────────────────────────────────────
// SLIDE 10 – 總結
// ─────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slide_border(s);
  sectionBar(s, "課程總結  Summary", 0);

  // Achievement box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 0.72, w: 9.4, h: 1.5,
    fill: { color: "254525" }, line: { color: "98FB98", width: 2 }
  });
  s.addText("🚀 FS8A 已於 2025年10月 成功以 SpaceX Falcon-9 發射升空", {
    x: 0.45, y: 0.82, w: 9.1, h: 0.5,
    fontSize: 18, fontFace: TITLE_FONT, color: "98FB98", bold: true, align: "center", margin: 0
  });
  s.addText("561 km 太陽同步軌道正常運行　·　各分系統電機介面健全　·　遙測遙控功能正常", {
    x: 0.45, y: 1.35, w: 9.1, h: 0.75,
    fontSize: 13, fontFace: BODY_FONT, color: DIM, align: "center", margin: 0
  });

  // Key takeaways
  const takes = [
    { icon: "📋", title: "計畫歷程", text: "從先導型(250kg)到福衛八號(400kg)，歷時8年完成首發" },
    { icon: "🛰", title: "系統架構", text: "雙備援(CAN Bus / OBC / TT&C)確保任務可靠性" },
    { icon: "📡", title: "通訊介面", text: "S-band TT&C + X-band 600Mbps 下傳，全面覆蓋需求" },
    { icon: "🔬", title: "自製能量", text: "OBC、PCU、GPSR-B、X-band TX 等核心元件國內自製" },
  ];
  takes.forEach((t, i) => {
    const x = (i % 2) * 4.8 + 0.3;
    const y = Math.floor(i / 2) * 1.5 + 2.45;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.4, h: 1.35,
      fill: { color: "254525" }, line: { color: CHALK_YELLOW, width: 1.5 }
    });
    s.addText(t.icon + " " + t.title, {
      x: x + 0.12, y: y + 0.08, w: 4.2, h: 0.4,
      fontSize: 14, fontFace: TITLE_FONT, color: CHALK_YELLOW, bold: true, margin: 0
    });
    s.addText(t.text, {
      x: x + 0.12, y: y + 0.52, w: 4.2, h: 0.78,
      fontSize: 12, fontFace: BODY_FONT, color: WHITE, valign: "top", margin: 0
    });
  });

  // Footer
  s.addText("下一步 → FS-8B　預計 2026年12月發射（Space-X Falcon-9）", {
    x: 0.3, y: 5.25, w: 9.4, h: 0.32,
    fontSize: 12, fontFace: BODY_FONT, color: CHALK_BLUE, align: "center", margin: 0
  });
}

// ─────────────────────────────────────────
// Border helper (defined before use via hoisting is not reliable in node — move up)
function slide_border(slide) {
  const c = "3A6A3A";
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.08, y: 0.08, w: 9.84, h: 5.46,
    fill: { type: "none" }, line: { color: c, width: 2, dashType: "dash" }
  });
}

// ─────────────────────────────────────────
const OUT = "C:/Users/kevin/Claude Cowork/class/FS8A_教學簡報_黑板粉筆風格.pptx";
pres.writeFile({ fileName: OUT })
  .then(() => console.log("✅ Saved: " + OUT))
  .catch(e => { console.error("❌ Error:", e); process.exit(1); });
