const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "XDS & TTC Operations FS-8A";

// ── 設計系統 ──────────────────────────────────────────────────
const BG      = "2D4A2D";
const BG2     = "1E3A1E";
const WHITE   = "FFFFFF";
const YELLOW  = "F5E642";
const BLUE    = "A8D8EA";
const DIM     = "C8C8C8";
const ORANGE  = "FFB347";
const GREEN   = "98FB98";
const RED     = "FF7070";

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.25 });

function addSlide() {
  const s = pres.addSlide();
  s.background = { color: BG };
  return s;
}

function addTitle(s, text, y = 0.32) {
  s.addText(text, {
    x: 0.55, y, w: 8.9, h: 0.62,
    fontFace: "Impact", fontSize: 32, color: YELLOW,
    charSpacing: 1, margin: 0
  });
  s.addShape(pres.shapes.LINE, {
    x: 0.55, y: y + 0.65, w: 8.9, h: 0,
    line: { color: "3A6A3A", width: 1.5, dashType: "dash" }
  });
}

function addFooter(s, text = "TASA · FS-8A · XDS&TTC Operations · 系工組") {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.3, w: 10, h: 0.325,
    fill: { color: "1A321A" }, line: { color: "1A321A" }
  });
  s.addText(text, {
    x: 0.4, y: 5.32, w: 9.2, h: 0.28,
    fontFace: "Calibri", fontSize: 11, color: "4A7A4A", align: "left", margin: 0
  });
}

function addChip(s, text, x, y, color = YELLOW, textColor = "1A1A1A") {
  const w = 0.14 * text.length + 0.3;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h: 0.28,
    fill: { color }, line: { color }, rectRadius: 0.05
  });
  s.addText(text, {
    x, y, w, h: 0.28,
    fontFace: "Calibri", fontSize: 10, bold: true, color: textColor,
    align: "center", valign: "middle", margin: 0
  });
  return w;
}

function addCard(s, x, y, w, h, title, body, accent = YELLOW) {
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: "254525" }, line: { color: "3A6A3A", width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.07, h,
    fill: { color: accent }, line: { color: accent }
  });
  s.addText(title, {
    x: x + 0.14, y: y + 0.08, w: w - 0.22, h: 0.3,
    fontFace: "Calibri", fontSize: 12, bold: true, color: accent, margin: 0
  });
  s.addText(body, {
    x: x + 0.14, y: y + 0.42, w: w - 0.22, h: h - 0.52,
    fontFace: "Calibri", fontSize: 11, color: WHITE, margin: 0
  });
}

// ══════════════════════════════════════════════════════════════
// Slide 1 — 封面
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 5.625,
    fill: { color: "1A321A" }, line: { color: "1A321A" }
  });

  s.addText("XDS & TTC", {
    x: 0.45, y: 0.75, w: 9.1, h: 1.1,
    fontFace: "Impact", fontSize: 68, color: YELLOW,
    charSpacing: 2, margin: 0
  });
  s.addText("Operations", {
    x: 0.45, y: 1.8, w: 9.1, h: 0.85,
    fontFace: "Impact", fontSize: 50, color: WHITE,
    charSpacing: 1, margin: 0
  });

  s.addText("X-band 下行系統 & S-band TTC 飛行操作訓練", {
    x: 0.45, y: 2.8, w: 8.5, h: 0.45,
    fontFace: "Calibri", fontSize: 18, color: BLUE,
    italic: true, margin: 0
  });

  s.addShape(pres.shapes.LINE, {
    x: 0.45, y: 3.38, w: 6.5, h: 0,
    line: { color: "3A6A3A", width: 1.5, dashType: "dash" }
  });

  addChip(s, "FS-8A",  0.45, 3.58, YELLOW);
  addChip(s, "TTC",    1.25, 3.58, BLUE, "1A1A1A");
  addChip(s, "XDS",    1.85, 3.58, ORANGE, "1A1A1A");
  addChip(s, "系工組",  2.45, 3.58, "3A6A3A", WHITE);

  s.addText("2025 / 04 / 02", {
    x: 7.0, y: 4.5, w: 2.8, h: 0.35,
    fontFace: "Calibri", fontSize: 13, color: DIM, align: "right", margin: 0
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 2 — 課程大綱
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();
  addTitle(s, "課程大綱  AGENDA");

  const topics = [
    { num: "01", title: "TTC 系統規格與架構", sub: "S-band 頻率 · 收發規格 · 路徑切換",  color: YELLOW },
    { num: "02", title: "TTC 操作程序",       sub: "RF Switch · 指令清單 · 遙測監視點", color: BLUE },
    { num: "03", title: "XDS X-band 系統架構", sub: "RSI EU · Cavity BPF · Horn Antenna", color: GREEN },
    { num: "04", title: "XDS 操作程序",        sub: "XTx-A On · XTx-B On · 遙測確認",   color: ORANGE },
  ];

  topics.forEach((t, i) => {
    const x = (i % 2) * 4.8 + 0.4;
    const y = Math.floor(i / 2) * 1.8 + 1.15;
    const w = 4.4, h = 1.55;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h,
      fill: { color: "254525" }, line: { color: "3A6A3A", width: 1 },
      shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 0.06,
      fill: { color: t.color }, line: { color: t.color }
    });
    s.addText(t.num, {
      x: x + 0.15, y: y + 0.12, w: 0.65, h: 0.7,
      fontFace: "Impact", fontSize: 40, color: t.color, margin: 0
    });
    s.addText(t.title, {
      x: x + 0.85, y: y + 0.15, w: w - 0.98, h: 0.4,
      fontFace: "Calibri", fontSize: 14, bold: true, color: WHITE, margin: 0
    });
    s.addText(t.sub, {
      x: x + 0.85, y: y + 0.58, w: w - 0.98, h: 0.4,
      fontFace: "Calibri", fontSize: 10, color: DIM, margin: 0
    });
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 3 — TTC 規格
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();
  addTitle(s, "01  TTC 系統規格  S-band Specifications");

  // Transmitter
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.45, h: 3.95,
    fill: { color: BG2 }, line: { color: YELLOW, width: 1.5 }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.45, h: 0.38,
    fill: { color: YELLOW }, line: { color: YELLOW }
  });
  s.addText("Transmitter (STx)", {
    x: 0.45, y: 1.1, w: 4.35, h: 0.38,
    fontFace: "Impact", fontSize: 16, color: "1A1A1A",
    align: "center", valign: "middle", margin: 0
  });

  const txRows = [
    ["Center Frequency", "2215 MHz"],
    ["Bandwidth", "≥ 4 MHz"],
    ["Modulation", "BPSK"],
    ["Data Rate (LEO)",  "50 kbps"],
    ["Data Rate (ASH)",  "250 kbps"],
    ["Data Rate (NM)",   "2 Mbps"],
  ];
  txRows.forEach(([k, v], i) => {
    const y = i * 0.52 + 1.58;
    s.addText(k, { x: 0.55, y, w: 2.5, h: 0.38, fontFace: "Calibri", fontSize: 11, color: DIM, margin: 0 });
    s.addText(v, { x: 3.1,  y, w: 1.6, h: 0.38, fontFace: "Consolas", fontSize: 12, bold: true, color: WHITE, align: "right", margin: 0 });
  });

  // Receiver
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.15, y: 1.1, w: 4.45, h: 3.95,
    fill: { color: BG2 }, line: { color: BLUE, width: 1.5 }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.15, y: 1.1, w: 4.45, h: 0.38,
    fill: { color: BLUE }, line: { color: BLUE }
  });
  s.addText("Receiver (SRx)", {
    x: 5.2, y: 1.1, w: 4.35, h: 0.38,
    fontFace: "Impact", fontSize: 16, color: "1A1A1A",
    align: "center", valign: "middle", margin: 0
  });

  const rxRows = [
    ["Center Frequency", "2039.645833 MHz"],
    ["Bandwidth", "≥ 2 MHz"],
    ["Modulation", "BPSK"],
    ["Data Rate", "32 kbps"],
    ["Received Power", "-50 ~ -112 dBm"],
    ["Carrier Sweep Rate", "≥ 32 kHz/s"],
  ];
  rxRows.forEach(([k, v], i) => {
    const y = i * 0.52 + 1.58;
    s.addText(k, { x: 5.3, y, w: 2.3, h: 0.38, fontFace: "Calibri", fontSize: 11, color: DIM, margin: 0 });
    s.addText(v, { x: 7.6, y, w: 1.85, h: 0.38, fontFace: "Consolas", fontSize: 11, bold: true, color: WHITE, align: "right", margin: 0 });
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 4 — TTC 架構與路徑切換
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();
  addTitle(s, "01  TTC 架構與 RF 路徑切換");

  // 左：架構說明
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.4, h: 3.9,
    fill: { color: BG2 }, line: { color: "3A6A3A" }
  });
  s.addText("系統架構", {
    x: 0.55, y: 1.15, w: 4.1, h: 0.3,
    fontFace: "Calibri", fontSize: 13, bold: true, color: YELLOW, margin: 0
  });

  const archItems = [
    { label: "TRxA (710)", color: YELLOW },
    { label: "TRxB (711)", color: BLUE },
    { label: "RF SW (730)", color: ORANGE },
    { label: "Diplexer A/B (740/741)", color: GREEN },
    { label: "Helix_1A (750)  Patch antennas", color: DIM },
  ];
  archItems.forEach((a, i) => {
    const y = i * 0.56 + 1.52;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.55, y, w: 0.12, h: 0.32,
      fill: { color: a.color }, line: { color: a.color }
    });
    s.addText(a.label, {
      x: 0.78, y, w: 3.85, h: 0.32,
      fontFace: "Calibri", fontSize: 11, color: WHITE, margin: 0
    });
  });

  s.addShape(pres.shapes.LINE, {
    x: 0.55, y: 4.35, w: 4.1, h: 0,
    line: { color: "3A6A3A", width: 1, dashType: "dash" }
  });
  s.addText("雙備援架構：TRxA 與 TRxB 可互為備援\n切換由 RF Switch (730) 控制", {
    x: 0.55, y: 4.45, w: 4.1, h: 0.48,
    fontFace: "Calibri", fontSize: 10, color: DIM, margin: 0
  });

  // 右：POS1 / POS2
  const posList = [
    {
      pos: "POS 1", color: YELLOW,
      rows: [
        ["TRxA (710)", "Diplexer A (740)", "Helix_1A (750) / Patch_2A (752)"],
        ["TRxB (711)", "Diplexer B (741)", "Patch_1B (751) / Patch_2B (753)"],
      ]
    },
    {
      pos: "POS 2", color: BLUE,
      rows: [
        ["TRxA (710)", "Diplexer B (741)", "Patch_2B (753) / Helix_1A (750)"],
        ["TRxB (711)", "Diplexer A (740)", "Patch_2A (752)"],
      ]
    },
  ];

  posList.forEach((pos, pi) => {
    const y = pi * 2.05 + 1.1;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.1, y, w: 4.5, h: 1.9,
      fill: { color: "254525" }, line: { color: pos.color, width: 1.5 }
    });
    s.addText(pos.pos, {
      x: 5.2, y: y + 0.08, w: 1.2, h: 0.32,
      fontFace: "Impact", fontSize: 16, color: pos.color, margin: 0
    });
    pos.rows.forEach((r, ri) => {
      const ry = ri * 0.62 + y + 0.48;
      s.addText(r[0], { x: 5.2, ry, y: ry, w: 1.3, h: 0.3, fontFace: "Consolas", fontSize: 9, color: WHITE, margin: 0 });
      s.addText("→", { x: 6.55, y: ry, w: 0.3, h: 0.3, fontFace: "Calibri", fontSize: 11, color: pos.color, align: "center", margin: 0 });
      s.addText(r[1], { x: 6.9, y: ry, w: 1.4, h: 0.3, fontFace: "Consolas", fontSize: 9, color: DIM, margin: 0 });
      s.addText("→ " + r[2], { x: 5.2, y: ry + 0.28, w: 4.25, h: 0.26, fontFace: "Calibri", fontSize: 9, color: DIM, margin: 0 });
    });
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 5 — TTC 指令清單
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();
  addTitle(s, "02  TTC 指令清單  EPS Commands");

  const cmds = [
    { mnemonic: "ESBTXAEN",    desc: "S-band Transmitter A Enable",        color: GREEN },
    { mnemonic: "ESBTXBEN",    desc: "S-band Transmitter B Enable",        color: GREEN },
    { mnemonic: "ESBTXADS",    desc: "S-band Transmitter A Disable",       color: RED },
    { mnemonic: "ESBTXBDS",    desc: "S-band Transmitter B Disable",       color: RED },
    { mnemonic: "ESBSWPOS1",   desc: "S-band RF Switch → Position 1 Pulse", color: YELLOW },
    { mnemonic: "ESBSWPOS2",   desc: "S-band RF Switch → Position 2 Pulse", color: YELLOW },
    { mnemonic: "ESBSWPOS1OC", desc: "RF Switch POS1 OC flag",             color: BLUE },
    { mnemonic: "ESBSWPOS2OC", desc: "RF Switch POS2 OC flag",             color: BLUE },
    { mnemonic: "ESBTXAON",    desc: "S-band Transmitter A On",            color: ORANGE },
    { mnemonic: "ESBTXBON",    desc: "S-band Transmitter B On",            color: ORANGE },
    { mnemonic: "CSTXOFF",     desc: "S-band Transmitter A/B OFF",         color: DIM },
  ];

  cmds.forEach((c, i) => {
    const col = i < 6 ? 0 : 1;
    const row = i < 6 ? i : i - 6;
    const x = col * 4.85 + 0.35;
    const y = row * 0.67 + 1.08;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.45, h: 0.57,
      fill: { color: "1E3A1E" }, line: { color: c.color, width: 0.75 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.07, h: 0.57,
      fill: { color: c.color }, line: { color: c.color }
    });
    s.addText(c.mnemonic, {
      x: x + 0.15, y: y + 0.05, w: 2.1, h: 0.26,
      fontFace: "Consolas", fontSize: 11, bold: true, color: c.color, margin: 0
    });
    s.addText(c.desc, {
      x: x + 0.15, y: y + 0.3, w: 4.2, h: 0.22,
      fontFace: "Calibri", fontSize: 10, color: DIM, margin: 0
    });
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 6 — TTC 遙測監視點
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();
  addTitle(s, "02  TTC 遙測監視點  Key Telemetry");

  const groups = [
    {
      title: "Transceiver A 監視點", color: YELLOW,
      items: [
        ["SBAND_A_RXBPSK_RATE",    "32KBPS"],
        ["SBAND_A_BPSK_CALOCK",    "LOCK"],
        ["SBAND_A_RXLOCK",         "LOCK"],
        ["SBAND_A_TX_POWER_GOOD",  "NOMINAL"],
        ["SBAND_A_TX_LOCK",        "LOCK"],
        ["SBAND_A_TXDATARATEINDEX","50/250/2000 kbps"],
      ]
    },
    {
      title: "Transceiver B 監視點", color: BLUE,
      items: [
        ["SBAND_B_RXBPSK_RATE",    "32KBPS"],
        ["SBAND_B_BPSK_CALOCK",    "LOCK"],
        ["SBAND_B_RXLOCK",         "LOCK"],
        ["SBAND_B_TX_LOCK",        "LOCK"],
        ["SBAND_B_TXDATARATEINDEX","50/250/2000 kbps"],
        ["SBAND_B_SSPA_TEMP",      "-20 ~ +60 °C"],
      ]
    },
  ];

  groups.forEach((g, gi) => {
    const x = gi * 4.85 + 0.35;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.08, w: 4.45, h: 3.95,
      fill: { color: BG2 }, line: { color: g.color, width: 1.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.08, w: 4.45, h: 0.35,
      fill: { color: g.color }, line: { color: g.color }
    });
    s.addText(g.title, {
      x: x + 0.1, y: 1.08, w: 4.25, h: 0.35,
      fontFace: "Calibri", fontSize: 12, bold: true, color: "1A1A1A",
      valign: "middle", margin: 0
    });
    g.items.forEach(([mnem, val], ii) => {
      const y = ii * 0.58 + 1.5;
      s.addText(mnem, {
        x: x + 0.15, y, w: 2.8, h: 0.28,
        fontFace: "Consolas", fontSize: 9.5, color: WHITE, margin: 0
      });
      s.addText(val, {
        x: x + 3.0, y, w: 1.3, h: 0.28,
        fontFace: "Consolas", fontSize: 9.5, color: g.color, align: "right", margin: 0
      });
      if (ii < g.items.length - 1) {
        s.addShape(pres.shapes.LINE, {
          x: x + 0.1, y: y + 0.32, w: 4.25, h: 0,
          line: { color: "3A6A3A", width: 0.5, dashType: "dot" }
        });
      }
    });
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 7 — XDS 系統架構
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();
  addTitle(s, "03  XDS X-band 下行系統架構");

  // 中央架構流程（左→右）
  const nodes = [
    { label: "OBC\n(511)", x: 0.35, color: DIM },
    { label: "RSI EU-A\n(820 / 850)", x: 2.3, color: YELLOW },
    { label: "X-band Tx A\n(860 / 870)", x: 4.45, color: ORANGE },
    { label: "Cavity BPF A\n(860)", x: 6.6, color: BLUE },
    { label: "Horn\nAntenna A", x: 8.6, color: GREEN },
  ];

  nodes.forEach((n, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: n.x, y: 1.15, w: 1.75, h: 0.95,
      fill: { color: "254525" }, line: { color: n.color, width: 1.5 }
    });
    s.addText(n.label, {
      x: n.x, y: 1.15, w: 1.75, h: 0.95,
      fontFace: "Calibri", fontSize: 10, bold: true, color: n.color,
      align: "center", valign: "middle", margin: 0
    });
    if (i < nodes.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: n.x + 1.75, y: 1.625, w: 0.55, h: 0,
        line: { color: WHITE, width: 1.5, endArrowType: "arrow" }
      });
    }
  });

  // Channel B
  const nodesB = [
    { label: "RSI EU-B\n(821 / 851)", x: 2.3, color: YELLOW },
    { label: "X-band Tx B\n(861 / 871)", x: 4.45, color: ORANGE },
    { label: "Cavity BPF B\n(861)", x: 6.6, color: BLUE },
    { label: "Horn\nAntenna B", x: 8.6, color: GREEN },
  ];
  nodesB.forEach((n, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: n.x, y: 2.6, w: 1.75, h: 0.95,
      fill: { color: "1E3A1E" }, line: { color: n.color, width: 1, dashType: "dash" }
    });
    s.addText(n.label, {
      x: n.x, y: 2.6, w: 1.75, h: 0.95,
      fontFace: "Calibri", fontSize: 10, color: n.color,
      align: "center", valign: "middle", margin: 0
    });
    if (i < nodesB.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: n.x + 1.75, y: 3.075, w: 0.55, h: 0,
        line: { color: DIM, width: 1, endArrowType: "arrow" }
      });
    }
  });

  // OBC 連到兩條路
  s.addShape(pres.shapes.LINE, {
    x: 2.1, y: 1.625, w: 0.2, h: 0,
    line: { color: WHITE, width: 1.5, endArrowType: "arrow" }
  });
  s.addShape(pres.shapes.LINE, {
    x: 2.1, y: 3.075, w: 0.2, h: 0,
    line: { color: DIM, width: 1, endArrowType: "arrow" }
  });
  s.addShape(pres.shapes.LINE, {
    x: 2.1, y: 1.625, w: 0, h: 1.45,
    line: { color: DIM, width: 1 }
  });

  // 介面說明
  const buses = [
    { label: "LVDS  Data & Clock", x: 2.9, y: 2.35 },
    { label: "UART / RS422", x: 0.35, y: 2.35 },
    { label: "CAN Bus", x: 0.35, y: 2.75 },
  ];
  buses.forEach(b => {
    s.addText(b.label, {
      x: b.x, y: b.y, w: 2.2, h: 0.22,
      fontFace: "Calibri", fontSize: 9, color: DIM, italic: true, margin: 0
    });
  });

  // 底部備註
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.0, w: 9.3, h: 0.9,
    fill: { color: BG2 }, line: { color: "3A6A3A" }
  });
  s.addText("Channel A（主）與 Channel B（備援）互為冗餘架構。OBC 透過 LVDS 傳送影像資料，RSI EU 負責資料格式化，X-band Tx 進行 RF 調變後經 Cavity BPF 濾波，最後由 Horn 天線下行傳輸至地面。", {
    x: 0.5, y: 4.05, w: 9.1, h: 0.8,
    fontFace: "Calibri", fontSize: 11, color: DIM, margin: 0
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 8 — XTx-A 開機程序
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();
  addTitle(s, "04  XTx-A 開機程序  XTx-A On");

  const steps = [
    {
      step: "Step 1", label: "Enable XTx-A Equipment",
      cmd: "CMD CXTXEAON",
      tms: [
        ["XTXA_MAIN_STATE", "IdlePower"],
        ["XTXA_TX_PWR_ONOFF", "Power Off"],
      ],
      color: YELLOW
    },
    {
      step: "Step 2", label: "Configure Channel A Input",
      cmd: 'CMD PXTXACMD  XTXCMDID="HBP_DATAIF_CONFIG"  XTXCMDDATA=0x00',
      tms: [
        ["XTXA_DATA_IF_CFG_CHANNEL", "Channel A"],
        ["XTXA_TEST_MODE_ENDIS", "Normal Mode"],
      ],
      color: BLUE
    },
    {
      step: "Step 3", label: "Enable Transmission",
      cmd: 'CMD PXTXACMD  XTXCMDID="TRANSMISSION_ON_OFF"  XTXCMDDATA=0x80',
      tms: [
        ["XTXA_MAIN_STATE", "Transmitting"],
        ["XTXA_TX_PWR_ONOFF", "Power On"],
      ],
      color: GREEN
    },
    {
      step: "Step 4", label: "Disable Transmission",
      cmd: 'CMD PXTXACMD  XTXCMDID="TRANSMISSION_ON_OFF"  XTXCMDDATA=0x00',
      tms: [
        ["XTXA_MAIN_STATE", "IdlePower"],
      ],
      color: RED
    },
  ];

  steps.forEach((st, i) => {
    const y = i * 1.02 + 1.08;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 9.3, h: 0.9,
      fill: { color: "1E3A1E" }, line: { color: st.color, width: 1 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 0.07, h: 0.9,
      fill: { color: st.color }, line: { color: st.color }
    });
    s.addText(st.step, {
      x: 0.5, y: y + 0.06, w: 0.9, h: 0.26,
      fontFace: "Impact", fontSize: 13, color: st.color, margin: 0
    });
    s.addText(st.label, {
      x: 1.45, y: y + 0.06, w: 4.0, h: 0.26,
      fontFace: "Calibri", fontSize: 11, bold: true, color: WHITE, margin: 0
    });
    s.addText(st.cmd, {
      x: 0.5, y: y + 0.36, w: 5.5, h: 0.46,
      fontFace: "Consolas", fontSize: 9.5, color: st.color, margin: 0
    });
    st.tms.forEach((tm, ti) => {
      s.addText(`${tm[0]}: ${tm[1]}`, {
        x: 6.2 + ti * 0.0, y: y + 0.06 + ti * 0.36, w: 3.35, h: 0.28,
        fontFace: "Calibri", fontSize: 9.5, color: DIM, align: "right", margin: 0
      });
    });
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 9 — XTx-B 操作程序
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();
  addTitle(s, "04  XTx-B 操作程序  XTx-B Operations");

  const ops = [
    {
      cmd: "CMD CTXTEBON",
      label: "XTx-B On（設備指令）",
      tm: "XTXB_BASEBAND_ENABLE = 0x03",
      note: "設備電源開啟",
      color: GREEN
    },
    {
      cmd: "CMD PXTXFON  XTX_CONF=\"XTXB\"",
      label: "功能指令 — 設定 FPGA Digital & RF",
      tm: "FPGA digital & RF 參數初始化",
      note: "設定 FPGA 組態",
      color: BLUE
    },
    {
      cmd: "CMD CXTXBRFON",
      label: "RF On",
      tm: "XTXB_PA_3W_ENABLE = ENABLED\nXTXB_PA_2W_ENABLE = ENABLED",
      note: "功率放大器啟動",
      color: YELLOW
    },
    {
      cmd: "CMD CXTXBRFOFF",
      label: "RF Off",
      tm: "XTXB_TX_PA_NEG_ENABLE = DISABLED\nXTXB_TX_PA_POS_ENABLE = DISABLED",
      note: "功率放大器關閉",
      color: ORANGE
    },
    {
      cmd: "CMD CXTXEBOFF",
      label: "XTx-B Off",
      tm: "XTXB_BASEBAND_ENABLE = 0x00",
      note: "設備完全關閉",
      color: RED
    },
  ];

  ops.forEach((op, i) => {
    const y = i * 0.82 + 1.08;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 9.3, h: 0.7,
      fill: { color: "1E3A1E" }, line: { color: op.color, width: 0.75 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 0.07, h: 0.7,
      fill: { color: op.color }, line: { color: op.color }
    });
    s.addText(op.cmd, {
      x: 0.5, y: y + 0.06, w: 4.8, h: 0.28,
      fontFace: "Consolas", fontSize: 11, bold: true, color: op.color, margin: 0
    });
    s.addText(op.label, {
      x: 0.5, y: y + 0.38, w: 4.5, h: 0.24,
      fontFace: "Calibri", fontSize: 10, color: WHITE, margin: 0
    });
    s.addText(op.tm, {
      x: 5.5, y: y + 0.06, w: 3.95, h: 0.42,
      fontFace: "Consolas", fontSize: 9, color: DIM, margin: 0
    });
    s.addText(op.note, {
      x: 5.5, y: y + 0.44, w: 3.95, h: 0.2,
      fontFace: "Calibri", fontSize: 9, color: op.color, italic: true, margin: 0
    });
  });

  addFooter(s);
}

// ══════════════════════════════════════════════════════════════
// Slide 10 — 總結
// ══════════════════════════════════════════════════════════════
{
  const s = addSlide();

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 5.625,
    fill: { color: "1A321A" }, line: { color: "1A321A" }
  });

  s.addText("SUMMARY", {
    x: 0.5, y: 0.3, w: 9, h: 0.55,
    fontFace: "Impact", fontSize: 14, color: "3A6A3A", charSpacing: 6, margin: 0
  });
  s.addText("總結", {
    x: 0.5, y: 0.75, w: 9, h: 0.8,
    fontFace: "Impact", fontSize: 52, color: YELLOW, charSpacing: 2, margin: 0
  });

  const pts = [
    { icon: "◈", text: "TTC S-band：雙備援 TRxA/B，透過 RF Switch (730) 切換 POS1/POS2 路徑", color: YELLOW },
    { icon: "◈", text: "TTC 操作關鍵：確認 BPSK_CALOCK / RXLOCK / TX_LOCK 遙測全部 LOCK 狀態", color: BLUE },
    { icon: "◈", text: "XDS X-band：OBC → RSI EU → X-band Tx → BPF → Horn Antenna 下行鏈路", color: GREEN },
    { icon: "◈", text: "XTx 操作：設備 On → FPGA 設定 → RF On，每步驟皆需確認遙測回報正常", color: ORANGE },
  ];

  pts.forEach((pt, i) => {
    s.addText([
      { text: pt.icon + "  ", options: { color: pt.color, bold: true } },
      { text: pt.text, options: { color: WHITE } }
    ], {
      x: 0.5, y: i * 0.75 + 1.72, w: 9.1, h: 0.6,
      fontFace: "Calibri", fontSize: 13, valign: "middle", margin: 0
    });
  });

  addFooter(s, "XDS & TTC Operations · TASA FS-8A · 系工組");
}

// ── 輸出 ──────────────────────────────────────────────────────
pres.writeFile({ fileName: "C:/Users/kevin/Claude Cowork/class/EDI/TTC/FS8A/XDS_TTC_FS8A.pptx" })
  .then(() => console.log("✅ PPTX 已生成：XDS_TTC_FS8A.pptx"))
  .catch(e => { console.error("❌ 錯誤：", e); process.exit(1); });
