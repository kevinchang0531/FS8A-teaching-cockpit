const pptxgen = require("pptxgenjs");

const prs = new pptxgen();
prs.layout = 'LAYOUT_16x9';
prs.title = 'FS9A SAT 教學簡報';

// ── 色彩常數（no # prefix）──
const C = {
  bg:     "2D4A2D",
  bg2:    "1C3A1C",
  card:   "254525",
  border: "3A6A3A",
  header: "1A301A",
  white:  "FFFFFF",
  yellow: "F5E642",
  blue:   "A8D8EA",
  orange: "FFB347",
  green:  "98FB98",
  dim:    "C8C8C8",
  dark:   "0D1F0D",
};

function makeShadow() {
  return { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.25 };
}

// 每張投影片共用的背景條
function addBase(slide) {
  slide.background = { color: C.bg };
  // 頂部深色標題條
  slide.addShape(prs.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.85,
    fill: { color: C.header }, line: { color: C.border, width: 1 }
  });
  // 底部細線
  slide.addShape(prs.shapes.RECTANGLE, {
    x: 0, y: 5.42, w: 10, h: 0.2,
    fill: { color: C.border }
  });
  // 底部角標
  slide.addText("TASA | FS9A SAT 操控系統驗收測試", {
    x: 0.2, y: 5.38, w: 6, h: 0.22,
    fontSize: 8, color: C.dim, italic: true
  });
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.3, y: 0, w: 9.4, h: 0.85,
    fontSize: 26, fontFace: "Impact", color: C.yellow,
    valign: "middle", margin: 0
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.3, y: 0.82, w: 9.4, h: 0.25,
      fontSize: 10, color: C.blue, valign: "top", margin: 0
    });
  }
}

// 卡片背景
function addCard(slide, x, y, w, h, color) {
  slide.addShape(prs.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: color || C.card },
    line: { color: C.border, width: 1 },
    shadow: makeShadow()
  });
}

// ════════════════════════════════════════════
// Slide 01 — 標題頁
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  sl.background = { color: C.bg2 };

  // 中央大卡
  sl.addShape(prs.shapes.RECTANGLE, {
    x: 1.0, y: 1.2, w: 8.0, h: 3.2,
    fill: { color: C.bg }, line: { color: C.yellow, width: 2 },
    shadow: makeShadow()
  });

  // 黃色左側縱條
  sl.addShape(prs.shapes.RECTANGLE, {
    x: 1.0, y: 1.2, w: 0.18, h: 3.2,
    fill: { color: C.yellow }
  });

  // SAT 徽章
  sl.addShape(prs.shapes.OVAL, {
    x: 8.2, y: 1.3, w: 0.7, h: 0.7,
    fill: { color: C.yellow }, line: { color: C.yellow, width: 1 }
  });
  sl.addText("SAT", {
    x: 8.2, y: 1.3, w: 0.7, h: 0.7,
    fontSize: 10, fontFace: "Impact", color: C.bg2,
    align: "center", valign: "middle", bold: true
  });

  sl.addText("FS9A 系統驗收測試訓練", {
    x: 1.4, y: 1.4, w: 6.6, h: 1.0,
    fontSize: 34, fontFace: "Impact", color: C.yellow,
    valign: "middle"
  });
  sl.addText("FORMOSAT-9A 衛星操控系統驗收測試審查", {
    x: 1.4, y: 2.4, w: 6.6, h: 0.5,
    fontSize: 15, color: C.blue, valign: "top"
  });
  sl.addText([
    { text: "台灣太空中心 TASA", options: { color: C.white } },
    { text: "   |   ", options: { color: C.border } },
    { text: "操控組", options: { color: C.white } },
    { text: "   |   ", options: { color: C.border } },
    { text: "2026", options: { color: C.yellow } },
  ], {
    x: 1.4, y: 3.2, w: 6.6, h: 0.4,
    fontSize: 12, valign: "top"
  });

  sl.addText("System Acceptance Test | SOCC · TT&C · FDF · MCC · DCC · GCN", {
    x: 0, y: 5.2, w: 10, h: 0.3,
    fontSize: 9, color: C.dim, align: "center"
  });
}

// ════════════════════════════════════════════
// Slide 02 — FS9A 任務目標
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "FS9A 任務目標", "FORMOSAT-9A Mission Objectives");

  // 左欄：任務要點
  addCard(sl, 0.3, 1.15, 5.9, 4.15);
  sl.addText("任務核心目標", {
    x: 0.3, y: 1.15, w: 5.9, h: 0.38,
    fontSize: 13, fontFace: "Impact", color: C.yellow,
    align: "center", valign: "middle"
  });
  sl.addText([
    { text: "台灣首個 X 頻段 SAR 衛星任務", options: { bullet: true, breakLine: true, color: C.white, bold: true } },
    { text: "建立自主 SAR 技術能力：Bus、酬載、關鍵元件", options: { bullet: true, breakLine: true, color: C.dim } },
    { text: "2027/2029 發射 9A/9B 兩星組成星系", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "全天候觀測（SAR 不受雲霧影響，24 小時可用）", options: { bullet: true, breakLine: true, color: C.dim } },
    { text: "應用：防救災、環境監測、資源管理、情報蒐集", options: { bullet: true, color: C.white } },
  ], {
    x: 0.45, y: 1.58, w: 5.6, h: 3.6,
    fontSize: 13, valign: "top", paraSpaceAfter: 6
  });

  // 右欄：關鍵數字
  addCard(sl, 6.35, 1.15, 3.35, 1.9, C.border);
  sl.addText("最高解析度", { x: 6.35, y: 1.15, w: 3.35, h: 0.35, fontSize: 11, color: C.blue, align: "center", valign: "middle" });
  sl.addText("1 m", { x: 6.35, y: 1.5, w: 3.35, h: 0.75, fontSize: 44, fontFace: "Impact", color: C.yellow, align: "center", valign: "middle" });
  sl.addText("Spotlight 模式（SRD 規範 0.7m）", { x: 6.35, y: 2.25, w: 3.35, h: 0.28, fontSize: 9, color: C.dim, align: "center" });

  addCard(sl, 6.35, 3.15, 3.35, 1.85, C.border);
  sl.addText("星系計畫", { x: 6.35, y: 3.15, w: 3.35, h: 0.35, fontSize: 11, color: C.blue, align: "center", valign: "middle" });
  sl.addText("2 顆", { x: 6.35, y: 3.5, w: 3.35, h: 0.75, fontSize: 40, fontFace: "Impact", color: C.green, align: "center", valign: "middle" });
  sl.addText("FS9A (2027) + FS9B (2029)", { x: 6.35, y: 4.25, w: 3.35, h: 0.28, fontSize: 9, color: C.dim, align: "center" });
  sl.addText("重訪週期：4天/單星  ·  2天/雙星", { x: 6.35, y: 4.55, w: 3.35, h: 0.28, fontSize: 9, color: C.orange, align: "center" });
}

// ════════════════════════════════════════════
// Slide 03 — 系統三大區段
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "FORMOSAT-9A 系統架構", "System Architecture — Three Segments");

  const segs = [
    {
      title: "太空段", sub: "Space Segment",
      color: C.blue,
      items: ["衛星本體 Bus（三軸穩定）", "主酬載：XPAA（X頻段相位陣列天線）", "次酬載：GNSS-RO/R", "頻率：9.65 GHz（X-Band SAR）"]
    },
    {
      title: "地面段", sub: "Ground Segment",
      color: C.yellow,
      items: ["SOCS：SOCC + TT&C + FDF + MCC + DCC + GCN", "TT&C 地面站：中壢 TS1 + 台南 TS2 + RTS", "IPS 影像處理系統", "SDC 科學數據中心"]
    },
    {
      title: "發射段", sub: "Launch Segment",
      color: C.orange,
      items: ["發射載具 LV", "發射支援服務（LSS）", "入軌服務（OIS）", "發射場介面支援"]
    },
  ];

  segs.forEach((s, i) => {
    const x = 0.25 + i * 3.25;
    addCard(sl, x, 1.1, 3.1, 4.3);
    sl.addShape(prs.shapes.RECTANGLE, { x, y: 1.1, w: 3.1, h: 0.48, fill: { color: C.header } });
    sl.addText(s.title, { x, y: 1.1, w: 3.1, h: 0.28, fontSize: 15, fontFace: "Impact", color: s.color, align: "center", valign: "middle", margin: 0 });
    sl.addText(s.sub, { x, y: 1.38, w: 3.1, h: 0.2, fontSize: 9, color: C.dim, align: "center", margin: 0 });
    sl.addText(
      s.items.map((item, j) => ({
        text: item,
        options: { bullet: true, breakLine: j < s.items.length - 1, color: j % 2 === 0 ? C.white : C.dim }
      })),
      { x: x + 0.12, y: 1.65, w: 2.86, h: 3.6, fontSize: 12, valign: "top", paraSpaceAfter: 5 }
    );
  });
}

// ════════════════════════════════════════════
// Slide 04 — SAR 觀測模式
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "SAR 三種觀測模式", "SAR Imaging Modes");

  const modes = [
    { name: "Stripmap 條帶", res: "3 – 20 m", swath: "≥ 50 km", use: "環境監測 / 農業 / 地形測繪", color: C.green, icon: "▬" },
    { name: "Spotlight 聚光", res: "最高 1 m", swath: "小幅寬", use: "偵察 / 安全 / 高精度目標", color: C.yellow, icon: "◎" },
    { name: "ScanSAR 掃描", res: "~ 10 m", swath: "超寬幅", use: "災害監測 / 海事 / 全球覆蓋", color: C.blue, icon: "⊞" },
  ];

  modes.forEach((m, i) => {
    const x = 0.25 + i * 3.25;
    addCard(sl, x, 1.1, 3.1, 4.3);
    // 頂色條
    sl.addShape(prs.shapes.RECTANGLE, { x, y: 1.1, w: 3.1, h: 0.55, fill: { color: C.header } });
    sl.addText(m.icon, { x, y: 1.1, w: 0.7, h: 0.55, fontSize: 22, color: m.color, align: "center", valign: "middle" });
    sl.addText(m.name, { x: x + 0.65, y: 1.1, w: 2.45, h: 0.55, fontSize: 13, fontFace: "Impact", color: m.color, valign: "middle" });

    // 規格格
    const specs = [
      { label: "空間解析度", val: m.res },
      { label: "觀測幅寬", val: m.swath },
    ];
    specs.forEach((sp, j) => {
      const sy = 1.75 + j * 0.85;
      sl.addShape(prs.shapes.RECTANGLE, { x: x + 0.15, y: sy, w: 2.8, h: 0.75, fill: { color: C.border } });
      sl.addText(sp.label, { x: x + 0.15, y: sy, w: 2.8, h: 0.3, fontSize: 9, color: C.dim, align: "center", valign: "middle" });
      sl.addText(sp.val, { x: x + 0.15, y: sy + 0.3, w: 2.8, h: 0.42, fontSize: 18, fontFace: "Impact", color: m.color, align: "center", valign: "middle" });
    });

    // 應用
    sl.addShape(prs.shapes.RECTANGLE, { x: x + 0.15, y: 3.5, w: 2.8, h: 0.25, fill: { color: C.border } });
    sl.addText("主要應用", { x: x + 0.15, y: 3.5, w: 2.8, h: 0.25, fontSize: 9, color: C.dim, align: "center", valign: "middle" });
    sl.addText(m.use, { x: x + 0.15, y: 3.78, w: 2.8, h: 0.55, fontSize: 10, color: C.white, align: "center", valign: "middle" });
  });

  sl.addText("MR120：X-Band 中心頻率 9.65 GHz  ·  MR20：三種模式均須支援  ·  MR40：入射角 25°~45°（特殊需求可至 20°~60°）", {
    x: 0.2, y: 5.38, w: 9, h: 0.2,
    fontSize: 8, color: C.dim
  });
}

// ════════════════════════════════════════════
// Slide 05 — 衛星操作模式
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "衛星操作模式（Bus Modes）", "Satellite Operating Modes");

  const modes = [
    { name: "Ground Mode", sub: "地面整測", desc: "地面組裝、整測及環境試驗期間使用。衛星在發射前的最終驗收狀態。", color: C.dim },
    { name: "ASH Mode", sub: "捕獲與安全保持", desc: "Acquisition & Safe Hold。發射後初始捕獲，確保通訊建立前維持安全姿態與電力。", color: C.orange },
    { name: "Normal Mode", sub: "正常任務操作", desc: "LVLH 姿態 + 偏航操縱（Yaw Steering）。指向精度 < 0.03°，執行 SAR 成像任務。", color: C.green },
    { name: "OCM Mode", sub: "軌道控制模式", desc: "Orbit Control Mode。執行軌道機動（Delta-V burn）時切換，軌道維持與碰撞規避。", color: C.blue },
  ];

  // 流程箭頭
  const modeX = [0.25, 2.65, 5.05, 7.45];
  const arrows = ["→", "→", "↔"];

  modes.forEach((m, i) => {
    const x = modeX[i];
    addCard(sl, x, 1.15, 2.25, 4.15);
    sl.addShape(prs.shapes.RECTANGLE, { x, y: 1.15, w: 2.25, h: 0.45, fill: { color: C.header } });
    sl.addText(m.name, { x, y: 1.15, w: 2.25, h: 0.45, fontSize: 12, fontFace: "Impact", color: m.color, align: "center", valign: "middle" });
    sl.addText(m.sub, { x, y: 1.62, w: 2.25, h: 0.28, fontSize: 10, color: m.color, align: "center", valign: "top" });
    sl.addShape(prs.shapes.LINE, { x: x + 0.2, y: 1.98, w: 1.85, h: 0, line: { color: C.border, width: 1 } });
    sl.addText(m.desc, { x: x + 0.12, y: 2.1, w: 2.01, h: 2.9, fontSize: 11, color: C.white, valign: "top" });
  });

  arrows.forEach((a, i) => {
    sl.addText(a, {
      x: modeX[i] + 2.25, y: 2.8, w: 0.35, h: 0.5,
      fontSize: 18, color: C.yellow, align: "center", valign: "middle"
    });
  });

  sl.addText("Normal Mode 最低每 48 小時接收一次指令（SOCSRD 00520）", {
    x: 0.3, y: 5.38, w: 9, h: 0.2,
    fontSize: 8, color: C.dim
  });
}

// ════════════════════════════════════════════
// Slide 06 — 關鍵性能規格
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "FS9A 關鍵性能規格摘要", "Key Performance Requirements");

  const specs = [
    { label: "軌道高度", val: "514 km", sub: "SSO，LTDN 11am–12pm", color: C.blue },
    { label: "軌道傾角", val: "97.46°", sub: "太陽同步軌道（SSO）", color: C.blue },
    { label: "設計壽命", val: "≥ 5 年", sub: "含燃料餘量", color: C.green },
    { label: "衛星濕重", val: "≤ 735 kg", sub: "含推進劑、平衡質量", color: C.orange },
    { label: "姿態精度", val: "< 0.03°", sub: "Normal Mode 指向精度", color: C.yellow },
    { label: "定位誤差", val: "< 50 m", sub: "GCP 校正後 3 pixels", color: C.yellow },
  ];

  const cols = 3, rows = 2;
  const cw = 3.1, ch = 1.9, gx = 0.25, gy = 1.1;

  specs.forEach((s, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = gx + col * (cw + 0.05);
    const y = gy + row * (ch + 0.1);
    addCard(sl, x, y, cw, ch);
    sl.addShape(prs.shapes.RECTANGLE, { x, y, w: 0.18, h: ch, fill: { color: s.color } });
    sl.addText(s.label, { x: x + 0.25, y: y + 0.15, w: cw - 0.3, h: 0.3, fontSize: 11, color: C.dim, valign: "top" });
    sl.addText(s.val, { x: x + 0.25, y: y + 0.42, w: cw - 0.3, h: 0.8, fontSize: 28, fontFace: "Impact", color: s.color, valign: "middle" });
    sl.addText(s.sub, { x: x + 0.25, y: y + 1.5, w: cw - 0.3, h: 0.3, fontSize: 9, color: C.dim, valign: "top" });
  });
}

// ════════════════════════════════════════════
// Slide 07 — MOC 任務操作中心
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "MOC — 任務操作中心", "Mission Operations Center");

  addCard(sl, 0.25, 1.1, 6.0, 4.3);
  sl.addText("MOC 核心功能", {
    x: 0.25, y: 1.1, w: 6.0, h: 0.38,
    fontSize: 13, fontFace: "Impact", color: C.blue,
    align: "center", valign: "middle"
  });
  sl.addText([
    { text: "24/7 全天候衛星健康監控（SOH 遙傳接收與處理）", options: { bullet: true, breakLine: true, color: C.white, bold: true } },
    { text: "指令上鏈處理（Command Processing）", options: { bullet: true, breakLine: true, color: C.dim } },
    { text: "地面站遠端監控：TT&C、XAS、RTS", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "危機管理（Crisis Management）：衛星與地面設備", options: { bullet: true, breakLine: true, color: C.dim } },
    { text: "遙傳資料庫管理（Database Management）", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "衛星遙傳資料流模擬驗證（Telemetry Stream Simulation）", options: { bullet: true, color: C.dim } },
  ], {
    x: 0.4, y: 1.53, w: 5.7, h: 3.7,
    fontSize: 12, valign: "top", paraSpaceAfter: 7
  });

  // 右側 SAT 測試項
  addCard(sl, 6.4, 1.1, 3.3, 4.3, C.border);
  sl.addText("SAT 測試項目", {
    x: 6.4, y: 1.1, w: 3.3, h: 0.38,
    fontSize: 11, fontFace: "Impact", color: C.yellow, align: "center", valign: "middle"
  });
  const tests = [
    "SAT_MOC_01：即時遙傳處理",
    "SAT_MOC_02：播放遙傳處理",
    "SAT_MOC_03：地面站遠端監控",
    "SAT_MOC_04：遙傳資料流模擬",
    "SAT_MOC_05：指令上鏈處理",
  ];
  sl.addText(tests.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < tests.length - 1, color: i % 2 === 0 ? C.white : C.dim }
  })), {
    x: 6.55, y: 1.55, w: 3.0, h: 3.6,
    fontSize: 10, valign: "top", paraSpaceAfter: 5
  });
}

// ════════════════════════════════════════════
// Slide 08 — MCC & FDF
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "MCC 任務控制中心 / FDF 飛行動態設施", "Mission Control Center / Flight Dynamics Facility");

  // MCC
  addCard(sl, 0.25, 1.1, 4.65, 4.3);
  sl.addShape(prs.shapes.RECTANGLE, { x: 0.25, y: 1.1, w: 4.65, h: 0.4, fill: { color: C.header } });
  sl.addText("MCC 任務控制中心", { x: 0.25, y: 1.1, w: 4.65, h: 0.4, fontSize: 13, fontFace: "Impact", color: C.yellow, align: "center", valign: "middle" });
  sl.addText([
    { text: "衛星與酬載活動排程（Scheduling）", options: { bullet: true, breakLine: true, color: C.white, bold: true } },
    { text: "時程衝突解決與重排程", options: { bullet: true, breakLine: true, color: C.dim } },
    { text: "Command Load Package 產生（TTQ / MPQ / GCQ）", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "海外支援站（RTS）操作時序確認", options: { bullet: true, breakLine: true, color: C.dim } },
    { text: "每軌 ≥ 2 分鐘酬載資料採集規劃（SOCSRD 00510）", options: { bullet: true, color: C.orange } },
  ], {
    x: 0.4, y: 1.55, w: 4.35, h: 3.7,
    fontSize: 12, valign: "top", paraSpaceAfter: 7
  });

  // FDF
  addCard(sl, 5.1, 1.1, 4.65, 4.3);
  sl.addShape(prs.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.65, h: 0.4, fill: { color: C.header } });
  sl.addText("FDF 飛行動態設施", { x: 5.1, y: 1.1, w: 4.65, h: 0.4, fontSize: 13, fontFace: "Impact", color: C.blue, align: "center", valign: "middle" });
  sl.addText([
    { text: "軌道判定（OD）& 預測（使用 OASYS 商用軟體）", options: { bullet: true, breakLine: true, color: C.white, bold: true } },
    { text: "TLE（Two-Line Element）產生", options: { bullet: true, breakLine: true, color: C.dim } },
    { text: "Delta-V 軌道調整規劃", options: { bullet: true, breakLine: true, color: C.white } },
    { text: "碰撞預警：距離 <50m & 概率 >1/10000", options: { bullet: true, breakLine: true, color: C.orange } },
    { text: "轉移至 514 km 操作軌道（精度 ±50m / 0.01°）", options: { bullet: true, color: C.dim } },
  ], {
    x: 5.25, y: 1.55, w: 4.35, h: 3.7,
    fontSize: 12, valign: "top", paraSpaceAfter: 7
  });
}

// ════════════════════════════════════════════
// Slide 09 — DCC / GCN / TT&C
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "DCC / GCN / TT&C 次系統", "Data Control Center / Ground Communications / TT&C Stations");

  const secs = [
    {
      title: "DCC 資料控制中心", color: C.orange,
      items: [
        "SOH 遙傳資料存檔與管理",
        "趨勢分析（Trending）：監控衛星健康長期趨勢",
        "趨勢統計報告產生（評估與預測）",
        "軌道數據與虛擬通道資料轉傳使用者",
      ]
    },
    {
      title: "GCN 地面通訊網路", color: C.blue,
      items: [
        "SOCC 內部 LAN + 對外介面（IPS / SDC / 發射場）",
        "效能、可靠性與資安保護",
        "行動內部網路支援（SOCS Mobile Intranet）",
        "測試項：SAT_GCN_01（全面連通性驗證）",
      ]
    },
    {
      title: "TT&C 遙傳追蹤指令站", color: C.green,
      items: [
        "S 頻段：中壢 TS1（含 MOC 功能）+ 台南 TS2 + RTS 備援",
        "功能：指令上鏈 / 遙傳下鏈 / 角度追蹤",
        "每日 ≥ 1 次接觸，每次 ≥ 4 分鐘（仰角 10°+）",
        "X 頻段接收站（XAS）：國內主站 + 備援站",
      ]
    },
  ];

  secs.forEach((s, i) => {
    const x = 0.25 + i * 3.25;
    addCard(sl, x, 1.1, 3.1, 4.3);
    sl.addShape(prs.shapes.RECTANGLE, { x, y: 1.1, w: 3.1, h: 0.42, fill: { color: C.header } });
    sl.addText(s.title, { x, y: 1.1, w: 3.1, h: 0.42, fontSize: 11, fontFace: "Impact", color: s.color, align: "center", valign: "middle" });
    sl.addText(s.items.map((item, j) => ({
      text: item, options: { bullet: true, breakLine: j < s.items.length - 1, color: j % 2 === 0 ? C.white : C.dim }
    })), {
      x: x + 0.12, y: 1.57, w: 2.86, h: 3.7,
      fontSize: 11, valign: "top", paraSpaceAfter: 8
    });
  });
}

// ════════════════════════════════════════════
// Slide 10 — S 頻段介面
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "S 頻段 TT&C 介面規格（GS_4）", "S-Band Interface Control Requirements");

  const rows = [
    ["項目", "規格", "備註"],
    ["下行頻率", "2215 MHz", "BPSK 調變"],
    ["上行頻率", "2039.645833 MHz", "BPSK 調變"],
    ["極化", "RHCP", "—"],
    ["遙傳速率", "50 / 250 kbps / 2 Mbps", "最高 2 Mbps"],
    ["指令速率", "32 / 250 kbps", "—"],
    ["遙傳編碼", "CCSDS RS(255,223) Interleave=5", "NRZ-M 線碼"],
    ["指令加密", "AES-256-CCM 模式", "FS-5 heritage"],
    ["鏈路餘裕", "≥ 6 dB @ 10° 仰角", "遙傳 & 指令"],
    ["BER", "遙傳 ≤ 10⁻⁶ / 指令 ≤ 10⁻⁷", "—"],
  ];

  const tableData = rows.map((r, ri) => r.map((cell, ci) => ({
    text: cell,
    options: {
      bold: ri === 0,
      color: ri === 0 ? C.bg2 : (ci === 1 ? C.yellow : (ci === 0 ? C.white : C.dim)),
      fill: { color: ri === 0 ? C.blue : (ri % 2 === 0 ? C.card : C.bg) },
      fontSize: ri === 0 ? 11 : 10,
      align: ci === 1 ? "center" : (ci === 0 ? "left" : "left"),
    }
  })));

  sl.addTable(tableData, {
    x: 0.4, y: 1.1, w: 9.2, h: 4.2,
    border: { pt: 1, color: C.border },
    colW: [2.5, 3.8, 2.9],
  });
}

// ════════════════════════════════════════════
// Slide 11 — X 頻段介面
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "X 頻段影像下鏈介面規格（GS_5）", "X-Band Interface Control Requirements");

  // 兩欄：XPAA vs 傳統 X-band
  addCard(sl, 0.25, 1.1, 4.65, 4.3, C.card);
  sl.addShape(prs.shapes.RECTANGLE, { x: 0.25, y: 1.1, w: 4.65, h: 0.4, fill: { color: C.header } });
  sl.addText("X 頻段相位陣列天線（XPAA）", { x: 0.25, y: 1.1, w: 4.65, h: 0.4, fontSize: 12, fontFace: "Impact", color: C.yellow, align: "center", valign: "middle" });

  const xpaaSpecs = [
    ["頻率", "8200 MHz"],
    ["下鏈速率", "800 Mbps"],
    ["調變", "QPSK / 16APSK / 32APSK (CCSDS 131.2-B-2)"],
    ["幅寬（3dB）", "250 / 270 / 300 MHz（roll-off 0.25/0.35/0.5）"],
    ["編碼", "LDPC (8160, 7136)，無交錯"],
    ["BER", "≤ 10⁻⁹"],
    ["鏈路餘裕", "≥ 6 dB @ 20° 仰角"],
    ["衛星 EIRP", "≥ 28 dBW（XPAA peak）"],
  ];
  sl.addText(xpaaSpecs.map(([label, val], i) => [
    { text: label + "：", options: { bold: true, color: C.dim, breakLine: false } },
    { text: val, options: { color: C.yellow, breakLine: i < xpaaSpecs.length - 1 } }
  ]).flat(), {
    x: 0.4, y: 1.57, w: 4.35, h: 3.7,
    fontSize: 10.5, valign: "top", paraSpaceAfter: 5
  });

  addCard(sl, 5.1, 1.1, 4.65, 4.3, C.card);
  sl.addShape(prs.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.65, h: 0.4, fill: { color: C.header } });
  sl.addText("傳統 X-band 天線（XTx）", { x: 5.1, y: 1.1, w: 4.65, h: 0.4, fontSize: 12, fontFace: "Impact", color: C.blue, align: "center", valign: "middle" });

  const xtxSpecs = [
    ["頻率", "8200 MHz"],
    ["下鏈速率", "600 Mbps"],
    ["調變", "8PSK-4D-TCM（ECSS-E-ST-50-05C）"],
    ["極化", "RHCP"],
    ["編碼", "LDPC (8160, 7136)，無交錯"],
    ["BER", "≤ 10⁻⁹"],
    ["衛星 EIRP", "≥ 20 dBW（peak）"],
    ["地面天線 G/T", "≥ 31.3 dB/K（TS2 ≥ 32.5 dB/K）"],
  ];
  sl.addText(xtxSpecs.map(([label, val], i) => [
    { text: label + "：", options: { bold: true, color: C.dim, breakLine: false } },
    { text: val, options: { color: C.blue, breakLine: i < xtxSpecs.length - 1 } }
  ]).flat(), {
    x: 5.25, y: 1.57, w: 4.35, h: 3.7,
    fontSize: 10.5, valign: "top", paraSpaceAfter: 5
  });
}

// ════════════════════════════════════════════
// Slide 12 — SAT 驗收測試方法
// ════════════════════════════════════════════
{
  const sl = prs.addSlide();
  addBase(sl);
  addTitle(sl, "系統驗收測試（SAT）方法與準則", "SAT Methodology — IADT Acceptance Criteria");

  // IADT 四格
  const iadt = [
    { code: "ii", name: "Inspection", zh: "檢查", desc: "文件審查、目視確認、記錄比對。無需實際操作設備。", color: C.blue },
    { code: "aa", name: "Analysis", zh: "分析", desc: "計算、模擬、工程推導。以既有資料或模型驗證符合性。", color: C.yellow },
    { code: "dd", name: "Demonstration", zh: "展示", desc: "功能操作演示，觀察結果符合預期。無詳細量測數據要求。", color: C.orange },
    { code: "tt", name: "Test", zh: "測試", desc: "實際量測、有完整數據記錄。最嚴格的驗收方式。", color: C.green },
  ];

  iadt.forEach((item, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.25 + col * 4.87;
    const y = 1.1 + row * 2.1;
    addCard(sl, x, y, 4.72, 1.95);
    sl.addShape(prs.shapes.RECTANGLE, { x, y, w: 0.8, h: 1.95, fill: { color: C.header } });
    sl.addText(`"${item.code}"`, { x, y, w: 0.8, h: 0.7, fontSize: 18, fontFace: "Impact", color: item.color, align: "center", valign: "middle" });
    sl.addText(item.zh, { x, y: y + 0.68, w: 0.8, h: 0.5, fontSize: 12, color: item.color, align: "center", valign: "top" });
    sl.addText(item.name, { x, y: y + 1.18, w: 0.8, h: 0.62, fontSize: 9, color: C.dim, align: "center", valign: "top" });
    sl.addText(item.desc, { x: x + 0.88, y: y + 0.15, w: 3.7, h: 1.65, fontSize: 11, color: C.white, valign: "top" });
  });

  // 底部三大測試範疇
  addCard(sl, 0.25, 5.3, 9.5, 0.0);
  sl.addText("測試範疇：①衛星操控系統需求 (SOCSRD)   ②地面↔衛星介面控制 (FS9-ICD-0004)   ③地面端介面控制 (FS9-ICD-0001)", {
    x: 0.3, y: 5.28, w: 9.4, h: 0.22,
    fontSize: 8.5, color: C.dim, align: "center"
  });
}

// ════════════════════════════════════════════
// 輸出
// ════════════════════════════════════════════
prs.writeFile({ fileName: "C:\\Users\\kevin\\Claude Cowork\\class\\EDI\\SAT\\FS9A\\FS9A_SAT_教學簡報.pptx" })
  .then(() => console.log("[OK] FS9A_SAT_教學簡報.pptx 已儲存"))
  .catch(e => { console.error("[ERR]", e); process.exit(1); });
