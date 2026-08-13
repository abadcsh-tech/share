// Build lecture.pptx — Apple-style 14-slide deck mirroring lecture.html
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" × 7.5"
const W = 13.3, H = 7.5;
pres.title = "MD 파일, 모든 것의 시작";
pres.author = "csh";

// === Tokens ============================================================
const C = {
  primary: "0066CC",
  primaryOnDark: "2997FF",
  canvas: "FFFFFF",
  parchment: "F5F5F7",
  pearl: "FAFAFC",
  tile1: "272729",
  tile2: "2A2A2C",
  tile3: "1D1D1F",
  ink: "1D1D1F",
  inkMuted80: "333333",
  inkMuted48: "7A7A7A",
  onDark: "FFFFFF",
  bodyMuted: "CCCCCC",
  hairline: "E0E0E0",
  divider: "F0F0F0",
};

const FONT_DISPLAY = "Apple SD Gothic Neo";
const FONT_TEXT = "Apple SD Gothic Neo";
const FONT_MONO = "Menlo";

// === Helpers ===========================================================
const PAD_X = 0.9;        // outer left/right padding
const PAD_TOP = 0.7;
const CONTENT_W = W - 2 * PAD_X;

function bg(slide, mode) {
  const map = {
    light: C.canvas,
    parchment: C.parchment,
    dark: C.tile1,
    "dark-2": C.tile2,
    "dark-3": C.tile3,
  };
  slide.background = { color: map[mode] || C.canvas };
}

function isDark(mode) {
  return mode && mode.startsWith("dark");
}

function eyebrow(slide, text, mode = "light", y = PAD_TOP) {
  slide.addText(text, {
    x: PAD_X, y, w: CONTENT_W, h: 0.4,
    fontFace: FONT_DISPLAY,
    fontSize: 14,
    bold: true,
    charSpacing: 4,
    color: isDark(mode) ? C.primaryOnDark : C.primary,
    margin: 0,
  });
}

function title(slide, runs, opts = {}) {
  const {
    y = PAD_TOP + 0.55,
    h = 2.2,
    align = "left",
    fontSize = 56,
    color = C.ink,
  } = opts;
  // runs: array of {text, color?, bold?}
  const text = runs.map((r, i) => ({
    text: r.text,
    options: {
      color: r.color || color,
      bold: r.bold !== false,
      breakLine: r.breakLine || false,
    },
  }));
  slide.addText(text, {
    x: PAD_X, y, w: CONTENT_W, h,
    fontFace: FONT_DISPLAY,
    fontSize,
    align,
    valign: "top",
    margin: 0,
    paraSpaceAfter: 0,
  });
}

function lead(slide, text, opts = {}) {
  const {
    y, w = CONTENT_W, x = PAD_X, h = 1.8, fontSize = 20, color, align = "left",
  } = opts;
  slide.addText(text, {
    x, y, w, h,
    fontFace: FONT_TEXT,
    fontSize,
    color: color || C.inkMuted80,
    align,
    valign: "top",
    margin: 0,
    paraSpaceAfter: 6,
  });
}

function pageNumber(slide, n, total, mode) {
  const c = isDark(mode) ? "888888" : C.inkMuted48;
  slide.addText(`${n} / ${total}`, {
    x: W - 1.2, y: H - 0.45, w: 0.9, h: 0.3,
    fontFace: FONT_TEXT, fontSize: 10, color: c, align: "right",
    margin: 0,
  });
  slide.addText("MD가 곧 원본이다 — 강의", {
    x: PAD_X, y: H - 0.45, w: 6, h: 0.3,
    fontFace: FONT_TEXT, fontSize: 10, color: c,
    margin: 0,
  });
  // Top progress bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.05,
    fill: { color: isDark(mode) ? "333333" : "EEEEEE" },
    line: { type: "none" },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: (W * n) / total, h: 0.05,
    fill: { color: C.primary },
    line: { type: "none" },
  });
}

function pillTag(slide, text, mode) {
  const dark = isDark(mode);
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: (W - 2.2) / 2, y: PAD_TOP + 0.1, w: 2.2, h: 0.4,
    fill: { color: dark ? "FFFFFF" : C.parchment, transparency: dark ? 88 : 0 },
    line: { type: "none" },
    rectRadius: 0.2,
  });
  slide.addText(text, {
    x: (W - 2.2) / 2, y: PAD_TOP + 0.1, w: 2.2, h: 0.4,
    fontFace: FONT_TEXT, fontSize: 12, bold: true, charSpacing: 2,
    color: dark ? C.onDark : C.ink, align: "center", valign: "middle",
    margin: 0,
  });
}

// Renders a panel that looks like a code editor / preview window
function renderPanel(slide, x, y, w, h, header, mode, contentRenderer) {
  const dark = isDark(mode);
  const panelFill = dark ? C.tile3 : C.canvas;
  const headerLine = dark ? "FFFFFF" : C.hairline;

  // Panel background
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: panelFill },
    line: { color: dark ? "FFFFFF" : C.hairline, width: dark ? 0.5 : 1 },
    rectRadius: 0.15,
  });
  // Header strip
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h: 0.45,
    fill: { color: panelFill },
    line: { color: headerLine, width: 0 },
  });
  // Three traffic-light dots
  ["FF5F57", "FEBC2E", "28C840"].forEach((dotColor, i) => {
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.18 + i * 0.22, y: y + 0.16, w: 0.13, h: 0.13,
      fill: { color: dotColor },
      line: { type: "none" },
    });
  });
  // Header label
  slide.addText(header.left, {
    x: x + 0.95, y: y + 0.05, w: w - 2.5, h: 0.36,
    fontFace: FONT_TEXT, fontSize: 10, bold: true,
    charSpacing: 2,
    color: dark ? C.bodyMuted : C.inkMuted48,
    align: "left", valign: "middle",
    margin: 0,
  });
  if (header.right) {
    slide.addText(header.right, {
      x: x + w - 2, y: y + 0.05, w: 1.85, h: 0.36,
      fontFace: FONT_TEXT, fontSize: 10, bold: true,
      charSpacing: 2,
      color: dark ? C.bodyMuted : C.inkMuted48,
      align: "right", valign: "middle",
      margin: 0,
    });
  }
  // Divider
  slide.addShape(pres.shapes.LINE, {
    x, y: y + 0.45, w, h: 0,
    line: { color: dark ? "444444" : C.divider, width: 0.5 },
  });
  // Content
  contentRenderer(x + 0.3, y + 0.6, w - 0.6, h - 0.8, dark);
}

// === Slides =============================================================
const TOTAL = 14;

// ----- Slide 1: Title --------------------------------------------------
{
  const s = pres.addSlide();
  bg(s, "light");
  pillTag(s, "강의 자료 · 2026", "light");
  // Centered title
  s.addText([
    { text: "MD 파일,", options: { bold: true, breakLine: true } },
    { text: "모든 것의 ", options: { bold: true, color: C.ink } },
    { text: "시작", options: { bold: true, color: C.primary } },
    { text: ".", options: { bold: true, color: C.ink } },
  ], {
    x: 0, y: 2.3, w: W, h: 2.2,
    fontFace: FONT_DISPLAY, fontSize: 76,
    align: "center", valign: "top", margin: 0,
    paraSpaceAfter: 0,
  });
  s.addText("Obsidian × Claude Code, 그리고 마크다운이 만드는 자유", {
    x: 0, y: 4.7, w: W, h: 0.6,
    fontFace: FONT_TEXT, fontSize: 24,
    color: C.inkMuted80, align: "center", valign: "top",
    margin: 0,
  });
  s.addText("→ 화살표 · Space · 클릭으로 진행", {
    x: 0, y: H - 1.0, w: W, h: 0.4,
    fontFace: FONT_TEXT, fontSize: 11, charSpacing: 4,
    color: C.inkMuted48, align: "center",
    margin: 0,
  });
  pageNumber(s, 1, TOTAL, "light");
}

// ----- Slide 2: 학습 목표 ----------------------------------------------
{
  const s = pres.addSlide();
  bg(s, "parchment");
  eyebrow(s, "이번 시간에", "parchment");
  title(s, [
    { text: "우리는 ", color: C.ink },
    { text: "이 세 가지", color: C.primary },
    { text: "를 이해하게 됩니다.", color: C.ink, breakLine: false },
  ], { fontSize: 44, h: 1.6 });

  const cardW = (CONTENT_W - 0.6) / 3;
  const cardY = 3.4;
  const cardH = 2.7;
  const objs = [
    { n: "1", t: "옵시디언이 하는 일", d: "MD 파일을 사람이 보기 편한 형태로 그려주는 도구라는 점." },
    { n: "2", t: "Claude Code가 하는 일", d: "그 MD 파일을 자동으로 만들어주는 생산자라는 점." },
    { n: "3", t: "왜 MD를 원본으로", d: "한 번 .md로 두면, 모든 포맷이 자유로워진다는 점." },
  ];
  objs.forEach((o, i) => {
    const x = PAD_X + i * (cardW + 0.3);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: cardY, w: cardW, h: cardH,
      fill: { color: C.canvas },
      line: { color: C.hairline, width: 1 },
      rectRadius: 0.18,
      shadow: { type: "outer", color: "000000", blur: 8, offset: 1, angle: 90, opacity: 0.04 },
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.4, y: cardY + 0.4, w: 0.55, h: 0.55,
      fill: { color: C.primary }, line: { type: "none" },
    });
    s.addText(o.n, {
      x: x + 0.4, y: cardY + 0.4, w: 0.55, h: 0.55,
      fontFace: FONT_DISPLAY, fontSize: 18, bold: true,
      color: C.onDark, align: "center", valign: "middle",
      margin: 0,
    });
    s.addText(o.t, {
      x: x + 0.4, y: cardY + 1.1, w: cardW - 0.8, h: 0.6,
      fontFace: FONT_DISPLAY, fontSize: 22, bold: true,
      color: C.ink, valign: "top", margin: 0,
    });
    s.addText(o.d, {
      x: x + 0.4, y: cardY + 1.75, w: cardW - 0.8, h: 0.95,
      fontFace: FONT_TEXT, fontSize: 15,
      color: C.inkMuted80, valign: "top", margin: 0,
      paraSpaceAfter: 4,
    });
  });
  pageNumber(s, 2, TOTAL, "parchment");
}

// ----- Slide 3: 옵시디언 정의 ------------------------------------------
{
  const s = pres.addSlide();
  bg(s, "light");
  eyebrow(s, "OBSIDIAN", "light");
  s.addText([
    { text: "옵시디언은", options: { bold: true, breakLine: true } },
    { text: "'", options: { bold: true } },
    { text: "디스플레이어", options: { bold: true, color: C.primary } },
    { text: "' 입니다.", options: { bold: true } },
  ], {
    x: PAD_X, y: PAD_TOP + 0.55, w: CONTENT_W, h: 2.2,
    fontFace: FONT_DISPLAY, fontSize: 64,
    color: C.ink, align: "left", valign: "top", margin: 0,
  });
  s.addText([
    { text: "내 컴퓨터에 저장된 ", options: {} },
    { text: "평문 .md 파일", options: { bold: true } },
    { text: "을 — 사람이 읽기 좋은 모습으로 ", options: {} },
    { text: "그려서 보여주는 것", options: { bold: true } },
    { text: ". 그 이상도, 그 이하도 아닙니다.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "파일은 옵시디언의 것이 아닙니다. 옵시디언이 사라져도, .md 파일은 그대로 남습니다.", options: {} },
  ], {
    x: PAD_X, y: 4.5, w: CONTENT_W, h: 2.2,
    fontFace: FONT_TEXT, fontSize: 22,
    color: C.inkMuted80, valign: "top", margin: 0,
    paraSpaceAfter: 8,
  });
  pageNumber(s, 3, TOTAL, "light");
}

// ----- Slide 4: 옵시디언 라이브 데모 (static side-by-side) -------------
{
  const s = pres.addSlide();
  bg(s, "parchment");
  eyebrow(s, "LIVE DEMO", "parchment");
  s.addText([
    { text: "왼쪽이 ", options: { bold: true } },
    { text: "실제 .md", options: { bold: true, color: C.primary } },
    { text: " · 오른쪽이 ", options: { bold: true } },
    { text: "옵시디언이 본 화면", options: { bold: true, color: C.primary } },
  ], {
    x: PAD_X, y: PAD_TOP + 0.55, w: CONTENT_W, h: 1.4,
    fontFace: FONT_DISPLAY, fontSize: 38,
    color: C.ink, align: "left", valign: "top", margin: 0,
  });
  const panelY = 3.0;
  const panelH = 3.7;
  const gap = 0.3;
  const panelW = (CONTENT_W - gap) / 2;

  // Left: raw markdown
  renderPanel(s, PAD_X, panelY, panelW, panelH, { left: "● ● ●  sermon.md", right: "MARKDOWN" }, "light", (cx, cy, cw, ch, dark) => {
    const md = `# 사랑의 본질

> "하나님이 세상을 이처럼 사랑하사" — *요 3:16*

## 세 가지 핵심
1. **무조건적** 사랑
2. **희생하는** 사랑
3. **지속되는** 사랑

### 적용
- 이웃을 향해 한 걸음 더
- \`오늘부터\` 작은 실천`;
    s.addText(md, {
      x: cx, y: cy, w: cw, h: ch,
      fontFace: FONT_MONO, fontSize: 11,
      color: C.ink, valign: "top", margin: 0,
      paraSpaceAfter: 0, lineSpacingMultiple: 1.4,
    });
  });

  // Right: rendered preview
  renderPanel(s, PAD_X + panelW + gap, panelY, panelW, panelH, { left: "◇  Obsidian Preview", right: "LIVE" }, "light", (cx, cy, cw, ch, dark) => {
    s.addText("사랑의 본질", {
      x: cx, y: cy, w: cw, h: 0.5,
      fontFace: FONT_DISPLAY, fontSize: 22, bold: true,
      color: C.ink, margin: 0,
    });
    // Blockquote
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy + 0.55, w: 0.04, h: 0.5,
      fill: { color: C.primary }, line: { type: "none" },
    });
    s.addText('"하나님이 세상을 이처럼 사랑하사" — 요 3:16', {
      x: cx + 0.15, y: cy + 0.55, w: cw - 0.15, h: 0.5,
      fontFace: FONT_TEXT, fontSize: 13, italic: true,
      color: C.inkMuted48, margin: 0,
    });
    s.addText("세 가지 핵심", {
      x: cx, y: cy + 1.2, w: cw, h: 0.4,
      fontFace: FONT_DISPLAY, fontSize: 17, bold: true,
      color: C.ink, margin: 0,
    });
    s.addText([
      { text: "1.  ", options: { color: C.inkMuted80 } },
      { text: "무조건적", options: { bold: true, color: C.ink } },
      { text: " 사랑", options: { color: C.inkMuted80, breakLine: true } },
      { text: "2.  ", options: { color: C.inkMuted80 } },
      { text: "희생하는", options: { bold: true, color: C.ink } },
      { text: " 사랑", options: { color: C.inkMuted80, breakLine: true } },
      { text: "3.  ", options: { color: C.inkMuted80 } },
      { text: "지속되는", options: { bold: true, color: C.ink } },
      { text: " 사랑", options: { color: C.inkMuted80 } },
    ], {
      x: cx + 0.1, y: cy + 1.65, w: cw - 0.1, h: 1.1,
      fontFace: FONT_TEXT, fontSize: 14,
      valign: "top", margin: 0, paraSpaceAfter: 4,
    });
    s.addText("적용", {
      x: cx, y: cy + 2.85, w: cw, h: 0.3,
      fontFace: FONT_DISPLAY, fontSize: 15, bold: true,
      color: C.ink, margin: 0,
    });
    s.addText([
      { text: "•  이웃을 향해 한 걸음 더", options: { breakLine: true } },
      { text: "•  ", options: {} },
      { text: "오늘부터", options: { fontFace: FONT_MONO } },
      { text: " 작은 실천", options: {} },
    ], {
      x: cx + 0.1, y: cy + 3.2, w: cw - 0.1, h: 0.8,
      fontFace: FONT_TEXT, fontSize: 13,
      color: C.inkMuted80, valign: "top", margin: 0, paraSpaceAfter: 2,
    });
  });

  pageNumber(s, 4, TOTAL, "parchment");
}

// ----- Slide 5: Claude Code 정의 ---------------------------------------
{
  const s = pres.addSlide();
  bg(s, "dark");
  eyebrow(s, "CLAUDE CODE", "dark");
  s.addText([
    { text: "Claude Code는", options: { bold: true, color: C.onDark, breakLine: true } },
    { text: "'", options: { bold: true, color: C.onDark } },
    { text: "생산자", options: { bold: true, color: C.primaryOnDark } },
    { text: "' 입니다.", options: { bold: true, color: C.onDark } },
  ], {
    x: PAD_X, y: PAD_TOP + 0.55, w: CONTENT_W, h: 2.2,
    fontFace: FONT_DISPLAY, fontSize: 64,
    align: "left", valign: "top", margin: 0,
  });
  s.addText([
    { text: "터미널에서 자연어로 부탁하면 — 글·노트·코드를 ", options: { color: C.bodyMuted } },
    { text: ".md 파일로 출력", options: { bold: true, color: C.onDark } },
    { text: "합니다.", options: { color: C.bodyMuted, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "그 출력은 곧바로 옵시디언 볼트에 떨어지고, 옵시디언이 즉시 그려줍니다.", options: { color: C.bodyMuted } },
  ], {
    x: PAD_X, y: 4.5, w: CONTENT_W, h: 2.2,
    fontFace: FONT_TEXT, fontSize: 22,
    valign: "top", margin: 0, paraSpaceAfter: 8,
  });
  pageNumber(s, 5, TOTAL, "dark");
}

// ----- Slide 6: 터미널 데모 --------------------------------------------
{
  const s = pres.addSlide();
  bg(s, "dark-2");
  eyebrow(s, "LIVE DEMO", "dark-2");
  s.addText([
    { text: '"', options: { bold: true, color: C.onDark } },
    { text: "자연어", options: { bold: true, color: C.primaryOnDark } },
    { text: '" 한 줄이', options: { bold: true, color: C.onDark, breakLine: true } },
    { text: "한 장의 ", options: { bold: true, color: C.onDark } },
    { text: ".md", options: { bold: true, color: C.primaryOnDark } },
    { text: "가 됩니다.", options: { bold: true, color: C.onDark } },
  ], {
    x: PAD_X, y: PAD_TOP + 0.55, w: CONTENT_W, h: 1.7,
    fontFace: FONT_DISPLAY, fontSize: 38,
    align: "left", valign: "top", margin: 0,
  });
  const panelY = 3.0;
  const panelH = 3.7;
  const gap = 0.3;
  const panelW = (CONTENT_W - gap) / 2;

  // Left: terminal
  renderPanel(s, PAD_X, panelY, panelW, panelH, { left: "● ● ●  Terminal", right: "~/CSH_REMOTE" }, "dark-2", (cx, cy, cw, ch, dark) => {
    const term = `$ claude
> 오늘 설교 핵심 세 가지를 정리해서 새 노트로 만들어줘
✓ 새 파일 생성: 2026-05-04.md
✓ 옵시디언 볼트에 저장됨
✓ 위키링크 3개 자동 연결
$ open obsidian://open?file=2026-05-04`;
    s.addText(term, {
      x: cx, y: cy, w: cw, h: ch,
      fontFace: FONT_MONO, fontSize: 11.5,
      color: C.onDark, valign: "top", margin: 0,
      paraSpaceAfter: 0, lineSpacingMultiple: 1.5,
    });
  });

  // Right: rendered output
  renderPanel(s, PAD_X + panelW + gap, panelY, panelW, panelH, { left: "새 파일 — 출력", right: "2026-05-04.MD" }, "dark-2", (cx, cy, cw, ch, dark) => {
    s.addText("2026-05-04 주일 설교", {
      x: cx, y: cy, w: cw, h: 0.5,
      fontFace: FONT_DISPLAY, fontSize: 22, bold: true,
      color: C.onDark, margin: 0,
    });
    s.addText("핵심 세 가지", {
      x: cx, y: cy + 0.6, w: cw, h: 0.4,
      fontFace: FONT_DISPLAY, fontSize: 17, bold: true,
      color: C.onDark, margin: 0,
    });
    s.addText([
      { text: "1.  ", options: { color: C.bodyMuted } },
      { text: "하나님의 사랑은 무조건적이다.", options: { bold: true, color: C.onDark, breakLine: true } },
      { text: "2.  ", options: { color: C.bodyMuted } },
      { text: "그 사랑은 십자가에서 증명된다.", options: { bold: true, color: C.onDark, breakLine: true } },
      { text: "3.  ", options: { color: C.bodyMuted } },
      { text: "우리는 그 사랑으로 이웃을 사랑한다.", options: { bold: true, color: C.onDark } },
    ], {
      x: cx + 0.1, y: cy + 1.05, w: cw - 0.1, h: 1.4,
      fontFace: FONT_TEXT, fontSize: 13.5,
      valign: "top", margin: 0, paraSpaceAfter: 4,
    });
    s.addText("본문", {
      x: cx, y: cy + 2.55, w: cw, h: 0.3,
      fontFace: FONT_DISPLAY, fontSize: 15, bold: true,
      color: C.onDark, margin: 0,
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy + 2.9, w: 0.04, h: 0.35,
      fill: { color: C.primaryOnDark }, line: { type: "none" },
    });
    s.addText("요한복음 3:16", {
      x: cx + 0.15, y: cy + 2.9, w: cw - 0.15, h: 0.35,
      fontFace: FONT_TEXT, fontSize: 13, italic: true,
      color: C.bodyMuted, margin: 0,
    });
    s.addText("연결", {
      x: cx, y: cy + 3.3, w: cw, h: 0.3,
      fontFace: FONT_DISPLAY, fontSize: 15, bold: true,
      color: C.onDark, margin: 0,
    });
    s.addText([
      { text: "•  [[하나님의 사랑]]", options: { breakLine: true } },
      { text: "•  [[십자가의 의미]]", options: { breakLine: true } },
      { text: "•  [[이웃 사랑의 실천]]", options: {} },
    ], {
      x: cx + 0.1, y: cy + 3.65, w: cw - 0.1, h: 1.1,
      fontFace: FONT_TEXT, fontSize: 12,
      color: C.primaryOnDark, valign: "top", margin: 0, paraSpaceAfter: 2,
    });
  });

  pageNumber(s, 6, TOTAL, "dark-2");
}

// ----- Slide 7: 핵심 통찰 ----------------------------------------------
{
  const s = pres.addSlide();
  bg(s, "dark-3");
  s.addText("핵심 통찰", {
    x: 0, y: 1.5, w: W, h: 0.5,
    fontFace: FONT_DISPLAY, fontSize: 14, bold: true,
    charSpacing: 4, color: C.primaryOnDark,
    align: "center", margin: 0,
  });
  s.addText([
    { text: "두 도구는 서로 다른 일을 하지만,", options: { color: C.onDark, breakLine: true } },
    { text: "같은 .md 파일", options: { color: C.primaryOnDark } },
    { text: "을 봅니다.", options: { color: C.onDark } },
  ], {
    x: 0, y: 2.3, w: W, h: 2.4,
    fontFace: FONT_DISPLAY, fontSize: 56, bold: true,
    align: "center", valign: "top", margin: 0,
  });
  s.addText([
    { text: "Claude Code가 만들고 → 옵시디언이 읽고 → 우리가 옵시디언에서 다듬고 → 다시 Claude Code가 변환합니다.", options: { color: C.bodyMuted, breakLine: true } },
    { text: "모두 ", options: { color: C.bodyMuted } },
    { text: ".md 한 장", options: { color: C.onDark, bold: true } },
    { text: "을 가운데 두고 돌아갑니다.", options: { color: C.bodyMuted } },
  ], {
    x: 1.5, y: 5.4, w: W - 3, h: 1.3,
    fontFace: FONT_TEXT, fontSize: 18,
    align: "center", valign: "top", margin: 0, paraSpaceAfter: 6,
  });
  pageNumber(s, 7, TOTAL, "dark-3");
}

// ----- Slide 8: 변환 6포맷 ---------------------------------------------
{
  const s = pres.addSlide();
  bg(s, "light");
  eyebrow(s, "CONVERT", "light");
  s.addText([
    { text: "하나의 원본,", options: { bold: true, color: C.ink, breakLine: true } },
    { text: "여섯 개", options: { bold: true, color: C.primary } },
    { text: "의 출력.", options: { bold: true, color: C.ink } },
  ], {
    x: PAD_X, y: PAD_TOP + 0.55, w: CONTENT_W, h: 1.8,
    fontFace: FONT_DISPLAY, fontSize: 50,
    align: "left", valign: "top", margin: 0,
  });
  s.addText("같은 .md가 어떻게 다른 형식이 되는가 — 여섯 가지 길.", {
    x: PAD_X, y: 3.55, w: CONTENT_W, h: 0.4,
    fontFace: FONT_TEXT, fontSize: 16, color: C.inkMuted80,
    margin: 0,
  });

  const cards = [
    { tag: "PPT", color: "D24726", title: "PowerPoint", desc: "설교 슬라이드" },
    { tag: "PDF", color: "D93025", title: "PDF", desc: "인쇄·배포" },
    { tag: "DOC", color: "2B579A", title: "Word DOCX", desc: "교회 양식" },
    { tag: "WEB", color: "E44D26", title: "인터랙티브 HTML", desc: "웹페이지·강의" },
    { tag: "IMG", color: "5E35B1", title: "인포그래픽", desc: "SNS·카드뉴스" },
    { tag: "DB",  color: "1A73E8", title: "구조화 데이터", desc: "JSON·DB 연동" },
  ];
  const cols = 3;
  const rows = 2;
  const cardW = (CONTENT_W - (cols - 1) * 0.25) / cols;
  const cardH = 1.3;
  const gridY = 4.2;
  cards.forEach((c, i) => {
    const r = Math.floor(i / cols);
    const col = i % cols;
    const x = PAD_X + col * (cardW + 0.25);
    const y = gridY + r * (cardH + 0.25);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.canvas },
      line: { color: C.hairline, width: 1 },
      rectRadius: 0.16,
    });
    // Tag pill
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.5,
      fill: { color: c.color }, line: { type: "none" },
      rectRadius: 0.08,
    });
    s.addText(c.tag, {
      x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.5,
      fontFace: FONT_DISPLAY, fontSize: 11, bold: true,
      color: C.onDark, align: "center", valign: "middle", margin: 0,
    });
    s.addText(c.title, {
      x: x + 1.15, y: y + 0.28, w: cardW - 1.3, h: 0.45,
      fontFace: FONT_DISPLAY, fontSize: 17, bold: true,
      color: C.ink, valign: "middle", margin: 0,
    });
    s.addText(c.desc, {
      x: x + 1.15, y: y + 0.7, w: cardW - 1.3, h: 0.4,
      fontFace: FONT_TEXT, fontSize: 13,
      color: C.inkMuted48, valign: "middle", margin: 0,
    });
  });
  pageNumber(s, 8, TOTAL, "light");
}

// ----- Comparison slide helper -----------------------------------------
function compareSlide(idx, eyebrowText, titleRuns, leftCard, rightCard) {
  const s = pres.addSlide();
  const mode = idx % 2 === 1 ? "dark-2" : "dark"; // 9: dark-2, 10: dark, 11: dark-2
  bg(s, mode);
  eyebrow(s, eyebrowText, mode);
  s.addText(titleRuns, {
    x: PAD_X, y: PAD_TOP + 0.55, w: CONTENT_W, h: 1.8,
    fontFace: FONT_DISPLAY, fontSize: 46, bold: true,
    color: C.onDark, valign: "top", margin: 0,
  });

  const cardY = 3.6;
  const cardH = 3.4;
  const gap = 0.3;
  const cardW = (CONTENT_W - gap) / 2;

  function drawCard(cx, card, isWin) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: C.tile2 },
      line: { color: isWin ? C.primaryOnDark : "FFFFFF", width: isWin ? 1.5 : 0 },
      rectRadius: 0.2,
    });
    s.addText(card.name, {
      x: cx + 0.5, y: cardY + 0.4, w: cardW - 1, h: 0.6,
      fontFace: FONT_DISPLAY, fontSize: 26, bold: true,
      color: C.onDark, margin: 0,
    });
    s.addText(card.meta, {
      x: cx + 0.5, y: cardY + 1.0, w: cardW - 1, h: 0.3,
      fontFace: FONT_TEXT, fontSize: 11, bold: true,
      charSpacing: 3, color: C.bodyMuted, margin: 0,
    });
    // Big stat
    s.addText(card.label, {
      x: cx + 0.5, y: cardY + 1.55, w: cardW - 1.7, h: 0.4,
      fontFace: FONT_TEXT, fontSize: 14,
      color: C.bodyMuted, valign: "middle", margin: 0,
    });
    s.addText(card.pct, {
      x: cx + cardW - 1.7, y: cardY + 1.4, w: 1.2, h: 0.7,
      fontFace: FONT_DISPLAY, fontSize: 36, bold: true,
      color: isWin ? C.primaryOnDark : C.onDark,
      align: "right", valign: "middle", margin: 0,
    });
    // Bar background
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx + 0.5, y: cardY + 2.1, w: cardW - 1, h: 0.12,
      fill: { color: "FFFFFF", transparency: 90 },
      line: { type: "none" },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx + 0.5, y: cardY + 2.1, w: (cardW - 1) * (card.barPct / 100), h: 0.12,
      fill: { color: isWin ? C.primaryOnDark : "888888" },
      line: { type: "none" },
    });
    // Bullets
    card.notes.forEach((n, i) => {
      s.addText(`—  ${n}`, {
        x: cx + 0.5, y: cardY + 2.5 + i * 0.32, w: cardW - 1, h: 0.3,
        fontFace: FONT_TEXT, fontSize: 12,
        color: C.bodyMuted, margin: 0,
      });
    });
  }
  drawCard(PAD_X, leftCard, true);
  drawCard(PAD_X + cardW + gap, rightCard, false);

  pageNumber(s, idx, TOTAL, mode);
}

// ----- Slide 9: AI 친화도 ----------------------------------------------
compareSlide(9, "왜 MD인가 · 1",
  [
    { text: "AI는 ", options: { color: C.onDark, bold: true } },
    { text: "평문", options: { color: C.primaryOnDark, bold: true } },
    { text: "을 가장 잘 읽습니다.", options: { color: C.onDark, bold: true } },
  ],
  {
    name: "Markdown", meta: "평문 · 텍스트", label: "AI 가독성", pct: "100%", barPct: 100,
    notes: ["AI가 구조를 한 번에 이해", "토큰 효율이 가장 높음", "변환·요약·재구성 자유"],
  },
  {
    name: "PPTX / DOCX", meta: "바이너리 · 폐쇄적", label: "AI 가독성", pct: "35%", barPct: 35,
    notes: ["구조 정보가 형식 안에 갇힘", "파싱이 깨지기 쉬움", "이미지·텍스트가 분리되지 않음"],
  }
);

// ----- Slide 10: 변환 자유도 -------------------------------------------
compareSlide(10, "왜 MD인가 · 2",
  [
    { text: "한 번 닫힌 포맷은,", options: { color: C.onDark, bold: true, breakLine: true } },
    { text: "다시 열리지 않습니다.", options: { color: C.primaryOnDark, bold: true } },
  ],
  {
    name: "Markdown → 모든 포맷", meta: "출구가 열려 있다", label: "변환 성공률", pct: "98%", barPct: 98,
    notes: ["→ PPT · PDF · DOCX · HTML 모두", "레이아웃 깨짐 거의 없음", "한 원본에서 N개 출력"],
  },
  {
    name: "PPT → 다른 포맷", meta: "출구가 막혀 있다", label: "변환 성공률", pct: "22%", barPct: 22,
    notes: ["PPT를 Word로 → 깨짐", "PPT를 HTML로 → 거의 불가", "매번 처음부터 다시 작업"],
  }
);

// ----- Slide 11: 장기 보존성 -------------------------------------------
compareSlide(11, "왜 MD인가 · 3",
  [
    { text: "MD가 사라질 가능성은", options: { color: C.onDark, bold: true, breakLine: true } },
    { text: "매우 낮습니다.", options: { color: C.primaryOnDark, bold: true } },
  ],
  {
    name: "Markdown", meta: "오픈 표준 · 30년+", label: "장기 보존", pct: "95%", barPct: 95,
    notes: ["특정 회사에 종속되지 않음", "설령 바뀌어도 변환이 쉬움", "가장 대중적 → 미래 호환 유리"],
  },
  {
    name: "독자 포맷", meta: "회사·앱 종속", label: "장기 보존", pct: "40%", barPct: 40,
    notes: ["앱이 사라지면 파일도 위태", "버전 업그레이드마다 호환 깨짐", "옮길 때마다 또 한 번의 변환"],
  }
);

// ----- Slide 12: PPT 함정 (인용) ---------------------------------------
{
  const s = pres.addSlide();
  bg(s, "dark-3");
  s.addText("메인 소스의 문제", {
    x: 0, y: 1.3, w: W, h: 0.5,
    fontFace: FONT_DISPLAY, fontSize: 14, bold: true,
    charSpacing: 4, color: C.primaryOnDark,
    align: "center", margin: 0,
  });
  s.addText([
    { text: "\"메인 소스가 ", options: { color: C.onDark } },
    { text: ".md", options: { color: C.primaryOnDark } },
    { text: "로 되어 있지 않으면,", options: { color: C.onDark, breakLine: true } },
    { text: "나머지 파일들은 그저 ", options: { color: C.onDark } },
    { text: "호환되지 않는 섬", options: { color: C.primaryOnDark } },
    { text: "으로 존재한다.\"", options: { color: C.onDark } },
  ], {
    x: 0.5, y: 2.2, w: W - 1, h: 2.6,
    fontFace: FONT_DISPLAY, fontSize: 44, bold: true,
    align: "center", valign: "top", margin: 0,
  });
  s.addText([
    { text: "PPT를 한 번 만들어두면 — 그것은 그저 PPT일 뿐입니다.", options: { color: C.bodyMuted, breakLine: true } },
    { text: "MD에서 PPT로 가는 길은 ", options: { color: C.bodyMuted } },
    { text: "열려 있고", options: { color: C.onDark, bold: true } },
    { text: ", 거꾸로 가는 길은 ", options: { color: C.bodyMuted } },
    { text: "막혀 있습니다.", options: { color: C.onDark, bold: true } },
  ], {
    x: 1.5, y: 5.5, w: W - 3, h: 1.5,
    fontFace: FONT_TEXT, fontSize: 17,
    align: "center", valign: "top", margin: 0, paraSpaceAfter: 6,
  });
  pageNumber(s, 12, TOTAL, "dark-3");
}

// ----- Slide 13: 워크플로우 --------------------------------------------
{
  const s = pres.addSlide();
  bg(s, "light");
  eyebrow(s, "WORKFLOW", "light");
  s.addText([
    { text: "전체 흐름은", options: { bold: true, color: C.ink, breakLine: true } },
    { text: "이렇게 작동합니다.", options: { bold: true, color: C.ink } },
  ], {
    x: PAD_X, y: PAD_TOP + 0.55, w: CONTENT_W, h: 1.8,
    fontFace: FONT_DISPLAY, fontSize: 44,
    align: "left", valign: "top", margin: 0,
  });

  const steps = [
    { t: "Claude Code로 만든다", d: "\"이번 주 설교 노트 정리해줘\" — 자연어 한 줄이면 정돈된 .md가 출력됩니다." },
    { t: "옵시디언이 그려준다", d: "같은 폴더(볼트)를 옵시디언이 읽어 — 위키링크·태그·그래프뷰까지 살아 있는 노트로." },
    { t: "옵시디언에서 직접 다듬는다", d: "위키링크 걸기, 메모 추가, 본문 다듬기 — 모두 옵시디언에서. 파일은 여전히 .md." },
    { t: "다시 Claude Code에게 변환을 부탁한다", d: "\"PPT로 만들어줘\", \"교사용 DOCX로 묶어줘\" — 같은 원본에서 필요한 형식만." },
    { t: "원본은 하나, 출력은 무한", d: "주일은 PPT, 월간은 DOCX, 공유는 HTML, SNS는 인포그래픽 — 모두 같은 .md." },
  ];
  const startY = 3.4;
  const itemH = 0.72;
  // Vertical line
  s.addShape(pres.shapes.LINE, {
    x: PAD_X + 0.32, y: startY + 0.32, w: 0, h: (steps.length - 1) * itemH,
    line: { color: C.hairline, width: 1 },
  });
  steps.forEach((step, i) => {
    const y = startY + i * itemH;
    // Numbered circle
    s.addShape(pres.shapes.OVAL, {
      x: PAD_X, y, w: 0.65, h: 0.65,
      fill: { color: C.canvas },
      line: { color: C.primary, width: 2 },
    });
    s.addText(String(i + 1), {
      x: PAD_X, y, w: 0.65, h: 0.65,
      fontFace: FONT_DISPLAY, fontSize: 17, bold: true,
      color: C.primary, align: "center", valign: "middle", margin: 0,
    });
    // Title + desc on the right
    s.addText(step.t, {
      x: PAD_X + 0.95, y, w: CONTENT_W - 0.95, h: 0.32,
      fontFace: FONT_DISPLAY, fontSize: 17, bold: true,
      color: C.ink, valign: "top", margin: 0,
    });
    s.addText(step.d, {
      x: PAD_X + 0.95, y: y + 0.32, w: CONTENT_W - 0.95, h: 0.36,
      fontFace: FONT_TEXT, fontSize: 13,
      color: C.inkMuted80, valign: "top", margin: 0,
    });
  });

  pageNumber(s, 13, TOTAL, "light");
}

// ----- Slide 14: 마무리 ------------------------------------------------
{
  const s = pres.addSlide();
  bg(s, "light");
  pillTag(s, "정리", "light");
  s.addText([
    { text: "모든 것은 ", options: { bold: true, color: C.ink } },
    { text: ".md", options: { bold: true, color: C.primary } },
    { text: "에서", options: { bold: true, color: C.ink, breakLine: true } },
    { text: "시작합니다.", options: { bold: true, color: C.ink } },
  ], {
    x: 0, y: 2.0, w: W, h: 2.4,
    fontFace: FONT_DISPLAY, fontSize: 76,
    align: "center", valign: "top", margin: 0,
  });
  s.addText("오늘부터, 노트 한 장을 .md로 — 미래의 모든 형식이 그 안에 들어갑니다.", {
    x: 0, y: 5.0, w: W, h: 0.6,
    fontFace: FONT_TEXT, fontSize: 22,
    color: C.inkMuted80, align: "center", valign: "top", margin: 0,
  });
  s.addText("감사합니다 · 질의응답", {
    x: 0, y: H - 1.2, w: W, h: 0.4,
    fontFace: FONT_TEXT, fontSize: 12, charSpacing: 4,
    color: C.inkMuted48, align: "center", margin: 0,
  });
  pageNumber(s, 14, TOTAL, "light");
}

// === Save ===============================================================
pres.writeFile({ fileName: "lecture.pptx" })
  .then(file => console.log("Saved:", file));
