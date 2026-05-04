// Build devotional.pptx — Figma-style 16-slide deck for 5-day devotional
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" × 7.5"
const W = 13.3, H = 7.5;
pres.title = "한 가족이 되신 예수님 — 5일 묵상";
pres.author = "csh";

// === Tokens ============================================================
const C = {
  primary: "000000",
  canvas: "FFFFFF",
  ink: "000000",
  inverseInk: "FFFFFF",
  surfaceSoft: "F5F5F5",
  hairline: "E5E5E5",
  blockCream: "F5EFE0",
  blockMint: "D4ECD8",
  blockLilac: "E0D9F2",
  blockLime: "DBE89A",
  blockCoral: "F5C5B8",
  blockNavy: "1F2451",
  accentMagenta: "FF2D87",
};

const FONT_DISPLAY = "Baskerville"; // serif italic for headlines
const FONT_SANS = "Apple SD Gothic Neo";
const FONT_MONO = "Menlo";

const PAD_X = 0.85;
const PAD_TOP = 0.7;
const CONTENT_W = W - 2 * PAD_X;

function bg(slide, color) { slide.background = { color }; }

function eyebrow(slide, text, opts = {}) {
  const { y = PAD_TOP, color = C.ink } = opts;
  slide.addText(text, {
    x: PAD_X, y, w: CONTENT_W, h: 0.35,
    fontFace: FONT_MONO, fontSize: 11, bold: true,
    charSpacing: 5, color,
    margin: 0,
  });
}

function progressBar(slide, n, total, isDark = false) {
  // Top bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.04,
    fill: { color: isDark ? "333333" : "EEEEEE" },
    line: { type: "none" },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: (W * n) / total, h: 0.04,
    fill: { color: C.accentMagenta },
    line: { type: "none" },
  });
  // Page number
  const c = isDark ? "888888" : "888888";
  slide.addText(`${String(n).padStart(2, "0")} / ${String(total).padStart(2, "0")}`, {
    x: W - 1.5, y: H - 0.5, w: 1.2, h: 0.3,
    fontFace: FONT_MONO, fontSize: 10, color: c, align: "right",
    charSpacing: 4, margin: 0,
  });
  slide.addText("한 가족이 되신 예수님 · 5일 묵상", {
    x: PAD_X, y: H - 0.5, w: 6, h: 0.3,
    fontFace: FONT_MONO, fontSize: 10, color: c, charSpacing: 4,
    margin: 0,
  });
}

// Day-tile helper (used in series intro)
function dayTile(slide, x, y, w, h, num, scrip, ttl, blockColor, textColor = C.ink) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: blockColor }, line: { type: "none" },
    rectRadius: 0.16,
  });
  slide.addText(num, {
    x: x + 0.25, y: y + 0.25, w: w - 0.5, h: 0.3,
    fontFace: FONT_MONO, fontSize: 9, bold: true, charSpacing: 4,
    color: textColor, margin: 0,
  });
  slide.addText(scrip, {
    x: x + 0.25, y: y + 0.55, w: w - 0.5, h: 0.3,
    fontFace: FONT_MONO, fontSize: 9, charSpacing: 3,
    color: textColor, margin: 0,
  });
  slide.addText(ttl, {
    x: x + 0.25, y: y + h - 1.0, w: w - 0.5, h: 0.9,
    fontFace: FONT_DISPLAY, fontSize: 18, italic: true,
    color: textColor, valign: "bottom", margin: 0,
  });
}

function scriptureBlock(slide, x, y, w, text, cite, textColor = C.ink) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.04, h: 1.4,
    fill: { color: textColor }, line: { type: "none" },
  });
  slide.addText(`"${text}"`, {
    x: x + 0.18, y, w: w - 0.18, h: 1.0,
    fontFace: FONT_DISPLAY, fontSize: 18, italic: true,
    color: textColor, valign: "top", margin: 0,
    paraSpaceAfter: 4,
  });
  slide.addText(cite, {
    x: x + 0.18, y: y + 1.05, w: w - 0.18, h: 0.3,
    fontFace: FONT_MONO, fontSize: 10, charSpacing: 4,
    color: textColor, margin: 0,
  });
}

const TOTAL = 16;

// =============== Slide 1: Series intro ===============
{
  const s = pres.addSlide();
  bg(s, C.canvas);
  eyebrow(s, "● 사도신경 강해 10 · 죽으시고, 장사된 지");
  s.addText([
    { text: "한 가족이 되신 예수님,", options: { italic: true, breakLine: true } },
    { text: "다섯 날의 묵상.", options: { italic: true } },
  ], {
    x: PAD_X, y: PAD_TOP + 0.5, w: CONTENT_W, h: 1.8,
    fontFace: FONT_DISPLAY, fontSize: 56, color: C.ink,
    valign: "top", margin: 0,
  });
  s.addText("매일 한 편씩 — 본문 한 구절, 핵심 한 문장, 짧은 묵상글과 기도. 한 주를 함께 살아갑니다.", {
    x: PAD_X, y: 3.0, w: CONTENT_W * 0.7, h: 0.5,
    fontFace: FONT_SANS, fontSize: 17, color: C.ink, margin: 0,
  });

  const tileW = (CONTENT_W - 0.5) / 5;
  const tileH = 2.5;
  const tileY = 4.0;
  const tiles = [
    { num: "DAY 01", scrip: "고후 8:9", ttl: "한 가족이 되신 예수님.", color: C.blockCream },
    { num: "DAY 02", scrip: "고후 8:9", ttl: "가난과 부요의 교환.", color: C.blockMint },
    { num: "DAY 03", scrip: "갈 3:13-14", ttl: "저주를 짊어지신 예수님.", color: C.blockLilac },
    { num: "DAY 04", scrip: "고후 5:21", ttl: "의의 전가, 새로운 신분.", color: C.blockLime },
    { num: "DAY 05", scrip: "롬 8:38-39", ttl: "죽음조차 끊지 못하는 사랑.", color: C.blockCoral },
  ];
  tiles.forEach((t, i) => {
    dayTile(s, PAD_X + i * (tileW + 0.125), tileY, tileW, tileH, t.num, t.scrip, t.ttl, t.color);
  });

  progressBar(s, 1, TOTAL);
}

// =============== Day-builder helper ===============
function buildDay({ num, label, ttl1, ttl2, subhead, scriptText, scriptCite, keyPull, medP1, medP2, appHeadline, q1, q2, prayer, dayColor, mediColor = C.canvas, applyColor }) {
  const totalSlideStart = 1 + (num - 1) * 3 + 1; // 2, 5, 8, 11, 14

  // Slide A: Hero (color block)
  {
    const s = pres.addSlide();
    bg(s, dayColor);
    eyebrow(s, `● DAY 0${num} · ${label}`);
    s.addText([
      { text: ttl1, options: { italic: true, breakLine: true } },
      { text: ttl2, options: { italic: true } },
    ], {
      x: PAD_X, y: PAD_TOP + 0.5, w: CONTENT_W, h: 2.4,
      fontFace: FONT_DISPLAY, fontSize: 80, color: C.ink,
      valign: "top", margin: 0,
    });
    s.addText(subhead, {
      x: PAD_X, y: 4.5, w: CONTENT_W * 0.75, h: 0.8,
      fontFace: FONT_DISPLAY, fontSize: 22, italic: true,
      color: C.ink, valign: "top", margin: 0,
    });
    scriptureBlock(s, PAD_X, 5.4, CONTENT_W * 0.85, scriptText, scriptCite);
    progressBar(s, totalSlideStart, TOTAL);
  }

  // Slide B: Meditation (white canvas)
  {
    const s = pres.addSlide();
    bg(s, mediColor);
    eyebrow(s, `DAY 0${num} · 묵상`);
    // Two-line key pull
    const [keyA, keyB] = keyPull;
    s.addText([
      { text: keyA, options: { italic: true, color: C.ink, breakLine: true } },
      { text: keyB, options: { italic: true, color: C.ink } },
    ], {
      x: PAD_X, y: PAD_TOP + 0.5, w: CONTENT_W, h: 2.0,
      fontFace: FONT_DISPLAY, fontSize: 48,
      valign: "top", margin: 0,
    });
    // Meditation paragraphs
    s.addText(medP1, {
      x: PAD_X, y: 3.7, w: CONTENT_W * 0.85, h: 1.3,
      fontFace: FONT_SANS, fontSize: 17, color: C.ink,
      valign: "top", margin: 0, paraSpaceAfter: 8,
      lineSpacingMultiple: 1.55,
    });
    s.addText(medP2, {
      x: PAD_X, y: 5.05, w: CONTENT_W * 0.85, h: 1.6,
      fontFace: FONT_SANS, fontSize: 17, color: C.ink,
      valign: "top", margin: 0,
      lineSpacingMultiple: 1.55,
    });
    progressBar(s, totalSlideStart + 1, TOTAL);
  }

  // Slide C: Application + Prayer (color block)
  {
    const s = pres.addSlide();
    bg(s, applyColor);
    eyebrow(s, `DAY 0${num} · 적용 + 기도`);
    s.addText(appHeadline, {
      x: PAD_X, y: PAD_TOP + 0.5, w: CONTENT_W, h: 1.4,
      fontFace: FONT_DISPLAY, fontSize: 38, italic: true,
      color: C.ink, valign: "top", margin: 0,
    });
    // Question cards
    const qY = 2.7;
    const qW = CONTENT_W * 0.85;
    [q1, q2].forEach((q, i) => {
      const y = qY + i * 1.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: PAD_X, y, w: qW, h: 0.95,
        fill: { color: "FFFFFF", transparency: 35 },
        line: { type: "none" }, rectRadius: 0.08,
      });
      s.addText(`0${i + 1}`, {
        x: PAD_X + 0.25, y: y + 0.2, w: 0.5, h: 0.55,
        fontFace: FONT_MONO, fontSize: 12, bold: true, color: C.ink,
        valign: "middle", margin: 0,
      });
      s.addText(q, {
        x: PAD_X + 0.85, y: y + 0.15, w: qW - 1.05, h: 0.7,
        fontFace: FONT_SANS, fontSize: 14, color: C.ink,
        valign: "middle", margin: 0,
        lineSpacingMultiple: 1.45,
      });
    });
    // Prayer
    s.addText(prayer, {
      x: PAD_X, y: 5.3, w: CONTENT_W * 0.85, h: 1.4,
      fontFace: FONT_DISPLAY, fontSize: 18, italic: true,
      color: C.ink, valign: "top", margin: 0,
      paraSpaceAfter: 4,
      lineSpacingMultiple: 1.55,
    });
    s.addText("— 아멘.", {
      x: PAD_X, y: H - 0.95, w: 2, h: 0.35,
      fontFace: FONT_MONO, fontSize: 11, charSpacing: 5,
      color: C.ink, margin: 0,
    });
    progressBar(s, totalSlideStart + 2, TOTAL);
  }
}

// =============== DAY 1 ===============
buildDay({
  num: 1,
  label: "한 가족 되심",
  ttl1: "한 가족이",
  ttl2: "되신 예수님.",
  subhead: "예수님은 우리를 도와주신 게 아니라, 우리와 한 가족이 되셨습니다.",
  scriptText: "그리스도께서는 부요하나, 여러분을 위해서 가난하게 되셨습니다. 그것은 그의 가난으로 여러분을 부요하게 하시려는 것입니다.",
  scriptCite: "고린도후서 8:9",
  keyPull: ["가족이 된다는 건,", "모든 것을 함께 나눈다는 뜻입니다."],
  medP1: "부모님 카드를 써본 적 있나요? 그 카드 속 돈은 부모님이 번 것이지만, 가족이라는 이유 하나로 내가 쓸 수 있습니다. 가족이 된다는 건 그런 겁니다 — 내가 일하지 않은 것까지 함께 누리는 것.",
  medP2: "예수님이 우리에게 오신 방식이 그렇습니다. 좋은 것만 주시는 게 아니라, 우리의 모든 것 — 가난도, 저주도, 죄도 — 자기 것으로 가져가시고, 그분의 부요와 축복과 의로움을 우리에게 선물로 주셨습니다. 이것이 \"그리스도와의 연합\"입니다.",
  appHeadline: "한 가지만 골라, 오늘 마음에 두기.",
  q1: "오늘 내가 예수님께 \"가족이라서\" 가져가시도록 맡길 수 있는 한 가지는 무엇입니까?",
  q2: "\"가족이 되어 주셨다\"는 표현을 들었을 때, 가장 먼저 떠오르는 장면이나 느낌은?",
  prayer: "예수님, 저를 돕는 분이 아니라 한 가족이 되어 주셔서 감사합니다.\n오늘 제 짐을 혼자 지지 않고, 당신께 가져가게 하소서.\n가족으로 함께 살아가는 하루가 되게 하소서.",
  dayColor: C.blockCream,
  applyColor: C.blockMint,
});

// =============== DAY 2 ===============
buildDay({
  num: 2,
  label: "부요와 가난의 교환",
  ttl1: "가난과",
  ttl2: "부요의 교환.",
  subhead: "그분이 가난해지셨기에, 나는 부요해졌습니다.",
  scriptText: "그리스도께서는 부요하나, 여러분을 위해서 가난하게 되셨습니다. 그것은 그의 가난으로 여러분을 부요하게 하시려는 것입니다.",
  scriptCite: "고린도후서 8:9",
  keyPull: ["도와주신 게 아니라,", "내 자리에 직접 내려오셨습니다."],
  medP1: "예수님은 후원자가 아니라 우리 자리에 직접 내려오신 분입니다. 그분이 가난하셨기에 우리는 부요해졌고, 그분이 낮아지셨기에 우리는 높여졌습니다.",
  medP2: "가난은 단지 돈이 없는 것만 아닙니다. 사랑받지 못함, 인정받지 못함, 안전감의 부재 — 이 모든 가난을 그분이 먼저 지나가셨습니다. 그래서 그분은 그 자리에서 손을 내미십니다 — \"나도 거기 있었다\"고.",
  appHeadline: "내 안의 \"가난\"을 한 단어로 적어보기.",
  q1: "내가 지금 느끼는 \"가난\"은 무엇입니까? 그것이 예수님이 이미 지나가신 자리임을 떠올려 보세요.",
  q2: "오늘 누군가에게 \"내가 가진 부요\"를 작은 형태로라도 나눌 수 있는 한 가지는?",
  prayer: "예수님, 저를 위해 가난해지신 그 사랑을 묵상합니다.\n오늘 제 결핍의 자리에 당신이 먼저 와 계셨음을 기억하게 하소서.\n그리고 저도 누군가에게 작은 부요를 흘려보내게 하소서.",
  dayColor: C.blockMint,
  applyColor: C.blockLilac,
});

// =============== DAY 3 ===============
buildDay({
  num: 3,
  label: "저주의 자리",
  ttl1: "저주를",
  ttl2: "짊어지신 예수님.",
  subhead: "가장 큰 저주의 자리에 그분이 가셨기에, 가장 큰 축복이 내게 흘러옵니다.",
  scriptText: "그리스도께서 우리를 위하여 저주를 받은 사람이 되심으로써, 우리를 율법의 저주에서 속량해 주셨습니다.",
  scriptCite: "갈라디아서 3:13",
  keyPull: ["예수님은 두려운 자리에", "먼저 가신 분입니다."],
  medP1: "십자가는 단순한 처형 도구가 아니었습니다. 당시 사람들에게 그것은 \"저 사람은 하나님께 버림받았다\"라는 가장 큰 모욕이자 저주였습니다.",
  medP2: "그 자리에 예수님이 가셨습니다. 자기 잘못이 하나도 없으셨는데도 — 우리가 받아야 할 저주를 자기 등에 지셨습니다. 우리가 두려워하는 것 — 거절, 실패, 버림받음 — 그 모든 그늘을 그분이 먼저 통과하셨기에, 우리는 이제 정죄가 아닌 축복 안에서 살게 됩니다.",
  appHeadline: "내 안의 \"버림받음\"의 그림자를 마주하기.",
  q1: "내가 두려워하는 \"버림받음\"의 형태는 무엇입니까? 그 자리에도 예수님이 먼저 가셨다는 사실이 위로가 됩니까?",
  q2: "오늘 내가 누군가의 \"저주\" — 외로움이나 수치 — 를 조금이라도 함께 짊어질 수 있다면 무엇입니까?",
  prayer: "예수님, 가장 깊은 어둠의 자리에 먼저 가셔서 감사합니다.\n제가 두려워하는 곳마다 당신의 발자국이 먼저 새겨져 있음을 보게 하소서.\n오늘 그 발자국을 따라 한 걸음 내딛게 하소서.",
  dayColor: C.blockLilac,
  applyColor: C.blockCoral,
});

// =============== DAY 4 ===============
buildDay({
  num: 4,
  label: "새 신분",
  ttl1: "의의 전가,",
  ttl2: "새로운 신분.",
  subhead: "내가 의로워진 게 아니라, 그분의 의로움이 내게 입혀졌습니다.",
  scriptText: "하나님께서는 죄를 모르시는 분에게 우리 대신으로 죄를 씌우셨습니다. 그것은 우리가 그리스도 안에서 하나님의 의가 되게 하시려는 것입니다.",
  scriptCite: "고린도후서 5:21",
  keyPull: ["하나님은 내가 한 일이 아니라,", "예수님이 하신 일을 보십니다."],
  medP1: "\"내가 노력해서 깨끗해지면 하나님이 받아주시겠지\" — 우리는 자주 이렇게 생각합니다. 그러나 본문은 정반대로 말합니다.",
  medP2: "죄가 전혀 없으신 분이 우리 자리에 죄인으로 서셨고, 그래서 죄 많은 우리가 그분 안에서 의로운 사람이 되었습니다. 이것은 거래가 아니라 가족의 교환입니다. 하나님은 이제 우리를 볼 때, 우리가 한 일이 아니라 예수님이 하신 일을 보십니다. 이 신분의 변화 위에 우리의 매일이 세워집니다.",
  appHeadline: "\"내 의로움\"을 내려놓는 한 순간.",
  q1: "오늘 내가 \"내 의로움\"으로 평가받지 않고 \"그분의 의로움\" 안에서 산다면, 무엇이 가장 가벼워질까요?",
  q2: "누군가를 그의 행동이 아니라 \"예수님 안의 신분\"으로 봐주는 한 번의 시선을 가져 보세요.",
  prayer: "예수님, 제가 한 일이 아니라 당신이 하신 일을 의지하게 하소서.\n오늘 누군가를 평가가 아닌 사랑의 시선으로 바라보게 하소서.\n당신 안의 새 신분으로 살게 하소서.",
  dayColor: C.blockLime,
  applyColor: C.blockCream,
});

// =============== DAY 5 ===============
buildDay({
  num: 5,
  label: "끊어지지 않는 사랑",
  ttl1: "죽음조차",
  ttl2: "끊지 못하는 사랑.",
  subhead: "그분이 죽음의 자리까지 가셨기에, 그곳도 더는 우리를 끊지 못합니다.",
  scriptText: "죽음도, 삶도… 그 어떤 피조물도, 우리를 우리 주 예수 그리스도 안에 있는 하나님의 사랑에서 끊을 수 없습니다.",
  scriptCite: "로마서 8:38-39",
  keyPull: ["예수님이 먼저 가신 자리는,", "더 이상 우리를 끊지 못합니다."],
  medP1: "예수님의 죽음은 연극이 아니었습니다. 진짜로 심장이 멈추고, 무덤에 묻히신 사건입니다. 우리와 완전히 하나가 되시기 위해서였습니다.",
  medP2: "외로움, 실패, 불안, 두려움 — 우리가 가는 모든 자리에 예수님은 먼저 가셨고, 거기서 우리에게 손을 내미십니다. 죽음조차도 그분의 사랑에서 우리를 끊지 못합니다. 이 한 주의 묵상이 끝이 아니라, 일상의 시작이 되기를.",
  appHeadline: "한 주의 묵상에서 한 가지를 챙겨가기.",
  q1: "이번 한 주 묵상을 통해 \"예수님이 먼저 가신 자리\"라는 표현이 가장 깊이 와닿은 순간은?",
  q2: "다음 한 주, 어떤 한 가지 영역에서 이 진리를 살아내고 싶습니까?",
  prayer: "예수님, 죽음의 자리까지 가신 그 사랑이 제 매일의 자리가 되게 하소서.\n이 한 주 묵상이 끝이 아니라, 일상의 시작이 되게 하소서.",
  dayColor: C.blockCoral,
  applyColor: C.blockLime,
});

// === Save ===============================================================
pres.writeFile({ fileName: "/Users/csh/Projects/claude-code-md-explainer/devotion-10/devotional.pptx" })
  .then(file => console.log("Saved:", file));
