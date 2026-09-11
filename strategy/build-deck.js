const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.3 x 7.5
pres.author = "長路 亮太";
pres.company = "株式会社アルソア長路";
pres.title = "ワンバー宣言｜クイーンシルバー再生提案";

const W = 13.3, H = 7.5, M = 0.7;

// ---- palette (ARSOA長路ブランド: navy / gold / crimson / silver) ----
const NAVY   = "1A2744";
const NAVY_D = "111B31";
const NAVY_L = "243460";
const GOLD   = "B8973E";
const GOLD_L = "D4AF50";
const CRIM   = "C0392B";
const WHITE  = "FFFFFF";
const SILVER = "EFF1F4";
const LINE   = "DCE0E6";
const TXT    = "1C1F26";
const TXT_M  = "4E545E";
const TXT_L  = "8B919B";

const JP  = "Meiryo";
const NUM = "Arial";

const sh = (o = {}) => Object.assign({ type: "outer", angle: 90, blur: 10, offset: 2, color: "9AA0A8", opacity: 0.25 }, o);

// ---------- helpers ----------
function darkBase(s) {
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, { x: 10.6, y: -2.2, w: 5.6, h: 5.6, fill: { color: NAVY_L }, line: { color: NAVY_L } });
  s.addShape(pres.ShapeType.ellipse, { x: -1.9, y: 5.0, w: 4.4, h: 4.4, fill: { color: NAVY_L }, line: { color: NAVY_L } });
}

function head(s, eyebrow, title, lead, dark) {
  const tc = dark ? WHITE : NAVY;
  s.addText(eyebrow, { x: M, y: 0.46, w: 11.9, h: 0.3, fontSize: 11, bold: true, color: dark ? GOLD_L : GOLD, fontFace: JP, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText(title, { x: M, y: 0.80, w: 11.9, h: 0.74, fontSize: 33, bold: true, color: tc, fontFace: JP, isTextBox: true, margin: 0, valign: "top" });
  if (lead) s.addText(lead, { x: M, y: 1.58, w: 11.3, h: 0.6, fontSize: 14, color: dark ? "C4CBD8" : TXT_M, fontFace: JP, lineSpacing: 22, isTextBox: true, margin: 0 });
}

function card(s, x, y, w, h, opt = {}) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: opt.fill || WHITE },
    line: { color: opt.line || LINE, width: 1 },
    shadow: opt.shadow === false ? undefined : sh(),
  });
}

function badge(s, x, y, label, opt = {}) {
  const d = opt.d || 0.52;
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: opt.fill || GOLD }, line: { color: opt.fill || GOLD } });
  s.addText(label, { x, y, w: d, h: d, fontSize: opt.fs || 15, bold: true, color: WHITE, fontFace: NUM, align: "center", valign: "middle", isTextBox: true, margin: 0 });
}

function note(s, txt) { s.addNotes(txt); }

function foot(s, txt, dark) {
  s.addText(txt, { x: M, y: 6.92, w: 11.9, h: 0.3, fontSize: 10, color: dark ? "7E8798" : TXT_L, fontFace: JP, isTextBox: true, margin: 0 });
}

function bullets(s, items, x, y, w, opt = {}) {
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: { code: "2022" }, breakLine: i !== items.length - 1 } })), {
    x, y, w, h: opt.h || 1.6, fontSize: opt.fs || 13, color: opt.color || TXT_M, fontFace: JP,
    lineSpacing: opt.ls || 20, paraSpaceAfter: opt.psa === undefined ? 7 : opt.psa, isTextBox: true, margin: 0, valign: "top",
  });
}

/* ============================ 1. TITLE ============================ */
{
  const s = pres.addSlide();
  darkBase(s);
  s.addText("ARSOA QUEEN SILVER ／ ブランド再生提案", { x: M, y: 1.55, w: 10, h: 0.32, fontSize: 12, bold: true, color: GOLD_L, fontFace: JP, charSpacing: 3, isTextBox: true, margin: 0 });
  s.addText("ワンバー宣言", { x: M, y: 2.0, w: 11, h: 1.5, fontSize: 68, bold: true, color: WHITE, fontFace: JP, isTextBox: true, margin: 0 });
  s.addText("クイーンシルバーを、もう一度ど真ん中に。", { x: M, y: 3.62, w: 11, h: 0.55, fontSize: 22, color: GOLD_L, fontFace: JP, isTextBox: true, margin: 0 });
  s.addText("顔も、体も、家族も。石鹸ひとつで洗う暮らしへ。", { x: M, y: 4.24, w: 11, h: 0.45, fontSize: 15, color: "C4CBD8", fontFace: JP, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.35, w: 5.3, h: 1.15, rectRadius: 0.06, fill: { color: NAVY_L }, line: { color: "3A4A78", width: 1 } });
  s.addText([
    { text: "株式会社アルソア長路　常務取締役", options: { breakLine: true, fontSize: 12, color: "A9B3C6" } },
    { text: "長路 亮太", options: { fontSize: 17, bold: true, color: WHITE } },
  ], { x: M + 0.3, y: 5.5, w: 4.7, h: 0.85, fontFace: JP, isTextBox: true, margin: 0, valign: "middle", lineSpacing: 22 });
  s.addText("2026.09", { x: 6.4, y: 5.75, w: 2, h: 0.4, fontSize: 12, color: "8894AB", fontFace: NUM, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "冒頭の一言：『今日は新商品の話をしません。クイーンシルバー1個の話だけをします。』\n狙いは、商品追加ではなく“主役を戻す”提案だと最初に宣言すること。");
}

/* ============================ 2. SUMMARY ============================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "EXECUTIVE SUMMARY", "この提案の結論", "商品を足すのではなく、クイーンシルバーの「担当面積」を取り戻す。それだけで釜は回り続ける。");
  const items = [
    ["01", "問題", "客数が半分になったのではない。\nクイーンシルバーが「顔だけの脇役」に\n降格し、担当面積が減った。", "客数ではなく、面積の問題", NAVY],
    ["02", "打ち手", "既存のお客様が全身に使うだけで、\n1人あたり年間消費は約2倍。\n新規ゼロでも売上は戻せる。", "新規ゼロで 1.3〜1.5倍", GOLD],
    ["03", "壁", "障害は品質でも価格でもなく\n「もったいない」の一語。\n引き算の家計で、これを壊す。", "「もったいない」を壊す", CRIM],
  ];
  items.forEach(([n, t, b, k, c], i) => {
    const x = M + i * 4.03;
    card(s, x, 2.42, 3.73, 3.45);
    badge(s, x + 0.34, 2.76, n, { fill: c });
    s.addText(t, { x: x + 1.02, y: 2.80, w: 2.5, h: 0.45, fontSize: 19, bold: true, color: c, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(b, { x: x + 0.34, y: 3.58, w: 3.1, h: 1.5, fontSize: 13.5, color: TXT_M, fontFace: JP, lineSpacing: 23, isTextBox: true, margin: 0, valign: "top" });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.34, y: 5.14, w: 3.05, h: 0.46, rectRadius: 0.08, fill: { color: "F2F4F7" }, line: { color: "F2F4F7" } });
    s.addText(k, { x: x + 0.34, y: 5.14, w: 3.05, h: 0.46, fontSize: 12.5, bold: true, color: c, fontFace: JP, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  });
  foot(s, "※ 本資料の数値は現時点の仮置きです。BCモニターによる実測値で差し替えてください。");
  note(s, "3枚のカードだけ読み上げる。詳細は次ページ以降。\n『新規を増やす話ではありません』を強調。");
}

/* ============================ 3. 現状 ============================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "現状認識", "クイーンシルバーの売上は、全盛期の半分以下", "そして、55年間この石鹸を焚き続けてきたミクニの釜が、稼働を保てない水準に近づいている。");

  card(s, M, 2.5, 4.5, 3.6, { fill: NAVY, line: NAVY, shadow: false });
  s.addText("全盛期比", { x: M + 0.42, y: 2.80, w: 3.6, h: 0.32, fontSize: 13, color: GOLD_L, fontFace: JP, bold: true, isTextBox: true, margin: 0 });
  s.addText("50%", { x: M + 0.34, y: 3.10, w: 3.9, h: 1.5, fontSize: 80, bold: true, color: WHITE, fontFace: NUM, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("以下", { x: M + 0.42, y: 4.64, w: 3.6, h: 0.38, fontSize: 18, bold: true, color: WHITE, fontFace: JP, isTextBox: true, margin: 0 });
  s.addText("創業以来55年、アルソアの命であり続けた石鹸の現在地。", { x: M + 0.42, y: 5.12, w: 3.66, h: 0.8, fontSize: 12, color: "B9C2D2", fontFace: JP, lineSpacing: 18, isTextBox: true, margin: 0 });

  s.addChart(pres.ChartType.bar, [{ name: "売上インデックス", labels: ["全盛期", "現　在"], values: [100, 48] }], {
    x: 5.5, y: 2.55, w: 7.1, h: 3.5,
    barDir: "bar", barGapWidthPct: 55,
    chartColors: [GOLD, CRIM],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 15, dataLabelFontBold: true, dataLabelFontFace: NUM,
    showLegend: false, showTitle: false,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 14, catAxisLabelFontFace: JP, catAxisLabelFontBold: true, catAxisLineShow: false,
    valAxisLabelColor: TXT_L, valAxisLabelFontSize: 10, valAxisLineShow: false, valAxisMaxVal: 115,
    valGridLine: { color: LINE, size: 1 }, catGridLine: { style: "none" },
    plotArea: { fill: { color: WHITE } },
  });
  foot(s, "※ 指数は説明用の仮置き。millcaの実データに差し替えること。");
  note(s, "ここは事実の共有。感情を入れずに淡々と。\n『半分』という数字より、釜が止まりうるという事実のほうが重い。");
}

/* ============================ 4. 問題の再定義 ============================ */
{
  const s = pres.addSlide();
  s.background = { color: SILVER };
  head(s, "問題の再定義", "減ったのは「客数」ではなく、石鹸の「担当面積」", "ラインナップが増えるほど、クイーンシルバーが洗う場所は狭くなっていった。");

  const col = [
    { t: "かつて", sub: "洗面所の主役", c: GOLD, k: "1個で、家じゅうを洗っていた", list: ["顔も、体も、手も、家族も洗う", "浴室と洗面所の両方に置かれていた", "1世帯で年間に何個も減っていった"] },
    { t: "いま", sub: "顔だけの脇役", c: CRIM, k: "棚に4本、石鹸は端に1個", list: ["洗顔だけ。担当は顔のわずかな面積", "体はボディソープ、メイクはクレンジング", "棚にボトルが並び、石鹸は端に追いやられた"] },
  ];
  col.forEach((o, i) => {
    const x = M + i * 6.12;
    card(s, x, 2.42, 5.82, 3.45, { fill: WHITE });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.36, y: 2.78, w: 1.28, h: 0.42, rectRadius: 0.08, fill: { color: o.c }, line: { color: o.c } });
    s.addText(o.t, { x: x + 0.36, y: 2.78, w: 1.28, h: 0.42, fontSize: 13, bold: true, color: WHITE, fontFace: JP, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(o.sub, { x: x + 1.8, y: 2.73, w: 3.8, h: 0.5, fontSize: 21, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    bullets(s, o.list, x + 0.36, 3.52, 5.1, { h: 1.55, fs: 13.5, ls: 21 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.36, y: 5.14, w: 5.1, h: 0.46, rectRadius: 0.08, fill: { color: "F2F4F7" }, line: { color: "F2F4F7" } });
    s.addText(o.k, { x: x + 0.36, y: 5.14, w: 5.1, h: 0.46, fontSize: 13, bold: true, color: o.c, fontFace: JP, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  });
  s.addText("主役を取り戻す = 担当面積を取り戻す。これが本提案のすべて。", { x: M, y: 6.15, w: 11.9, h: 0.5, fontSize: 16, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "『商品が悪くなったわけでも、お客様が離れたわけでもない』とはっきり言う。\n責任の所在を人ではなく構造に置くと、BCが前を向く。");
}

/* ============================ 5. 売上の分解 ============================ */
{
  const s = pres.addSlide();
  darkBase(s);
  head(s, "構造", "売上のどこが減ったのか", "三つの掛け算に分解すると、手を打つべき場所は一つしかない。", true);
  const terms = [["人　数", "大きくは減っていない", false], ["頻　度", "1日2回のまま", false], ["1回の量", "ここが激減した", true]];
  terms.forEach(([t, d, hot], i) => {
    const x = M + i * 4.2;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.8, w: 3.5, h: 2.5, rectRadius: 0.06, fill: { color: hot ? GOLD : NAVY_L }, line: { color: hot ? GOLD : "3A4A78", width: 1 } });
    s.addText(t, { x: x + 0.25, y: 3.2, w: 3.0, h: 0.7, fontSize: 26, bold: true, color: WHITE, fontFace: JP, align: "center", isTextBox: true, margin: 0, valign: "middle" });
    s.addText(d, { x: x + 0.2, y: 4.05, w: 3.1, h: 0.9, fontSize: 13, color: hot ? "3D2F0C" : "AEB9CE", fontFace: JP, align: "center", lineSpacing: 19, isTextBox: true, margin: 0, valign: "top" });
    if (i < 2) s.addText("×", { x: x + 3.5, y: 3.75, w: 0.7, h: 0.6, fontSize: 28, bold: true, color: GOLD_L, fontFace: NUM, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  });
  s.addText("つまり、新規獲得ではなく「使用量」を戻せば、売上は戻る。", { x: M, y: 5.75, w: 11.9, h: 0.6, fontSize: 19, bold: true, color: GOLD_L, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "この式を口頭でなぞる。『人数を増やす施策は一つも入っていません』と言い切ると、現場の負担感が消える。");
}

/* ============================ 6. 試算 ============================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "試算", "全身に使うだけで、1人あたりの消費は約2倍になる", "顔だけなら1個で約2か月。全身に使えば3〜4週間。家族で共有すればさらに短くなる。");
  s.addChart(pres.ChartType.bar, [{ name: "1人あたり年間消費個数", labels: ["顔だけ（現状）", "顔＋全身", "＋家族で共有"], values: [6, 12, 15] }], {
    x: M - 0.1, y: 2.45, w: 7.5, h: 3.6,
    barDir: "col", barGapWidthPct: 60,
    chartColors: [TXT_L, GOLD, NAVY],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TXT, dataLabelFontSize: 16, dataLabelFontBold: true, dataLabelFontFace: NUM,
    showLegend: false, showTitle: false,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 12, catAxisLabelFontFace: JP, catAxisLabelFontBold: true, catAxisLineShow: false,
    valAxisLabelColor: TXT_L, valAxisLabelFontSize: 10, valAxisLineShow: false, valAxisMaxVal: 18,
    valGridLine: { color: LINE, size: 1 }, catGridLine: { style: "none" },
    plotArea: { fill: { color: WHITE } },
  });
  card(s, 8.4, 2.5, 4.2, 1.7, { fill: NAVY, line: NAVY, shadow: false });
  s.addText("既存客の半数が移行した場合", { x: 8.72, y: 2.74, w: 3.6, h: 0.35, fontSize: 12, bold: true, color: GOLD_L, fontFace: JP, isTextBox: true, margin: 0 });
  s.addText("1.3〜1.5倍", { x: 8.66, y: 3.12, w: 3.8, h: 0.85, fontSize: 40, bold: true, color: WHITE, fontFace: NUM, isTextBox: true, margin: 0, valign: "middle" });
  card(s, 8.4, 4.38, 4.2, 1.7);
  s.addText("新規獲得", { x: 8.72, y: 4.62, w: 3.6, h: 0.35, fontSize: 12, bold: true, color: CRIM, fontFace: JP, isTextBox: true, margin: 0 });
  s.addText("0人でよい", { x: 8.66, y: 5.0, w: 3.8, h: 0.85, fontSize: 36, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  foot(s, "※ 消費ペースは仮定値。BC10名×30日のモニターで実測し、自社データに差し替えること（本提案の最重要アクション）。");
  note(s, "ここが提案の心臓部。数字が仮置きであることは正直に言い、だからこそ最初の30日で実測すると続ける。");
}

/* ============================ 7. 壁：もったいない ============================ */
{
  const s = pres.addSlide();
  s.background = { color: SILVER };
  head(s, "最大の障害", "壁は価格ではない。「もったいない」という一語だ", "高価な石鹸で体を洗うことへの抵抗は、足し算で考えている限り絶対に消えない。置き換えの総額で語る。");

  card(s, M, 2.45, 5.4, 3.35, { fill: WHITE });
  s.addText("いまの洗面所と浴室", { x: M + 0.36, y: 2.75, w: 4.6, h: 0.45, fontSize: 17, bold: true, color: TXT, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  bullets(s, ["クレンジング", "洗顔料", "ボディソープ", "ハンドソープ"], M + 0.36, 3.38, 4.6, { h: 1.9, fs: 14, ls: 22, psa: 9 });
  s.addText("ボトル4本ぶんの出費", { x: M + 0.36, y: 5.2, w: 4.6, h: 0.4, fontSize: 14, bold: true, color: CRIM, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });

  s.addText("→", { x: 6.28, y: 3.82, w: 0.8, h: 0.6, fontSize: 30, bold: true, color: GOLD, fontFace: NUM, align: "center", valign: "middle", isTextBox: true, margin: 0 });

  card(s, 7.2, 2.45, 5.4, 3.35, { fill: NAVY, line: NAVY, shadow: false });
  s.addText("ワンバー生活", { x: 7.56, y: 2.75, w: 4.6, h: 0.45, fontSize: 17, bold: true, color: WHITE, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("クイーンシルバー　1個", { x: 7.5, y: 3.45, w: 4.8, h: 0.9, fontSize: 27, bold: true, color: GOLD_L, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("4本を1個に置き換える。だから\n「もったいない」ではなく「まとまる」。", { x: 7.56, y: 4.5, w: 4.6, h: 1.0, fontSize: 13.5, color: "C4CBD8", fontFace: JP, lineSpacing: 22, isTextBox: true, margin: 0, valign: "top" });

  s.addText("語る順序を変える。「石鹸を買ってください」ではなく「棚のボトルを、いくつ減らせますか」。", { x: M, y: 6.05, w: 11.9, h: 0.5, fontSize: 15, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "BCが一番つまずくのがここ。『体に使うなんてもったいない』と言われた時の返し方を、この1枚で共有する。");
}

/* ============================ 8. 戦略4本柱 ============================ */
{
  const s = pres.addSlide();
  darkBase(s);
  head(s, "STRATEGY", "ムーブメントを起こす4本の柱", "宣言をつくり、見せて納得させ、物語で燃やし、現場を動かす。", true);
  const pil = [
    ["A", "宣言をつくる", "ワンバー宣言 ─ 30日間、顔も体も石鹸1個で洗う"],
    ["B", "見せて納得させる", "落ちる／落ちないの実演 ─ オイルフリーから逆算する"],
    ["C", "物語で燃やす", "ミクニ55年 ─「この釜を、止めない」"],
    ["D", "現場を動かす", "クイーンシルバー・マイスター ─ 話を1本に絞る"],
  ];
  pil.forEach(([n, t, d], i) => {
    const x = M + (i % 2) * 6.12, y = 2.6 + Math.floor(i / 2) * 1.95;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.82, h: 1.65, rectRadius: 0.06, fill: { color: NAVY_L }, line: { color: "36456F", width: 1 } });
    badge(s, x + 0.34, y + 0.33, n, { fill: GOLD, fs: 16 });
    s.addText(t, { x: x + 1.04, y: y + 0.28, w: 4.4, h: 0.45, fontSize: 18, bold: true, color: WHITE, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(d, { x: x + 1.04, y: y + 0.82, w: 4.55, h: 0.62, fontSize: 12.5, color: "AEB9CE", fontFace: JP, lineSpacing: 18, isTextBox: true, margin: 0, valign: "top" });
  });
  note(s, "4本を一度に走らせる。AとBが集客と納得、Cが燃料、Dが実行部隊。");
}

/* ============================ 9. 施策A ============================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "施策 A", "ワンバー宣言 ─ 洗面所から、ボトルを消す", "30日間、顔・体・手をクイーンシルバー1個だけで洗う。参加者がSNSに上げるのは、肌ではなく「棚の写真」。");
  const rows = [
    ["01", "撮るのは肌ではなく、棚", "ボトルが5本並んだ棚 → 石鹸1個だけの棚。肌の変化は個人差があり表現も制限されるが、棚の写真は誰も否定できず、真似したくなる。"],
    ["02", "置き場所を先に用意する", "受け皿と泡立てネットをセットで配布。浴室に置けなければ習慣は変わらない。「洗面所に1個、浴室に1個」の2個持ちを標準に。"],
    ["03", "合言葉をひとつだけ決める", "#ワンバー生活　#洗面所が空になった　── ハッシュタグは2つまで。増やすほど束にならない。"],
  ];
  rows.forEach(([n, t, d], i) => {
    const y = 2.45 + i * 1.36;
    card(s, M, y, 11.9, 1.18, { fill: i % 2 ? WHITE : SILVER });
    badge(s, M + 0.32, y + 0.33, n, { fill: NAVY, fs: 14 });
    s.addText(t, { x: M + 1.02, y: y + 0.14, w: 3.9, h: 0.45, fontSize: 15.5, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(d, { x: M + 4.25, y: y + 0.16, w: 7.35, h: 0.9, fontSize: 12, color: TXT_M, fontFace: JP, lineSpacing: 18, isTextBox: true, margin: 0, valign: "middle" });
  });
  foot(s, "肌の変化・効能を断定する表現は使わない（薬機法）。訴求は「本数が減る」という事実に限定する。");
  note(s, "棚のビフォーアフターという切り口がこの施策の肝。肌写真は炎上と法務の両方のリスクがある。");
}

/* ============================ 10. 施策B ============================ */
{
  const s = pres.addSlide();
  s.background = { color: SILVER };
  head(s, "施策 B", "落ちるかどうかで、メイクを選ぶ", "血管を見せて納得させてきたやり方を、そのまま石鹸に持ち込む。長路の武器は「可視化」。");
  const c = [
    ["実演する", "腕に2本ラインを引き、泡で洗い流す。片方は落ち、片方は残る。説明ではなく、その場で見せる。", "まず1本、長路の腕で撮る", GOLD],
    ["話法を逆にする", "「アルソアのメイクを買ってください」ではなく「クレンジングを卒業したいなら、落ちる側のメイクに替えるしかない」。", "主語は「クレンジング卒業」", NAVY],
    ["境界線を守る", "他社名・他社商品が特定できる表現は出さない。「一般的なオイルベース／フィルムタイプ」と表記する。", "表記は「一般的なタイプ」で統一", CRIM],
  ];
  c.forEach(([t, d, k, col], i) => {
    const x = M + i * 4.03;
    card(s, x, 2.42, 3.73, 3.25, { fill: WHITE });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.34, y: 2.75, w: 0.16, h: 0.16, rectRadius: 0.08, fill: { color: col }, line: { color: col } });
    s.addText(t, { x: x + 0.66, y: 2.63, w: 2.9, h: 0.45, fontSize: 17, bold: true, color: col, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(d, { x: x + 0.34, y: 3.28, w: 3.1, h: 1.6, fontSize: 13, color: TXT_M, fontFace: JP, lineSpacing: 21, isTextBox: true, margin: 0, valign: "top" });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.34, y: 4.98, w: 3.05, h: 0.46, rectRadius: 0.08, fill: { color: "F2F4F7" }, line: { color: "F2F4F7" } });
    s.addText(k, { x: x + 0.34, y: 4.98, w: 3.05, h: 0.46, fontSize: 12, bold: true, color: col, fontFace: JP, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  });
  s.addText("落ちない側を悪者にしない。「落ちにくいものは、落とすために強いものが要る」という因果だけを語る。", { x: M, y: 5.95, w: 11.9, h: 0.5, fontSize: 14.5, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  foot(s, "前提：アルソアのポイントメイクが油剤不使用であることを、商品部に事実確認すること。ここが崩れると施策全体が崩れる。");
  note(s, "実験動画は長路さん自身の腕で撮る。人柄と一貫性が説得力になる。");
}

/* ============================ 11. 施策C ============================ */
{
  const s = pres.addSlide();
  darkBase(s);
  head(s, "施策 C", "この釜を、止めない", "ムーブメントには顔と締切が要る。55年、この石鹸を焚き続けてきた工場の話を、正面から語る。", true);
  const c = [
    ["語るのは成分ではなく、人", "職人の手、釜、55年ぶんの道具を撮る。BCが伝えるべきは処方や原価ではなく、そこで働いてきた人の話。"],
    ["ゴールを見える化する", "「年◯万個で釜は回り続ける」というメーターを公開し、毎月更新する。数字が動くと、人は参加する。"],
  ];
  c.forEach(([t, d], i) => {
    const x = M + i * 6.12;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.55, w: 5.82, h: 1.95, rectRadius: 0.06, fill: { color: NAVY_L }, line: { color: "36456F", width: 1 } });
    s.addText(t, { x: x + 0.36, y: 2.82, w: 5.1, h: 0.5, fontSize: 17, bold: true, color: GOLD_L, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(d, { x: x + 0.36, y: 3.42, w: 5.1, h: 0.9, fontSize: 13, color: "C4CBD8", fontFace: JP, lineSpacing: 20, isTextBox: true, margin: 0, valign: "top" });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.78, w: 11.94, h: 1.55, rectRadius: 0.06, fill: { color: "3A1614" }, line: { color: CRIM, width: 1 } });
  s.addText("前提条件：ミクニ様の合意なしに、この施策は動かせない", { x: M + 0.36, y: 5.0, w: 11.2, h: 0.42, fontSize: 16, bold: true, color: "F2A9A0", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("取引先の経営状況を外部に出す話である以上、社名・撮影・「苦境」という表現のどこまでを語ってよいか、先に握る。\n主語は「救ってください」ではなく「守りたい」。ここだけは順序を間違えると取り返しがつかない。", { x: M + 0.36, y: 5.5, w: 11.2, h: 0.72, fontSize: 12.5, color: "E8BDB8", fontFace: JP, lineSpacing: 19, isTextBox: true, margin: 0, valign: "top" });
  note(s, "最初に会いに行くのはミクニさん。ここの合意が取れるまでCは公開しない。");
}

/* ============================ 12. 施策D ============================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "施策 D", "クイーンシルバー・マイスター ─ 話を1本に絞る", "BCが売れないのは、覚えることが多すぎて話が散るから。教える内容を、ひとつに減らす。");
  const rows = [
    ["01", "90秒トーク1本だけ", "台本は1本。覚えることを増やさない。"],
    ["02", "実験キットを配る", "腕・泡立てネット・受け皿。見せられる状態で渡す。"],
    ["03", "KPIは売上ではなく人数", "「ワンバー転換人数」で競う。売上を追うと、結局セット販売の話に戻る。"],
    ["04", "自社BCで先行実証", "アルソア長路で数字を作ってから、本社・他販社へ横展開する。"],
  ];
  rows.forEach(([n, t, d], i) => {
    const x = M + (i % 2) * 6.12, y = 2.5 + Math.floor(i / 2) * 1.85;
    card(s, x, y, 5.82, 1.6, { fill: SILVER });
    badge(s, x + 0.34, y + 0.3, n, { fill: GOLD, fs: 14 });
    s.addText(t, { x: x + 1.04, y: y + 0.24, w: 4.5, h: 0.42, fontSize: 16, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(d, { x: x + 1.04, y: y + 0.74, w: 4.55, h: 0.68, fontSize: 12.5, color: TXT_M, fontFace: JP, lineSpacing: 19, isTextBox: true, margin: 0, valign: "top" });
  });
  s.addText("覚えることを1つに減らすと、現場は動く。増やすほど、誰も話さなくなる。", { x: M, y: 6.25, w: 11.9, h: 0.45, fontSize: 14.5, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "BC向けの説明会は1回60分。商品説明は一切せず、実演練習だけをやる。");
}

/* ============================ 13. ラインナップ ============================ */
{
  const s = pres.addSlide();
  s.background = { color: SILVER };
  head(s, "論点", "ラインナップ整理は「廃止」ではなく「入口に降格」", "メイクオフやボディソープを今すぐ全廃する案には、率直に反対します。主役を戻す決断のほうが、速く効く。");

  card(s, M, 2.42, 5.82, 3.5, { fill: WHITE });
  s.addText("いま全廃した場合のリスク", { x: M + 0.36, y: 2.72, w: 5.1, h: 0.45, fontSize: 16.5, bold: true, color: CRIM, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  bullets(s, ["既存愛用者が離反し、クイーンシルバーごと離れる", "BCの当座の売上が落ち、現場が動かなくなる", "他社メイクを使う新規客の受け皿が消える"], M + 0.36, 3.35, 5.1, { h: 1.8, fs: 13, ls: 21, psa: 10 });

  card(s, 7.2, 2.42, 5.82, 3.5, { fill: NAVY, line: NAVY, shadow: false });
  s.addText("代わりに：主役を戻し、他は降格", { x: 7.56, y: 2.72, w: 5.1, h: 0.45, fontSize: 16.5, bold: true, color: GOLD_L, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  bullets(s, ["カタログの一等地・BCトーク・LPの主語を全てQSに戻す", "メイクオフは「他社メイクの方向けの移行アイテム」として棚の端へ", "3か月ごとに構成比を見て、自然に縮んだものから畳む"], 7.56, 3.35, 5.1, { h: 1.8, fs: 13, ls: 21, psa: 10, color: "C4CBD8" });

  s.addShape(pres.ShapeType.roundRect, { x: M + 0.36, y: 5.24, w: 5.1, h: 0.46, rectRadius: 0.08, fill: { color: "FBF3F2" }, line: { color: "FBF3F2" } });
  s.addText("失うものが、得るものより大きい", { x: M + 0.36, y: 5.24, w: 5.1, h: 0.46, fontSize: 13, bold: true, color: CRIM, fontFace: JP, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: 7.56, y: 5.24, w: 5.1, h: 0.46, rectRadius: 0.08, fill: { color: "2A3A63" }, line: { color: "2A3A63" } });
  s.addText("3か月ごとに見直し、静かに畳む", { x: 7.56, y: 5.24, w: 5.1, h: 0.46, fontSize: 13, bold: true, color: GOLD_L, fontFace: JP, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  s.addText("それでも全廃で行くと決まれば、在庫消化とBCへの説明順序を含む撤退設計を別途組みます。", { x: M, y: 6.22, w: 11.9, h: 0.5, fontSize: 14, color: TXT_M, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "ここは唯一の“反対意見”スライド。逃げずに先に出しておくと、後の議論が早い。決めるのは長路さん。");
}

/* ============================ 14. 最初の30日 ============================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "ACTION", "最初の30日でやること", "大きな発表より先に、足元の事実を固める。ここを飛ばすと、全部が精神論になる。");
  const steps = [
    ["1", "実測する", "BC10名にワンバー生活30日。消費個数を記録し「年6個→年◯個」の自社データを作る。"],
    ["2", "会いに行く", "ミクニ様へ。語ってよい範囲の合意と、撮影の許可を取る。"],
    ["3", "裏を取る", "ポイントメイクが本当に油剤不使用か、商品部に事実確認する。"],
    ["4", "置き場所を作る", "受け皿と泡立てネットのセットを用意する。浴室に置けなければ習慣は変わらない。"],
    ["5", "1本だけ撮る", "落ちる／落ちないの実験動画を、長路が自分の腕で撮る。"],
  ];
  steps.forEach(([n, t, d], i) => {
    const y = 2.4 + i * 0.85;
    badge(s, M, y, n, { fill: i === 0 ? CRIM : NAVY, d: 0.46, fs: 14 });
    s.addText(t, { x: M + 0.66, y: y - 0.02, w: 2.7, h: 0.5, fontSize: 15.5, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(d, { x: M + 3.5, y: y - 0.04, w: 8.4, h: 0.55, fontSize: 12.5, color: TXT_M, fontFace: JP, lineSpacing: 18, isTextBox: true, margin: 0, valign: "middle" });
    if (i < steps.length - 1) s.addShape(pres.ShapeType.line, { x: M + 0.23, y: y + 0.46, w: 0, h: 0.4, line: { color: LINE, width: 1.5 } });
  });
  s.addText("1が最優先。この実測値が、以降すべての説得材料になる。", { x: M, y: 6.55, w: 11.9, h: 0.45, fontSize: 14.5, bold: true, color: CRIM, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "30日後に再度集まり、実測値を見てから全社展開の可否を決める、という進め方を提案する。");
}

/* ============================ 15. KPI ============================ */
{
  const s = pres.addSlide();
  s.background = { color: SILVER };
  head(s, "KPI", "追う数字は4つだけ", "売上を直接追わない。売上は、下の4つが動いた結果として後からついてくる。");
  const k = [
    ["ワンバー転換人数", "全身に切り替えたお客様の実数", "最重要"],
    ["1人あたり年間個数", "年6個 → 年12個へ", "目標2倍"],
    ["浴室への設置率", "2個持ちになっている世帯の比率", "習慣の指標"],
    ["定期継続率", "3か月後も続いている比率", "本物かどうか"],
  ];
  k.forEach(([t, d, tag], i) => {
    const x = M + (i % 2) * 6.12, y = 2.5 + Math.floor(i / 2) * 1.92;
    card(s, x, y, 5.82, 1.7, { fill: WHITE });
    s.addText(t, { x: x + 0.36, y: y + 0.22, w: 4.0, h: 0.45, fontSize: 17, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addShape(pres.ShapeType.roundRect, { x: x + 4.42, y: y + 0.26, w: 1.1, h: 0.36, rectRadius: 0.08, fill: { color: GOLD }, line: { color: GOLD } });
    s.addText(tag, { x: x + 4.42, y: y + 0.26, w: 1.1, h: 0.36, fontSize: 10, bold: true, color: WHITE, fontFace: JP, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(d, { x: x + 0.36, y: y + 0.78, w: 5.1, h: 0.6, fontSize: 13, color: TXT_M, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addText("売上をKPIに置いた瞬間、話は「セットで買ってください」に戻る。転換人数で競う。", { x: M, y: 6.4, w: 11.9, h: 0.45, fontSize: 14, color: TXT_M, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "月次でこの4つだけを見る会議体を作る。議題を増やさない。");
}

/* ============================ 16. 注意点 ============================ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "RISK", "越えてはいけない3本の線", "この3つを踏むと、ムーブメントは一夜で止まります。先に共有しておきます。");
  const r = [
    ["薬機法", "「洗うだけで治る」「クレンジングは肌に悪い」は言えない。語れるのは “落ちる／落ちない” と “本数が減る” という事実だけ。強いのはそこなので、制約にはならない。"],
    ["景品表示法・比較広告", "他社名や他社商品が特定できる表現を出さない。下取りキャンペーンは景品規制に触れうるため、実施前に確認する。"],
    ["取引先との合意", "ミクニ様の経営状況に触れる以上、公開範囲の合意が先。社名・写真・表現のどこまでが可か、書面で握る。"],
  ];
  r.forEach(([t, d], i) => {
    const y = 2.45 + i * 1.42;
    card(s, M, y, 11.9, 1.28, { fill: "FBF3F2", line: "EFD6D3" });
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.34, y: y + 0.34, w: 2.3, h: 0.5, rectRadius: 0.08, fill: { color: CRIM }, line: { color: CRIM } });
    s.addText(t, { x: M + 0.34, y: y + 0.34, w: 2.3, h: 0.5, fontSize: 13.5, bold: true, color: WHITE, fontFace: JP, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(d, { x: M + 2.94, y: y + 0.18, w: 8.7, h: 0.9, fontSize: 12.5, color: TXT_M, fontFace: JP, lineSpacing: 19, isTextBox: true, margin: 0, valign: "middle" });
  });
  note(s, "制約を先に言うと、BCが安心して話せるようになる。禁止ではなく“ここまでは言える”という伝え方をする。");
}

/* ============================ 17. CLOSING ============================ */
{
  const s = pres.addSlide();
  darkBase(s);
  s.addText("CLOSING", { x: M, y: 1.9, w: 10, h: 0.32, fontSize: 12, bold: true, color: GOLD_L, fontFace: JP, charSpacing: 3, isTextBox: true, margin: 0 });
  s.addText("洗面所から、\nボトルが消える。", { x: M, y: 2.35, w: 11.5, h: 2.1, fontSize: 50, bold: true, color: WHITE, fontFace: JP, lineSpacing: 62, isTextBox: true, margin: 0, valign: "top" });
  s.addText("残るのは、55年変わらない石鹸ひとつ。\nクイーンシルバーを、もう一度ど真ん中に。", { x: M, y: 4.7, w: 11.5, h: 0.9, fontSize: 17, color: GOLD_L, fontFace: JP, lineSpacing: 28, isTextBox: true, margin: 0, valign: "top" });
  s.addShape(pres.ShapeType.line, { x: M, y: 5.95, w: 3.0, h: 0, line: { color: "3A4A78", width: 1 } });
  s.addText("株式会社アルソア長路　常務取締役　長路 亮太　／　金沢市松村2-177", { x: M, y: 6.25, w: 11.5, h: 0.4, fontSize: 12, color: "8894AB", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  note(s, "締めは一言だけ。『まず30日、10人で試させてください』と具体的な依頼で終える。");
}

pres.writeFile({ fileName: "./queen-silver-onebar.pptx" })
  .then(f => console.log("written:", f));
