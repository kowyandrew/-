// さくらハウジング 9月進捗報告（改訂版）を生成するスクリプト
// 使い方: npm install && node build.js
// 数字・担当者・写真が決まったら、下の「設定値」だけを書き換えて再実行する。

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

// ============================================================
// 設定値（ここだけ書き換える）
// ============================================================

// --- 月別実績（単位：円）---
// 出典：Googleスプレッドシート「さくらハウジング_実績数値管理シート」55期タブ
// 55期 = 2026年3月〜2027年2月（3月始まり）
const MONTHS = ["3月", "4月", "5月", "6月", "7月", "8月"];
const ACTUALS = {
  total:      [2319745, 2517686, 2413512, 2287661, 2535588, 2337010], // 65行 売上利益合計（売買を除く）
  baibai:     [0, 110000, 0, 330000, 455400, 1452000],                // 69行 売買仲介収入計
  ad:         [400000, 404200, 180200, 239000, 423500, 289000],       // 51行 受取AD
  hoken:      [0, 35200, 95420, 168340, 141376, 131280],              // 26行 保険手数料
  hosho:      [6192, 14379, 19400, 14000, 41220, 48600],              // 28行 保証会社手数料
  futai:      [0, 0, 0, 0, 0, 6600],                                  // 52行 付帯商品
};

const ANNUAL_TARGET = 2500; // 今年の年間売上目標（万円）
const LAST_YEAR_ANNUAL = 2400; // 昨年の年間売上（万円）。社内の聞き取り値
const INCLUDE_BAIBAI = true; // 売買仲介を売上に含める

// --- 期間 ---
const FY_START_MONTH = 3; // 年度は3月始まり（55期 = 2026年3月〜2027年2月）
const MONTHS_IN_YEAR = 12;
const ACTION_MONTHS = [10, 11, 12]; // 施策を動かす月

// --- 人と写真 ---
// 写真は photos/ フォルダに置いてファイル名を書く。空なら「写真を入れる枠」を表示する。
const PHOTOS = {
  store: "", // 店舗・現場の写真（表紙）
  site: "", // 現地（案内・立会いなど）の写真
  usui: "", // 臼井さん
  kawakami: "", // 川上さん
};
const PART_STAFF = ["臼井さん", "川上さん"];
const PART_SCHEDULE = "週2日"; // パートさんの勤務（聞き取り。違えば書き換える）

// 空文字なら「要記入」と表示
const OWNERS = {
  "01": PART_STAFF.join("・"), // 物件入力・写真掲載
  "02": "", // 会計をDFEへ
  "03": "", // クレーム一次対応を駆けつけサービスへ
};

const OUTPUT = "さくらハウジング_進捗報告_改訂版.pptx";

// ============================================================
// 計算（表示する数字はすべてここから作る）
// ============================================================

const sum = (a) => a.reduce((x, y) => x + y, 0);
const man = (yen) => yen / 10000; // 円 → 万円
const round1 = (v) => Math.round(v * 10) / 10;
const round10 = (v) => Math.round(v / 10) * 10;

const monthsElapsed = MONTHS.length; // 6
const monthlySales = ACTUALS.total.map((v, i) => v + (INCLUDE_BAIBAI ? ACTUALS.baibai[i] : 0));
const cumulative = Math.round(man(sum(monthlySales))); // 1,676
const cumulativeExBaibai = Math.round(man(sum(ACTUALS.total))); // 1,441
const achievementPct = round1((man(sum(monthlySales)) / ANNUAL_TARGET) * 100); // 67.0
const planPacePct = round1((monthsElapsed / MONTHS_IN_YEAR) * 100); // 50.0
const aheadPt = round1(achievementPct - planPacePct); // 17.0
const monthlyAverage = Math.round(man(sum(monthlySales)) / monthsElapsed); // 279
const monthsLeft = MONTHS_IN_YEAR - monthsElapsed; // 6
const requiredMonthly = Math.round((ANNUAL_TARGET - man(sum(monthlySales))) / monthsLeft); // 137
const projected = round10((man(sum(monthlySales)) / monthsElapsed) * MONTHS_IN_YEAR); // 約3,350
const projectedExBaibai = round10((man(sum(ACTUALS.total)) / monthsElapsed) * MONTHS_IN_YEAR); // 約2,880

// 付帯系 = 保険手数料 + 保証会社手数料 + 付帯商品
const futaiMonthly = MONTHS.map((_, i) => ACTUALS.hoken[i] + ACTUALS.hosho[i] + ACTUALS.futai[i]);
const adTotal = round1(man(sum(ACTUALS.ad))); // 193.6
const adAvg = Math.round(man(sum(ACTUALS.ad)) / monthsElapsed); // 32
const futaiFirst = round1(man(futaiMonthly[0])); // 0.6
const futaiLast = round1(man(futaiMonthly[futaiMonthly.length - 1])); // 18.6

const calMonth = (offset) => ((FY_START_MONTH - 1 + offset) % 12) + 1;
const firstMonth = calMonth(0); // 3
const lastClosedMonth = calMonth(monthsElapsed - 1); // 8
const nextMonth = calMonth(monthsElapsed); // 9
const fyEndMonth = calMonth(MONTHS_IN_YEAR - 1); // 2
const [A0, A1, A2] = ACTION_MONTHS;

// 表示用の書式
const yen = (v) => String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
const pct = (v) => v.toFixed(1).replace(/\.0$/, "");
const TBD = "要記入";
const orTbd = (v) => (v && String(v).trim() ? String(v).trim() : TBD);

// ============================================================
// デザイン
// ============================================================

const FONT = "Meiryo";
const C = {
  main: "C8315F",
  mainPale: "FBE9EF",
  mainLight: "F2B8CA",
  mainMid: "E88AA8",
  blue: "3D6FB6",
  ink: "2B2D42",
  inkSoft: "3A3D55",
  sub: "6B6E80",
  green: "2E8B62",
  greenPale: "E6F4ED",
  orange: "C77700",
  orangePale: "FFF3DD",
  card: "F4F5F8",
  line: "E3E4EA",
  gray: "D5D7DF",
  white: "FFFFFF",
};
const W = 13.333;
const M = 0.6;
const CW = W - M * 2;
const TOTAL = 5;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "さくらハウジング";
pres.title = `さくらハウジング ${nextMonth}月 進捗報告`;
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

// pptxgenjs は options を書き換えるので、呼び出しごとに新しいオブジェクトを作る
const T = (opts) => ({ fontFace: FONT, color: C.ink, margin: 0, valign: "top", isTextBox: true, ...opts });
const box = (s, x, y, w, h, fill, extra = {}) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12, ...extra });

// 写真：ファイルがあれば貼る。なければ「写真を入れる枠」
function photo(s, key, x, y, w, h, label, dark = false) {
  const file = PHOTOS[key] ? path.join(__dirname, "photos", PHOTOS[key]) : "";
  if (file && fs.existsSync(file)) {
    s.addImage({ path: file, x, y, w, h, sizing: { type: "cover", w, h }, altText: label });
    return;
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: dark ? C.inkSoft : C.card },
    line: { color: dark ? C.mainLight : C.sub, width: 1.25, dashType: "dash" },
  });
  s.addText(
    [
      { text: "写真", options: { bold: true, breakLine: true } },
      { text: label, options: { fontSize: 11 } },
    ],
    T({ x, y, w, h, fontSize: 13, color: dark ? C.mainLight : C.sub, align: "center", valign: "middle" })
  );
}

function addHeader(slide, section, title) {
  slide.background = { color: C.white };
  slide.addText(section, T({ x: M, y: 0.5, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.main }));
  slide.addText(title, T({ x: M, y: 0.85, w: CW, h: 0.65, fontSize: 28, bold: true, valign: "middle" }));
}
function addPageNumber(slide, n) {
  slide.addText(`${n} / ${TOTAL}`, T({ x: W - M - 1.2, y: 6.85, w: 1.2, h: 0.25, fontSize: 11, color: C.sub, align: "right" }));
}

// ============================================================
// スライド1：表紙＋結論
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.ink };
  s.addText(`さくらハウジング　2026年${nextMonth}月 進捗報告`, T({ x: 0.8, y: 0.8, w: 7.8, h: 0.4, fontSize: 16, bold: true, color: C.mainLight }));
  s.addText(
    [
      { text: `半年で目標の${Math.round(achievementPct)}%。`, options: { breakLine: true } },
      { text: "年間3,000万円超えのペース。", options: { breakLine: true } },
      { text: "営業を現地に集中させて、", options: { breakLine: true } },
      { text: "さらに伸ばします。" },
    ],
    T({ x: 0.8, y: 1.4, w: 7.9, h: 2.9, fontSize: 34, bold: true, color: C.white })
  );
  const rows = [
    ["現状", `${firstMonth}〜${lastClosedMonth}月で ${yen(cumulative)}万円（年間目標 ${yen(ANNUAL_TARGET)}万円）`],
    ["伸び", "AD と 付帯系の手数料"],
    ["打ち手", "入力→パートさん／会計→DFE／クレーム→駆けつけ"],
  ];
  rows.forEach(([label, body], i) => {
    const y = 4.65 + i * 0.7;
    box(s, 0.8, y, 1.2, 0.44, C.main, { rectRadius: 0.08 });
    s.addText(label, T({ x: 0.8, y, w: 1.2, h: 0.44, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" }));
    s.addText(body, T({ x: 2.25, y, w: 6.5, h: 0.44, fontSize: 15, color: C.white, valign: "middle" }));
  });
  photo(s, "store", 9.2, 0.8, W - M - 9.2, 5.9, "店舗・現場", true);

  s.addNotes(
    `最初に結論です。${firstMonth}月から${lastClosedMonth}月の半年で売上は${yen(cumulative)}万円、年間目標${yen(ANNUAL_TARGET)}万円の${Math.round(achievementPct)}%まで来ました。` +
      `このペースなら年間で3,000万円を超えます。伸びているのはADと付帯系の手数料です。` +
      "この勢いを続けるために、営業が現地でしかできない仕事に集中できる仕組みをつくります。"
  );
}

// ============================================================
// スライド2：① 現状（目標・見込み・昨年のバー比較）
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "① 現状", `半年で目標の${Math.round(achievementPct)}%。年間3,000万円を超えるペース`);
  addPageNumber(s, 2);

  const lx = M, lw = 1.9; // 行ラベル
  const bx = lx + lw + 0.2, maxW = 5.2, scaleMax = 3500;
  const wOf = (v) => (maxW * v) / scaleMax;
  const rows = [
    { label: "昨年", sub: "年間実績", y: 2.0 },
    { label: "今年の目標", sub: "年間", y: 3.25 },
    { label: "今年の見込み", sub: "このペースなら", y: 4.5 },
  ];
  const bh = 0.8;
  rows.forEach((r) => {
    s.addText(
      [
        { text: r.label, options: { fontSize: 16, bold: true, breakLine: true } },
        { text: r.sub, options: { fontSize: 11, color: C.sub } },
      ],
      T({ x: lx, y: r.y, w: lw, h: bh, valign: "middle" })
    );
  });
  // 昨年
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: rows[0].y, w: wOf(LAST_YEAR_ANNUAL), h: bh, fill: { color: C.gray }, line: { color: C.gray } });
  // 目標ラインと重ならないよう、昨年の数字はバーの内側に置く
  s.addText(`${yen(LAST_YEAR_ANNUAL)}万円`, T({ x: bx + 0.15, y: rows[0].y, w: wOf(LAST_YEAR_ANNUAL) - 0.3, h: bh, fontSize: 18, bold: true, align: "right", valign: "middle" }));
  // 目標
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: rows[1].y, w: wOf(ANNUAL_TARGET), h: bh, fill: { color: C.ink }, line: { color: C.ink } });
  s.addText(`${yen(ANNUAL_TARGET)}万円`, T({ x: bx + wOf(ANNUAL_TARGET) + 0.15, y: rows[1].y, w: 2, h: bh, fontSize: 18, bold: true, valign: "middle" }));
  // 見込み = 実績 + 残りの見込み
  const wAct = wOf(cumulative), wProj = wOf(projected);
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: rows[2].y, w: wAct, h: bh, fill: { color: C.main }, line: { color: C.main } });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + wAct, y: rows[2].y, w: wProj - wAct, h: bh, fill: { color: C.mainLight }, line: { color: C.mainLight } });
  s.addText(
    [
      { text: `${firstMonth}〜${lastClosedMonth}月 実績`, options: { fontSize: 11, breakLine: true } },
      { text: `${yen(cumulative)}万円`, options: { fontSize: 15, bold: true } },
    ],
    T({ x: bx + 0.15, y: rows[2].y, w: wAct - 0.3, h: bh, color: C.white, valign: "middle" })
  );
  // 薄い部分の説明はバーの下に（目標ラインと重ならないように）
  const ly2 = rows[2].y + bh + 0.12;
  s.addShape(pres.shapes.RECTANGLE, { x: bx + wAct, y: ly2 + 0.07, w: 0.18, h: 0.18, fill: { color: C.mainLight }, line: { color: C.mainLight } });
  s.addText(`${nextMonth}〜${fyEndMonth}月の見込み（同じペースなら）`, T({ x: bx + wAct + 0.28, y: ly2, w: 4.5, h: 0.32, fontSize: 12, color: C.sub, valign: "middle" }));
  s.addText(`約${yen(projected)}万円`, T({ x: bx + wProj + 0.15, y: rows[2].y, w: 2, h: bh, fontSize: 20, bold: true, color: C.main, valign: "middle" }));
  // 目標ライン
  const tx = bx + wOf(ANNUAL_TARGET);
  s.addShape(pres.shapes.LINE, { x: tx, y: rows[0].y - 0.2, w: 0, h: rows[2].y + bh + 0.05 - (rows[0].y - 0.2), line: { color: C.ink, width: 1.25, dashType: "dash" } });

  // 右：ひとこと
  const rx = 9.95, rw = W - M - rx;
  box(s, rx, 2.0, rw, 1.6, C.mainPale);
  s.addText(
    [
      { text: "半年で達成率", options: { fontSize: 13, color: C.sub, breakLine: true } },
      { text: `${pct(achievementPct)}%`, options: { fontSize: 36, bold: true, color: C.main, breakLine: true } },
      { text: `計画 ${pct(planPacePct)}%より +${pct(aheadPt)}pt`, options: { fontSize: 13, bold: true } },
    ],
    T({ x: rx + 0.25, y: 2.0, w: rw - 0.4, h: 1.6, valign: "middle" })
  );
  box(s, rx, 3.9, rw, 1.4, C.greenPale);
  s.addText(
    [
      { text: "残り6ヶ月は", options: { fontSize: 13, breakLine: true } },
      { text: `月${requiredMonthly}万円で達成`, options: { fontSize: 18, bold: true, color: C.green, breakLine: true } },
      { text: `今の月平均 ${monthlyAverage}万円`, options: { fontSize: 12, color: C.sub } },
    ],
    T({ x: rx + 0.25, y: 3.9, w: rw - 0.4, h: 1.4, valign: "middle" })
  );

  s.addText(
    `※売買仲介を含む。見込み＝${firstMonth}〜${lastClosedMonth}月の月平均×12ヶ月。売買仲介を除くと ${firstMonth}〜${lastClosedMonth}月 ${yen(cumulativeExBaibai)}万円・見込み 約${yen(projectedExBaibai)}万円。昨年は年間実績（社内聞き取り）`,
    T({ x: M, y: 5.85, w: CW, h: 0.5, fontSize: 11, color: C.sub })
  );

  s.addNotes(
    `年度は${firstMonth}月始まりなので、${lastClosedMonth}月末でちょうど半年です。半年で${yen(cumulative)}万円、年間目標${yen(ANNUAL_TARGET)}万円の${pct(achievementPct)}%で、計画の${pct(planPacePct)}%を大きく上回っています。` +
      `このペースが続けば年間で約${yen(projected)}万円、昨年の${yen(LAST_YEAR_ANNUAL)}万円を大きく超えます。` +
      `8月は売買の大きな案件があったので、売買を除いた見込みでも約${yen(projectedExBaibai)}万円です。残り6ヶ月は月${requiredMonthly}万円で目標に届きます。`
  );
}

// ============================================================
// スライド3：② 伸びている理由（AD・付帯系の積み上げ棒グラフ）
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "② 伸びている理由", "AD に加えて、付帯系の手数料が大きく伸びている");
  addPageNumber(s, 3);

  const toMan = (arr) => arr.map((v) => round1(man(v)));
  s.addChart(
    pres.charts.BAR,
    [
      { name: "AD", labels: MONTHS, values: toMan(ACTUALS.ad) },
      { name: "付帯系（保険・保証・付帯商品）", labels: MONTHS, values: toMan(futaiMonthly) },
    ],
    {
      x: M, y: 1.75, w: 8.2, h: 4.85,
      barDir: "col", barGrouping: "stacked", barGapWidthPct: 60,
      chartColors: [C.main, C.mainMid],
      showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 12, dataLabelFontFace: FONT, dataLabelColor: C.white,
      dataLabelFormatCode: "0",
      showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 13, legendColor: C.ink,
      catAxisLabelFontFace: FONT, catAxisLabelFontSize: 14, catAxisLabelColor: C.ink,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      catAxisLineShow: true, catAxisLineColor: C.line,
      showTitle: true, title: "AD・付帯系の月別売上（万円）", titleFontFace: FONT, titleFontSize: 14, titleColor: C.sub,
    }
  );

  const rx = M + 8.2 + 0.4, rw = W - M - rx;
  const cards = [
    ["AD（受取広告料）", `${pct(adTotal)}`, "万円", `${firstMonth}〜${lastClosedMonth}月の合計（月平均 約${adAvg}万円）`, C.main],
    ["付帯系の手数料", `${pct(futaiFirst)} → ${pct(futaiLast)}`, "万円", `${firstMonth}月 → ${lastClosedMonth}月（1ヶ月あたり）`, C.mainMid],
  ];
  cards.forEach(([name, num, unit, foot, dot], i) => {
    const y = 1.85 + i * 1.95;
    box(s, rx, y, rw, 1.7, C.card);
    s.addShape(pres.shapes.OVAL, { x: rx + 0.3, y: y + 0.3, w: 0.2, h: 0.2, fill: { color: dot }, line: { color: dot } });
    s.addText(name, T({ x: rx + 0.6, y: y + 0.2, w: rw - 0.8, h: 0.4, fontSize: 15, bold: true }));
    s.addText(
      [
        { text: num, options: { fontSize: 30, bold: true, color: C.main } },
        { text: ` ${unit}`, options: { fontSize: 14 } },
      ],
      T({ x: rx + 0.3, y: y + 0.6, w: rw - 0.5, h: 0.6, valign: "bottom" })
    );
    s.addText(foot, T({ x: rx + 0.3, y: y + 1.25, w: rw - 0.5, h: 0.3, fontSize: 11, color: C.sub }));
  });
  s.addText("※付帯系＝火災保険・保証会社の紹介手数料＋付帯商品", T({ x: rx, y: 5.85, w: rw, h: 0.5, fontSize: 11, color: C.sub }));

  s.addNotes(
    `伸びを支えているのはADと付帯系の手数料です。ADは${firstMonth}月から${lastClosedMonth}月で${pct(adTotal)}万円、月平均で約${adAvg}万円入っています。` +
      `火災保険や保証会社の紹介手数料などの付帯系は、${firstMonth}月の${pct(futaiFirst)}万円から${lastClosedMonth}月は${pct(futaiLast)}万円まで増えました。` +
      "これは営業が接客のたびにきちんと提案できているからで、この時間をもっと増やしたいと考えています。"
  );
}

// ============================================================
// スライド4：③ 仕組みの図解
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "③ 仕組み", "営業は「現地でしかできないこと」に集中する");
  addPageNumber(s, 4);

  // 左：営業がやること
  const lx = M, ly = 1.8, lw = 5.0, lh = 4.75;
  box(s, lx, ly, lw, lh, C.mainPale);
  box(s, lx + 0.3, ly + 0.3, 2.6, 0.45, C.main, { rectRadius: 0.08 });
  s.addText("営業がやる", T({ x: lx + 0.3, y: ly + 0.3, w: 2.6, h: 0.45, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" }));
  s.addText("現地でしかできないこと", T({ x: lx + 0.3, y: ly + 0.9, w: lw - 0.6, h: 0.45, fontSize: 20, bold: true, color: C.main }));
  s.addText(
    [
      { text: "物件の案内・接客", options: { bullet: true, breakLine: true } },
      { text: "オーナーへの提案", options: { bullet: true, breakLine: true } },
      { text: "現地の確認・立会い", options: { bullet: true } },
    ],
    T({ x: lx + 0.3, y: ly + 1.45, w: lw - 0.6, h: 1.2, fontSize: 16, paraSpaceAfter: 4 })
  );
  photo(s, "site", lx + 0.3, ly + 2.8, lw - 0.6, lh - 3.1, "現地（案内・立会い）");

  // 中央の矢印
  const ax = lx + lw + 0.2;
  s.addShape(pres.shapes.RIGHT_ARROW, { x: ax, y: ly + lh / 2 - 0.35, w: 1.0, h: 0.7, fill: { color: C.mainLight }, line: { color: C.mainLight } });
  s.addText("外に出す", T({ x: ax - 0.15, y: ly + lh / 2 + 0.45, w: 1.3, h: 0.3, fontSize: 12, bold: true, color: C.sub, align: "center" }));

  // 右：外に出す3つ
  const rx = ax + 1.2, rw = W - M - rx;
  const cards = [
    { tag: "営業でなくてもできる", what: "物件入力・写真掲載", to: `パートさん（${PART_SCHEDULE}）`, h: 1.95, staff: true },
    { tag: "現地でなくてもできる", what: "会計", to: "DFE", h: 1.2 },
    { tag: "現地に行く前の一次対応", what: "クレーム対応", to: "駆けつけサービス", h: 1.2 },
  ];
  let y = ly;
  cards.forEach((c) => {
    box(s, rx, y, rw, c.h, C.card);
    s.addText(c.tag, T({ x: rx + 0.3, y: y + 0.2, w: 3.6, h: 0.3, fontSize: 12, bold: true, color: C.sub }));
    s.addText(
      [
        { text: c.what, options: { color: C.ink } },
        { text: "  →  ", options: { color: C.main } },
        { text: c.to, options: { color: C.main } },
      ],
      T({ x: rx + 0.3, y: y + 0.55, w: c.staff ? rw - 2.9 : rw - 0.6, h: c.staff ? 0.9 : 0.45, fontSize: 18, bold: true })
    );
    if (c.staff) {
      const pw = 1.05, ph = 1.2, gap = 0.2;
      [["usui", PART_STAFF[0]], ["kawakami", PART_STAFF[1]]].forEach(([key, name], k) => {
        const px = rx + rw - 0.25 - (2 - k) * pw - (1 - k) * gap;
        photo(s, key, px, y + 0.2, pw, ph, name);
        s.addText(name, T({ x: px - 0.1, y: y + 0.2 + ph + 0.08, w: pw + 0.2, h: 0.3, fontSize: 11, bold: true, align: "center" }));
      });
    }
    y += c.h + 0.35;
  });

  s.addNotes(
    "考え方はシンプルです。営業の仕事を「現地でしかできないこと」と「そうでないこと」に分けます。" +
      "物件の案内や接客、オーナーへの提案、現地の確認は営業にしかできません。" +
      `一方で、物件入力や写真掲載は営業でなくてもできるので、パートの${PART_STAFF.join("と")}にお願いします。` +
      "会計は現地でなくてもできるのでDFEに、クレームの一次対応は駆けつけサービスに任せます。"
  );
}

// ============================================================
// スライド5：④ やること＋決めてほしいこと
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "④ やること", `誰が・いつまでに・何をやるか（${A0}〜${A2}月）`);
  addPageNumber(s, 5);

  const border = { type: "solid", pt: 1, color: C.line };
  const base = { fontFace: FONT, fontSize: 15, color: C.ink, valign: "middle", border: [border, border, border, border], margin: [0.08, 0.15, 0.08, 0.15] };
  const hdr = (text) => ({ text, options: { ...base, bold: true, color: C.white, fill: { color: C.ink }, fontSize: 14 } });
  const cell = (text, extra = {}) => ({ text, options: { ...base, ...extra } });
  const no = (n) => cell(n, { bold: true, color: C.main, align: "center", fontSize: 18 });
  const owner = (n) => {
    const v = orTbd(OWNERS[n]);
    return v === TBD ? cell(TBD, { bold: true, color: C.orange, fill: { color: C.orangePale }, align: "center" }) : cell(v, { align: "center" });
  };
  const task = (title) => cell(title, { bold: true, fontSize: 16 });

  const rows = [
    [hdr("No"), hdr("やること"), hdr("担当"), hdr("期限（案）"), hdr("完了の基準")],
    [no("01"), task("物件入力・写真掲載をパートさんへ"), owner("01"), cell(`${A0}月末 独り立ち`), cell("営業が入力・掲載をしない")],
    [no("02"), task("会計をDFEへ"), owner("02"), cell(`${A0}月末 契約\n${A1}月 移管`), cell("社内の会計作業がほぼゼロ")],
    [no("03"), task("クレーム一次対応を駆けつけサービスへ"), owner("03"), cell(`${A1}月末 契約\n${A2}月 稼働`), cell("営業時間中のクレームを社員が受けない")],
  ];
  s.addTable(rows, { x: M, y: 1.8, w: CW, colW: [0.8, 4.3, 2.0, 2.0, CW - 0.8 - 4.3 - 2.0 - 2.0], rowH: [0.5, 0.9, 0.9, 0.9] });

  s.addText(
    [
      { text: TBD, options: { bold: true, color: C.orange } },
      { text: "＝発表までに埋める欄。期限は案。毎月末に売上と進み具合を確認します。", options: { color: C.sub } },
    ],
    T({ x: M, y: 5.35, w: CW, h: 0.3, fontSize: 12 })
  );

  const dy = 5.85, dh = 0.85;
  box(s, M, dy, CW, dh, C.orangePale);
  s.addText("会社に決めてほしいこと", T({ x: M + 0.35, y: dy, w: 3.3, h: dh, fontSize: 17, bold: true, color: C.orange, valign: "middle" }));
  s.addText("外注の予算枠（DFE・駆けつけサービス）", T({ x: M + 3.8, y: dy, w: CW - 4.1, h: dh, fontSize: 17, bold: true, valign: "middle" }));

  s.addNotes(
    `具体的にやることは3つです。${A0}月中にパートさんへの引き継ぎとDFEとの契約、${A1}月から会計を移し、${A2}月に駆けつけサービスを稼働させます。` +
      "毎月末に売上と進み具合を確認します。最後に1点、DFEと駆けつけサービスの外注予算についてご判断ください。"
  );
}

pres.writeFile({ fileName: path.join(__dirname, OUTPUT) }).then((f) => {
  console.log(`書き出し: ${f}`);
  console.log({ cumulative, cumulativeExBaibai, achievementPct, planPacePct, aheadPt, monthlyAverage, requiredMonthly, projected, projectedExBaibai, adTotal, adAvg, futaiFirst, futaiLast });
});
