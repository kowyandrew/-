// さくらハウジング 9月進捗報告（改訂版）を生成するスクリプト
// 使い方: npm install && node build.js
// 数字・担当者が決まったら、下の「設定値」だけを書き換えて再実行する。

const pptxgen = require("pptxgenjs");

// ============================================================
// 設定値（ここだけ書き換える）
// ============================================================

// --- 売上（単位：万円。売買仲介手数料を含む） ---
const ANNUAL_TARGET = 2500; // 今年の年間売上目標
const CUMULATIVE_SALES = 1675; // 8月末 累計売上（仲介込み）
const LAST_YEAR_ANNUAL = 2400; // 昨年の年間売上

// --- AD・付帯商品の積み上げ棒グラフ（単位：万円） ---
// null のままだとグラフの代わりに「要記入」の枠を表示する。
// 「その他の売上」は total − AD − 付帯商品 で自動計算する。
const MIX = [
  { label: "昨年 1〜8月", total: null, ad: null, futai: null },
  { label: "今年 1〜8月", total: CUMULATIVE_SALES, ad: null, futai: null },
];

// --- 期間 ---
// 年度の開始月。1〜12月の年度と仮定している（未確定）。違う場合はここを変える。
const FISCAL_YEAR_START_MONTH = 1;
const MONTHS_ELAPSED = 8; // 報告時点で締まっている月数（8月末）
const MONTHS_IN_YEAR = 12;
// 施策を実行する月（暦の月）。年度の区切りとは関係なく 10〜12月
const ACTION_MONTHS = [10, 11, 12];

// --- 発表前に埋める欄（空文字なら「要記入」と表示） ---
const OWNERS = {
  "01": "", // 会計をDFEへ
  "02": "", // 写真撮影・入力をパートさんへ
  "03": "", // クレーム一次対応を外注
};

const OUTPUT = "さくらハウジング_進捗報告_改訂版.pptx";

// ============================================================
// 計算（表示する数字はすべてここから作る）
// ============================================================

const round1 = (v) => Math.round(v * 10) / 10;
const round10 = (v) => Math.round(v / 10) * 10;
const achievementPct = round1((CUMULATIVE_SALES / ANNUAL_TARGET) * 100); // 67.0
const planPacePct = round1((MONTHS_ELAPSED / MONTHS_IN_YEAR) * 100); // 66.7
const aheadPt = round1(achievementPct - planPacePct); // 0.3
const monthsLeft = MONTHS_IN_YEAR - MONTHS_ELAPSED; // 4
const monthlyAverage = Math.round(CUMULATIVE_SALES / MONTHS_ELAPSED); // 209
const remaining = ANNUAL_TARGET - CUMULATIVE_SALES; // 825
const requiredMonthly = Math.round(remaining / monthsLeft); // 206
const projectedAnnual = round10((CUMULATIVE_SALES / MONTHS_ELAPSED) * MONTHS_IN_YEAR); // 約2,510
const vsLastYear = projectedAnnual - LAST_YEAR_ANNUAL; // 約+110

const calMonth = (offset) => ((FISCAL_YEAR_START_MONTH - 1 + offset) % 12) + 1;
const lastClosedMonth = calMonth(MONTHS_ELAPSED - 1); // 8
const reportMonth = calMonth(MONTHS_ELAPSED); // 9
const fyEnd = calMonth(MONTHS_IN_YEAR - 1); // 12
const [A0, A1, A2] = ACTION_MONTHS;

const mixReady = MIX.every((m) => [m.total, m.ad, m.futai].every((v) => typeof v === "number"));
const mixRows = MIX.map((m) => ({ ...m, other: mixReady ? m.total - m.ad - m.futai : null }));
const growth = (key) => (mixReady ? mixRows[1][key] - mixRows[0][key] : null);

// 表示用の書式
const yen = (v) => String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
const pct = (v) => v.toFixed(1);
const signed = (v) => (v >= 0 ? `+${yen(v)}` : `−${yen(-v)}`);
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
  ink: "2B2D42",
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
pres.title = `さくらハウジング ${reportMonth}月 進捗報告`;
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

// pptxgenjs は options を書き換えるので、呼び出しごとに新しいオブジェクトを作る
const T = (opts) => ({ fontFace: FONT, color: C.ink, margin: 0, valign: "top", isTextBox: true, ...opts });
const box = (s, x, y, w, h, fill, extra = {}) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12, ...extra });

function addHeader(slide, section, title) {
  slide.background = { color: C.white };
  slide.addText(section, T({ x: M, y: 0.5, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.main }));
  slide.addText(title, T({ x: M, y: 0.85, w: CW, h: 0.65, fontSize: 30, bold: true, valign: "middle" }));
}
function addPageNumber(slide, n) {
  slide.addText(`${n} / ${TOTAL}`, T({ x: W - M - 1.2, y: 6.8, w: 1.2, h: 0.25, fontSize: 11, color: C.sub, align: "right" }));
}

// ============================================================
// スライド1：表紙＋結論
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.ink };
  s.addText(`さくらハウジング　2026年${reportMonth}月 進捗報告`, T({ x: 0.8, y: 0.8, w: 11, h: 0.4, fontSize: 16, bold: true, color: C.mainLight }));
  s.addText(
    [
      { text: "昨年を上回るペースで進んでいます。", options: { breakLine: true } },
      { text: "営業に集中できる仕組みをつくり、", options: { breakLine: true } },
      { text: `年間${yen(ANNUAL_TARGET)}万円を達成します。` },
    ],
    T({ x: 0.8, y: 1.5, w: 11.7, h: 2.7, fontSize: 40, bold: true, color: C.white, lineSpacingMultiple: 1.05 })
  );
  const rows = [
    ["現状", `${lastClosedMonth}月末 ${yen(CUMULATIVE_SALES)}万円（達成率 ${pct(achievementPct)}%）　計画ペース ${pct(planPacePct)}%を上回る`],
    ["伸び", "AD・付帯商品の上がり幅が大きい"],
    ["打ち手", "会計はDFEへ／写真撮影・入力はパートさんへ／クレーム一次対応は外注"],
  ];
  rows.forEach(([label, body], i) => {
    const y = 4.45 + i * 0.75;
    box(s, 0.8, y, 1.3, 0.46, C.main, { rectRadius: 0.08 });
    s.addText(label, T({ x: 0.8, y, w: 1.3, h: 0.46, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" }));
    s.addText(body, T({ x: 2.4, y, w: 10.1, h: 0.46, fontSize: 16, color: C.white, valign: "middle" }));
  });
  s.addNotes(
    `最初に結論です。${lastClosedMonth}月末の売上は${yen(CUMULATIVE_SALES)}万円、達成率${pct(achievementPct)}%で、計画ペースを上回っています。` +
      `このままいけば昨年の${yen(LAST_YEAR_ANNUAL)}万円を超えるペースです。特にAD・付帯商品が大きく伸びています。` +
      `この流れを確実にするために、営業が営業に集中できる仕組みを${A0}月から${A2}月でつくります。`
  );
}

// ============================================================
// スライド2：① 現状（数字3つ）
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "① 現状", "計画を上回り、昨年を超えるペース");
  addPageNumber(s, 2);

  const gap = 0.4, cw = (CW - gap * 2) / 3, cy = 1.95, ch = 2.75;
  const cards = [
    [`${lastClosedMonth}月末 達成率`, pct(achievementPct), "%", `計画ペース ${pct(planPacePct)}%（+${pct(aheadPt)}pt）`, true],
    ["このペースでの年間見込み", yen(projectedAnnual), "万円", `年間目標 ${yen(ANNUAL_TARGET)}万円`, false],
    ["昨年の年間売上", yen(LAST_YEAR_ANNUAL), "万円", `見込みは昨年より 約${signed(vsLastYear)}万円`, false],
  ];
  cards.forEach(([label, num, unit, foot, hot], i) => {
    const x = M + i * (cw + gap);
    box(s, x, cy, cw, ch, hot ? C.main : C.card);
    const fg = hot ? C.white : C.ink;
    s.addText(label, T({ x: x + 0.35, y: cy + 0.3, w: cw - 0.7, h: 0.4, fontSize: 16, bold: true, color: hot ? C.white : C.sub }));
    s.addText(
      [
        ...(i === 1 ? [{ text: "約", options: { fontSize: 20, bold: true } }] : []),
        { text: num, options: { fontSize: 48, bold: true } },
        { text: ` ${unit}`, options: { fontSize: 20, bold: true } },
      ],
      T({ x: x + 0.35, y: cy + 0.8, w: cw - 0.5, h: 1.1, color: fg, valign: "bottom" })
    );
    s.addText(foot, T({ x: x + 0.35, y: cy + 2.05, w: cw - 0.7, h: 0.4, fontSize: 15, color: fg }));
  });

  const by = cy + ch + 0.4;
  box(s, M, by, CW, 0.75, C.greenPale);
  s.addText(
    [
      { text: `残り${monthsLeft}ヶ月は 月${requiredMonthly}万円 で目標達成`, options: { bold: true, color: C.green } },
      { text: `（これまでの月平均 ${monthlyAverage}万円）`, options: { color: C.ink } },
    ],
    T({ x: M + 0.35, y: by, w: CW - 0.7, h: 0.75, fontSize: 18, valign: "middle" })
  );
  s.addText(`※売上は売買仲介手数料を含む。計画ペースは年度を${FISCAL_YEAR_START_MONTH}〜${fyEnd}月と仮定して算出`, T({ x: M, y: by + 0.95, w: CW, h: 0.25, fontSize: 11, color: C.sub }));

  s.addNotes(
    `年間目標${yen(ANNUAL_TARGET)}万円に対し、${lastClosedMonth}月末で${yen(CUMULATIVE_SALES)}万円、達成率${pct(achievementPct)}%です。` +
      `${MONTHS_ELAPSED}ヶ月経過時点の計画ペース${pct(planPacePct)}%を上回っています。` +
      `このペースで行けば年間で約${yen(projectedAnnual)}万円、昨年の${yen(LAST_YEAR_ANNUAL)}万円を約${yen(vsLastYear)}万円上回る見込みです。` +
      `残り${monthsLeft}ヶ月は月${requiredMonthly}万円で目標達成で、これまでの月平均${monthlyAverage}万円なら届く水準です。`
  );
}

// ============================================================
// スライド3：② AD・付帯商品の伸び（積み上げ棒グラフ）
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "② 伸びている理由", "伸びを引っ張っているのは AD・付帯商品");
  addPageNumber(s, 3);

  const gx = M, gy = 1.8, gw = 7.6, gh = 4.7;
  if (mixReady) {
    s.addChart(
      pres.charts.BAR,
      [
        { name: "その他の売上", labels: mixRows.map((m) => m.label), values: mixRows.map((m) => m.other) },
        { name: "付帯商品", labels: mixRows.map((m) => m.label), values: mixRows.map((m) => m.futai) },
        { name: "AD", labels: mixRows.map((m) => m.label), values: mixRows.map((m) => m.ad) },
      ],
      {
        x: gx, y: gy, w: gw, h: gh,
        barDir: "col", barGrouping: "stacked", barGapWidthPct: 80,
        chartColors: [C.gray, C.mainMid, C.main],
        showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 13, dataLabelFontFace: FONT, dataLabelColor: C.ink,
        dataLabelFormatCode: "#,##0",
        showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 13, legendColor: C.ink,
        catAxisLabelFontFace: FONT, catAxisLabelFontSize: 14, catAxisLabelColor: C.ink,
        valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
        catAxisLineShow: true, catAxisLineColor: C.line,
        showTitle: true, title: "1〜8月の売上の内訳（万円）", titleFontFace: FONT, titleFontSize: 14, titleColor: C.sub,
      }
    );
  } else {
    box(s, gx, gy, gw, gh, C.orangePale);
    s.addText(
      [
        { text: `積み上げ棒グラフ：${TBD}`, options: { bold: true, color: C.orange, fontSize: 18, breakLine: true } },
        { text: "build.js の MIX に、昨年・今年（1〜8月）の", options: { breakLine: true } },
        { text: "売上合計・AD・付帯商品の金額を入れて再生成" },
      ],
      T({ x: gx + 0.5, y: gy, w: gw - 1, h: gh, fontSize: 14, color: C.ink, valign: "middle" })
    );
  }

  // 右：上がり幅
  const rx = gx + gw + 0.5, rw = W - M - rx;
  s.addText("昨年同期（1〜8月）からの上がり幅", T({ x: rx, y: gy, w: rw, h: 0.35, fontSize: 14, bold: true, color: C.sub }));
  const items = [
    ["AD", growth("ad"), C.main],
    ["付帯商品", growth("futai"), C.mainMid],
  ];
  items.forEach(([name, g, dot], i) => {
    const y = gy + 0.55 + i * 1.5;
    box(s, rx, y, rw, 1.2, C.card);
    s.addShape(pres.shapes.OVAL, { x: rx + 0.3, y: y + 0.24, w: 0.2, h: 0.2, fill: { color: dot }, line: { color: dot } });
    s.addText(name, T({ x: rx + 0.6, y: y + 0.15, w: rw - 0.9, h: 0.4, fontSize: 16, bold: true }));
    s.addText(
      g === null
        ? [{ text: TBD, options: { color: C.orange, fontSize: 22, bold: true } }]
        : [{ text: signed(g), options: { fontSize: 36, bold: true, color: C.main } }, { text: " 万円", options: { fontSize: 16 } }],
      T({ x: rx + 0.3, y: y + 0.5, w: rw - 0.6, h: 0.6, valign: "bottom" })
    );
  });
  s.addText("この伸びを続けるため、営業が接客・提案に使える時間を増やす", T({ x: rx, y: gy + 3.65, w: rw, h: 0.9, fontSize: 15, bold: true, lineSpacingMultiple: 1.2 }));

  s.addNotes(
    "売上の伸びを引っ張っているのはADと付帯商品です。昨年の同じ時期と比べて、この2つの上がり幅が大きくなっています。" +
      "この伸びを続けるには、営業がお客様と向き合う時間を増やすことが一番の近道です。"
  );
}

// ============================================================
// スライド4：③ やること
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "③ やること", "営業に集中できる仕組みを3つつくる");
  addPageNumber(s, 4);

  const border = { type: "solid", pt: 1, color: C.line };
  const base = { fontFace: FONT, fontSize: 15, color: C.ink, valign: "middle", border: [border, border, border, border], margin: [0.08, 0.15, 0.08, 0.15] };
  const hdr = (text) => ({ text, options: { ...base, bold: true, color: C.white, fill: { color: C.ink }, fontSize: 14 } });
  const cell = (text, extra = {}) => ({ text, options: { ...base, ...extra } });
  const no = (n) => cell(n, { bold: true, color: C.main, align: "center", fontSize: 18 });
  const owner = (n) => {
    const v = orTbd(OWNERS[n]);
    return v === TBD ? cell(TBD, { bold: true, color: C.orange, fill: { color: C.orangePale }, align: "center" }) : cell(v, { align: "center" });
  };
  const task = (title, sub) =>
    cell([
      { text: title, options: { bold: true, fontSize: 17, breakLine: true } },
      { text: sub, options: { fontSize: 12, color: C.sub } },
    ]);

  const rows = [
    [hdr("No"), hdr("やること"), hdr("担当"), hdr("期限（案）"), hdr("完了の基準")],
    [no("01"), task("会計をDFEに外注", "見積 → 契約 → 10月分から移管"), owner("01"), cell(`${A0}月末 契約\n${A1}月 移管完了`), cell("社内の会計作業がほぼゼロ")],
    [no("02"), task("写真撮影・入力をパートさんへ", "9/18採用のパートさんに手順を引き継ぐ"), owner("02"), cell(`${A0}月末 独り立ち`), cell("営業が撮影・入力をしない")],
    [no("03"), task("クレーム一次対応を外注", "コールセンター・駆けつけ業者と契約"), owner("03"), cell(`${A1}月末 契約\n${A2}月 稼働`), cell("営業時間中のクレームを社員が受けない")],
  ];
  s.addTable(rows, { x: M, y: 1.9, w: CW, colW: [0.8, 4.5, 1.6, 2.2, CW - 0.8 - 4.5 - 1.6 - 2.2], rowH: [0.5, 1.05, 1.05, 1.05] });

  s.addText(
    [
      { text: TBD, options: { bold: true, color: C.orange } },
      { text: "＝発表までに埋める欄。期限は案なので、実態に合わせて修正してください。", options: { color: C.sub } },
    ],
    T({ x: M, y: 5.95, w: CW, h: 0.3, fontSize: 12 })
  );
  s.addNotes(
    "具体的にやることは3つです。会計はDFEに外注、写真撮影と入力は9月に採用したパートさんにお願いし、クレームの一次対応はコールセンターと駆けつけ業者に任せます。" +
      "それぞれ担当・期限・完了の基準を決め、毎月の会議で確認します。"
  );
}

// ============================================================
// スライド5：④ スケジュール＋決めてほしいこと
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "④ スケジュール", `${A0}〜${A2}月で、3つとも動かす`);
  addPageNumber(s, 5);

  const lx = M, lw = 3.4;
  const gx = lx + lw + 0.2, gw = W - M - gx, mw = gw / 3;
  const hy = 1.85, hh = 0.45;
  ACTION_MONTHS.forEach((m, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: gx + i * mw, y: hy, w: mw, h: hh, fill: { color: C.ink }, line: { color: C.white, width: 1.5 } });
    s.addText(`${m}月`, T({ x: gx + i * mw, y: hy, w: mw, h: hh, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" }));
  });
  const rowY0 = hy + hh + 0.1, rowH = 0.7;
  [["01", "会計 → DFE"], ["02", "写真撮影・入力 → パート"], ["03", "クレーム一次対応 → 外注"]].forEach(([n, name], i) => {
    const y = rowY0 + i * rowH;
    s.addText([{ text: `${n}  `, options: { color: C.main } }, { text: name }], T({ x: lx, y, w: lw, h: rowH, fontSize: 15, bold: true, valign: "middle" }));
    s.addShape(pres.shapes.LINE, { x: lx, y: y + rowH, w: W - M - lx, h: 0, line: { color: C.line, width: 1 } });
  });
  const bars = [
    [0, 0, 1, "見積・契約", false], [0, 1, 2, "DFEで運用", true],
    [1, 0, 1, "引き継ぎ", false], [1, 1, 2, "パートさんが担当", true],
    [2, 0, 2, "業者選び・契約・周知", false], [2, 2, 1, "稼働", true],
  ];
  bars.forEach(([r, start, len, text, dark]) => {
    const x = gx + start * mw + 0.06, w = len * mw - 0.12, y = rowY0 + r * rowH + 0.12, h = rowH - 0.24;
    box(s, x, y, w, h, dark ? C.main : C.mainLight, { rectRadius: 0.06 });
    s.addText(text, T({ x: x + 0.15, y, w: w - 0.3, h, fontSize: 13, bold: true, color: dark ? C.white : C.ink, valign: "middle" }));
  });
  const legY = rowY0 + 3 * rowH + 0.15;
  s.addText("毎月末に売上と進み具合を確認", T({ x: gx, y: legY, w: 6, h: 0.3, fontSize: 13, color: C.sub }));

  const dy = legY + 0.65, dh = 1.0;
  box(s, M, dy, CW, dh, C.orangePale);
  s.addText("会社に決めてほしいこと", T({ x: M + 0.35, y: dy, w: 3.2, h: dh, fontSize: 18, bold: true, color: C.orange, valign: "middle" }));
  s.addText("外注の予算枠（DFE・コールセンター・駆けつけ）", T({ x: M + 3.7, y: dy, w: CW - 4.0, h: dh, fontSize: 18, bold: true, valign: "middle" }));

  s.addNotes(
    `${A0}月に契約と引き継ぎ、${A1}月から会計はDFE、写真撮影・入力はパートさんで回し、${A2}月にクレーム一次対応の外注を稼働させます。` +
      "最後に1点ご判断ください。DFEとコールセンター・駆けつけの外注の予算枠です。"
  );
}

pres.writeFile({ fileName: OUTPUT }).then((f) => {
  console.log(`書き出し: ${f}`);
  console.log({ achievementPct, planPacePct, aheadPt, monthlyAverage, remaining, requiredMonthly, projectedAnnual, vsLastYear, mixReady });
});
