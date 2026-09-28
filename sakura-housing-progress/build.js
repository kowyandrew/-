// さくらハウジング 9月進捗報告（改訂版）を生成するスクリプト
// 使い方: npm install && node build.js
// 数字・担当者・目標値が決まったら、下の「設定値」だけを書き換えて再実行する。

const pptxgen = require("pptxgenjs");

// ============================================================
// 設定値（ここだけ書き換える）
// ============================================================

// --- 売上の元データ（単位：万円） ---
const ANNUAL_TARGET = 2500; // 年間売上目標
const CUMULATIVE_SALES = 1441; // 8月末 累計売上
const CUMULATIVE_WITH_BROKERAGE = 1675; // 売買仲介手数料込みの累計
const ORIGINAL_MONTHLY_TARGET = 210; // 当初の月目標

// --- 期間 ---
// 年度の開始月。1〜12月の年度と仮定している（未確定）。違う場合はここを変える。
const FISCAL_YEAR_START_MONTH = 1;
const MONTHS_ELAPSED = 8; // 報告時点で締まっている月数（8月末）
const MONTHS_IN_YEAR = 12;
// 施策を実行する月（暦の月）。年度の区切りとは関係なく 10〜12月
const ACTION_MONTHS = [10, 11, 12];

// --- 発表前に埋める欄（空文字なら「要記入」と表示） ---
const OWNERS = {
  "01": "", // 会計業務の外注
  "02": "", // ポータル掲載パートの戦力化
  "03": "", // コールセンター・駆けつけの外注
  "04": "", // 売上の進捗確認
};
const PORTAL_MONTHLY_TARGET = ""; // 例: "掲載30件・反響15件"

const OUTPUT = "さくらハウジング_進捗報告_改訂版.pptx";

// ============================================================
// 計算（表示する数字はすべてここから作る）
// ============================================================

const round1 = (v) => Math.round(v * 10) / 10;
const achievementPct = round1((CUMULATIVE_SALES / ANNUAL_TARGET) * 100); // 57.6
const withBrokeragePct = round1((CUMULATIVE_WITH_BROKERAGE / ANNUAL_TARGET) * 100); // 67.0
const planPacePct = round1((MONTHS_ELAPSED / MONTHS_IN_YEAR) * 100); // 66.7
// 差は画面に出す丸めた値どうしで引く（66.7 − 57.6 と読み手が検算できるように）
const gapPt = round1(planPacePct - achievementPct); // 9.1
const monthsLeft = MONTHS_IN_YEAR - MONTHS_ELAPSED; // 4
const monthlyAverage = Math.round(CUMULATIVE_SALES / MONTHS_ELAPSED); // 180
const remaining = ANNUAL_TARGET - CUMULATIVE_SALES; // 1,059
const requiredMonthly = Math.round(remaining / monthsLeft); // 265
const requiredMonthlyWithBrokerage = Math.round(
  (ANNUAL_TARGET - CUMULATIVE_WITH_BROKERAGE) / monthsLeft
); // 206
const uplift = requiredMonthly - monthlyAverage; // 85

// 暦の月（1〜12）に直す。offset=0 が年度の最初の月
const calMonth = (offset) => ((FISCAL_YEAR_START_MONTH - 1 + offset) % 12) + 1;
const lastClosedMonth = calMonth(MONTHS_ELAPSED - 1); // 8
const reportMonth = calMonth(MONTHS_ELAPSED); // 9
const leftFrom = calMonth(MONTHS_ELAPSED); // 9
const leftTo = calMonth(MONTHS_IN_YEAR - 1); // 12

// 表示用の書式
const yen = (v) => String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
const pct = (v) => v.toFixed(1);
const MINUS = "−"; // U+2212
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
  ink: "2B2D42",
  sub: "6B6E80",
  green: "2E8B62",
  greenPale: "E6F4ED",
  orange: "C77700",
  orangePale: "FFF3DD",
  card: "F4F5F8",
  line: "E3E4EA",
  white: "FFFFFF",
};
const W = 13.333;
const M = 0.6; // 外周の余白
const CW = W - M * 2; // 本文幅
const TOTAL = 5;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "さくらハウジング";
pres.title = `さくらハウジング ${reportMonth}月 進捗報告`;
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

// 共通のテキスト既定値。呼び出しごとに新しいオブジェクトを作る（pptxgenjs は options を書き換えるため）
const T = (opts) => ({
  fontFace: FONT,
  color: C.ink,
  margin: 0,
  valign: "top",
  isTextBox: true,
  ...opts,
});

function addHeader(slide, section, title) {
  slide.background = { color: C.white };
  slide.addText(section, T({ x: M, y: 0.5, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.main }));
  slide.addText(title, T({ x: M, y: 0.85, w: CW, h: 0.6, fontSize: 28, bold: true, valign: "middle" }));
}

function addPageNumber(slide, n) {
  slide.addText(`${n} / ${TOTAL}`, T({ x: W - M - 1.2, y: 6.75, w: 1.2, h: 0.25, fontSize: 11, color: C.sub, align: "right" }));
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
      { text: `目標まであと${yen(remaining)}万円。`, options: { breakLine: true } },
      { text: "「時間を取り戻す」3つの施策を", options: { breakLine: true } },
      { text: `${ACTION_MONTHS[0]}〜${ACTION_MONTHS[2]}月で実行します。` },
    ],
    T({ x: 0.8, y: 1.5, w: 11.7, h: 2.7, fontSize: 40, bold: true, color: C.white, lineSpacingMultiple: 1.05 })
  );

  const rows = [
    ["現状", [
      { text: `${lastClosedMonth}月末 達成率 ` },
      { text: `${pct(achievementPct)}%`, options: { bold: true } },
      { text: `（計画ペース ${pct(planPacePct)}%に対し ${MINUS}${pct(gapPt)}pt）` },
    ]],
    ["必要", [
      { text: `残り${monthsLeft}ヶ月は ` },
      { text: `月${requiredMonthly}万円`, options: { bold: true } },
      { text: ` の売上が必要（これまでの月平均は${monthlyAverage}万円）` },
    ]],
    ["打ち手", [{ text: "会計の外注／ポータル掲載パートの戦力化／クレーム一次対応の外注" }]],
  ];
  rows.forEach(([label, body], i) => {
    const y = 4.45 + i * 0.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y, w: 1.3, h: 0.46, fill: { color: C.main }, rectRadius: 0.08, line: { color: C.main } });
    s.addText(label, T({ x: 0.8, y, w: 1.3, h: 0.46, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" }));
    s.addText(body, T({ x: 2.4, y, w: 10.1, h: 0.46, fontSize: 16, color: C.white, valign: "middle" }));
  });

  s.addNotes(
    `最初に結論です。${lastClosedMonth}月末の達成率は${pct(achievementPct)}%で、年間の計画ペースより約${Math.round(gapPt)}ポイント遅れています。` +
      `残り${monthsLeft}ヶ月で月${requiredMonthly}万円が必要です。その原因は営業に使う時間が足りないことなので、` +
      `時間を取り戻す3つの施策を${ACTION_MONTHS[0]}月から${ACTION_MONTHS[2]}月で実行します。`
  );
}

// ============================================================
// スライド2：① 現状
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "① 現状", `計画より${Math.round(gapPt)}ポイント遅れ。残り${monthsLeft}ヶ月は月${requiredMonthly}万円が必要`);
  addPageNumber(s, 2);

  // 左カード
  const cx = M, cy = 1.85, cw = 7.5, ch = 4.8;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: cw, h: ch, fill: { color: C.card }, line: { color: C.card }, rectRadius: 0.12 });
  const ix = cx + 0.4, iw = cw - 0.8; // カード内側
  s.addText(`年間目標 ${yen(ANNUAL_TARGET)}万円 に対する達成率`, T({ x: ix, y: cy + 0.3, w: iw, h: 0.35, fontSize: 14, color: C.sub }));
  s.addText(
    [
      { text: pct(achievementPct), options: { fontSize: 66, bold: true, color: C.main } },
      { text: " %", options: { fontSize: 30, bold: true, color: C.main } },
    ],
    T({ x: ix, y: cy + 0.65, w: 3.9, h: 1.3, valign: "bottom" })
  );
  s.addText(
    [
      { text: "累計 ", options: { fontSize: 16, color: C.sub } },
      { text: yen(CUMULATIVE_SALES), options: { fontSize: 24, bold: true } },
      { text: " 万円", options: { fontSize: 16, color: C.sub } },
    ],
    T({ x: ix + 4.0, y: cy + 1.35, w: 2.7, h: 0.55, valign: "bottom" })
  );

  // 進捗バー
  const bx = ix, bw = iw, by = cy + 2.6, bh = 0.4;
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: bh, fill: { color: C.white }, line: { color: C.line, width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: (bw * achievementPct) / 100, h: bh, fill: { color: C.main }, line: { color: C.main, width: 1 } });
  const px = bx + (bw * planPacePct) / 100; // 計画ペースの位置
  s.addShape(pres.shapes.LINE, { x: px, y: by - 0.25, w: 0, h: bh + 0.5, line: { color: C.ink, width: 1.5, dashType: "dash" } });
  // 差ラベル：バーの上、点線の左に右寄せ
  s.addText(`差 ${MINUS}${pct(gapPt)}pt`, T({ x: px - 1.9, y: by - 0.52, w: 1.8, h: 0.3, fontSize: 13, bold: true, color: C.main, align: "right" }));
  // 計画ペースラベル：バーの下、点線の右に左寄せ
  s.addText(
    [
      { text: `計画ペース ${pct(planPacePct)}%`, options: { bold: true, breakLine: true } },
      { text: `（${MONTHS_ELAPSED}/${MONTHS_IN_YEAR}ヶ月）` },
    ],
    T({ x: px + 0.1, y: by + bh + 0.05, w: bx + bw - px - 0.1, h: 0.55, fontSize: 12 })
  );

  // 参考
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ix, y: cy + 3.75, w: iw, h: 0.5, fill: { color: C.white }, line: { color: C.line, width: 1 }, rectRadius: 0.06 });
  s.addText(`参考：売買仲介手数料込みなら ${yen(CUMULATIVE_WITH_BROKERAGE)}万円（${pct(withBrokeragePct)}%）`, T({ x: ix + 0.2, y: cy + 3.75, w: iw - 0.4, h: 0.5, fontSize: 13, color: C.sub, valign: "middle" }));
  s.addText(`※計画ペースは年度を${FISCAL_YEAR_START_MONTH}〜${calMonth(11)}月と仮定して算出`, T({ x: ix, y: cy + 4.37, w: iw, h: 0.25, fontSize: 11, color: C.sub }));

  // 右：3段の数字カード
  const rx = cx + cw + 0.4, rw = W - M - rx;
  s.addText(`残り${monthsLeft}ヶ月（${leftFrom}〜${leftTo}月）に必要な売上`, T({ x: rx, y: cy, w: rw, h: 0.3, fontSize: 13, color: C.sub }));
  const stats = [
    ["これまでの月平均", monthlyAverage, false],
    ["当初の月目標", ORIGINAL_MONTHLY_TARGET, false],
    ["これから必要な月額", requiredMonthly, true],
  ];
  stats.forEach(([label, value, hot], i) => {
    const y = cy + 0.45 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx, y, w: rw, h: 0.75, fill: { color: hot ? C.main : C.card }, line: { color: hot ? C.main : C.card }, rectRadius: 0.1 });
    s.addText(label, T({ x: rx + 0.25, y, w: 2.2, h: 0.75, fontSize: 14, bold: true, color: hot ? C.white : C.ink, valign: "middle" }));
    s.addText(
      [
        { text: String(value), options: { fontSize: 34, bold: true } },
        { text: " 万円", options: { fontSize: 13 } },
      ],
      T({ x: rx + 2.3, y, w: rw - 2.5, h: 0.75, color: hot ? C.white : C.ink, align: "right", valign: "middle" })
    );
  });
  s.addText(
    [
      { text: `残り ${yen(remaining)}万円 ÷ ${monthsLeft}ヶ月 ＝ 月${requiredMonthly}万円`, options: { breakLine: true } },
      { text: "→ 月平均より " },
      { text: `約${uplift}万円`, options: { color: C.main } },
      { text: " の上積みが必要" },
    ],
    T({ x: rx, y: cy + 3.5, w: rw, h: 0.8, fontSize: 14, bold: true, lineSpacingMultiple: 1.2 })
  );

  s.addNotes(
    `年間目標${yen(ANNUAL_TARGET)}万円に対し、${lastClosedMonth}月末で${yen(CUMULATIVE_SALES)}万円、${pct(achievementPct)}%です。` +
      `${MONTHS_ELAPSED}ヶ月経過時点の計画ペースは${pct(planPacePct)}%なので、約${Math.round(gapPt)}ポイント遅れています。` +
      `残り${yen(remaining)}万円を${monthsLeft}ヶ月で取るには月${requiredMonthly}万円が必要で、` +
      `これまでの月平均${monthlyAverage}万円から約${uplift}万円の上積みが必要です。` +
      `なお売買仲介手数料を含めると${pct(withBrokeragePct)}%で、ほぼ計画どおりです。`
  );
}

// ============================================================
// スライド3：② なぜ遅れているか
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "② なぜ遅れているか", "営業に使える時間が3つの業務に奪われている");
  addPageNumber(s, 3);

  const gap = 0.6;
  const colW = (CW - gap * 2) / 3;
  const xs = [M, M + colW + gap, M + (colW + gap) * 2];
  const heads = ["課題（時間を奪っているもの）", "打ち手", "狙う効果"];
  heads.forEach((h, i) => s.addText(h, T({ x: xs[i], y: 1.8, w: colW, h: 0.3, fontSize: 13, bold: true, color: C.sub })));

  const rows = [
    ["会計業務に追われている", "毎月 約40時間が会計作業に", "01", "会計業務を外注する", "月40時間を営業に回す"],
    ["ポータル掲載が追いつかない", "集客の要なのに更新の時間がない", "02", "掲載専任パートを戦力化\n（9/18採用済）", "掲載スピードUP → 反響UP"],
    ["クレームで仕事が中断", "突発対応で1日の計画が崩れる", "03", "コールセンター・\n駆けつけを外注", "一次対応を外に出し、\n計画どおり動ける"],
  ];
  const rh = 0.95, rgap = 0.3;
  rows.forEach(([issue, detail, no, action, effect], i) => {
    const y = 2.25 + i * (rh + rgap);
    const box = (x, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: colW, h: rh, fill: { color: fill }, line: { color: fill }, rectRadius: 0.1 });
    box(xs[0], C.card);
    box(xs[1], C.mainPale);
    box(xs[2], C.greenPale);
    s.addText(
      [
        { text: issue, options: { fontSize: 15, bold: true, breakLine: true } },
        { text: detail, options: { fontSize: 12, color: C.sub } },
      ],
      T({ x: xs[0] + 0.25, y, w: colW - 0.5, h: rh, valign: "middle", paraSpaceAfter: 2 })
    );
    s.addText(
      [
        { text: `${no}  `, options: { color: C.main } },
        { text: action },
      ],
      T({ x: xs[1] + 0.25, y, w: colW - 0.5, h: rh, fontSize: 15, bold: true, valign: "middle" })
    );
    s.addText(effect, T({ x: xs[2] + 0.25, y, w: colW - 0.5, h: rh, fontSize: 15, bold: true, color: C.green, valign: "middle" }));
    // 列の間の矢印
    [0, 1].forEach((k) => {
      const ax = xs[k] + colW + (gap - 0.3) / 2;
      s.addShape(pres.shapes.RIGHT_ARROW, { x: ax, y: y + rh / 2 - 0.15, w: 0.3, h: 0.3, fill: { color: C.mainLight }, line: { color: C.mainLight } });
    });
  });

  const py = 2.25 + 3 * (rh + rgap);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: py, w: CW, h: 0.55, fill: { color: C.white }, line: { color: C.line, width: 1 }, rectRadius: 0.08 });
  s.addText(
    [
      { text: "準備済み", options: { bold: true, color: C.green } },
      { text: "　広告料・販促物の見直し ／ 専門業者との協議 ／ 現場課題の洗い出し", options: { color: C.ink } },
    ],
    T({ x: M + 0.25, y: py, w: CW - 0.5, h: 0.55, fontSize: 13, valign: "middle" })
  );

  s.addNotes(
    "遅れの原因は、営業に使える時間が3つの業務に奪われていることです。会計業務に月40時間、ポータル掲載の手が回らない、クレームの突発対応で計画が崩れる。" +
      "それぞれに対して、外注・専任パート・外部サービスで手を打ちます。広告費の見直しや業者との協議、課題の洗い出しはすでに終わっています。"
  );
}

// ============================================================
// スライド4：③ やること
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "③ やること", "誰が・いつまでに・何をやれば完了か");
  addPageNumber(s, 4);

  const border = { type: "solid", pt: 1, color: C.line };
  const base = { fontFace: FONT, fontSize: 13, color: C.ink, valign: "middle", border: [border, border, border, border], margin: [0.06, 0.12, 0.06, 0.12] };
  const hdr = (text) => ({ text, options: { ...base, bold: true, color: C.white, fill: { color: C.ink }, fontSize: 13 } });
  const cell = (text, extra = {}) => ({ text, options: { ...base, ...extra } });
  const ownerCell = (no) => {
    const v = orTbd(OWNERS[no]);
    return v === TBD
      ? cell(TBD, { bold: true, color: C.orange, fill: { color: C.orangePale }, align: "center" })
      : cell(v, { align: "center" });
  };
  const taskCell = (title, steps) =>
    cell([
      { text: title, options: { bold: true, fontSize: 14, breakLine: true } },
      { text: steps, options: { fontSize: 11, color: C.sub } },
    ]);

  const portal = orTbd(PORTAL_MONTHLY_TARGET);
  const rows = [
    [hdr("No"), hdr("やること"), hdr("担当"), hdr("期限（案）"), hdr("完了の基準")],
    [
      cell("01", { bold: true, color: C.main, align: "center", fontSize: 14 }),
      taskCell("会計業務の外注", "見積比較 → 外注先決定 → 10月分から移管"),
      ownerCell("01"),
      cell("10月末 契約\n11月 移管完了"),
      cell("社内の会計作業 月40h → ほぼゼロ"),
    ],
    [
      cell("02", { bold: true, color: C.main, align: "center", fontSize: 14 }),
      taskCell("ポータル掲載パートの戦力化", "掲載手順を引き継ぎ、単独で更新できる状態に"),
      ownerCell("02"),
      cell("10月末 独り立ち"),
      cell([
        { text: "掲載件数・反響数が月目標に到達", options: { breakLine: true } },
        portal === TBD
          ? { text: `（目標値：${TBD}）`, options: { bold: true, color: C.orange } }
          : { text: `（目標値：${portal}）` },
      ]),
    ],
    [
      cell("03", { bold: true, color: C.main, align: "center", fontSize: 14 }),
      taskCell("コールセンター・駆けつけの外注", "業者比較 → 契約 → 入居者・オーナーへ周知"),
      ownerCell("03"),
      cell("11月末 契約\n12月 稼働"),
      cell("営業時間中のクレーム一次対応を社員が受けない"),
    ],
    [
      cell("04", { bold: true, color: C.main, align: "center", fontSize: 14 }),
      taskCell("売上の進捗確認", `月${requiredMonthly}万円に届いているかを毎月チェック`),
      ownerCell("04"),
      cell("毎月末"),
      cell("月次で目標との差と次の手を報告"),
    ],
  ];
  // 02 の目標値が要記入のときは、セル全体も薄いオレンジにして埋める欄だとわかるようにする
  if (portal === TBD) rows[2][4].options.fill = { color: C.orangePale };

  s.addTable(rows, {
    x: M, y: 1.8, w: CW,
    colW: [0.75, 4.3, 1.5, 2.2, CW - 0.75 - 4.3 - 1.5 - 2.2],
    rowH: [0.45, 0.85, 0.85, 0.85, 0.85],
  });

  s.addText(
    [
      { text: TBD, options: { bold: true, color: C.orange } },
      { text: "＝発表までに埋める欄。期限は案なので、実態に合わせて修正してください。", options: { color: C.sub } },
    ],
    T({ x: M, y: 6.0, w: CW, h: 0.3, fontSize: 12 })
  );

  s.addNotes("具体的にやることは4つです。それぞれ担当・期限・完了の基準を決めました。「完了の基準」を満たしたかどうかで毎月確認します。");
}

// ============================================================
// スライド5：④ スケジュール＋決めてほしいこと
// ============================================================
{
  const s = pres.addSlide();
  addHeader(s, "④ スケジュール", `${ACTION_MONTHS[0]}〜${ACTION_MONTHS[2]}月の3ヶ月で、3施策をすべて稼働させる`);
  addPageNumber(s, 5);

  const lx = M, lw = 3.0;
  const gx = lx + lw + 0.2, gw = W - M - gx, mw = gw / 3;
  const hy = 1.8, hh = 0.42;
  ACTION_MONTHS.forEach((m, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: gx + i * mw, y: hy, w: mw, h: hh, fill: { color: C.ink }, line: { color: C.white, width: 1.5 } });
    s.addText(`${m}月`, T({ x: gx + i * mw, y: hy, w: mw, h: hh, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" }));
  });

  const rowY0 = hy + hh + 0.1, rowH = 0.6;
  const labels = [["01", "会計の外注"], ["02", "掲載パート戦力化"], ["03", "コールセンター等"], ["04", "進捗確認"]];
  labels.forEach(([no, name], i) => {
    const y = rowY0 + i * rowH;
    s.addText([{ text: `${no}  `, options: { color: C.main } }, { text: name }], T({ x: lx, y, w: lw, h: rowH, fontSize: 14, bold: true, valign: "middle" }));
    s.addShape(pres.shapes.LINE, { x: lx, y: y + rowH, w: W - M - lx, h: 0, line: { color: C.line, width: 1 } });
  });

  // バー：[行, 開始月index, 月数, 文言, 濃い色か]
  const bars = [
    [0, 0, 1, "見積・契約", false],
    [0, 1, 2, "外注で運用（月40h創出）", true],
    [1, 0, 1, "引き継ぎ", false],
    [1, 1, 2, "単独で掲載・反響を追う", true],
    [2, 0, 2, "業者比較・契約・周知", false],
    [2, 2, 1, "稼働", true],
  ];
  bars.forEach(([r, start, len, text, dark]) => {
    const x = gx + start * mw + 0.06, w = len * mw - 0.12, y = rowY0 + r * rowH + 0.1, h = rowH - 0.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: dark ? C.main : C.mainLight }, line: { color: dark ? C.main : C.mainLight }, rectRadius: 0.06 });
    s.addText(text, T({ x: x + 0.15, y, w: w - 0.3, h, fontSize: 12, bold: true, color: dark ? C.white : C.ink, valign: "middle" }));
  });
  // 04：各月末に緑のマーカー
  [0, 1, 2].forEach((i) => {
    const d = 0.26;
    s.addShape(pres.shapes.OVAL, { x: gx + (i + 1) * mw - d - 0.1, y: rowY0 + 3 * rowH + (rowH - d) / 2, w: d, h: d, fill: { color: C.green }, line: { color: C.green } });
  });
  const legY = rowY0 + 4 * rowH + 0.15;
  s.addShape(pres.shapes.OVAL, { x: gx, y: legY + 0.06, w: 0.18, h: 0.18, fill: { color: C.green }, line: { color: C.green } });
  s.addText("月末に売上と進捗を確認", T({ x: gx + 0.28, y: legY, w: 4, h: 0.3, fontSize: 12, color: C.green, valign: "middle" }));

  // 会社に決めてほしいこと
  const dy = legY + 0.55, dh = 1.15;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: dy, w: CW, h: dh, fill: { color: C.orangePale }, line: { color: C.orangePale }, rectRadius: 0.1 });
  s.addText("会社に決めてほしいこと", T({ x: M + 0.3, y: dy, w: 2.4, h: dh, fontSize: 16, bold: true, color: C.orange, valign: "middle" }));
  s.addText(
    [
      { text: "1.  外注3件（会計・コールセンター・駆けつけ）の予算枠", options: { breakLine: true } },
      { text: `2.  年間目標に売買仲介手数料を含めるか（含めれば月${requiredMonthlyWithBrokerage}万円、含めなければ月${requiredMonthly}万円が必要）` },
    ],
    T({ x: M + 2.9, y: dy, w: CW - 3.2, h: dh, fontSize: 14, valign: "middle", paraSpaceAfter: 6 })
  );

  s.addNotes(
    `${ACTION_MONTHS[0]}月に契約と引き継ぎ、${ACTION_MONTHS[1]}月から会計外注とパートの単独稼働、${ACTION_MONTHS[2]}月にコールセンター等を稼働させます。` +
      "最後に2点ご判断ください。外注3件の予算枠と、年間目標に仲介手数料を含めるかどうかです。" +
      `含めるかどうかで、残り${monthsLeft}ヶ月に必要な月額が${requiredMonthlyWithBrokerage}万円か${requiredMonthly}万円かに変わります。`
  );
}

pres.writeFile({ fileName: OUTPUT }).then((f) => {
  console.log(`書き出し: ${f}`);
  console.log({ achievementPct, planPacePct, gapPt, withBrokeragePct, monthlyAverage, remaining, requiredMonthly, requiredMonthlyWithBrokerage, uplift });
});
