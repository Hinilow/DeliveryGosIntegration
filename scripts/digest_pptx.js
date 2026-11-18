// Еженедельный дайджест в PowerPoint из digest/digest_data.json.
// Запуск: node scripts/digest_pptx.js (после python scripts/build.py)
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
// Если задан путь к модулю apply_theme, цвета темы прописываются в файл. Без него слайды выглядят так же.
const applyTheme = process.env.PPTX_THEME_MODULE ? require(process.env.PPTX_THEME_MODULE).applyTheme : async () => {};

const ROOT = path.resolve(__dirname, "..");
const D = JSON.parse(fs.readFileSync(path.join(ROOT, "digest", "digest_data.json"), "utf8"));
const OUT = path.join(ROOT, "digest", "2026-W47_digest.pptx");

const THEME = {
  name: "Delivery Digest",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1F2A37", lt1: "FFFFFF", dk2: "334155", lt2: "F1F4F7",
    accent1: "0F766E", accent2: "B7791F", accent3: "C0392B", accent4: "2E7D4F",
    accent5: "64748B", accent6: "CBD5E1", hlink: "0F766E", folHlink: "334155",
  },
};
const HEX = THEME.colors;
const pct = (v) => `${Math.round(v * 100)}%`;
const pp = (v) => `${v >= 0 ? "+" : "−"}${Math.abs(v * 100).toFixed(1).replace(".", ",")} п.п.`;
const shortOwner = (o) => o.replace("Руководитель группы ", "РГ ");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = `Дайджест кластера госинтеграций ${D.week}`;
pres.author = "Алексей Степин";
pres.company = "Демо-проект";
pres.subject = "Миграция кластера госинтеграций на DropApp";

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { text: { text: `Дайджест кластера госинтеграций, ${D.week}. Демо на вымышленных данных. РГ – руководитель группы.`,
              options: { x: 0.6, y: 6.95, w: 10.8, h: 0.3, fontSize: 10, color: HEX.accent5, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.8, fontSize: 26,
                                bold: true, color: HEX.dk1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: HEX.accent5, align: "right" },
});

const TONE = { neutral: HEX.dk1, good: HEX.accent4, warn: HEX.accent2, bad: HEX.accent3, main: HEX.accent1 };

function card(slide, x, y, w, h, value, label, sub, tone, valueSize = 34) {
  slide.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: HEX.lt2 },
    line: { color: HEX.lt2 }, objectName: `card ${label}` });
  slide.addText(value, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.7, fontSize: valueSize, bold: true,
    color: TONE[tone], margin: 0, isTextBox: true, objectName: `card value ${label}` });
  slide.addText([{ text: label, options: { bold: true, color: HEX.dk1, breakLine: true } },
                 { text: sub, options: { color: HEX.accent5 } }],
    { x: x + 0.2, y: y + 0.88, w: w - 0.4, h: h - 0.98, fontSize: 12, margin: 0, valign: "top", isTextBox: true,
      objectName: `card label ${label}` });
}

function box(slide, x, y, w, h, title, titleColor, body, name, fontSize = 12) {
  slide.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: HEX.lt2 },
    line: { color: HEX.lt2 }, objectName: `${name} box` });
  slide.addText(title, { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.35, fontSize: 14, bold: true, color: titleColor,
    margin: 0, isTextBox: true, objectName: `${name} title` });
  slide.addText(body, { x: x + 0.25, y: y + 0.55, w: w - 0.5, h: h - 0.65, fontSize, color: HEX.dk2,
    paraSpaceAfter: fontSize > 12 ? 10 : 3, margin: 0, valign: "top", isTextBox: true, objectName: `${name} body` });
}

const bullets = (items) => items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } }));

// ---------------------------------------------------------------- слайд 1: статус
pres.addSection({ title: "Статус" });
let s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Статус" });
s.addText(`${D.in_prod} из ${D.n} сервисов в проде DropApp, за неделю +${D.in_prod - D.in_prod_prev}`,
  { placeholder: "title" });
const cw = 2.3, cg = 0.15, cy = 1.45;
const cx = (k) => 0.6 + k * (cw + cg);
card(s, cx(0), cy, cw, 1.6, `${D.in_prod}/${D.n}`, "В проде DropApp", `из них ${D.prod_debts.length} с долгами`, "main");
card(s, cx(1), cy, cw, 1.6, pct(D.progress), "Готовность", `${pp(D.progress - D.progress_prev)} за неделю`, "neutral");
card(s, cx(2), cy, cw, 1.6, String(D.by_status["Просрочено"]), "Просрочено", "план выхода прошёл", "bad");
card(s, cx(3), cy, cw, 1.6, String(D.by_status["Под риском"]), "Под риском",
  `${D.risk_no_blocker} без названного блокера`, "warn");
card(s, cx(4), cy, cw, 1.6, `${D.rcod_ready}/${D.n}`, "Готовы в РЦОД", "учения пройдены", "good");

s.addChart(pres.ChartType.bar, [{ name: "Готовность", labels: D.teams.map((t) => t.team), values: D.teams.map((t) => t.pct) }], {
  x: 0.6, y: 3.3, w: 7.2, h: 3.45, barDir: "bar", chartColors: [HEX.accent1],
  showTitle: true, title: "Готовность по командам", titleFontSize: 14, titleColor: HEX.dk1, titleFontFace: "+mj-lt",
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0%", dataLabelFontSize: 12,
  dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
  valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0%", valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" },
  catAxisLabelColor: HEX.dk2, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
  catAxisLineShow: false, showLegend: false, barGapWidthPct: 60, objectName: "chart team readiness",
});
s.addText("Что сдвинулось за неделю", { x: 8.2, y: 3.35, w: 4.5, h: 0.4, fontSize: 14, bold: true, color: HEX.dk1,
  margin: 0, isTextBox: true });
s.addText(D.moved.map((m, i) => ({
  text: `${m.service}: ${m.now}`,
  options: { bullet: true, breakLine: i < D.moved.length - 1, color: HEX.dk2 },
})), { x: 8.2, y: 3.8, w: 4.5, h: 2.95, fontSize: 13, paraSpaceAfter: 4, valign: "top", margin: 0, isTextBox: true,
  objectName: "moved list" });

// ---------------------------------------------------------------- слайд 2: прогноз
pres.addSection({ title: "Прогноз" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Прогноз" });
const F = D.forecast;
s.addText(`При текущем темпе к последнему окну ${F.last_window} не успеваем`, { placeholder: "title" });
s.addChart(pres.ChartType.line, [{ name: "Готовность", labels: D.history.dates, values: D.history.ready }], {
  x: 0.6, y: 1.4, w: 6.3, h: 3.3, chartColors: [HEX.accent1], lineSize: 3, lineDataSymbol: "circle",
  lineDataSymbolSize: 8, showTitle: true, title: "Готовность программы по неделям", titleFontSize: 14,
  titleColor: HEX.dk1, titleFontFace: "+mj-lt", showValue: true, dataLabelPosition: "t", dataLabelFormatCode: "0%",
  dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
  valAxisMinVal: 0, valAxisMaxVal: 1, valAxisHidden: true, valGridLine: { style: "none" },
  catAxisLabelColor: HEX.dk2, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", showLegend: false,
  objectName: "chart weekly readiness",
});
s.addText([
  { text: `Темп за 3 недели: ${pp(F.tempo_pp).replace("+", "")} и ${String(F.tempo_cnt.toFixed(1)).replace(".", ",")} `
          + "сервиса в прод в неделю.", options: { breakLine: true } },
  { text: `Последнее окно изменений ${F.last_window}, мораторий ${F.moratorium}. Срок программы ${D.deadline.slice(0, 5)} `
          + "попадает в мораторий." },
], { x: 0.6, y: 4.85, w: 6.3, h: 0.8, fontSize: 12, color: HEX.dk2, margin: 0, valign: "top", paraSpaceAfter: 4,
  isTextBox: true, objectName: "tempo note" });
card(s, 7.3, 1.45, 2.6, 1.75, F.ready_real, "90% готовности", `по темпу ${F.ready}, но это мораторий`, "warn");
card(s, 10.1, 1.45, 2.6, 1.75, F.count_real, `Все ${D.n} в проде`, "по числу сервисов, 1 в неделю", "bad");
box(s, 7.3, 3.4, 5.4, 3.3, "Предлагаю две волны", HEX.accent1, [
  { text: `Первая до ${F.last_window}: `, options: { bold: true, color: HEX.dk1 } },
  { text: `${D.wave1.join(", ")}. Без блокеров, зависимости в проде или в этой же волне. К мораторию в проде будет `
          + `${D.in_prod + D.wave1.length} из ${D.n}.`, options: { breakLine: true } },
  { text: "Вторая после моратория: ", options: { bold: true, color: HEX.dk1 } },
  { text: `${D.wave2.length} сервисов с блокерами и те, кто от них зависит.`, options: { breakLine: true } },
  { text: "Нужно решение руководителя кластера: ", options: { bold: true, color: HEX.dk1 } },
  { text: "согласовать волны и новый срок второй волны." },
], "waves", 14);

// ---------------------------------------------------------------- слайд 3: риски
pres.addSection({ title: "Риски" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Риски" });
s.addText(`${D.help.length} блокера висят дольше 10 рабочих дней`, { placeholder: "title" });
const head = ["Сервис", "Статус", "План", "Сигналы", "Эскалация"].map((t) => ({
  text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk1 } } }));
const rows = D.risks.map((r, i) => {
  const fill = { color: i % 2 ? HEX.lt1 : HEX.lt2 };
  const tone = r.status === "Просрочено" ? HEX.accent3 : HEX.accent2;
  return [
    { text: r.service, options: { bold: true, fill } },
    { text: r.status, options: { color: tone, bold: true, fill } },
    { text: r.plan, options: { fill } },
    { text: r.signals, options: { fill } },
    { text: r.escalation || "–", options: { fill, color: r.escalation === "Руководитель кластера" ? HEX.accent3 : HEX.dk2,
                                             bold: r.escalation === "Руководитель кластера" } },
  ];
});
s.addTable([head, ...rows], { x: 0.6, y: 1.4, w: 12.1, colW: [2.5, 1.4, 0.8, 4.9, 2.5], fontSize: 12,
  color: HEX.dk2, border: { type: "none" }, margin: [0.04, 0.1, 0.04, 0.1], rowH: 0.34, valign: "middle",
  objectName: "risk table" });
const fmt = (iso) => `${iso.slice(8, 10)}.${iso.slice(5, 7)}`;
box(s, 0.6, 5.0, 5.95, 1.75, `${D.conflicts.length} конфликта плана: зависимость выходит позже`, HEX.accent3,
  bullets(D.conflicts.map((c) => `${c.service} (${fmt(c.plan)}) ждёт ${c.depends_on} (${fmt(c.dep_plan)})`)),
  "conflicts");
box(s, 6.75, 5.0, 5.95, 1.75, `${D.prod_debts.length} сервиса в проде, но не до конца`, HEX.accent2,
  bullets(D.prod_debts.map((p) => `${p.service}: ${p.debts}`)), "debts");

// ---------------------------------------------------------------- слайд 4: решения и задачи
pres.addSection({ title: "Решения" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Решения" });
s.addText(`Решения встречи РГ ${D.report_date.slice(0, 5)} и задачи на неделю`, { placeholder: "title" });
s.addText("Решения", { x: 0.6, y: 1.4, w: 4.9, h: 0.4, fontSize: 14, bold: true, color: HEX.dk1, margin: 0, isTextBox: true });
s.addText(D.decisions.map((t, i) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < D.decisions.length - 1 } })),
  { x: 0.6, y: 1.85, w: 4.9, h: 4.9, fontSize: 11, color: HEX.dk2, paraSpaceAfter: 6, valign: "top", margin: 0,
    isTextBox: true, objectName: "decisions" });
const aHead = ["Задача", "Ответственный", "Срок"].map((t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk1 } } }));
const aRows = D.actions.map((a, i) => {
  const fill = { color: i % 2 ? HEX.lt1 : HEX.lt2 };
  return [
    { text: [{ text: `${a.service}. `, options: { bold: true } }, { text: a.action }], options: { fill } },
    { text: shortOwner(a.owner), options: { fill } },
    { text: a.due, options: { fill, bold: true } },
  ];
});
s.addTable([aHead, ...aRows], { x: 5.8, y: 1.4, w: 6.9, colW: [4.3, 1.85, 0.75], fontSize: 11, color: HEX.dk2,
  border: { type: "none" }, margin: [0.04, 0.08, 0.04, 0.08], valign: "middle", objectName: "actions table" });

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
