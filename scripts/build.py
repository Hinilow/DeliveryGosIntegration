"""Сборка delivery-kit из data/*.csv.

Запуск: python scripts/build.py
Пишет: трекер Excel, базу SQLite и результаты запросов, протокол встречи РГ, импорт задач в Jira,
черновик дайджеста (Markdown и данные для слайдов), статусы для схем.
"""
from __future__ import annotations

import csv
import json
import math
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
REPORT_DATE = date(2026, 11, 18)
DEADLINE = date(2026, 12, 25)

STAGES = [
    ("Не начата", 0.0),
    ("Анализ", 0.15),
    ("Контейнеризация", 0.40),
    ("Тест в DropApp", 0.65),
    ("Прод в DropApp", 0.90),
    ("Старый контур выведен", 1.0),
]
WEIGHT = dict(STAGES)
PROD = 0.9
STUCK_WEEKS = 3          # столько недель на одном этапе без движения считаю риском
TAIL_WEEKS = 2           # столько держим старый контур после выхода в прод ради отката
SOON_DAYS = 14           # до плана меньше двух недель, а чек-лист не закрыт
TEMPO_WEEKS = 3          # окно для темпа
TARGET_READINESS = 0.9   # все сервисы в проде дают примерно 90%

CHECK_ITEMS = [  # колонка в checklist.csv, заголовок в трекере, название в базе и текстах
    ("ci_scan", "CI и сканер образа", "сборка в CI и сканер образа"),
    ("secrets", "Секреты в хранилище", "секреты в хранилище"),
    ("monitoring", "Метрики, логи, алерты", "мониторинг и алерты"),
    ("load_test", "Нагрузка ×1,5 от пика", "нагрузочный тест"),
    ("ib_approval", "Согласование ИБ", "согласование ИБ"),
    ("runbook", "Runbook и откат", "runbook и план отката"),
    ("ops_acceptance", "Приёмка сопровождением", "приёмка сопровождением"),
]

# Решения встречи 18.11. Их пишу руками после встречи, всё остальное в протоколе считается из данных.
DECISIONS = [
    "smev-adapter: новый план 04.12, если сертификат СМЭВ будет до 27.11. smev-queue-router зависит от него "
    "и стоял на тот же день, поэтому его план сдвигаем на 11.12.",
    "ebs-gateway: требования ИБ к биометрии в контейнерах согласуем на отдельной встрече 20.11. "
    "Если не договоримся, вопрос уходит руководителю кластера.",
    "gosuslugi-status-push: сетевой доступ к ЕПГУ ждём с 02.11, эскалируем руководителю кластера.",
    "sign-orchestrator: 20.11 в прод не выходит, doc-archive и ukep-adapter не готовы. До 25.11 команда Подписи "
    "выбирает межконтурный доступ на переходный период или перенос на 18.12.",
    "consent-registry: в проде с 13.11, но учения РЦОД не пройдены. Повтор 27.11 после исправлений из разбора.",
    "audit-log и observability-agent: к 02.12 нужен план вывода старого контура.",
]
OPEN_QUESTIONS = [
    "Нужен ли esia-auth-adapter слот на тестовой среде ЕСИА на следующей неделе. Уточняю у РГ Идентификации в чате.",
    "24.11 регламентные работы на тестовой среде СМЭВ: тесты сервисов СМЭВ в этот день не планируем.",
]
# Что нужно от руководителя кластера по блокерам, которые висят дольше 10 рабочих дней
HELP_NEEDED = {
    "S11": "позиция ИБ по обработке биометрии в контейнерах, если 20.11 не договоримся сами",
    "S18": "ускорить выпуск сертификата СМЭВ для нового контура",
    "S20": "согласование заявки на сетевой доступ к ЕПГУ у ИБ",
    "S09": "срок, к которому платформа DropApp выделит GPU-узлы",
}


def read(name: str) -> list[dict]:
    with open(DATA / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def d(value: str) -> date | None:
    return datetime.strptime(value, "%Y-%m-%d").date() if value else None


def workdays(start: date, end: date) -> int:
    """Как NETWORKDAYS в Excel: рабочие дни включая обе даты, без праздников."""
    return sum((start + timedelta(k)).weekday() < 5 for k in range((end - start).days + 1))


def ru_date(x: date) -> str:
    return f"{x:%d.%m.%Y}"


def pp(x: float) -> str:
    return f"{'+' if x >= 0 else '−'}{abs(x) * 100:.1f}".replace(".", ",") + " п.п."


# ------------------------------------------------------------------ данные
teams = read("teams.csv")
services = read("services.csv")
deps = read("dependencies.csv")
migration = {m["service_id"]: m for m in read("migration.csv")}
rcod = {r["service_id"]: r for r in read("rcod.csv")}
checklist = {c["service_id"]: c for c in read("checklist.csv")}
calendar = read("calendar.csv")
actions = read("meeting_actions.csv")
with open(DATA / "history.csv", encoding="utf-8") as f:
    reader = csv.reader(f)
    snap_dates = [d(x) for x in next(reader)[1:]]
    history = {row[0]: row[1:] for row in reader}
assert snap_dates[-1] == REPORT_DATE, "последний срез в history.csv должен совпадать с датой отчёта"
weeks = [f"W{x.isocalendar()[1]}" for x in snap_dates]
WEEK = f"{weeks[-1]} ({REPORT_DATE - timedelta(2):%d}–{REPORT_DATE + timedelta(2):%d.%m.%Y})"

team_name = {t["team_id"]: t["team"] for t in teams}
svc = {s["service_id"]: s for s in services}
name_of = {s["service_id"]: s["service"] for s in services}
last_window = next(d(c["date_from"]) for c in calendar if c["kind"] == "Окно")
moratorium = next((d(c["date_from"]), d(c["date_to"])) for c in calendar if c["kind"] == "Мораторий")


def rcod_status(r: dict) -> str:
    if r["rcod"] == "Да" and r["dr_result"] == "Успех" and int(r["rto_fact_min"]) <= int(r["rto_target_min"]):
        return "Готов"
    if r["rcod"] == "Да" and r["dr_result"] == "Неуспех":
        return "Учения не пройдены"
    if r["rcod"] == "Нет":
        return "Не развёрнут"
    return "В работе"


stage_now = {sid: h[-1] for sid, h in history.items()}


def weeks_on_stage(h: list[str]) -> int:
    n = 1
    while n < len(h) and h[-1 - n] == h[-1]:
        n += 1
    return n


def conflicts_of(sid: str) -> list[str]:
    """Зависимости, которые не в проде и по плану выходят не раньше самого сервиса."""
    plan = d(migration[sid]["plan_prod"])
    return [x["depends_on"] for x in deps if x["service_id"] == sid
            and WEIGHT[stage_now[x["depends_on"]]] < PROD and plan <= d(migration[x["depends_on"]]["plan_prod"])]


def dependents_of(sid: str) -> list[str]:
    """Все, кто прямо или через цепочку зависит от сервиса."""
    out, todo = [], [sid]
    while todo:
        cur = todo.pop()
        for x in deps:
            if x["depends_on"] == cur and x["service_id"] not in out:
                out.append(x["service_id"])
                todo.append(x["service_id"])
    return out


rows = []
for s in services:
    sid, m, h = s["service_id"], migration[s["service_id"]], history[s["service_id"]]
    pct, in_prod_now = WEIGHT[h[-1]], WEIGHT[h[-1]] >= PROD
    plan, fact = d(m["plan_prod"]), d(m["prod_fact"])
    wos = weeks_on_stage(h)
    conf = conflicts_of(sid)
    blocked_wd = workdays(d(m["blocker_since"]), REPORT_DATE) if m["blocker"] and not in_prod_now else None
    signals = []
    if not in_prod_now:
        if plan < REPORT_DATE:
            signals.append("план прошёл")
        if m["blocker"]:
            signals.append("блокер")
        if conf:
            signals.append("зависимость не готова")
        if wos >= STUCK_WEEKS:
            signals.append(f"{wos} нед. на этапе")
    if in_prod_now:
        status = "В проде"
    elif plan < REPORT_DATE:
        status = "Просрочено"
    elif m["blocker"] or conf or wos >= STUCK_WEEKS:
        status = "Под риском"
    else:
        status = "В графике"
    escalation = "" if blocked_wd is None else (
        "Руководитель кластера" if blocked_wd > 10 else "Встреча РГ" if blocked_wd > 5 else "Команда")
    rc = rcod_status(rcod[sid])
    open_items = [label for key, _, label in CHECK_ITEMS if checklist[sid][key] != "Да"]
    debts = open_items + ([] if rc == "Готов" else ["РЦОД не готов" if rc != "Учения не пройдены" else "учения РЦОД не пройдены"])
    rows.append(dict(
        id=sid, service=s["service"], team=team_name[s["team_id"]], team_id=s["team_id"],
        stage=h[-1], stage_prev=h[-2], pct=pct, pct_prev=WEIGHT[h[-2]], weeks_on_stage=wos,
        plan=plan, prod_fact=fact, blocker=m["blocker"], blocked_wd=blocked_wd, conflicts=conf,
        owner=m["owner_role"], status=status, signals=signals, escalation=escalation, rcod=rc,
        open_items=open_items, debts=debts if in_prod_now else [], comment=m["comment"],
        two_contours=(REPORT_DATE - fact).days // 7 if h[-1] == "Прод в DropApp" and fact else None,
        slip=(fact - plan).days if fact else None,
        checklist_done=(len(CHECK_ITEMS) + 1 - len(debts)) / (len(CHECK_ITEMS) + 1),
    ))
by_id = {r["id"]: r for r in rows}

plan_conflicts = [dict(service=r["service"], plan=f"{r['plan']:%Y-%m-%d}", depends_on=name_of[c],
                       dep_plan=migration[c]["plan_prod"]) for r in rows for c in r["conflicts"]]
plan_conflicts.sort(key=lambda c: c["plan"])

n = len(rows)
in_prod = sum(r["pct"] >= PROD for r in rows)
in_prod_prev = sum(r["pct_prev"] >= PROD for r in rows)
progress = sum(r["pct"] for r in rows) / n
progress_prev = sum(r["pct_prev"] for r in rows) / n
by_status = {k: sum(r["status"] == k for r in rows) for k in ("В проде", "В графике", "Под риском", "Просрочено")}
risk_no_blocker = [r for r in rows if r["status"] == "Под риском" and not r["blocker"]]
rcod_ready = sum(r["rcod"] == "Готов" for r in rows)
prod_debts = [r for r in rows if r["debts"]]
tail = [r for r in rows if r["two_contours"] is not None and r["two_contours"] > TAIL_WEEKS]
long_blockers = sorted([r for r in rows if r["escalation"] == "Руководитель кластера"], key=lambda r: -r["blocked_wd"])
moved = [r for r in rows if r["stage"] != r["stage_prev"]]
risks = sorted([r for r in rows if r["status"] in ("Просрочено", "Под риском")],
               key=lambda r: (r["status"] != "Просрочено", r["plan"]))
team_progress = []
for t in teams:
    tr = [r for r in rows if r["team_id"] == t["team_id"]]
    team_progress.append(dict(team=t["team"], services=len(tr), in_prod=sum(r["pct"] >= PROD for r in tr),
                              pct=sum(r["pct"] for r in tr) / len(tr),
                              risk=sum(r["status"] in ("Под риском", "Просрочено") for r in tr)))

# Темп и прогноз
weekly_ready = [sum(WEIGHT[history[s][k]] for s in history) / n for k in range(len(snap_dates))]
weekly_prod = [sum(WEIGHT[history[s][k]] >= PROD for s in history) for k in range(len(snap_dates))]
tempo_pp = (weekly_ready[-1] - weekly_ready[-1 - TEMPO_WEEKS]) / TEMPO_WEEKS
tempo_cnt = (weekly_prod[-1] - weekly_prod[-1 - TEMPO_WEEKS]) / TEMPO_WEEKS


def forecast(remaining: float, tempo: float) -> tuple[date, date]:
    raw = REPORT_DATE + timedelta(days=math.ceil(remaining / tempo * 7))
    real = moratorium[1] + timedelta(1) if moratorium[0] <= raw <= moratorium[1] else raw
    return raw, real


fc_ready, fc_ready_real = forecast(TARGET_READINESS - progress, tempo_pp)
fc_count, fc_count_real = forecast(n - in_prod, tempo_cnt)

# Две волны: в первую идут сервисы без блокеров, у которых все зависимости в проде или тоже в первой волне
wave1 = [r["id"] for r in rows if r["status"] != "В проде" and not r["blocker"]]
changed = True
while changed:
    changed = False
    for sid in list(wave1):
        if any(x["service_id"] == sid and WEIGHT[stage_now[x["depends_on"]]] < PROD and x["depends_on"] not in wave1
               for x in deps):
            wave1.remove(sid)
            changed = True
wave2 = [r["id"] for r in rows if r["status"] != "В проде" and r["id"] not in wave1]

# ------------------------------------------------------------------ SQLite и SQL-запросы
db_path = ROOT / "sql" / "delivery.db"
db_path.unlink(missing_ok=True)
con = sqlite3.connect(db_path)
con.executescript((ROOT / "sql" / "schema.sql").read_text(encoding="utf-8"))
con.executemany("INSERT INTO stages VALUES (?,?,?)", [(name, i, w) for i, (name, w) in enumerate(STAGES)])
con.executemany("INSERT INTO teams VALUES (:team_id,:team,:lead_role,:focus)", teams)
con.executemany("INSERT INTO services VALUES (:service_id,:service,:team_id,:purpose,:criticality,:peak_rps,"
                "NULLIF(:external_system,''))", services)
con.executemany("INSERT INTO dependencies VALUES (:service_id,:depends_on)", deps)
con.executemany("INSERT INTO history VALUES (?,?,?,?)",
                [(weeks[k], f"{snap_dates[k]:%Y-%m-%d}", sid, h[k]) for sid, h in history.items() for k in range(len(h))])
con.executemany("INSERT INTO migration VALUES (:service_id,:plan_prod,NULLIF(:prod_fact,''),:owner_role,"
                "NULLIF(:blocker,''),NULLIF(:blocker_since,''),NULLIF(:comment,''))", migration.values())
con.executemany("INSERT INTO checklist VALUES (?,?,?)",
                [(sid, label, int(c[key] == "Да")) for sid, c in checklist.items() for key, _, label in CHECK_ITEMS])
con.executemany("INSERT INTO rcod VALUES (:service_id,:main_dc,:rcod,NULLIF(:dr_test_date,''),NULLIF(:dr_result,''),"
                ":rto_target_min,NULLIF(:rto_fact_min,''))", rcod.values())
con.executemany("INSERT INTO calendar VALUES (:date_from,:date_to,:event,:kind)", calendar)
con.commit()

queries = (ROOT / "sql" / "queries.sql").read_text(encoding="utf-8").split("-- @")[1:]
md = ["# Результаты SQL-запросов", "",
      f"База `sql/delivery.db` (SQLite), срез на {ru_date(REPORT_DATE)}. Сами запросы с комментариями: "
      "[`sql/queries.sql`](queries.sql).", ""]
for block in queries:
    title, sql = block.split("\n", 1)
    cur = con.execute(sql)
    cols = [c[0] for c in cur.description]
    result = cur.fetchall()
    md += [f"## {title.strip()}", "", "```sql", sql.strip(), "```", "",
           "| " + " | ".join(cols) + " |", "|" + "---|" * len(cols)]
    md += ["| " + " | ".join("" if v is None else str(v) for v in row) + " |" for row in result]
    md.append("")
(ROOT / "sql" / "results.md").write_text("\n".join(md), encoding="utf-8")
con.close()

# ------------------------------------------------------------------ Excel-трекер
FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F2A37")
HEAD_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY = Font(name=FONT, size=10)
INPUT = Font(name=FONT, size=10, color="0000FF")
LINK = Font(name=FONT, size=10, color="008000")
NOTE = Font(name=FONT, size=9, italic=True, color="5F6B7A")
THIN = Side(style="thin", color="D5DAE1")
BORDER = Border(bottom=THIN)
FILLS = {"green": "DCF2E3", "amber": "FFF1CC", "red": "FADBD8", "grey": "EEF0F3"}
M, R, H, CL = "'Миграция DropApp'", "РЦОД", "'История этапов'", "'Чек-лист выхода'"
first, last = 2, n + 1
REPORT = "Сводка!$B$2"


def header(ws, cols: list[tuple[str, int]], row: int = 1) -> None:
    for i, (title, width) in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=title)
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row].height = 42
    ws.freeze_panes = ws.cell(row=row + 1, column=3)


def style_rows(ws, r0: int, r1: int, ncols: int, wrap: tuple[int, ...] = ()) -> None:
    for r in range(r0, r1 + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            if cell.font == Font():
                cell.font = BODY
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=c in wrap)


def color_rule(ws, rng: str, cell: str, text: str, color: str) -> None:
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{cell}="{text}"'],
                                                   fill=PatternFill("solid", fgColor=FILLS[color])))


def status_colors(ws, rng: str, cell: str) -> None:
    for text, color in (("В проде", "green"), ("Готов", "green"), ("ОК", "green"), ("В графике", "grey"),
                        ("В работе", "amber"), ("Под риском", "amber"), ("Просрочено", "red"),
                        ("Не развёрнут", "red"), ("Учения не пройдены", "red"), ("Конфликт плана", "red"),
                        ("Прод с долгами", "red"), ("Скоро выход, не закрыт", "amber"),
                        ("Руководитель кластера", "red"), ("Встреча РГ", "amber")):
        color_rule(ws, rng, cell, text, color)


wb = Workbook()
wb.properties.creator = wb.properties.lastModifiedBy = "Алексей Степин"
wb.properties.title = "Трекер миграции кластера госинтеграций на DropApp"

# Как пользоваться
ws = wb.active
ws.title = "Как пользоваться"
guide = [
    ("Трекер миграции кластера госинтеграций на DropApp", "title"),
    ("Демо на вымышленных данных: 6 команд, 20 микросервисов. Автор: Алексей Степин.", ""),
    ("", ""),
    ("Обновление раз в неделю, по средам до встречи РГ (10–15 минут)", "h"),
    ("1. «Сводка»: поменять дату среза в B2.", ""),
    ("2. «История этапов»: скопировать столбцы D:F в C:E, затем значения столбца G (сейчас) вставить в F как значения. "
     "История сдвинется на неделю.", ""),
    ("3. «Миграция DropApp»: выбрать новые этапы из списка, обновить план, блокер и дату, с которой он висит.", ""),
    ("4. «Чек-лист выхода» и «РЦОД»: отметить, что закрыли за неделю.", ""),
    ("5. Всё остальное считается само. Для протокола и дайджеста: python scripts/build.py.", ""),
    ("", ""),
    ("Как считается статус", "h"),
    ("В проде: этап «Прод в DropApp» или «Старый контур выведен».", ""),
    ("Просрочено: план выхода в прод прошёл, а сервис не в проде.", ""),
    (f"Под риском: есть блокер, ИЛИ зависимость не в проде и выходит не раньше сервиса, ИЛИ сервис "
     f"{STUCK_WEEKS} недели и дольше стоит на одном этапе. Последние два сигнала ловят риск, который команда "
     "не назвала блокером.", ""),
    ("В графике: всё остальное.", ""),
    ("Я специально не даю командам ставить статус руками: «зелёный» статус обычно держится до последнего дня.", ""),
    ("", ""),
    ("Эскалация блокеров (рабочие дни с даты появления блокера)", "h"),
    ("До 5 дней решает команда. 6–10 дней: вопрос на встрече РГ. Больше 10: руководитель кластера, "
     "отдельным пунктом в дайджесте.", ""),
    ("", ""),
    ("Прогноз", "h"),
    (f"Лист «Динамика»: темп за последние {TEMPO_WEEKS} недели по двум метрикам, готовности и числу сервисов в проде. "
     "Прогноз с учётом моратория: если дата попадает в мораторий, она сдвигается на первый день после него.", ""),
    ("Две метрики дают разные даты, и это нормально: первые этапы дешёвые, поэтому готовность растёт быстрее, "
     "чем число сервисов в проде. Правда где-то между, смотреть надо на обе.", ""),
    ("", ""),
    ("Обозначения", "h"),
    ("Синий текст: вводится руками. Чёрный: формулы. Зелёный: ссылки на другие листы.", ""),
]
for i, (text, kind) in enumerate(guide, start=1):
    c = ws.cell(row=i, column=1, value=text)
    c.font = Font(name=FONT, size=14 if kind == "title" else 10, bold=kind in ("title", "h"))
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 130

# Справочник этапов
st = wb.create_sheet("Этапы")
header(st, [("Этап", 26), ("Готовность", 12), ("Порядок", 10)])
for i, (name, w) in enumerate(STAGES, start=2):
    st.cell(row=i, column=1, value=name)
    st.cell(row=i, column=2, value=w).number_format = "0%"
    st.cell(row=i, column=3, value=i - 2)
style_rows(st, 2, len(STAGES) + 1, 3)
st.cell(row=len(STAGES) + 3, column=1, value="Веса этапов мои: прод это 90%, а не 100%, потому что пока жив старый "
        "контур, работа не закончена.").font = NOTE
stage_range = f"Этапы!$A$2:$A${len(STAGES) + 1}"
weight_range = f"Этапы!$B$2:$B${len(STAGES) + 1}"

# Справочник команд
tm = wb.create_sheet("Команды")
header(tm, [("ID", 8), ("Команда", 26), ("Руководитель", 36), ("Зона ответственности", 60)])
for i, t in enumerate(teams, start=2):
    for j, k in enumerate(("team_id", "team", "lead_role", "focus"), start=1):
        tm.cell(row=i, column=j, value=t[k])
style_rows(tm, 2, len(teams) + 1, 4)

# Справочник сервисов
sv = wb.create_sheet("Сервисы")
header(sv, [("ID", 7), ("Сервис", 24), ("Команда", 22), ("Назначение", 50), ("Критичность", 12),
            ("Пиковая нагрузка, RPS", 12), ("Внешняя система", 16)])
for i, s in enumerate(services, start=2):
    sv.cell(row=i, column=1, value=s["service_id"])
    sv.cell(row=i, column=2, value=s["service"])
    sv.cell(row=i, column=3, value=f'=INDEX(Команды!$B$2:$B${len(teams) + 1},'
                                   f'MATCH("{s["team_id"]}",Команды!$A$2:$A${len(teams) + 1},0))').font = LINK
    sv.cell(row=i, column=4, value=s["purpose"])
    sv.cell(row=i, column=5, value=s["criticality"])
    sv.cell(row=i, column=6, value=int(s["peak_rps"])).number_format = "#,##0"
    sv.cell(row=i, column=7, value=s["external_system"] or "")
style_rows(sv, 2, last, 7)

# История этапов: 5 срезов, последний берётся из листа «Миграция DropApp»
hs = wb.create_sheet("История этапов")
hist_titles = ["4 недели назад", "3 недели назад", "2 недели назад", "Неделю назад", "Сейчас"]
header(hs, [("ID", 7), ("Сервис", 24)] + [(t, 20) for t in hist_titles])
dv_hist = DataValidation(type="list", formula1=f"={stage_range}", allow_blank=False)
hs.add_data_validation(dv_hist)
for i, s in enumerate(services, start=2):
    hs.cell(row=i, column=1, value=s["service_id"])
    hs.cell(row=i, column=2, value=f"=INDEX(Сервисы!$B:$B,MATCH(A{i},Сервисы!$A:$A,0))").font = LINK
    for k in range(4):
        hs.cell(row=i, column=3 + k, value=history[s["service_id"]][k]).font = INPUT
        dv_hist.add(f"{get_column_letter(3 + k)}{i}")
    hs.cell(row=i, column=7, value=f"={M}!D{i}").font = LINK
style_rows(hs, 2, last, 7)

# Миграция
mg = wb.create_sheet("Миграция DropApp", 1)
mg_cols = [("ID", 6), ("Сервис", 21), ("Команда", 18), ("Этап сейчас", 20), ("Этап неделю назад", 20),
           ("Готов-ность", 8), ("За неделю", 8), ("Недель на этапе", 8), ("План выхода в прод", 11),
           ("Дней до плана", 8), ("Блокер", 34), ("Блокер с", 11), ("Раб. дней с блокером", 9),
           ("Зависимостей с поздним планом", 11), ("Статус", 12), ("Сигналы риска", 28), ("Эскалация", 15),
           ("Вышел в прод, факт", 11), ("Отклонение от плана, дн.", 10), ("Недель в двух контурах", 10),
           ("Ответственный", 30), ("Комментарий", 50)]
header(mg, mg_cols)
dv = DataValidation(type="list", formula1=f"={stage_range}", allow_blank=False)
mg.add_data_validation(dv)
last_dp = len(deps) + 1
for i, s in enumerate(services, start=2):
    sid, m = s["service_id"], migration[s["service_id"]]
    hist = lambda col: f"{H}!{col}{i}"
    mg.cell(row=i, column=1, value=sid)
    mg.cell(row=i, column=2, value=f"=INDEX(Сервисы!$B:$B,MATCH(A{i},Сервисы!$A:$A,0))").font = LINK
    mg.cell(row=i, column=3, value=f"=INDEX(Сервисы!$C:$C,MATCH(A{i},Сервисы!$A:$A,0))").font = LINK
    mg.cell(row=i, column=4, value=history[sid][-1]).font = INPUT
    dv.add(f"D{i}")
    mg.cell(row=i, column=5, value=f"={hist('F')}").font = LINK
    mg.cell(row=i, column=6, value=f"=IFERROR(INDEX({weight_range},MATCH(D{i},{stage_range},0)),0)").number_format = "0%"
    mg.cell(row=i, column=7, value=f"=F{i}-IFERROR(INDEX({weight_range},MATCH(E{i},{stage_range},0)),0)"
            ).number_format = '+0%;-0%;"–"'
    mg.cell(row=i, column=8, value=f'=IF(F{i}>={PROD},"",IF({hist("G")}<>{hist("F")},1,IF({hist("F")}<>{hist("E")},2,'
                                   f'IF({hist("E")}<>{hist("D")},3,IF({hist("D")}<>{hist("C")},4,5)))))')
    c = mg.cell(row=i, column=9, value=d(m["plan_prod"]))
    c.number_format, c.font = "DD.MM.YYYY", INPUT
    mg.cell(row=i, column=10, value=f'=IF(F{i}>={PROD},"",I{i}-{REPORT})').number_format = "0"
    mg.cell(row=i, column=11, value=m["blocker"] or None).font = INPUT
    c = mg.cell(row=i, column=12, value=d(m["blocker_since"]))
    c.number_format, c.font = "DD.MM.YYYY", INPUT
    mg.cell(row=i, column=13, value=f'=IF(OR(F{i}>={PROD},L{i}=""),"",NETWORKDAYS(L{i},{REPORT}))')
    mg.cell(row=i, column=14, value=f'=IF(F{i}>={PROD},"",COUNTIFS(Зависимости!$A$2:$A${last_dp},B{i},'
                                    f'Зависимости!$F$2:$F${last_dp},"Конфликт плана"))')
    mg.cell(row=i, column=15, value=f'=IF(F{i}>={PROD},"В проде",IF(I{i}<{REPORT},"Просрочено",'
                                    f'IF(OR(K{i}<>"",N{i}>0,H{i}>={STUCK_WEEKS}),"Под риском","В графике")))')
    mg.cell(row=i, column=16, value=f'=IF(F{i}>={PROD},"",MID(IF(I{i}<{REPORT},"; план прошёл","")'
                                    f'&IF(K{i}<>"","; блокер","")&IF(N{i}>0,"; зависимость не готова","")'
                                    f'&IF(H{i}>={STUCK_WEEKS},"; "&H{i}&" нед. на этапе",""),3,200))')
    mg.cell(row=i, column=17, value=f'=IF(M{i}="","",IF(M{i}>10,"Руководитель кластера",'
                                    f'IF(M{i}>5,"Встреча РГ","Команда")))')
    c = mg.cell(row=i, column=18, value=d(m["prod_fact"]))
    c.number_format, c.font = "DD.MM.YYYY", INPUT
    mg.cell(row=i, column=19, value=f'=IF(R{i}="","",R{i}-I{i})').number_format = '+0;-0;0'
    mg.cell(row=i, column=20, value=f'=IF(AND(D{i}="Прод в DropApp",R{i}<>""),INT(({REPORT}-R{i})/7),"")')
    mg.cell(row=i, column=21, value=m["owner_role"]).font = INPUT
    mg.cell(row=i, column=22, value=m["comment"] or None).font = INPUT
style_rows(mg, first, last, len(mg_cols), wrap=(11, 16, 22))
status_colors(mg, f"O2:O{last}", "$O2")
status_colors(mg, f"Q2:Q{last}", "$Q2")
mg.conditional_formatting.add(f"T2:T{last}", FormulaRule(formula=[f'AND(ISNUMBER($T2),$T2>{TAIL_WEEKS})'],
                                                          fill=PatternFill("solid", fgColor=FILLS["amber"])))
mg.conditional_formatting.add(f"H2:H{last}", FormulaRule(formula=[f'AND(ISNUMBER($H2),$H2>={STUCK_WEEKS})'],
                                                          fill=PatternFill("solid", fgColor=FILLS["amber"])))
mg.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=["$G2>0"], font=Font(name=FONT, color="137333", bold=True)))
mg.auto_filter.ref = f"A1:V{last}"

# РЦОД
rc = wb.create_sheet("РЦОД", 2)
header(rc, [("ID", 7), ("Сервис", 22), ("Команда", 20), ("Основной ЦОД", 11), ("РЦОД", 11),
            ("Учения по переключению", 13), ("Результат учений", 12), ("RTO цель, мин", 10),
            ("RTO факт, мин", 10), ("Статус", 18)])
dv2 = DataValidation(type="list", formula1='"Да,Частично,Нет"')
dv3 = DataValidation(type="list", formula1='"Успех,Неуспех"', allow_blank=True)
rc.add_data_validation(dv2)
rc.add_data_validation(dv3)
for i, s in enumerate(services, start=2):
    r = rcod[s["service_id"]]
    rc.cell(row=i, column=1, value=s["service_id"])
    rc.cell(row=i, column=2, value=f"=INDEX(Сервисы!$B:$B,MATCH(A{i},Сервисы!$A:$A,0))").font = LINK
    rc.cell(row=i, column=3, value=f"=INDEX(Сервисы!$C:$C,MATCH(A{i},Сервисы!$A:$A,0))").font = LINK
    rc.cell(row=i, column=4, value=r["main_dc"]).font = INPUT
    rc.cell(row=i, column=5, value=r["rcod"]).font = INPUT
    c = rc.cell(row=i, column=6, value=d(r["dr_test_date"]))
    c.number_format, c.font = "DD.MM.YYYY", INPUT
    rc.cell(row=i, column=7, value=r["dr_result"] or None).font = INPUT
    rc.cell(row=i, column=8, value=int(r["rto_target_min"])).font = INPUT
    rc.cell(row=i, column=9, value=int(r["rto_fact_min"]) if r["rto_fact_min"] else None).font = INPUT
    rc.cell(row=i, column=10, value=f'=IF(AND(E{i}="Да",G{i}="Успех",I{i}<=H{i}),"Готов",'
                                     f'IF(AND(E{i}="Да",G{i}="Неуспех"),"Учения не пройдены",'
                                     f'IF(E{i}="Нет","Не развёрнут","В работе")))')
    dv2.add(f"E{i}")
    dv3.add(f"G{i}")
style_rows(rc, first, last, 10)
status_colors(rc, f"J2:J{last}", "$J2")
rc.auto_filter.ref = f"A1:J{last}"

# Чек-лист выхода в прод
cl = wb.create_sheet("Чек-лист выхода", 3)
header(cl, [("ID", 7), ("Сервис", 21), ("Этап", 18)] + [(t, 11) for _, t, _ in CHECK_ITEMS]
       + [("РЦОД готов", 9), ("Закрыто", 9), ("Проверка", 22)])
dv4 = DataValidation(type="list", formula1='"Да,Нет"')
cl.add_data_validation(dv4)
nk = len(CHECK_ITEMS)
c_rcod, c_done, c_check = 4 + nk, 5 + nk, 6 + nk
L_first, L_rcod, L_done = "D", get_column_letter(c_rcod), get_column_letter(c_done)
for i, s in enumerate(services, start=2):
    sid = s["service_id"]
    cl.cell(row=i, column=1, value=sid)
    cl.cell(row=i, column=2, value=f"={M}!B{i}").font = LINK
    cl.cell(row=i, column=3, value=f"={M}!D{i}").font = LINK
    for k, (key, _, _) in enumerate(CHECK_ITEMS):
        cl.cell(row=i, column=4 + k, value=checklist[sid][key]).font = INPUT
        dv4.add(f"{get_column_letter(4 + k)}{i}")
    cl.cell(row=i, column=c_rcod, value=f'=IF({R}!J{i}="Готов","Да","Нет")').font = LINK
    cl.cell(row=i, column=c_done, value=f'=COUNTIF({L_first}{i}:{L_rcod}{i},"Да")/{nk + 1}').number_format = "0%"
    cl.cell(row=i, column=c_check, value=f'=IF({M}!F{i}>={PROD},IF({L_done}{i}<1,"Прод с долгами","ОК"),'
                                         f'IF(AND({M}!J{i}<={SOON_DAYS},{L_done}{i}<1),"Скоро выход, не закрыт",""))')
style_rows(cl, first, last, c_check)
L_check = get_column_letter(c_check)
status_colors(cl, f"{L_check}2:{L_check}{last}", f"${L_check}2")
for k in range(nk + 1):
    col = get_column_letter(4 + k)
    color_rule(cl, f"{col}2:{col}{last}", f"{col}2", "Да", "green")
cl.cell(row=last + 2, column=1, value=f"«Прод с долгами»: сервис в проде, но пункт чек-листа или РЦОД не закрыт. "
        f"«Скоро выход, не закрыт»: до плана {SOON_DAYS} дней или меньше, а чек-лист не закрыт.").font = NOTE

# Зависимости
dp = wb.create_sheet("Зависимости", 4)
header(dp, [("Сервис", 22), ("Зависит от", 22), ("План сервиса", 13), ("План зависимости", 13),
            ("Готовность зависимости", 12), ("Проверка", 16)])
for i, dep in enumerate(deps, start=2):
    a, b = dep["service_id"], dep["depends_on"]
    mrow = lambda sid: f"MATCH(\"{sid}\",{M}!$A:$A,0)"
    dp.cell(row=i, column=1, value=f"=INDEX({M}!$B:$B,{mrow(a)})").font = LINK
    dp.cell(row=i, column=2, value=f"=INDEX({M}!$B:$B,{mrow(b)})").font = LINK
    dp.cell(row=i, column=3, value=f"=INDEX({M}!$I:$I,{mrow(a)})").number_format = "DD.MM.YYYY"
    dp.cell(row=i, column=4, value=f"=INDEX({M}!$I:$I,{mrow(b)})").number_format = "DD.MM.YYYY"
    dp.cell(row=i, column=5, value=f"=INDEX({M}!$F:$F,{mrow(b)})").number_format = "0%"
    dp.cell(row=i, column=6, value=f'=IF(AND(E{i}<{PROD},C{i}<=D{i}),"Конфликт плана","ОК")')
style_rows(dp, 2, last_dp, 6)
status_colors(dp, f"F2:F{last_dp}", "$F2")
dp.cell(row=last_dp + 2, column=1, value="Конфликт плана: зависимость не в проде и по плану выходит не раньше самого "
        "сервиса. Варианты: межконтурный доступ на переходный период или перенос срока.").font = NOTE

# Календарь
cal = wb.create_sheet("Календарь")
header(cal, [("С", 12), ("По", 12), ("Событие", 52), ("Тип", 14)])
for i, c in enumerate(calendar, start=2):
    for j, v in enumerate((d(c["date_from"]), d(c["date_to"]), c["event"], c["kind"]), start=1):
        cell = cal.cell(row=i, column=j, value=v)
        cell.font = INPUT
        if j <= 2:
            cell.number_format = "DD.MM.YYYY"
style_rows(cal, 2, len(calendar) + 1, 4)
cal.cell(row=len(calendar) + 3, column=1, value="Окна изменений в проде: ночь с четверга на пятницу. "
         "В мораторий в прод ничего не выкатываем.").font = NOTE
row_window = 2 + next(k for k, c in enumerate(calendar) if c["kind"] == "Окно")
row_mor = 2 + next(k for k, c in enumerate(calendar) if c["kind"] == "Мораторий")

# Задачи
tk = wb.create_sheet("Задачи")
header(tk, [("ID", 6), ("Сервис", 22), ("Задача", 70), ("Ответственный", 34), ("Срок", 12), ("Статус", 12)])
for i, a in enumerate(actions, start=2):
    tk.cell(row=i, column=1, value=a["action_id"])
    tk.cell(row=i, column=2, value=name_of[a["service_id"]] if a["service_id"] else "Весь кластер")
    tk.cell(row=i, column=3, value=a["action"]).font = INPUT
    tk.cell(row=i, column=4, value=a["owner_role"]).font = INPUT
    c = tk.cell(row=i, column=5, value=d(a["due"]))
    c.number_format, c.font = "DD.MM.YYYY", INPUT
    tk.cell(row=i, column=6, value=a["status"]).font = INPUT
style_rows(tk, 2, len(actions) + 1, 6, wrap=(3,))

# Сводка
sm = wb.create_sheet("Сводка", 0)
sm.column_dimensions["A"].width = 50
for col, w in zip("BCDEFG", (14, 12, 12, 12, 12, 12)):
    sm.column_dimensions[col].width = w
sm["A1"] = "Сводка по миграции кластера на DropApp"
sm["A1"].font = Font(name=FONT, size=14, bold=True)
DY = "Динамика"
kpis = [  # ключ, подпись, значение, формат, шрифт
    ("report", "Дата среза", REPORT_DATE, "DD.MM.YYYY", INPUT),
    ("deadline", "Срок программы", DEADLINE, "DD.MM.YYYY", INPUT),
    ("window", "Последнее окно изменений до моратория", f"=Календарь!$A${row_window}", "DD.MM.YYYY", LINK),
    ("mor_from", "Мораторий с", f"=Календарь!$A${row_mor}", "DD.MM.YYYY", LINK),
    ("mor_to", "Мораторий по", f"=Календарь!$B${row_mor}", "DD.MM.YYYY", LINK),
    ("n", "Сервисов в программе", f"=COUNTA({M}!$A$2:$A${last})", "0", BODY),
    ("prod", "В проде DropApp", f'=COUNTIF({M}!$F$2:$F${last},">={PROD}")', "0", BODY),
    ("debt", "  из них с долгами (чек-лист или РЦОД)", f'=COUNTIF({CL}!${L_check}$2:${L_check}${last},"Прод с долгами")', "0", BODY),
    ("ready", "Готовность (взвешенная по этапам)", f"=AVERAGE({M}!$F$2:$F${last})", "0.0%", BODY),
    ("gain", "Прирост за неделю", None, '+0.0%;-0.0%;"–"', BODY),
    ("ok", "В графике", f'=COUNTIF({M}!$O$2:$O${last},"В графике")', "0", BODY),
    ("risk", "Под риском", f'=COUNTIF({M}!$O$2:$O${last},"Под риском")', "0", BODY),
    ("risk_nb", "  из них без названного блокера",
     f'=SUMPRODUCT(({M}!$O$2:$O${last}="Под риском")*({M}!$K$2:$K${last}=""))', "0", BODY),
    ("late", "Просрочено", f'=COUNTIF({M}!$O$2:$O${last},"Просрочено")', "0", BODY),
    ("esc", "Блокеров дольше 10 раб. дней (к руководителю кластера)",
     f'=COUNTIF({M}!$Q$2:$Q${last},"Руководитель кластера")', "0", BODY),
    ("conf", "Конфликты плана по зависимостям", f'=COUNTIF(Зависимости!$F$2:$F${last_dp},"Конфликт плана")', "0", BODY),
    ("tail", f"Старый контур дольше {TAIL_WEEKS} недель после прода", f'=COUNTIF({M}!$T$2:$T${last},">{TAIL_WEEKS}")', "0", BODY),
    ("rcod", "Готовы в РЦОД (учения пройдены)", f'=COUNTIF({R}!$J$2:$J${last},"Готов")', "0", BODY),
    ("fc1", "Прогноз 90% готовности, с учётом моратория", f"={DY}!$B$16", "DD.MM.YYYY", LINK),
    ("fc2", "Прогноз: все сервисы в проде, с учётом моратория", f"={DY}!$B$17", "DD.MM.YYYY", LINK),
    ("days", "Дней до последнего окна", None, "0", BODY),
]
kr = {key: i for i, (key, *_) in enumerate(kpis, start=2)}
for key, label, value, fmt, font in kpis:
    i = kr[key]
    if key == "gain":
        value = f"=SUM({M}!$G$2:$G${last})/B{kr['n']}"
    if key == "days":
        value = f"=B{kr['window']}-B{kr['report']}"
    sm.cell(row=i, column=1, value=label).font = BODY
    c = sm.cell(row=i, column=2, value=value)
    c.number_format, c.font = fmt, font
    for col in (1, 2):
        sm.cell(row=i, column=col).border = BORDER
sm.cell(row=2, column=3, value="меняется раз в неделю").font = NOTE
sm.conditional_formatting.add(f"B{kr['fc1']}:B{kr['fc2']}", FormulaRule(
    formula=[f"B{kr['fc1']}>$B${kr['window']}"], fill=PatternFill("solid", fgColor=FILLS["red"])))
assert REPORT == f"Сводка!$B${kr['report']}"

t0 = len(kpis) + 4
for j, title in enumerate(("Команда", "Сервисов", "В проде", "Готовность", "Под риском", "Просрочено"), start=1):
    c = sm.cell(row=t0, column=j, value=title)
    c.font, c.fill = HEAD_FONT, HEAD_FILL
for i, t in enumerate(teams, start=t0 + 1):
    sm.cell(row=i, column=1, value=t["team"]).font = BODY
    sm.cell(row=i, column=2, value=f"=COUNTIF({M}!$C$2:$C${last},A{i})").font = BODY
    sm.cell(row=i, column=3, value=f'=COUNTIFS({M}!$C$2:$C${last},A{i},{M}!$F$2:$F${last},">={PROD}")').font = BODY
    c = sm.cell(row=i, column=4, value=f"=AVERAGEIF({M}!$C$2:$C${last},A{i},{M}!$F$2:$F${last})")
    c.number_format, c.font = "0%", BODY
    sm.cell(row=i, column=5, value=f'=COUNTIFS({M}!$C$2:$C${last},A{i},{M}!$O$2:$O${last},"Под риском")').font = BODY
    sm.cell(row=i, column=6, value=f'=COUNTIFS({M}!$C$2:$C${last},A{i},{M}!$O$2:$O${last},"Просрочено")').font = BODY
    for col in range(1, 7):
        sm.cell(row=i, column=col).border = BORDER
t1 = t0 + len(teams)

chart = BarChart()
chart.type = "bar"
chart.title = "Готовность по командам"
chart.y_axis.number_format = "0%"
chart.y_axis.scaling.min, chart.y_axis.scaling.max = 0, 1
chart.add_data(Reference(sm, min_col=4, min_row=t0, max_row=t1), titles_from_data=True)
chart.set_categories(Reference(sm, min_col=1, min_row=t0 + 1, max_row=t1))
chart.legend = None
chart.roundedCorners = False
chart.title.overlay = False
chart.x_axis.delete = False
chart.y_axis.delete = False
chart.series[0].graphicalProperties.solidFill = "0F766E"
chart.height, chart.width = 7.5, 14
sm.add_chart(chart, "I2")

# Динамика и прогноз
dy = wb.create_sheet(DY, 1)
header(dy, [("Срез", 14), ("В проде", 10), ("Готовность", 12)])
dy.freeze_panes = None
hist_cols = "CDEFG"
for k in range(5):
    i = 2 + k
    col = f"{H}!{hist_cols[k]}$2:{hist_cols[k]}${last}"
    c = dy.cell(row=i, column=1, value=f"=Сводка!$B$2-{7 * (4 - k)}" if k < 4 else "=Сводка!$B$2")
    c.number_format, c.font = "DD.MM.YYYY", LINK
    dy.cell(row=i, column=2, value=f'=COUNTIF({col},"Прод в DropApp")+COUNTIF({col},"Старый контур выведен")')
    dy.cell(row=i, column=3, value=f"=SUMPRODUCT(SUMIF({stage_range},{col},{weight_range}))/COUNTA({col})"
            ).number_format = "0.0%"
style_rows(dy, 2, 6, 3)
dy.column_dimensions["A"].width = 52
forecast_rows = [
    (8, "Прогноз", None, None, True),
    (9, f"Темп готовности за {TEMPO_WEEKS} недели, в неделю", f"=(C6-C{6 - TEMPO_WEEKS})/{TEMPO_WEEKS}", "0.0%", False),
    (10, f"Темп: сервисов в прод в неделю, за {TEMPO_WEEKS} недели", f"=(B6-B{6 - TEMPO_WEEKS})/{TEMPO_WEEKS}", "0.0", False),
    (11, "Цель по готовности (все в проде)", TARGET_READINESS, "0%", False),
    (12, "Прогноз по готовности", '=IF(B9<=0,"темп нулевой",A6+ROUNDUP((B11-C6)/B9*7,0))', "DD.MM.YYYY", False),
    (13, "Прогноз по числу сервисов", f'=IF(B10<=0,"темп нулевой",A6+ROUNDUP((Сводка!$B${kr["n"]}-B6)/B10*7,0))',
     "DD.MM.YYYY", False),
    (14, "Мораторий с", f"=Сводка!$B${kr['mor_from']}", "DD.MM.YYYY", False),
    (15, "Мораторий по", f"=Сводка!$B${kr['mor_to']}", "DD.MM.YYYY", False),
    (16, "Прогноз по готовности с учётом моратория", "=IF(AND(B12>=B14,B12<=B15),B15+1,B12)", "DD.MM.YYYY", False),
    (17, "Прогноз по числу сервисов с учётом моратория", "=IF(AND(B13>=B14,B13<=B15),B15+1,B13)", "DD.MM.YYYY", False),
    (18, "Отставание от последнего окна, дней (по готовности)", f"=B16-Сводка!$B${kr['window']}", "+0;-0;0", False),
    (19, "Отставание от последнего окна, дней (по числу сервисов)", f"=B17-Сводка!$B${kr['window']}", "+0;-0;0", False),
]
for r_, label, value, fmt, bold in forecast_rows:
    dy.cell(row=r_, column=1, value=label).font = Font(name=FONT, size=10, bold=bold)
    if value is not None:
        c = dy.cell(row=r_, column=2, value=value)
        c.number_format = fmt
        c.font = INPUT if r_ == 11 else (LINK if r_ in (14, 15) else BODY)
        for col in (1, 2):
            dy.cell(row=r_, column=col).border = BORDER
dy.cell(row=21, column=1, value="Прогноз линейный, по последним трём неделям. Для 20 сервисов сложнее модель не нужна: "
        "важнее понять порядок отставания.").font = NOTE
dy.column_dimensions["B"].width = 14
line = LineChart()
line.title = "Готовность программы по неделям"
line.add_data(Reference(dy, min_col=3, min_row=1, max_row=6), titles_from_data=True)
line.set_categories(Reference(dy, min_col=1, min_row=2, max_row=6))
line.y_axis.number_format = "0%"
line.y_axis.scaling.min, line.y_axis.scaling.max = 0, 1
line.x_axis.number_format = "DD.MM"
line.legend = None
line.roundedCorners = False
line.title.overlay = False
line.x_axis.delete = False
line.y_axis.delete = False
line.series[0].graphicalProperties.line.solidFill = "0F766E"
line.series[0].graphicalProperties.line.width = 28000
line.height, line.width = 7.5, 14
dy.add_chart(line, "E2")

xlsx_path = ROOT / "tracker" / "delivery_tracker.xlsx"
wb.save(xlsx_path)

# ------------------------------------------------------------------ тексты: общие куски
short = lambda role: role.replace("Руководитель группы ", "РГ ")
svc_list = lambda ids: ", ".join(name_of[x] for x in ids)
team_lines = "\n".join(f"| {t['team']} | {t['services']} | {t['in_prod']} | {t['pct']:.0%} | {t['risk'] or ''} |"
                       for t in team_progress)
risk_lines = "\n".join(
    f"| {r['service']} | {r['status']} | {r['plan']:%d.%m} | {'; '.join(r['signals'])} | "
    f"{r['blocker'] or r['comment'] or ''} | {r['escalation']} |" for r in risks)
action_lines = "\n".join(
    f"| {a['action_id']} | {name_of[a['service_id']] if a['service_id'] else 'Весь кластер'} | {a['action']} | "
    f"{short(a['owner_role'])} | {d(a['due']):%d.%m} |" for a in actions)
tempo_cnt_txt = f"{tempo_cnt:.1f}".replace(".", ",")
tempo_line = f"Средний темп за {TEMPO_WEEKS} недели: {pp(tempo_pp).lstrip('+')} готовности и {tempo_cnt_txt} сервиса в прод в неделю."
weekly_line = ", ".join(f"{weeks[k]} {weekly_ready[k]:.0%}" for k in range(len(weeks)))

# ------------------------------------------------------------------ протокол встречи
notes = [
    f"**Темп.** По неделям готовность шла так: {weekly_line}. {tempo_line} "
    f"При таком темпе 90% готовности будет к {fc_ready:%d.%m}, а это уже мораторий, то есть реально не раньше "
    f"{fc_ready_real:%d.%m.%Y}. По числу сервисов в проде ещё хуже: все {n} будут к {fc_count:%d.%m.%Y}. "
    f"Последнее окно изменений {last_window:%d.%m}. Это не решается на уровне РГ, выношу в дайджест.",
    "**Блокеры дольше 10 рабочих дней:** " + ", ".join(
        f"{r['service']} ({r['blocked_wd']} дн.)" for r in long_blockers) + ". Ни один из них команда не снимет "
    "сама: это ИБ, платформа и внешние согласования.",
    "**Под риском без названного блокера.** " + " ".join(
        f"{r['service']} ({', '.join(r['signals'])}): {r['comment'][0].lower() + r['comment'][1:]}."
        for r in risk_no_blocker),
    "**РЦОД.** " + "; ".join(f"{r['service']} в проде, но учения не пройдены (RTO {rcod[r['id']]['rto_fact_min']} мин "
                             f"при цели {rcod[r['id']]['rto_target_min']})" for r in rows
                             if r["rcod"] == "Учения не пройдены") + ". Разбор: `docs/razbor_ucheniy_consent-registry.md`.",
    f"**Старый контур.** Дольше {TAIL_WEEKS} недель после прода в двух контурах живут "
    + ", ".join(f"{r['service']} ({r['two_contours']} нед.)" for r in tail)
    + ". Две недели держим ради отката, дальше платим за два контура без пользы. Нужен план вывода.",
]
protocol = f"""# Встреча руководителей групп: {ru_date(REPORT_DATE)}

Еженедельно по средам, 45 минут, ведёт Delivery Lead. Были РГ всех шести команд. ИБ на встрече не было,
поэтому вопрос по ebs-gateway разбираем отдельно 20.11.

## Цифры на начало встречи

В проде {in_prod} из {n} (+{in_prod - in_prod_prev} за неделю), готовность {progress:.0%} ({pp(progress - progress_prev)}).
Под риском {by_status['Под риском']}, из них {len(risk_no_blocker)} без названного блокера. Просрочен {by_status['Просрочено']}.
Готовы в РЦОД {rcod_ready} из {n}.

| Команда | Сервисов | В проде | Готовность | Под риском и просрочено |
|---|---|---|---|---|
{team_lines}

## Что обсудили

""" + "\n".join(f"{i}. {x}" for i, x in enumerate(notes, start=1)) + """

## Решения

""" + "\n".join(f"{i}. {x}" for i, x in enumerate(DECISIONS, start=1)) + f"""

## Задачи

| ID | Сервис | Задача | Ответственный | Срок |
|---|---|---|---|---|
{action_lines}

Задачи в Jira (`jira/jira_import.csv`) и на листе «Задачи» трекера.

## Не успели

""" + "\n".join(f"- {x}" for x in OPEN_QUESTIONS) + """

Следующая встреча 25.11.2026.
"""
(ROOT / "meetings" / "2026-11-18_vstrecha_RG.md").write_text(protocol, encoding="utf-8")

# ------------------------------------------------------------------ Jira
risk_ids = {r["id"] for r in risks}
with open(ROOT / "jira" / "jira_import.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Summary", "Issue Type", "Priority", "Assignee", "Due Date", "Component", "Labels", "Description"])
    for a in actions:
        service = name_of[a["service_id"]] if a["service_id"] else "cluster"
        team = team_name[svc[a["service_id"]]["team_id"]] if a["service_id"] else ""
        prio = "High" if a["service_id"] in risk_ids else "Medium"
        w.writerow([f"[{service}] {a['action']}", "Task", prio, a["owner_role"], d(a["due"]).strftime("%d/%m/%Y"),
                    team, "dropapp-migration rg-2026-11-18",
                    f"Решение встречи РГ {ru_date(REPORT_DATE)}. Протокол: meetings/2026-11-18_vstrecha_RG.md"])

# ------------------------------------------------------------------ дайджест
help_items = [
    f"**{r['service']}**: {r['blocker'][0].lower() + r['blocker'][1:]}, {r['blocked_wd']} раб. дней. "
    + (f"От него зависят: {svc_list(dependents_of(r['id']))}. " if dependents_of(r["id"]) else "")
    + f"Нужно: {HELP_NEEDED.get(r['id'], 'решение по блокеру')}."
    for r in long_blockers]
help_items.append(
    f"**Срок программы.** {ru_date(DEADLINE)} попадает в мораторий, последнее окно {last_window:%d.%m}. "
    f"Предлагаю две волны. Первая до {last_window:%d.%m}: {svc_list(wave1)}. У них нет блокеров, и все зависимости "
    f"уже в проде или в этой же волне. Это {len(wave1)} сервиса за {(last_window - REPORT_DATE).days // 7} недели, "
    f"как раз под нынешний темп. Тогда к мораторию в проде будет {in_prod + len(wave1)} из {n}. Остальные "
    f"{len(wave2)} сервисов уходят во вторую волну после {moratorium[1]:%d.%m}.")
moved_lines = "\n".join(f"| {r['service']} | {r['stage_prev']} | {r['stage']} |" for r in moved)
debt_lines = "\n".join(f"- {r['service']}: {', '.join(r['debts'])}." for r in prod_debts)
tail_lines = "\n".join(f"- {r['service']}: {r['two_contours']} нед. в двух контурах. {r['comment']}" for r in tail)
conflict_lines = "\n".join(f"- {c['service']} (план {d(c['plan']):%d.%m}) ждёт {c['depends_on']} "
                           f"(план {d(c['dep_plan']):%d.%m})" for c in plan_conflicts)
failed_dr = [r for r in rows if r["rcod"] == "Учения не пройдены"]
digest = f"""# Дайджест кластера госинтеграций, неделя {WEEK}

Черновик собран из трекера, правлю руками перед отправкой руководителю кластера в пятницу.

**Коротко.** В проде {in_prod} из {n} сервисов, за неделю +{in_prod - in_prod_prev}: {svc_list([r['id'] for r in moved if r['pct'] >= PROD and r['pct_prev'] < PROD])}.
Готовность {progress:.0%} ({pp(progress - progress_prev)} за неделю). Неделя хорошая, но средний темп за {TEMPO_WEEKS} недели
ниже: {pp(tempo_pp).lstrip('+')} в неделю. При нём до моратория не успеваем, нужно решение по сроку (пункт {len(help_items)} ниже).

## Нужна помощь руководителя кластера

""" + "\n".join(f"{i}. {x}" for i, x in enumerate(help_items, start=1)) + f"""

## Прогноз

| Метрика | Сейчас | Темп в неделю | Прогноз | С учётом моратория |
|---|---|---|---|---|
| Готовность до 90% | {progress:.0%} | {pp(tempo_pp).lstrip('+')} | {fc_ready:%d.%m.%Y} | {fc_ready_real:%d.%m.%Y} |
| Все сервисы в проде | {in_prod} из {n} | {tempo_cnt_txt} сервиса | {fc_count:%d.%m.%Y} | {fc_count_real:%d.%m.%Y} |

Две метрики расходятся на {(fc_count_real - fc_ready_real).days} дней. Первые этапы дешёвые, поэтому готовность растёт быстрее, чем число
сервисов в проде, а впереди самые тяжёлые: биометрия, stateful-хранилище, СМЭВ. Я бы ориентировался на вторую цифру.

## Что сдвинулось

| Сервис | Было | Стало |
|---|---|---|
{moved_lines}

## Под риском и просрочено

| Сервис | Статус | План | Сигналы | Причина или комментарий | Эскалация |
|---|---|---|---|---|---|
{risk_lines}

Конфликты плана: сервис выходит в прод не позже своей зависимости.

{conflict_lines}

## В проде, но не до конца

{debt_lines}

Старый контур дольше {TAIL_WEEKS} недель после выхода в прод:

{tail_lines}

## РЦОД

Готовы {rcod_ready} из {n}: развёрнуты и прошли учения. """ + " ".join(
    f"{r['service']} не прошёл учения: RTO {rcod[r['id']]['rto_fact_min']} мин при цели {rcod[r['id']]['rto_target_min']}, "
    f"повтор 27.11." for r in failed_dr) + """

## Решения встречи РГ

""" + "\n".join(f"- {x}" for x in DECISIONS) + "\n"
(ROOT / "digest" / "2026-W47_digest.md").write_text(digest, encoding="utf-8")

payload = dict(
    week=WEEK, report_date=ru_date(REPORT_DATE), deadline=ru_date(DEADLINE), n=n, in_prod=in_prod,
    in_prod_prev=in_prod_prev, progress=round(progress, 4), progress_prev=round(progress_prev, 4),
    by_status=by_status, risk_no_blocker=len(risk_no_blocker), rcod_ready=rcod_ready,
    teams=[dict(team=t["team"], pct=round(t["pct"], 4), in_prod=t["in_prod"], services=t["services"]) for t in team_progress],
    risks=[dict(service=r["service"], status=r["status"], plan=f"{r['plan']:%d.%m}", signals="; ".join(r["signals"]),
                reason=r["blocker"] or r["comment"] or "", escalation=r["escalation"]) for r in risks],
    moved=[dict(service=r["service"], prev=r["stage_prev"], now=r["stage"]) for r in moved],
    conflicts=plan_conflicts, decisions=DECISIONS,
    actions=[dict(service=name_of[a["service_id"]] if a["service_id"] else "Весь кластер",
                  action=a["action"], owner=a["owner_role"], due=f"{d(a['due']):%d.%m}") for a in actions],
    history=dict(weeks=weeks, dates=[f"{x:%d.%m}" for x in snap_dates], ready=[round(x, 4) for x in weekly_ready],
                 prod=weekly_prod),
    forecast=dict(tempo_pp=round(tempo_pp, 4), tempo_cnt=round(tempo_cnt, 2), ready=f"{fc_ready:%d.%m}",
                  ready_real=f"{fc_ready_real:%d.%m}", count=f"{fc_count:%d.%m}", count_real=f"{fc_count_real:%d.%m}",
                  last_window=f"{last_window:%d.%m}", moratorium=f"{moratorium[0]:%d.%m}–{moratorium[1]:%d.%m}"),
    wave1=[name_of[x] for x in wave1], wave2=[name_of[x] for x in wave2],
    help=[dict(service=r["service"], days=r["blocked_wd"], need=HELP_NEEDED.get(r["id"], "")) for r in long_blockers],
    prod_debts=[dict(service=r["service"], debts=", ".join(r["debts"])) for r in prod_debts],
    tail=[dict(service=r["service"], weeks=r["two_contours"]) for r in tail],
)
(ROOT / "digest" / "digest_data.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

(ROOT / "diagrams" / "status.json").write_text(json.dumps(
    {r["id"]: dict(service=r["service"], team_id=r["team_id"], stage=r["stage"], status=r["status"],
                   debt=bool(r["debts"])) for r in rows}, ensure_ascii=False, indent=2), encoding="utf-8")

print(f"В проде {in_prod}/{n} (с долгами {len(prod_debts)}), готовность {progress:.2%} (было {progress_prev:.2%}), "
      f"статусы {by_status}, без блокера {len(risk_no_blocker)}, РЦОД {rcod_ready}, конфликтов {len(plan_conflicts)}, "
      f"хвост {len(tail)}, к РК {len(long_blockers)}")
print(f"темп {tempo_pp:.4f}/нед и {tempo_cnt:.2f} серв/нед; прогноз {fc_ready} ({fc_ready_real}) и {fc_count} "
      f"({fc_count_real}); волна 1: {svc_list(wave1)}; волна 2: {len(wave2)}")
