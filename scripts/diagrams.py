"""Карта команд и систем и BPMN-схема миграции: файлы draw.io и PNG из одной геометрии.

Запуск: python scripts/diagrams.py (после scripts/build.py, который пишет diagrams/status.json).
"""
from __future__ import annotations

import csv
import json
import subprocess
from html import escape
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "diagrams"
FONT = "Arial, Helvetica, sans-serif"
INK, MUTED, LINE = "#1F2A37", "#5F6B7A", "#9AA5B1"
STATUS_STYLE = {
    "В проде": ("#DCF2E3", "#2E7D4F"),
    "В графике": ("#FFFFFF", "#9AA5B1"),
    "Под риском": ("#FFF1CC", "#B7791F"),
    "Просрочено": ("#FADBD8", "#C0392B"),
}


class Diagram:
    """Минимальная модель схемы: фигуры, тексты и рёбра с точками излома."""

    def __init__(self, width: int, height: int):
        self.w, self.h = width, height
        self.shapes, self.edges = [], []

    def shape(self, sid, kind, x, y, w, h, label="", fill="#FFFFFF", stroke=INK, font=12, bold=False,
              color=INK, label_pos="center", sub="", stroke_w=1.2, align="center", dashed=False):
        self.shapes.append(dict(id=sid, kind=kind, x=x, y=y, w=w, h=h, label=label, fill=fill, stroke=stroke,
                                font=font, bold=bold, color=color, label_pos=label_pos, sub=sub,
                                stroke_w=stroke_w, align=align, dashed=dashed))
        return self.shapes[-1]

    def edge(self, src, dst, points, label="", label_at=None, color=INK, dashed=False):
        self.edges.append(dict(src=src, dst=dst, points=points, label=label, label_at=label_at,
                               color=color, dashed=dashed))

    # ---------------------------------------------------------------- draw.io
    def drawio(self, name: str) -> str:
        cells = ['<mxCell id="0"/>', '<mxCell id="1" parent="0"/>']
        for s in self.shapes:
            base = {"rect": "rounded=0;", "round": "rounded=1;arcSize=14;", "rhombus": "rhombus;",
                    "ellipse": "ellipse;", "text": "text;strokeColor=none;fillColor=none;"}[s["kind"]]
            pos = {"top": "verticalLabelPosition=top;verticalAlign=bottom;",
                   "bottom": "verticalLabelPosition=bottom;verticalAlign=top;",
                   "center": "verticalAlign=middle;"}[s["label_pos"]]
            style = (f"{base}whiteSpace=wrap;html=1;fillColor={s['fill']};strokeColor={s['stroke']};"
                     f"strokeWidth={s['stroke_w']};fontSize={s['font']};fontColor={s['color']};"
                     f"fontStyle={1 if s['bold'] else 0};align={s['align']};spacingLeft=6;{pos}"
                     f"{'dashed=1;' if s['dashed'] else ''}")
            value = escape(s["label"]) + (f"<br><font style=\"font-size:10px\" color=\"{MUTED}\">{escape(s['sub'])}</font>"
                                          if s["sub"] else "")
            cells.append(f'<mxCell id="{s["id"]}" value="{escape(value)}" style="{escape(style)}" vertex="1" parent="1">'
                         f'<mxGeometry x="{s["x"]}" y="{s["y"]}" width="{s["w"]}" height="{s["h"]}" as="geometry"/></mxCell>')
        for i, e in enumerate(self.edges):
            pts = "".join(f'<mxPoint x="{x}" y="{y}"/>' for x, y in e["points"][1:-1])
            style = (f"edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;endArrow=block;endFill=1;"
                     f"strokeColor={e['color']};fontSize=11;{'dashed=1;' if e['dashed'] else ''}")
            cells.append(f'<mxCell id="e{i}" value="{escape(e["label"])}" style="{escape(style)}" edge="1" parent="1" '
                         f'source="{e["src"]}" target="{e["dst"]}"><mxGeometry relative="1" as="geometry">'
                         f'<Array as="points">{pts}</Array></mxGeometry></mxCell>')
        return (f'<mxfile host="app.diagrams.net"><diagram name="{escape(name)}" id="{name}">'
                f'<mxGraphModel dx="{self.w}" dy="{self.h}" grid="1" gridSize="10" page="1" pageWidth="{self.w}" '
                f'pageHeight="{self.h}" background="#FFFFFF"><root>{"".join(cells)}</root></mxGraphModel></diagram></mxfile>')

    # ---------------------------------------------------------------- SVG
    @staticmethod
    def wrap(text: str, width: float, size: int) -> list[str]:
        limit = max(4, int(width / (size * 0.56)))
        lines, cur = [], ""
        for word in text.split():
            if cur and len(cur) + 1 + len(word) > limit:
                lines.append(cur)
                cur = word
            else:
                cur = f"{cur} {word}".strip()
        return lines + ([cur] if cur else [])

    def svg(self) -> str:
        out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.w}" height="{self.h}" '
               f'viewBox="0 0 {self.w} {self.h}" font-family="{FONT}">',
               f'<rect width="{self.w}" height="{self.h}" fill="#FFFFFF"/>',
               '<defs><marker id="arr" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" '
               'orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="context-stroke"/></marker></defs>']
        for s in self.shapes:
            x, y, w, h = s["x"], s["y"], s["w"], s["h"]
            dash = ' stroke-dasharray="5 4"' if s["dashed"] else ""
            common = f'fill="{s["fill"]}" stroke="{s["stroke"]}" stroke-width="{s["stroke_w"]}"{dash}'
            if s["kind"] in ("rect", "round"):
                r = 10 if s["kind"] == "round" else 0
                out.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" {common}/>')
            elif s["kind"] == "rhombus":
                out.append(f'<polygon points="{x + w / 2},{y} {x + w},{y + h / 2} {x + w / 2},{y + h} {x},{y + h / 2}" {common}/>')
            elif s["kind"] == "ellipse":
                out.append(f'<ellipse cx="{x + w / 2}" cy="{y + h / 2}" rx="{w / 2}" ry="{h / 2}" {common}/>')
            if not s["label"]:
                continue
            size = s["font"]
            lines = self.wrap(s["label"], w - 12 if s["label_pos"] == "center" else w + 60, size)
            subs = self.wrap(s["sub"], w - 12, 10) if s["sub"] else []
            block = len(lines) * size * 1.25 + len(subs) * 12.5
            if s["label_pos"] == "top":
                ty = y - 8 - block + size
            elif s["label_pos"] == "bottom":
                ty = y + h + 6 + size
            else:
                ty = y + (h - block) / 2 + size * 0.95
            anchor, tx = ("start", x + 10) if s["align"] == "left" else ("middle", x + w / 2)
            weight = "700" if s["bold"] else "400"
            for i, line in enumerate(lines):
                out.append(f'<text x="{tx}" y="{ty + i * size * 1.25:.1f}" font-size="{size}" font-weight="{weight}" '
                           f'fill="{s["color"]}" text-anchor="{anchor}">{escape(line)}</text>')
            for j, line in enumerate(subs):
                out.append(f'<text x="{tx}" y="{ty + len(lines) * size * 1.25 + j * 12.5:.1f}" font-size="10" '
                           f'fill="{MUTED}" text-anchor="{anchor}">{escape(line)}</text>')
        for e in self.edges:
            pts = " ".join(f"{x},{y}" for x, y in e["points"])
            dash = ' stroke-dasharray="5 4"' if e["dashed"] else ""
            out.append(f'<polyline points="{pts}" fill="none" stroke="{e["color"]}" stroke-width="1.4"{dash} '
                       f'marker-end="url(#arr)"/>')
            if e["label"]:
                lx, ly = e["label_at"]
                out.append(f'<text x="{lx}" y="{ly}" font-size="11" font-weight="700" fill="{e["color"]}">'
                           f'{escape(e["label"])}</text>')
        out.append("</svg>")
        return "\n".join(out)

    def save(self, stem: str, name: str) -> None:
        (OUT / f"{stem}.drawio").write_text(self.drawio(name), encoding="utf-8")
        (OUT / f"{stem}.svg").write_text(self.svg(), encoding="utf-8")


def right(s): return s["x"] + s["w"], s["y"] + s["h"] / 2
def left(s): return s["x"], s["y"] + s["h"] / 2
def top(s): return s["x"] + s["w"] / 2, s["y"]
def bottom(s): return s["x"] + s["w"] / 2, s["y"] + s["h"]


# ====================================================================== карта команд и систем
def cluster_map() -> None:
    status = json.loads((OUT / "status.json").read_text(encoding="utf-8"))
    teams = list(csv.DictReader(open(ROOT / "data" / "teams.csv", encoding="utf-8")))
    services = list(csv.DictReader(open(ROOT / "data" / "services.csv", encoding="utf-8")))
    order = {  # сервис с внешней системой стоит справа, шлюз Платформы слева
        "PLAT": ["S01", "S02", "S03", "S04"], "ID": ["S05", "S06", "S07"], "BIO": ["S08", "S10", "S09", "S11"],
        "SIGN": ["S12", "S13", "S14", "S15"], "SMEV": ["S16", "S17", "S18"], "NOTIFY": ["S19", "S20"],
    }
    ext = {s["service_id"]: s["external_system"] for s in services if s["external_system"]}
    ext_full = {"ЕСИА": "ЕСИА", "ЕБС": "ЕБС (биометрия)", "Госключ": "Госключ / УКЭП", "СМЭВ": "СМЭВ 4",
                "ЕПГУ": "ЕПГУ (Госуслуги)"}
    lane_h, top0, svc_w, svc_h, gap = 112, 96, 196, 54, 22
    x0, ext_x = 236, 236 + 4 * (svc_w + gap) + 70
    width, height = ext_x + 220, top0 + len(teams) * lane_h + 150
    dg = Diagram(width, height)
    dg.shape("title", "text", 24, 18, 1100, 30, "Карта команд и систем кластера госинтеграций", font=20, bold=True,
             align="left")
    dg.shape("subtitle", "text", 24, 50, 1300, 22,
             "Срез на 18.11.2026. Цвет сервиса показывает статус миграции на DropApp. Данные вымышленные.",
             font=12, color=MUTED, align="left")
    dg.shape("channels", "round", 24, top0, 170, len(teams) * lane_h - 10, "Продукты банка",
             sub="СберБанк Онлайн, СберБизнес, офисы", fill="#EEF2F6", stroke="#7B8794", font=13, bold=True)
    for i, t in enumerate(teams):
        ly = top0 + i * lane_h
        dg.shape(f"lane_{t['team_id']}", "rect", 214, ly, ext_x - 214 - 40, lane_h - 10, fill="#F7F8FA",
                 stroke="#E1E5EA")
        ids = order[t["team_id"]]
        k = sum(status[x]["status"] == "В проде" for x in ids)
        dg.shape(f"lane_label_{t['team_id']}", "text", 222, ly + 4, 700, 22,
                 f"{t['team']}: в проде {k} из {len(ids)}", font=12, bold=True, color=INK, align="left")
        for j, sid in enumerate(order[t["team_id"]]):
            st = status[sid]
            fill, stroke = STATUS_STYLE[st["status"]]
            debt = st["debt"]
            box = dg.shape(sid, "round", x0 + j * (svc_w + gap), ly + 36, svc_w, svc_h, st["service"],
                           sub="Прод в DropApp, есть долги" if debt else st["stage"], fill=fill, stroke=stroke,
                           font=12, bold=True, stroke_w=2 if st["status"] == "Просрочено" or debt else 1.2,
                           dashed=debt)
            if sid in ext:
                e = dg.shape(f"ext_{sid}", "rect", ext_x, ly + 36, 190, svc_h, ext_full[ext[sid]], sub="Внешняя ГИС",
                             fill="#FFFFFF", stroke=INK, font=12, bold=True, stroke_w=1.6)
                dg.edge(sid, e["id"], [right(box), left(e)], color=INK)
            if sid == "S01":
                ch = dg.shapes[2]
                y = box["y"] + svc_h / 2
                dg.edge("channels", sid, [(ch["x"] + ch["w"], y), left(box)], color=INK)
    ly = top0 + len(teams) * lane_h + 14
    dg.shape("legend_title", "text", 24, ly, 300, 22, "Статус миграции на DropApp", font=12, bold=True, align="left")
    legend = list(STATUS_STYLE.items()) + [("В проде, но открыт чек-лист или РЦОД", STATUS_STYLE["В проде"])]
    for k, (name, (fill, stroke)) in enumerate(legend):
        x = 24 + k * 190
        last_item = k == len(legend) - 1
        dg.shape(f"leg_{k}", "round", x, ly + 30, 34, 22, fill=fill, stroke=stroke,
                 stroke_w=2 if name == "Просрочено" or last_item else 1.2, dashed=last_item)
        dg.shape(f"leg_t_{k}", "text", x + 40, ly + 30, 330 if last_item else 140, 22, name, font=12, align="left")
    dg.shape("legend_note", "text", 24, ly + 64, 1400, 22,
             "Связи между внутренними сервисами и конфликты плана по зависимостям: лист «Зависимости» "
             "трекера и SQL-запрос 4.", font=11, color=MUTED, align="left")
    dg.save("cluster_map", "Карта команд и систем")


# ====================================================================== BPMN миграции
def bpmn() -> None:
    lanes = ["Delivery Lead", "Команда сервиса", "ИБ и архитектура", "Платформа DropApp", "Сопровождение"]
    lane_h, top0, col0, step = 124, 92, 236, 150
    tw, th = 126, 64
    width, height = col0 + 15 * step + 70, top0 + len(lanes) * lane_h + 30
    dg = Diagram(width, height)
    dg.shape("title", "text", 24, 18, 1200, 30, "Процесс миграции сервиса на DropApp (BPMN)", font=20, bold=True,
             align="left")
    dg.shape("subtitle", "text", 24, 50, 1500, 22,
             "От включения сервиса в план до вывода старого контура. Роли по дорожкам.", font=12,
             color=MUTED, align="left")
    for i, name in enumerate(lanes):
        y = top0 + i * lane_h
        dg.shape(f"lane{i}", "rect", 24, y, width - 48, lane_h, fill="#F7F8FA" if i % 2 == 0 else "#FFFFFF",
                 stroke="#D5DAE1")
        dg.shape(f"lane{i}_h", "rect", 24, y, 140, lane_h, name, fill="#EEF2F6", stroke="#D5DAE1", font=12,
                 bold=True)

    def cx(col): return col0 + col * step
    def cy(lane): return top0 + lane * lane_h + lane_h / 2

    def task(sid, lane, col, label):
        return dg.shape(sid, "round", cx(col) - tw / 2, cy(lane) - th / 2, tw, th, label, font=11,
                        fill="#FFFFFF", stroke=INK)

    def gate(sid, lane, col, label, pos="bottom"):
        return dg.shape(sid, "rhombus", cx(col) - 26, cy(lane) - 26, 52, 52, label, label_pos=pos, font=11,
                        bold=True, fill="#FFF7E0", stroke="#B7791F")

    def event(sid, lane, col, label, end=False):
        return dg.shape(sid, "ellipse", cx(col) - 18, cy(lane) - 18, 36, 36, label, label_pos="bottom", font=11,
                        fill="#FFFFFF", stroke="#C0392B" if end else "#2E7D4F", stroke_w=3 if end else 1.6)

    def note(sid, x, y, w, h, label):
        return dg.shape(sid, "rect", x, y, w, h, label, font=11, fill="#FFFDF5", stroke="#B7791F", dashed=True,
                        color=INK)

    s = event("start", 0, 0, "Сервис включён в план")
    t1 = task("t1", 1, 1, "Анализ: зависимости, данные, сети")
    t2 = task("t2", 1, 2, "Контейнеризация, Helm-чарты, CI/CD")
    t3 = task("t3", 0, 3, "Заявки на доступы, сертификаты и ресурсы")
    t4 = task("t4", 2, 4, "Проверка требований ИБ и архитектуры")
    g1 = gate("g1", 2, 5, "Согласовано?")
    t5 = task("t5", 1, 5, "Устранение замечаний")
    t6 = task("t6", 3, 6, "Развёртывание в тестовом контуре")
    t7 = task("t7", 1, 7, "Нагрузочное и регрессионное тестирование")
    g2 = gate("g2", 0, 8, "Чек-лист закрыт?", pos="top")
    t8 = task("t8", 3, 9, "Выход в прод в окно изменений")
    t9 = task("t9", 3, 10, "Развёртывание в РЦОД, учения по переключению")
    g3 = gate("g3", 3, 11, "RTO в норме?")
    t10 = task("t10", 1, 11, "Разбор учений, исправления")
    t11 = task("t11", 4, 12, "Приёмка после 2 недель без инцидентов")
    t12 = task("t12", 1, 13, "Вывод старого контура: данные, доступы, ВМ")
    t13 = task("t13", 0, 14, "Обновить трекер, карту и дайджест")
    e = event("end", 0, 15, "Миграция завершена", end=True)
    n1 = note("n1", cx(4) - 60, cy(0) - 34, 270, 68,
              "Блокер дольше 5 раб. дней: на встречу РГ. Дольше 10: руководителю кластера")
    n2 = note("n2", cx(8) - 40, cy(4) - 30, 300, 60,
              "Окно: ночь с четверга на пятницу. С 21.12 по 11.01 мораторий, в прод ничего не выкатываем")

    def link(a, b, label="", label_at=None):
        (ax, ay), (bx, by) = right(a), left(b)
        pts = [(ax, ay), (bx, by)] if abs(ay - by) < 1 else [(ax, ay), ((ax + bx) / 2, ay), ((ax + bx) / 2, by), (bx, by)]
        dg.edge(a["id"], b["id"], pts, label, label_at)

    link(s, t1)
    link(t1, t2)
    link(t2, t3)
    link(t3, t4)
    link(t4, g1)
    dg.edge("g1", "t5", [top(g1), bottom(t5)], "Нет", (top(g1)[0] + 8, top(g1)[1] - 22))
    dg.edge("t5", "t4", [left(t5), (top(t4)[0], left(t5)[1]), top(t4)])
    gx, gy = right(g1)
    dg.edge("g1", "t6", [(gx, gy), (left(t6)[0] - 30, gy), (left(t6)[0] - 30, left(t6)[1]), left(t6)], "Да",
            (gx + 8, gy - 8))
    link(t6, t7)
    dg.edge("t7", "g2", [top(t7), (top(t7)[0], cy(0)), left(g2)])
    dg.edge("g2", "t7", [bottom(g2), (bottom(g2)[0], cy(1)), right(t7)], "Нет", (bottom(g2)[0] + 8, bottom(g2)[1] + 24))
    gx, gy = right(g2)
    dg.edge("g2", "t8", [(gx, gy), (gx + 24, gy), (gx + 24, cy(3)), left(t8)], "Да", (gx + 6, gy - 8))
    link(t8, t9)
    link(t9, g3)
    dg.edge("g3", "t10", [top(g3), bottom(t10)], "Нет", (top(g3)[0] + 8, top(g3)[1] - 22))
    dg.edge("t10", "t9", [left(t10), (top(t9)[0], left(t10)[1]), top(t9)])
    gx, gy = right(g3)
    dg.edge("g3", "t11", [(gx, gy), (gx + 24, gy), (gx + 24, cy(4)), left(t11)], "Да", (gx + 6, gy - 8))
    link(t11, t12)
    link(t12, t13)
    link(t13, e)
    dg.edge("t3", "n1", [top(t3), (top(t3)[0], cy(0) - 52), (n1["x"] + 40, cy(0) - 52), (n1["x"] + 40, n1["y"])],
            color=LINE, dashed=True)
    dg.edge("t8", "n2", [bottom(t8), (bottom(t8)[0], n2["y"])], color=LINE, dashed=True)
    dg.save("migration_bpmn", "Миграция на DropApp")


def render_png() -> None:
    js = """
const {chromium} = require('playwright');
(async () => {
  const b = await chromium.launch();
  for (const stem of process.argv.slice(2)) {
    const p = await b.newPage({deviceScaleFactor: 2});
    await p.goto('file://' + stem + '.svg');
    const box = await p.locator('svg').boundingBox();
    await p.setViewportSize({width: Math.ceil(box.width), height: Math.ceil(box.height)});
    await p.screenshot({path: stem + '.png', clip: {x: 0, y: 0, width: box.width, height: box.height}});
  }
  await b.close();
})();
"""
    script = OUT / "_render.js"
    script.write_text(js, encoding="utf-8")
    subprocess.run(["node", str(script), str(OUT / "cluster_map"), str(OUT / "migration_bpmn")], check=True)
    script.unlink()


if __name__ == "__main__":
    cluster_map()
    bpmn()
    render_png()
    for stem in ("cluster_map", "migration_bpmn"):
        (OUT / f"{stem}.svg").unlink()
    print("ok")
