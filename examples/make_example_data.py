# -*- coding: utf-8 -*-
"""生成两份「完全虚构」的成绩报表，用于演示技能，不含任何真实数据。

结构刻意做成最常见的样子：
  第 1 行 = 分组名（合并单元格，只有首格有值）
  第 2 行 = 子表头（分数(原始) / 赋分 / 班名 / 校名 …）
  第 3 行起 = 数据

用法：
    python examples/make_example_data.py
产出：
    examples/样例_A期.xlsx    examples/样例_B期.xlsx
"""
import os
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
RNG = np.random.RandomState(20260301)

SCHOOLS = [("示例一中", "甲区", 0.9), ("示例二中", "乙区", 0.4), ("示例三中", "甲区", 0.0),
           ("示例四中", "丙区", -0.5), ("示例五中", "乙区", -1.0)]
PER_SCHOOL = 120

CORE = [("语文", 150, 95.0, 12.0), ("数学", 150, 75.0, 20.0),
        ("英语", 150, 85.0, 17.0), ("物理", 100, 55.0, 16.0)]
ELEC = [("化学", 100, 60.0, 15.0), ("生物", 100, 62.0, 14.0)]
XUANKE = ["物化生"]   # 示例统一选科，便于跑通；真实数据可能混合多种选科
SHIFT = {"语文": 8.0, "数学": 2.0, "英语": -1.0, "物理": 0.0, "化学": 0.0, "生物": 0.0}


def cohort():
    students = []
    seq = 0
    for si, (school, county, eff) in enumerate(SCHOOLS):
        for _ in range(PER_SCHOOL):
            seq += 1
            students.append({
                "county": county, "school": school, "cls": RNG.randint(1, 7),
                "name": f"同学{seq:03d}",
                "sid": f"9{si + 1}{20260000 + seq:08d}"[:12],
                "xk": RNG.choice(XUANKE),
                "ability": RNG.normal(0, 1) + eff,
            })
    return students


def scores(students, period):
    for st in students:
        st[period] = {}
        for name, full, mean, sd in CORE + ELEC:
            if name in ("化学", "生物") and name[0] not in st["xk"]:
                st[period][name] = None
                continue
            z = 0.72 * st["ability"] + 0.69 * RNG.normal(0, 1)
            v = mean + sd * z
            if period == "B":
                v += SHIFT[name] + RNG.normal(0, 0.3) * sd
            st[period][name] = float(np.clip(round(v), 0, full))
        for name in ("化学", "生物"):
            st[period][name + "_asg"] = None

    # 赋分：按选考该科者内的百分位折算 30~100
    for name in ("化学", "生物"):
        vals = [(st[period][name], st) for st in students if st[period][name] is not None]
        vals.sort(key=lambda t: t[0])
        n = len(vals)
        for i, (_, st) in enumerate(vals):
            st[period][name + "_asg"] = int(round(30 + 70 * (i + 0.5) / n))

    for st in students:
        p = st[period]
        raw = sum(p[n] for n, *_ in CORE)
        raw += sum(p[n] for n, *_ in ELEC if p[n] is not None)
        asg = sum(p[n] for n, *_ in CORE)
        asg += sum(p[n + "_asg"] for n, *_ in ELEC if p[n + "_asg"] is not None)
        p["raw_total"] = int(round(raw))
        p["asg_total"] = int(round(asg))


def ranks(students, period):
    def rk(key, scope=None):
        groups = {}
        for st in students:
            groups.setdefault(st[scope] if scope else 0, []).append(st)
        out = {}
        for g in groups.values():
            g = [s for s in g if key(s) is not None]
            g.sort(key=lambda s: -key(s))
            for i, s in enumerate(g, 1):
                out[id(s)] = i
        return out

    for name in ("语文", "数学", "英语", "物理", "化学", "生物"):
        key = (lambda s, n=name: s[period].get(n + "_asg") if n in ("化学", "生物")
               else s[period].get(n))
        for scope, field in ((None, "all"), ("cls", "cls"), ("school", "sch")):
            m = rk(key, scope)
            for st in students:
                st[period][f"{name}_{field}"] = m.get(id(st))

    m = rk(lambda s: s[period]["asg_total"])
    mc = rk(lambda s: s[period]["asg_total"], "cls")
    ms = rk(lambda s: s[period]["asg_total"], "school")
    for st in students:
        st[period]["tot"] = {"all": m.get(id(st)), "cls": mc.get(id(st)), "sch": ms.get(id(st))}


def write(path, students, period):
    wb = Workbook()
    ws = wb.active
    ws.title = "成绩"

    spec = [("区", 1, None), ("学校", 1, None), ("班级", 1, None),
            ("姓名", 1, None), ("学号", 1, None), ("选科", 1, None)]
    for name, *_ in CORE:
        spec.append((name, 5, ["分数(原始)", "班名", "校名", "市名", "联名"]))
    for name, *_ in ELEC:
        spec.append((name, 6, ["分数(原始)", "分数(赋分)", "班名", "校名", "市名", "联名"]))
    spec.append(("总分", 6, ["原始总分", "赋分总分", "班名", "校名", "市名", "联名"]))

    col = 1
    for group, n, subs in spec:
        if subs is None:
            ws.cell(row=1, column=col, value=group)
            ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        else:
            ws.cell(row=1, column=col, value=group)
            if n > 1:
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + n - 1)
            for k, sub in enumerate(subs):
                ws.cell(row=2, column=col + k, value=sub)
        col += n

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="DDEBF7")

    r = 3
    for st in students:
        p = st[period]
        vals = [st["county"], st["school"], st["cls"], st["name"], st["sid"], st["xk"]]
        for name, *_ in CORE:
            vals += [p[name], p[f"{name}_cls"], p[f"{name}_sch"], p[f"{name}_all"], None]
        for name, *_ in ELEC:
            vals += [p[name], p[name + "_asg"], p[f"{name}_cls"], p[f"{name}_sch"],
                     p[f"{name}_all"], None]
        vals += [p["raw_total"], p["asg_total"], p["tot"]["cls"], p["tot"]["sch"],
                 p["tot"]["all"], None]
        for c, v in enumerate(vals, 1):
            ws.cell(row=r, column=c, value=v)
        r += 1

    for c in range(1, col):
        ws.column_dimensions[get_column_letter(c)].width = 10
    ws.freeze_panes = "G3"
    wb.save(path)
    return r - 3


def main():
    students = cohort()
    scores(students, "A")
    ranks(students, "A")
    n1 = write(os.path.join(HERE, "样例_A期.xlsx"), students, "A")
    scores(students, "B")
    ranks(students, "B")
    n2 = write(os.path.join(HERE, "样例_B期.xlsx"), students, "B")
    print(f"已生成合成数据：A 期 {n1} 行，B 期 {n2} 行（姓名、学号、学校全部虚构）")


if __name__ == "__main__":
    main()
