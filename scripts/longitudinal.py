# -*- coding: utf-8 -*-
"""
纵向增值分析脚本（可复用）
=========================================================
用法：
    python longitudinal.py <A期报表.xlsx> <B期报表.xlsx> [输出.xlsx]

背景：
    同一批学生的两次考试。原始分差受试卷难度污染，不能直接解读为进退步；
    本脚本同时给出「原始分口径」与「标准分口径」两套结果。

两种口径回答的问题不同：
    · 原始分 Δ ：总分变化由哪些科目的变化构成（含难度）——可加，恒等式成立
    · 标准分 Δ ：学生的**相对位置**怎么变（已消除难度）——可加，但只反映位次

标准分构造（正态化标准分）：
    PR = (秩 - 0.5) / n            每科每期独立计算
    z  = Φ⁻¹(PR)
    SS = 满分×0.5 + z×(满分/6)      截断到 [0, 满分]
    → 各科回到原尺度，ΣΔSS = ΔSS总 精确成立

配对规则：
    主键  姓名 + 学校
    消歧  同校同名者，用「学号前 8 位」区分（学号后几位每场重排，不可用全号）

输出：多工作表 xlsx，含说明、两种口径的变动与分解、校级、个人明细。
"""

import os
import sys
import numpy as np
import pandas as pd
from statistics import NormalDist
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------- 可调配置 ----------------
CORE = {"语文": 150, "数学": 150, "英语": 150, "物理": 100}
ELEC = {"化学": 100, "生物": 100, "政治": 100, "地理": 100}
MATCH_PREFIX = 8          # 学号用于消歧的前缀长度
MIN_SCHOOL_N = 30         # 校级表的最小样本量
SHARE_MIN_TOTAL = 5.0     # 平均总分变动小于此值时，不报告占比

ND = NormalDist()


# ---------------- 一、读取与扁平化 ----------------
def read_and_flatten(path):
    from openpyxl import load_workbook as lw
    wb = lw(path, data_only=True, read_only=True)
    best = None
    for ws in wb.worksheets[:10]:
        head = []
        for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
            head.append(list(row))
        r_name = name_col = None
        for ri, row in enumerate(head, start=1):
            for ci, v in enumerate(row):
                if isinstance(v, str) and v.strip() == "姓名":
                    r_name, name_col = ri, ci
                    break
            if r_name:
                break
        if not r_name:
            continue
        groups_raw = head[r_name - 1]
        subs = head[r_name] if r_name < len(head) else []
        is_two_row = True
        if name_col < len(subs):
            sv = subs[name_col]
            if isinstance(sv, str) and sv.strip() and sv.strip() != "姓名":
                is_two_row = False
        if r_name + 1 > len(head) - 1:
            is_two_row = False
        filled, last = [], None
        for ci in range(len(groups_raw)):
            g = groups_raw[ci]
            if isinstance(g, str) and g.strip():
                last = g.strip()
            filled.append(last)
        ncol = max(len(filled), len(subs))
        names = []
        for ci in range(ncol):
            g = filled[ci] if ci < len(filled) else None
            s = subs[ci] if is_two_row and ci < len(subs) else None
            if is_two_row and isinstance(s, str) and s.strip():
                names.append(f"{g}_{s.strip()}" if g else s.strip())
            else:
                names.append(g if g else f"列{ci+1}")
        best = (ws.title, r_name, is_two_row, names, r_name + (2 if is_two_row else 1))
        break
    if best is None:
        wb.close()
        raise RuntimeError(f"未找到「姓名」列: {path}")
    sheet, r_name, is_two_row, names, data_start = best
    ws = wb[sheet]
    rows = []
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        if row is None:
            continue
        if not any(v is not None and str(v).strip() != "" for v in row):
            continue
        rows.append(list(row) + [None] * (len(names) - len(row)))
    df = pd.DataFrame(rows, columns=names)
    wb.close()
    df.columns = [str(c).strip() for c in df.columns]
    for c in ("姓名", "学校", "班级", "学号", "选科"):
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df


# ---------------- 二、配对 ----------------
def pair_students(a, b):
    for df in (a, b):
        df["_k"] = df["姓名"] + "|" + df["学校"]
        df["_p"] = df["学号"].astype(str).str[:MATCH_PREFIX]
    ca, cb = a["_k"].value_counts(), b["_k"].value_counts()
    uniq = set(ca[ca == 1].index) & set(cb[cb == 1].index)
    m1 = a[a["_k"].isin(uniq)].merge(b[b["_k"].isin(uniq)], on=["_k"], suffixes=("_A", "_B"))

    dupk = set(ca[ca > 1].index) & set(cb[cb > 1].index)
    a2 = a[a["_k"].isin(dupk)].copy()
    b2 = b[b["_k"].isin(dupk)].copy()
    for df in (a2, b2):
        df["_k2"] = df["_k"] + "|" + df["_p"]
    c2a, c2b = a2["_k2"].value_counts(), b2["_k2"].value_counts()
    ok2 = set(c2a[c2a == 1].index) & set(c2b[c2b == 1].index)
    m2 = a2[a2["_k2"].isin(ok2)].merge(b2[b2["_k2"].isin(ok2)], on=["_k2"], suffixes=("_A", "_B"))

    M = pd.concat([m1, m2], ignore_index=True)
    info = {"直接配对": len(m1), "消歧配对": len(m2), "同校同名组数": len(dupk)}
    return M, info


# ---------------- 三、标准分 ----------------
def norm_ss(series, full):
    s = pd.to_numeric(series, errors="coerce")
    ss = pd.Series(np.nan, index=s.index)
    pr = pd.Series(np.nan, index=s.index)
    m = s.notna()
    n = int(m.sum())
    if n < 10:
        return ss, pr
    r = s[m].rank(method="average")
    p = (r - 0.5) / n
    z = p.apply(lambda v: ND.inv_cdf(min(max(v, 1e-6), 1 - 1e-6)))
    ss[m] = (0.5 * full + z * (full / 6.0)).clip(0, full)
    pr[m] = p * 100
    return ss, pr


def build_merged(M):
    recs = []
    for _, r in M.iterrows():
        rec = {"姓名": r["姓名_A"], "学校": r["学校_A"],
               "班级_A": r["班级_A"], "班级_B": r["班级_B"],
               "学号_A": r.get("学号_A"), "学号_B": r.get("学号_B"),
               "选科_A": r["选科_A"], "选科_B": r["选科_B"]}
        for s in list(CORE) + list(ELEC):
            kind = "原始" if s in CORE else "赋分"
            for tag in ("A", "B"):
                col = f"{s}_分数({kind})_{tag}"
                rec[f"{s}_{tag}"] = pd.to_numeric(
                    pd.Series([r.get(col)]), errors="coerce").iloc[0]
        for tag in ("A", "B"):
            rec[f"总分_{tag}"] = pd.to_numeric(
                pd.Series([r.get(f"总分_赋分总分_{tag}")]), errors="coerce").iloc[0]
        recs.append(rec)
    return pd.DataFrame(recs)


def add_standard_scores(X):
    n = len(X)
    for s, full in {**CORE, **ELEC}.items():
        for tag in ("A", "B"):
            ss, pr = norm_ss(X[f"{s}_{tag}"], full)
            X[f"SS_{s}_{tag}"] = ss
            X[f"PR_{s}_{tag}"] = pr
    for s in list(CORE) + list(ELEC):
        X[f"SS_{s}"] = X[f"SS_{s}_B"] - X[f"SS_{s}_A"]
        X[f"PR_{s}"] = X[f"PR_{s}_B"] - X[f"PR_{s}_A"]
        X[f"RAW_{s}"] = X[f"{s}_B"] - X[f"{s}_A"]
    return X


def elective_pair(xk):
    s = str(xk).replace("物", "")
    names = [{"化": "化学", "生": "生物", "政": "政治", "地": "地理"}[c]
             for c in s if c in "化生政地"]
    return (names[0], names[1]) if len(names) >= 2 else (None, None)


def build_elective_totals(X):
    e1 = [elective_pair(x)[0] for x in X["选科_A"]]
    e2 = [elective_pair(x)[1] for x in X["选科_A"]]

    def gather(pre, tag):
        return {s: X[f"{pre}{s}_{tag}"].values for s in ELEC}

    def build(D1, D2):
        out = []
        for i in range(len(e1)):
            if e1[i] is None or e2[i] is None:
                out.append(np.nan); continue
            v1, v2 = D1[e1[i]][i], D2[e2[i]][i]
            out.append(v1 + v2 if not (np.isnan(v1) or np.isnan(v2)) else np.nan)
        return np.array(out, dtype=float)

    X["选考SS_A"] = build(gather("SS_", "A"), gather("SS_", "A"))
    X["选考SS_B"] = build(gather("SS_", "B"), gather("SS_", "B"))
    X["选考PR_A"] = build(gather("PR_", "A"), gather("PR_", "A"))
    X["选考PR_B"] = build(gather("PR_", "B"), gather("PR_", "B"))
    X["SS_选考"] = X["选考SS_B"] - X["选考SS_A"]
    X["PR_选考"] = X["选考PR_B"] - X["选考PR_A"]
    return X


def stat(v):
    v = pd.Series(v).dropna()
    if len(v) == 0:
        return None
    return {"n": len(v), "均值": v.mean(), "中位数": v.median(), "SD": v.std(ddof=1),
            "P10": v.quantile(.10), "P25": v.quantile(.25), "P75": v.quantile(.75),
            "P90": v.quantile(.90), "上升率": (v > 0).mean() * 100}


def safe_text(df):
    """把以 = + - @ 开头的文本加前导单引号，避免 Excel 公式注入（个人明细含姓名/校名）。"""
    out = df.copy()
    for c in out.columns:
        try:
            out[c] = out[c].map(
                lambda v: "'" + v if isinstance(v, str) and v[:1] in "=+-@" else v)
        except Exception:
            pass
    return out


# ---------------- 主流程 ----------------
def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    srcA, srcB = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else \
        os.path.splitext(srcA)[0] + "_纵向增值分析.xlsx"

    print("=" * 60)
    A = read_and_flatten(srcA)
    B = read_and_flatten(srcB)
    print(f"A 期 {len(A)} 行   B 期 {len(B)} 行")

    M, info = pair_students(A, B)
    for k, v in info.items():
        print(f"  {k}: {v}")
    print(f"配对样本 N = {len(M)}")

    X = build_merged(M)
    X = add_standard_scores(X)
    X = build_elective_totals(X)

    SUBS = list(CORE) + ["选考"]
    X["ΔSS_总"] = X[[f"SS_{s}" for s in SUBS]].sum(axis=1, min_count=len(SUBS))
    X["ΔPR_总和"] = X[[f"PR_{s}" for s in list(CORE)]].sum(axis=1, min_count=4) + X["PR_选考"]
    X["ΔRAW_总"] = X["总分_B"] - X["总分_A"]

    cc_cols = [f"SS_{s}" for s in SUBS] + ["ΔRAW_总"]
    CC = X.dropna(subset=cc_cols).copy()
    print(f"完全案例 N = {len(CC)}")

    # 两期全样本均值（难度指示）
    whole = []
    for tag, df in (("A 期", A), ("B 期", B)):
        rec = {"期": tag, "n": len(df)}
        for s in CORE:
            c = f"{s}_分数(原始)"
            rec[s] = pd.to_numeric(df[c], errors="coerce").mean() if c in df.columns else np.nan
        rec["总分(赋分)"] = pd.to_numeric(df["总分_赋分总分"], errors="coerce").mean()
        whole.append(rec)
    whole = pd.DataFrame(whole)
    d = {"期": "差值(B-A)", "n": len(B) - len(A)}
    for s in list(CORE) + ["总分(赋分)"]:
        d[s] = whole.loc[1, s] - whole.loc[0, s]
    whole = pd.concat([whole, pd.DataFrame([d])], ignore_index=True)

    # 三张变动表
    def tab(specs):
        rows = []
        for col, name in specs:
            st = stat(CC[col])
            if st:
                st["项目"] = name
                rows.append(st)
        return pd.DataFrame(rows)[["项目", "n", "均值", "中位数", "SD",
                                   "P10", "P25", "P75", "P90", "上升率"]]

    ss_tab = tab([(f"SS_{s}", f"{s}（标准分）") for s in SUBS] + [("ΔSS_总", "合计（标准总分）")])
    pr_tab = tab([(f"PR_{s}", f"{s}（百分位）") for s in SUBS] + [("ΔPR_总和", "合计（百分位和）")])
    raw_tab = tab([(f"RAW_{s}", f"{s}（原始分）") for s in CORE] + [("ΔRAW_总", "总分（原始）")])

    # 分解
    mean_ss = CC["ΔSS_总"].mean()
    comp = []
    for s in SUBS:
        m = CC[f"SS_{s}"].mean()
        comp.append({"项目": s, "标准分平均变动": m,
                     "占标准总分变动": m / mean_ss * 100 if mean_ss else np.nan})
    comp.append({"项目": "合计", "标准分平均变动": mean_ss, "占标准总分变动": 100.0})
    comp = pd.DataFrame(comp)
    if abs(mean_ss) < SHARE_MIN_TOTAL:
        comp["占标准总分变动"] = np.nan
        share_note = (f"平均标准总分变动仅 {mean_ss:.2f} 分，各科接近 0，占比不稳定，故不予报告。")
    else:
        share_note = f"平均标准总分变动 {mean_ss:.2f} 分，占比可解读。"

    # 校级
    sch = []
    for name, g in CC.groupby("学校"):
        if len(g) < MIN_SCHOOL_N:
            continue
        rec = {"学校": name, "n": len(g),
               "标准总分变动": g["ΔSS_总"].mean(),
               "原始总分变动": g["ΔRAW_总"].mean(),
               "上升率": (g["ΔSS_总"] > 0).mean() * 100}
        for s in SUBS:
            rec[s] = g[f"SS_{s}"].mean()
        sch.append(rec)
    sch = pd.DataFrame(sch).sort_values("标准总分变动", ascending=False).reset_index(drop=True)

    notes = pd.DataFrame({"项目": [
        "A 期", "B 期", "A 行数", "B 行数", "配对规则", "配对样本", "消歧配对数",
        "完全案例", "原始分口径", "标准分口径", "标准分构造",
        "占比说明", "重要限制", "解读建议"],
        "说明": [
            os.path.basename(srcA), os.path.basename(srcB), len(A), len(B),
            f"主键 姓名+学校；同校同名用学号前 {MATCH_PREFIX} 位消歧",
            len(X), info["消歧配对"], len(CC),
            "Δ = B期原始分 − A期原始分；含试卷难度，ΣΔ = Δ总分 精确成立",
            "每科每期独立正态化，消除难度；ΣΔSS = ΔSS总 精确成立",
            "PR=(秩-0.5)/n → z=Φ⁻¹(PR) → SS=满分×0.5+z×(满分/6)，截断到[0,满分]",
            share_note,
            "标准分只反映相对位置，不反映绝对掌握程度；若全体一起变难，标准分不变。"
            "两期参试人数不同会让各科标准分Δ系统性偏负。",
            "看原始分：分数变了多少；看标准分：人变了多少。两者必须一起报。",
        ]})

    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        notes.to_excel(xw, sheet_name="说明", index=False)
        whole.to_excel(xw, sheet_name="两期全样本对比", index=False)
        raw_tab.to_excel(xw, sheet_name="原始分变动", index=False)
        ss_tab.to_excel(xw, sheet_name="标准分变动", index=False)
        pr_tab.to_excel(xw, sheet_name="百分位变动", index=False)
        comp.to_excel(xw, sheet_name="标准分分解", index=False)
        sch.to_excel(xw, sheet_name="校级变动", index=False)
        safe_text(X).to_excel(xw, sheet_name="个人明细", index=False)

    wb = load_workbook(out)
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400)):
            for cell in row:
                cell.border = bd
        for col in ws.columns:
            w = max((len(str(c.value)) if c.value is not None else 0) for c in col[:200])
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(9, w + 3), 28)
    wb.save(out)

    print("\n=== 两期全样本均值（难度指示）===")
    print(whole.to_string(index=False))
    print("\n=== 标准分变动 ===")
    print(ss_tab.to_string(index=False))
    print("\n=== 原始分变动（对照）===")
    print(raw_tab.to_string(index=False))
    print("\n=== 标准分分解 ===")
    print(comp.to_string(index=False))
    print("注:", share_note)
    print("\n=== 校级（前5 / 后5）===")
    print(pd.concat([sch.head(5), sch.tail(5)]).to_string(index=False))
    print("\n已输出:", out)


if __name__ == "__main__":
    main()
