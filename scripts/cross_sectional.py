# -*- coding: utf-8 -*-
"""
成绩报表批量分析脚本（可复用）
=========================================================
用法：
    python cross_sectional.py <成绩报表.xlsx> [输出.xlsx]

适用表结构（会自动识别，无需改代码）：
    · 两行合并表头：第 1 行为分组名（县/学校/姓名/语文/数学…），第 2 行为子表头
      （原始总分/赋分总分/班名/校名/市名/联名、分数(原始)/分数(赋分)/班名…）
    · 单行表头：分组行即子表头行，下一行起为数据
    · 三行表头：标题行 + 分组行 + 子表头行（也支持）

自动完成的分析：
    ① 扁平化（拆分两行表头、横向填充合并单元格）
    ② 样本概况（县 / 学校 / 选科）
    ③ 各科描述统计（n/均值/中位数/偏值/SD/P25/P75/IQR/CV/峰度）
    ④ 分数段结构（按得分率 40/60/75/90% 分档）
    ⑤ 频数分布（组距 10）+ 单峰性提示
    ⑥ 相对位置（z 分分位数、正态 vs 实际百分位对照）
    ⑦ 相关矩阵（Pearson / Spearman / 差值）
    ⑧ 异常值（IQR 法则 / z 法则 / 交集）
    ⑨ 赋分效应（原始 vs 赋分的均值、离散、偏度、峰度对比）
    ⑩ 方差贡献率（各科对小计方差的贡献，合计 100%）
    ⑪ 分差分解（各校 vs 全市，各科各占多少）

输出：一个多工作表的 xlsx + 控制台打印校验锚点。
"""

import os
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------- 可调配置 ----------------
# 各科满分（用于得分率分档）。未列出的科目默认 100。
FULL_MARKS = {"语文": 150, "数学": 150, "英语": 150,
              "物理": 100, "化学": 100, "生物": 100,
              "政治": 100, "地理": 100, "历史": 100}
# 不计入“小计”的科目（本例中四科全部计入）
CORE_SUBJECTS = ["语文", "数学", "英语", "物理"]
# 组距（频数分布）
BIN_WIDTH = 10
# 得分率分档
CUTS = [0.40, 0.60, 0.75, 0.90]


# ---------------- 一、读取与扁平化 ----------------
def read_and_flatten(path, max_sheet_try=10):
    """返回 (flattened_df, meta)；自动识别表头结构。"""
    wb = load_workbook(path, data_only=True, read_only=True)
    best = None

    for ws in wb.worksheets[:max_sheet_try]:
        head = []
        for r, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), start=1):
            head.append(list(row))
        # 找“姓名”所在行
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
        # 判定是否为“分组行 + 子表头行”结构：
        # 若分组行中“姓名”所在列在子表头行为空，则说明是两行结构
        is_two_row = True
        if name_col < len(subs):
            sv = subs[name_col]
            if isinstance(sv, str) and sv.strip() and sv.strip() != "姓名":
                is_two_row = False
        if r_name + 1 > len(head) - 1:
            is_two_row = False

        # 横向填充分组名（合并单元格只有首格有值）
        filled = []
        last = None
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

        data_start = r_name + (2 if is_two_row else 1)
        score = (ws.title, r_name, is_two_row, names, data_start)
        best = score
        break

    if best is None:
        wb.close()
        raise RuntimeError("未在任何工作表中找到「姓名」列")

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

    meta = {"sheet": sheet, "header_row": r_name, "two_row": is_two_row,
            "data_start": data_start, "n_raw": len(df)}
    return df, meta


# ---------------- 二、识别科目列 ----------------
def detect_subjects(df):
    """返回 {科目: {'raw': 列名, 'asg': 列名或None, 'rank_cols': [...]}}"""
    subs = {}
    for col in df.columns:
        if not isinstance(col, str):
            continue
        if "_分数(原始)" in col:
            subj = col.split("_分数(原始)")[0]
            subs.setdefault(subj, {})["raw"] = col
        elif "_分数(赋分)" in col:
            subj = col.split("_分数(赋分)")[0]
            subs.setdefault(subj, {})["asg"] = col
    # 排除标题行误识别
    return {k: v for k, v in subs.items() if "raw" in v}


def to_num(s):
    return pd.to_numeric(s, errors="coerce")


def safe_text(df):
    """把以 = + - @ 开头的文本加一个前导单引号，避免 Excel 公式注入。

    源报表里若出现以 = 开头的姓名/校名（导出异常或被构造），openpyxl 会把它当公式写进
    输出工作簿，别人打开时会被执行。这里统一转成文本。
    """
    out = df.copy()
    for c in out.columns:
        try:
            out[c] = out[c].map(
                lambda v: "'" + v if isinstance(v, str) and v[:1] in "=+-@" else v)
        except Exception:
            pass
    return out


def desc_stats(v):
    v = v.dropna()
    n = len(v)
    if n == 0:
        return None
    q1 = v.quantile(.25)
    q3 = v.quantile(.75)
    return {
        "n": n, "均值": v.mean(), "中位数": v.median(), "偏值": v.skew(),
        "SD": v.std(ddof=1), "P25": q1, "P75": q3, "IQR": q3 - q1,
        "CV": v.std(ddof=1) / v.mean() if v.mean() else np.nan,
        "峰度": v.kurt(), "min": v.min(), "max": v.max(),
    }


# ---------------- 三、各项分析 ----------------
def sample_tables(df):
    """各维度分布。县区列在不同报表中可能叫「县」「县区」「区域」。"""
    out = {}
    county_key = None
    for cand in ("县区", "县", "区域", "区县"):
        if cand in df.columns:
            county_key = cand
            break
    plan = []
    if county_key:
        plan.append(("县区", county_key))
    plan += [("学校", "学校"), ("选科", "选科")]
    for name, key in plan:
        if key not in df.columns:
            continue
        vc = df[key].astype(str).value_counts()
        t = pd.DataFrame({"人数": vc.values, "占比": vc.values / len(df)}, index=vc.index)
        t.index.name = name
        out[name] = t.reset_index()
    return out


def subject_table(df, subs):
    rows = []
    for k, v in subs.items():
        st = desc_stats(to_num(df[v["raw"]]))
        if st:
            st.update({"科目": k, "口径": "原始"})
            rows.append(st)
        if "asg" in v:
            st = desc_stats(to_num(df[v["asg"]]))
            if st:
                st.update({"科目": k, "口径": "赋分"})
                rows.append(st)
    cols = ["科目", "口径", "n", "均值", "中位数", "偏值", "SD", "P25", "P75", "IQR", "CV", "峰度"]
    return pd.DataFrame(rows)[cols]


def band_table(df, subs):
    rows = []
    for k, v in subs.items():
        s = to_num(df[v["raw"]])
        full = FULL_MARKS.get(k, 100)
        cuts = [c * full for c in CUTS]
        n = s.notna().sum()
        if n == 0:
            continue
        cnt = [
            int((s < cuts[0]).sum()),
            int(((s >= cuts[0]) & (s < cuts[1])).sum()),
            int(((s >= cuts[1]) & (s < cuts[2])).sum()),
            int(((s >= cuts[2]) & (s < cuts[3])).sum()),
            int((s >= cuts[3]).sum()),
        ]
        rows.append({
            "科目": k, "满分": full,
            "后40%": cnt[0], "40~60%": cnt[1], "60~75%": cnt[2],
            "75~90%": cnt[3], ">=90%": cnt[4],
            "合计": sum(cnt), "校核": int(n),
        })
    return pd.DataFrame(rows)


def freq_table(df, subs):
    """组距 BIN_WIDTH 的频数分布，各科并排。
    满分不足该档的科目留存空；包含满分的档用右闭区间，避免漏掉满分者。"""
    maxfull = max(FULL_MARKS.get(k, 100) for k in subs)
    edges = list(range(0, maxfull + BIN_WIDTH, BIN_WIDTH))
    rows = []
    for i in range(len(edges) - 1):
        lo, hi = edges[i], edges[i + 1]
        rec = {"组下限": lo, "组上限": hi}
        for k, v in subs.items():
            full = FULL_MARKS.get(k, 100)
            s = to_num(df[v["raw"]])
            if lo >= full:
                rec[k] = None
            elif hi >= full:
                rec[k] = int((s >= lo).sum())
            else:
                rec[k] = int(((s >= lo) & (s < hi)).sum())
        rows.append(rec)
    return pd.DataFrame(rows)


def correlation_table(df, subs):
    from itertools import combinations
    rows = []
    names = list(subs)
    for a, b in combinations(names, 2):
        x = to_num(df[subs[a]["raw"]])
        y = to_num(df[subs[b]["raw"]])
        m = x.notna() & y.notna()
        if m.sum() < 10:
            continue
        x2, y2 = x[m], y[m]
        r = x2.corr(y2)
        rho = x2.rank().corr(y2.rank())
        rows.append({"配对": f"{a}~{b}", "n": int(m.sum()),
                     "Pearson": r, "Spearman": rho, "差值(P-S)": r - rho})
    return pd.DataFrame(rows)


def outlier_table(df, subs):
    rows = []
    for k, v in subs.items():
        s = to_num(df[v["raw"]]).dropna()
        if len(s) < 10:
            continue
        q1, q3 = s.quantile(.25), s.quantile(.75)
        iqr = q3 - q1
        lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        mu, sd = s.mean(), s.std(ddof=1)
        in_iqr = (s < lo) | (s > hi)
        in_z = ((s - mu) / sd).abs() > 3
        rows.append({
            "科目": k, "下栅栏": lo, "上栅栏": hi,
            "IQR法则": int(in_iqr.sum()), "z法则": int(in_z.sum()),
            "仅IQR": int((in_iqr & ~in_z).sum()),
            "仅z": int((in_z & ~in_iqr).sum()),
            "交集": int((in_iqr & in_z).sum()),
        })
    return pd.DataFrame(rows)


def contribution_table(df, subs):
    """方差贡献率：contrib_i = r_i * SD_i / SD_总分（四科小计）。"""
    rows = []
    if not all(c in subs for c in CORE_SUBJECTS):
        return pd.DataFrame()
    parts = {c: to_num(df[subs[c]["raw"]]) for c in CORE_SUBJECTS}
    m = pd.Series(True, index=df.index)
    for s in parts.values():
        m &= s.notna()
    if m.sum() < 10:
        return pd.DataFrame()
    total = sum(parts[c][m] for c in CORE_SUBJECTS)
    sd_tot = total.std(ddof=1)
    acc = 0.0
    for c in CORE_SUBJECTS:
        x = parts[c][m]
        r = x.corr(total)
        share = r * x.std(ddof=1) / sd_tot
        acc += share
        rows.append({"科目": c, "SD": x.std(ddof=1), "r(科,小计)": r,
                     "基线(若独立)": x.std(ddof=1) / sd_tot,
                     "贡献率": share})
    rows.append({"科目": "合计", "SD": np.nan, "r(科,小计)": np.nan,
                 "基线(若独立)": np.nan, "贡献率": acc})
    return pd.DataFrame(rows)


def gap_table(df, subs, min_n=80):
    """各校 vs 全市：四科分差 + 占比。"""
    if "学校" not in df.columns or not all(c in subs for c in CORE_SUBJECTS):
        return pd.DataFrame()
    city = {c: to_num(df[subs[c]["raw"]]).mean() for c in CORE_SUBJECTS}
    rows = []
    for sch, g in df.groupby("学校"):
        if len(g) < min_n:
            continue
        d = {c: to_num(g[subs[c]["raw"]]).mean() - city[c] for c in CORE_SUBJECTS}
        gap = sum(d.values())
        rec = {"学校": sch, "n": len(g), "小计差": gap}
        for c in CORE_SUBJECTS:
            rec[c] = d[c]
        if abs(gap) >= 10:
            for c in CORE_SUBJECTS:
                rec[f"{c}占比"] = d[c] / gap * 100
        else:
            for c in CORE_SUBJECTS:
                rec[f"{c}占比"] = np.nan
        rows.append(rec)
    t = pd.DataFrame(rows)
    if len(t):
        t = t.sort_values("小计差", ascending=False).reset_index(drop=True)
    return t


def percentile_table(df, subs, subject=None):
    """正态 vs 实际百分位对照（默认第一个可识别科目）。"""
    if subject is None:
        subject = next(iter(subs))
    s = to_num(df[subs[subject]["raw"]]).dropna()
    mu, sd = s.mean(), s.std(ddof=1)
    from math import erf, sqrt
    rows = []
    for z in [-2, -1.5, -1, -0.5, 0, 0.5, 1, 1.5]:
        x = mu + z * sd
        normal = (1 + erf(z / sqrt(2))) / 2 * 100
        actual = (s < x).mean() * 100
        rows.append({"z": z, "分数": x, "正态PR": normal,
                     "实际PR": actual, "差值": normal - actual})
    return pd.DataFrame(rows)


# ---------------- 四、输出 ----------------
def style_sheet(ws, title):
    """第 1 行作表头，不插入标题行（便于 pandas 直接读取）。"""
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            cell.border = bd
    for col in ws.columns:
        width = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(10, width + 4), 34)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else \
        os.path.splitext(src)[0] + "_分析结果.xlsx"

    print("=" * 66)
    print("源文件:", src)

    df, meta = read_and_flatten(src)
    print(f"工作表={meta['sheet']}  表头行={meta['header_row']}  "
          f"两行表头={meta['two_row']}  原始行数={meta['n_raw']}  列数={len(df.columns)}")

    subs = detect_subjects(df)
    print("识别到科目:", ", ".join(f"{k}({'原始+赋分' if 'asg' in v else '原始'})"
                                for k, v in subs.items()))
    if not subs:
        raise RuntimeError("未识别到任何科目列（形如「语文_分数(原始)」）")

    results = {}
    results["样本概况"] = None
    samp = sample_tables(df)
    results["科目指标"] = subject_table(df, subs)
    results["分数段"] = band_table(df, subs)
    results["频数分布"] = freq_table(df, subs)
    results["相关矩阵"] = correlation_table(df, subs)
    results["异常值"] = outlier_table(df, subs)
    results["贡献分解"] = contribution_table(df, subs)
    results["分差分解"] = gap_table(df, subs)
    results["百分位对照"] = percentile_table(df, subs)

    # 校验锚点
    print("\n--- 校验锚点 ---")
    print("样本量:", len(df))
    for name, t in results.items():
        if t is not None and len(t):
            print(f"  {name}: {len(t)} 行")
    bt = results["分数段"]
    if len(bt):
        bad = bt[bt["合计"] != bt["校核"]]
        print("  分数段合计与有效人数不一致的科目:",
              "无" if len(bad) == 0 else ", ".join(bad["科目"]))
    ct = results["贡献分解"]
    if len(ct):
        body = ct[ct["科目"] != "合计"]
        print(f"  贡献率合计 = {body['贡献率'].sum()*100:.2f}%（应为 100%）")
        print("  各项贡献率:", "  ".join(
            f"{r['科目']}={r['贡献率']*100:.1f}%" for _, r in body.iterrows()))

    # 写出
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for name, t in samp.items():
            safe_text(t).to_excel(xw, sheet_name=f"样本-{name}", index=False)
        for name in ["科目指标", "分数段", "频数分布", "相关矩阵",
                     "异常值", "贡献分解", "分差分解", "百分位对照"]:
            t = results[name]
            if t is not None and len(t):
                safe_text(t).to_excel(xw, sheet_name=name, index=False)

        # 赋分效应对比
        if any("asg" in v for v in subs.values()):
            rows = []
            for k, v in subs.items():
                if "asg" not in v:
                    continue
                for tag, col in (("原始", v["raw"]), ("赋分", v["asg"])):
                    st = desc_stats(to_num(df[col]))
                    if st:
                        st.update({"科目": k, "口径": tag})
                        rows.append(st)
            if rows:
                safe_text(pd.DataFrame(rows)[["科目", "口径", "n", "均值", "中位数", "偏值",
                                              "SD", "IQR", "CV", "峰度"]]).to_excel(
                    xw, sheet_name="赋分对比", index=False)

    # 统一样式
    wb = load_workbook(out)
    for ws in wb.worksheets:
        style_sheet(ws, ws.title)

    # 追加说明表
    if "说明" in wb.sheetnames:
        del wb["说明"]
    wsn = wb.create_sheet("说明", 0)
    notes = [
        ("项目", "内容"),
        ("源文件", os.path.basename(src)),
        ("工作表", meta["sheet"]),
        ("表头行", f"第 {meta['header_row']} 行（两行表头={meta['two_row']}）"),
        ("数据起始行", meta["data_start"]),
        ("样本量", len(df)),
        ("识别科目", "、".join(subs.keys())),
        ("选考科目", "、".join(k for k, v in subs.items() if "asg" in v) or "无"),
        ("", ""),
        ("口径规则", "内部描述用原始分；总分与排名用赋分；跨科比分化只能用位次"),
        ("分数段阈值", "满分×40%/60%/75%/90%；满分见 FULL_MARKS 配置"),
        ("频数分布组距", BIN_WIDTH),
        ("相关矩阵说明", "选考科目两两相关仅对同时选考该两科的学生成立；重叠样本<10 时自动跳过"),
        ("异常值说明", "IQR 法则必 z 法则宽松的结论并非普遍成立，需逐科核对「仅z」列"),
        ("贡献率说明", "贡献率=r×SD/SD总分，四项合计 100%；属方差归因而非因果"),
        ("分差分解说明", "小计差绝对值<10 分时占比不稳定，程序已置空"),
        ("", ""),
        ("重要限制", "本表为横断面描述，不含纵向对比；各校均分未做生源校正，不能评价教学水平"),
    ]
    for r, (a, b) in enumerate(notes, start=1):
        wsn.cell(row=r, column=1, value=a)
        wsn.cell(row=r, column=2, value=b)
    for cell in wsn[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    wsn.column_dimensions["A"].width = 18
    wsn.column_dimensions["B"].width = 72

    wb.save(out)

    print("\n已输出:", out)


if __name__ == "__main__":
    main()
