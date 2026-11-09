# -*- coding: utf-8 -*-
"""跑一遍示例，把控制台输出存成 expected_output.txt（供 clone 后对照）。"""
import io
import os
import re
import subprocess
import sys

D = os.path.dirname(os.path.abspath(__file__))          # examples/
ROOT = os.path.dirname(D)                               # 仓库根
PY = sys.executable


def clean(text):
    """把本机绝对路径换成仓库内的相对写法，避免把路径带进仓库。"""
    text = text.replace(ROOT + os.sep, "").replace(ROOT, "")
    text = re.sub(r"[A-Za-z]:\\[^\s]*?score-analysis-general\\?", "", text)
    text = text.replace("\\", "/")
    return text


def run(args):
    p = subprocess.run([PY] + args, capture_output=True, text=True, encoding="utf-8")
    out = p.stdout + ("\n[stderr]\n" + p.stderr if p.stderr.strip() else "")
    return clean(out)


print("1) 生成合成数据 …")
print(run([os.path.join(D, "make_example_data.py")]))

print("2) 横断面 …")
out1 = run([os.path.join(ROOT, "scripts", "cross_sectional.py"),
            os.path.join(D, "样例_A期.xlsx"),
            os.path.join(D, "_输出_横断面.xlsx")])
print(out1)

print("3) 纵向增值 …")
out2 = run([os.path.join(ROOT, "scripts", "longitudinal.py"),
            os.path.join(D, "样例_A期.xlsx"),
            os.path.join(D, "样例_B期.xlsx"),
            os.path.join(D, "_输出_纵向.xlsx")])
print(out2)

head = """# 示例输出（在 600 人合成数据上跑出来的，可用来对照）
#
# 生成数据：python examples/make_example_data.py
# 横断面  ：python scripts/cross_sectional.py examples/样例_A期.xlsx
# 纵向增值：python scripts/longitudinal.py examples/样例_A期.xlsx examples/样例_B期.xlsx
#
# 数据全部虚构（同学001…同学600，示例一中…示例五中）。

"""
io.open(os.path.join(D, "expected_output.txt"), "w", encoding="utf-8", newline="\n").write(
    head + "========== 横断面 ==========\n" + out1 + "\n\n========== 纵向增值 ==========\n" + out2)
print("已写入 expected_output.txt")
