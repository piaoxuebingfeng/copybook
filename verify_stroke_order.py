# -*- coding: utf-8 -*-
"""校验 --stroke-order（笔顺练习模式）：田字格里每一格到底放了什么。

为什么要单独一个脚本
--------------------
普通模式下田字格只贴一张整字描红，肉眼或数颜色就能看出对不对。
笔顺模式下每格的内容都不一样（第 1 格第 1 笔、第 2 格前 2 笔……），
「排布算法算对了」和「真的按这个顺序铺进格子了」是两件事，
必须回到文件里逐格核对。

怎么核对的
----------
不信任文件名，也不重复算法：把 traces/ 里的原图按**像素内容**建反查表，
再读 xlsx 里每张图片的嵌入内容反查它是「哪个字的第几阶段」，
最后按锚点 (行, 列) 还原出阅读顺序，和 stroke_order_plan() 算出的期望值逐格比对。
所以它能抓到「图铺错格」「顺序颠倒」「留白没留」这类问题。

注意：第 N 阶段（画满全部笔画）与整字描红的像素**完全相同**，两者语义等价，
校验时按同一个符号 'F' 对待，不算差异。

用法
----
    python3 verify_stroke_order.py output/字帖_水火日月
    python3 verify_stroke_order.py output/字帖_xxx/字帖_xxx.xlsx --cols=5   # 横向版式
"""
import glob
import json
import os
import sys

import numpy as np
from PIL import Image
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from generate_zitie import stroke_order_plan        # 复用同一条规则，不抄第二份

DATA = json.load(open(os.path.join(HERE, "stroke_data_cn.json"), encoding="utf-8"))
TRACE_PX = 300                                       # 与 render_trace.SIZE 一致
COLS = 8                                             # --cols=N，横向版式是 5

_pos = [a for a in sys.argv[1:] if not a.startswith("--")]
for _a in sys.argv[1:]:
    if _a.startswith("--cols="):
        COLS = int(_a.split("=", 1)[1])
if not _pos:
    raise SystemExit("用法：python3 verify_stroke_order.py <字帖目录或 xlsx> [--cols=8]")

target = _pos[0]
if os.path.isdir(target):
    xlsxs = sorted(glob.glob(os.path.join(target, "*.xlsx")))
    xlsxs = [x for x in xlsxs if not os.path.basename(x).startswith("_")]
    if not xlsxs:
        raise SystemExit(f"{target} 里没有 .xlsx")
    xlsx = xlsxs[0]
else:
    xlsx = target
work = os.path.dirname(os.path.abspath(xlsx))
traces = os.path.join(work, "traces")

# 1) 用 traces/ 的原图建立「像素内容 -> 文件名」反查表，顺便记下每个字有几个阶段文件
lookup = {}
stages_max = {}
for f in glob.glob(os.path.join(traces, "*.png")):
    nm = os.path.basename(f)[:-4]
    lookup[np.array(Image.open(f).convert("RGBA")).tobytes()] = nm
    ch, _, k = nm.partition("_")
    if k:
        stages_max[ch] = max(stages_max.get(ch, 0), int(k))
if not lookup:
    raise SystemExit(f"{traces} 里没有描红原图，无法反查")


def strokes_of(ch):
    """这个字几画。(笔画数, 来源)

    优先用参照数据 stroke_data_cn.json；不在其中（参照数据只覆盖 34 个字）时，
    退回到 traces/ 里 {字}_{k}.png 的最大 k —— 阶段图本来就是按笔画数渲染的。
    两条路都拿不到就返回 0，让校验报出来而不是猜。
    """
    if ch in DATA:
        n = len(DATA[ch]["strokes"])
        alt = stages_max.get(ch)
        if alt and alt != n:
            problems_src.append(f"{ch}: 参照数据说 {n} 画，但阶段图有 {alt} 张，两者矛盾")
        return n, "参照数据"
    if ch in stages_max:
        return stages_max[ch], "阶段图张数"
    return 0, "无"


problems_src = []

# 2) 读 xlsx 里每张图的锚点与内容
ws = openpyxl.load_workbook(xlsx).active
per, strips, unknown = {}, 0, 0
for im in ws._images:
    a = getattr(im.anchor, "_from", None)
    if a is None:
        continue
    pic = Image.open(im.ref).convert("RGBA")
    if pic.width != TRACE_PX:                 # 教学条比描红图宽得多
        strips += 1
        continue
    nm = lookup.get(np.array(pic).tobytes())
    if nm is None:
        unknown += 1
        continue
    ch, _, k = nm.partition("_")
    per.setdefault(ch, []).append((a.row, a.col, int(k) if k else None))

print(f"字帖：{os.path.relpath(xlsx, HERE)}")
print(f"每行 {COLS} 格　识别到描红图 {sum(len(v) for v in per.values())} 张"
      f"（教学条 {strips} 张跳过，未识别 {unknown} 张）")
print()

problems = list(problems_src)
for ch in sorted(per):
    n, src = strokes_of(ch)
    if not n:
        problems.append(f"{ch}: 拿不到笔画数（既不在参照数据里，traces/ 里也没有阶段图）")
        continue
    items = sorted(per[ch], key=lambda t: (t[0], t[1]))
    norm = lambda k: "F" if (k is None or k == n) else k        # 第 n 阶段 == 整字
    seq = [norm(k) for _r, _c, k in items]
    rows = sorted({r for r, _c, _k in items})
    cells = len(rows) * COLS
    plan = stroke_order_plan(n, cells)
    exp = [norm(x) if x != "full" else "F" for x in plan if x is not None]

    # 列号：同一行内每格恰好相隔 2 个子列（一个田字格占 2 个子列），而且必须落在
    # 真正的田字格左边界上 —— 直接读单元格边框判断，不能靠猜「偶数列」：
    # 横向版式右半页的起点是第 14 列（0-based 13，奇数），偶数判据会误报。
    byrow = {}
    for r, c, k in items:
        byrow.setdefault(r, []).append(c)
    for r, cs in sorted(byrow.items()):
        cs = sorted(cs)
        if any(b - a != 2 for a, b in zip(cs, cs[1:])):
            problems.append(f"{ch}: 行 {r} 的格子间距不是 2 个子列 {cs}")
        if len(cs) > COLS:
            problems.append(f"{ch}: 行 {r} 放了 {len(cs)} 格，超过每行 {COLS} 格")
        for c in cs:
            if ws.cell(row=r + 1, column=c + 1).border.left.style != "thin":
                problems.append(
                    f"{ch}: 行 {r} 的图锚在第 {c + 1} 列，但那一列不是田字格左边界")

    mark = "✓" if seq == exp else "✗"
    print(f"  {ch}　{n:2d} 画 / {cells:2d} 格（{len(rows)} 行 × {COLS}）  {mark}"
          f"　[笔画数来自{src}]")
    print(f"      实际 {seq}")
    print(f"      期望 {exp}")
    print(f"      留白 {cells - len(items)} 格（期望 {plan.count(None)}）")
    if seq != exp:
        problems.append(f"{ch}: 铺出来的序列和期望不一致")
    if cells - len(items) != plan.count(None):
        problems.append(f"{ch}: 留白格数不对")

print()
if problems:
    print(f"发现 {len(problems)} 个问题：")
    for p in problems:
        print("  -", p)
    raise SystemExit(1)
print(f"共 {len(per)} 个字：每一格的顺序、位置、留白都与 stroke_order_plan() 一致 ✓")