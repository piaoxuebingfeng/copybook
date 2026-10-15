# -*- coding: utf-8 -*-
"""对导出的 PDF 做验收（几何 + 内容 + 颜色）。

判据说明
--------
* 田字格横线：该行绿色总量 > 12cm 且最长连续段 > 1.5cm
  （教学条里的小方格每格只有 2cm，会被这条规则排除）。
* 字符数：用描红粉色区块来数 —— 每个字恰好有 TRACE_CELLS 个描红字，
  横向连成一条；纵向按 > 1cm 间隔切分成不同字。
"""
import sys
import pymupdf
import numpy as np

PDF = sys.argv[1]
NCHAR = int(sys.argv[2]) if len(sys.argv) > 2 else 34   # 预期字数（每个字 1 行描红）
TRACE_CELLS_ARG = int(sys.argv[3]) if len(sys.argv) > 3 else None
# 格子边长（cm）：a4 版式 2.01，a5 版式 2.38。不同版式必须传对应的值，
# 否则「格高筛选范围」「格位偏移」全部对不上，会误报。
CELL_CM = float(sys.argv[5]) if len(sys.argv) > 5 else 2.01083
# 每行几个田字格（纵向/a4 = 8；横向 = 5）
COLS = next((int(a.split("=",1)[1]) for a in sys.argv[1:]
             if a.startswith("--cols=")), 8)
DPI = 300
PX_CM = DPI / 2.54
TRACE_CELLS = TRACE_CELLS_ARG or 3
# --stroke-order：笔顺练习模式。此时每个字的**所有**田字格都贴了图（第 1 格第 1 笔、
# 第 2 格前 2 笔……），不再是"每字一行 3 格描红"，所以：
#   · 一行里只要有 1 格有粉色就算"带描红的行"（原来是要求 3 格全有）
#   · 预期行数 = 全部田字格行（原来 = 字数），精确的逐格顺序交给 verify_stroke_order.py
STROKE_ORDER = "--stroke-order" in sys.argv
NEED_CELLS = 1 if STROKE_ORDER else TRACE_CELLS
doc = pymupdf.open(PDF)
PW, PH = doc[0].rect.width / 72 * 2.54, doc[0].rect.height / 72 * 2.54
MARGIN = 0.8


def bands(mask_1d, gap_px, min_len=1):
    """把布尔/索引数组切成段，间隔 > gap_px 视为断开。输入为布尔数组。"""
    idx = np.flatnonzero(mask_1d)
    if idx.size == 0:
        return []
    br = np.flatnonzero(np.diff(idx) > gap_px)
    out = []
    for s in np.split(idx, br + 1):
        if len(s) >= min_len:
            out.append((int(s[0]), int(s[-1])))
    return out


problems = []
total_rows = total_clusters = 0
expect_rows = 0          # 笔顺模式：全部田字格行数（每行都该有图）
all_per = []
print(f"文件 {PDF}")
PAGES = int(sys.argv[4]) if len(sys.argv) > 4 else None
print(f"页数 {doc.page_count}" + (f"（预期 {PAGES}）" if PAGES else "") +
      f"   页面 {PW:.1f} x {PH:.1f} cm")
if PAGES and doc.page_count != PAGES:
    problems.append(f"页数 {doc.page_count} != {PAGES}")
print()
print(f"{'页':<4}{'内容 x(cm)':<17}{'内容 y(cm)':<17}{'字数':<6}{'描红格数':<11}{'格宽cm':<9}{'格高cm'}")
for pno, page in enumerate(doc):
    pm = page.get_pixmap(dpi=DPI)
    a = np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.width, pm.n)[..., :3].astype(int)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    ink = a.sum(axis=2) < 720
    ys, xs = np.where(ink)
    x0, x1 = xs.min()/PX_CM, xs.max()/PX_CM
    y0, y1 = ys.min()/PX_CM, ys.max()/PX_CM

    # --- 田字格几何 ---
    green = (g > r + 40) & (g > b + 20) & (g > 110)
    rowok = np.zeros(a.shape[0], dtype=bool)
    for i in range(a.shape[0]):
        idx = np.flatnonzero(green[i])
        if idx.size == 0:
            continue
        br = np.flatnonzero(np.diff(idx) > 1)
        mrun = max(len(s) for s in np.split(idx, br + 1))
        # 用 4 个格宽做总长阈值：描红图会盖住前 3 个格子的格线，
        # 所以一条田字格横线可能只剩 5 个格宽的绿色，不能按整页宽 12cm 卡。
        if green[i].sum() > PX_CM * 4 * CELL_CM and mrun > PX_CM * 1.5:
            rowok[i] = True
    hl = [(s + e) / 2 / PX_CM for s, e in bands(rowok, 2)]

    # 田字格左边界 = 页面最左边的长竖直绿线（教学条也左对齐，起点相同）
    colg = green.sum(axis=0)
    longcols = np.flatnonzero(colg > PX_CM * 1.5)
    origin = longcols.min() / PX_CM if longcols.size else float("nan")
    # 一行里可能有多个"田字格块"（横向版式 = 左右各一个）。
    # 直接看**横线的分段**最稳：一条横向格线在裁切线处会被间隔列断开，
    # 每一段就是一个田字格块的宽度，段起点就是块起点。
    BLOCKS, grid_w = [], np.nan
    _best = 0.0
    for _i in range(a.shape[0]):
        _idx = np.flatnonzero(green[_i])
        if _idx.size == 0:
            continue
        _br = np.flatnonzero(np.diff(_idx) > 1)
        _runs = [r for r in np.split(_idx, _br + 1)
                 if (r[-1] - r[0]) > PX_CM * 2 * CELL_CM]
        if _runs and (_runs[0][-1] - _runs[0][0]) > _best:
            _best = _runs[0][-1] - _runs[0][0]
            BLOCKS = [r[0] / PX_CM for r in _runs]
            grid_w = _best / PX_CM

    # --- 描红：逐"田字格行"检查前 3 个格子是否有描红 ---
    # 按格子的固定 x 位置判断（不能按色块数，因为"八""小"等字本身会横向分成多块）
    pink = (r > 200) & (g > 140) & (r - g > 25) & (np.abs(g - b) < 20)
    # 教学条里"这一笔"的深红色（描红行只有粉色，没有深红）
    red = (r > 150) & (g < 110) & (b < 110)
    per = []
    grid_rows_here = 0          # 本页去掉教学条之后的真田字格行数
    trace_band = None
    # 教学条图片用的是和格子**同一个绿色**画边框（风格统一），它的上下边缘也是长绿线；
    # 而"字头行高 + 教学条高"恰好也是一个格子高 —— 所以整条教学条会被当成一行田字格，
    # 里面的粉色描红示范格被算成"描红行"，验收就多报一行。
    # 用灰色内部十字区分不行（教学条里的田字格框本来就有灰十字）；
    # 靠结构也不行（笔画多的字教学条占 3 行，半页只剩 1 行田字格，同样是 2 条线）。
    # 可靠的办法看颜色：教学条里有"这一笔"的**深红色**笔画，田字格描红行只有粉色。
    stack = [(y0, y1) for y0, y1 in zip(hl, hl[1:])
             if 0.92 * CELL_CM < y1 - y0 < 1.10 * CELL_CM]
    for y0c, y1c in stack:
        _red_n = red[int((y0c + 0.10) * PX_CM):int((y1c - 0.10) * PX_CM)].sum()
        if _red_n > 500:            # 这是教学条，不是田字格行
            continue
        grid_rows_here += 1         # 跳过教学条后剩下的才是真田字格行
        s = int((y0c + 0.10) * PX_CM)
        e = int((y1c - 0.10) * PX_CM)
        band = pink[s:e, :]
        for _ox in (BLOCKS or [origin]):        # 每个田字格块各查一次
            hit = 0
            for k in range(NEED_CELLS if STROKE_ORDER else TRACE_CELLS):
                xa = int((_ox + CELL_CM * k + 0.10 * CELL_CM) * PX_CM)
                xb = int((_ox + CELL_CM * k + 0.90 * CELL_CM) * PX_CM)
                if band[:, xa:xb].sum() > 300:  # 阈值滤掉红色笔画的浅色反锯齿边缘
                    hit += 1
            if hit >= NEED_CELLS:
                per.append(hit)
                if trace_band is None:
                    trace_band = (y0c, y1c)
    nchar = len(per)
    all_per.extend(per)
    expect_rows += grid_rows_here * (len(BLOCKS) if BLOCKS else 1)

    # --- 几何：格高取格线间距；格宽在真正的田字格行内量竖直绿线 ---
    cw = ch = float("nan")
    if hl:
        d = np.diff(hl)
        sel = d[(d > 0.92 * CELL_CM) & (d < 1.10 * CELL_CM)]
        ch = float(sel.mean()) if sel.size else float("nan")
    # 格宽：田字格共 8 格。取**横向绿线的最大跨距**/8 ——
    # 总是存在不含描红的格子行，其横线横跨整个网格；而描红行会被图片截断，
    # 所以对每一行取 max 就得到完整网格宽度。比"数竖线"稳得多。
    if not np.isnan(grid_w):
        cw = grid_w / COLS

    print(f"{pno+1:<4}{f'{x0:.2f}~{x1:.2f}':<17}{f'{y0:.2f}~{y1:.2f}':<17}"
          f"{nchar:<6}{str(per):<11}{cw:.3f}    {ch:.3f}")

    total_rows += nchar
    total_clusters += sum(per)

    if x0 < MARGIN - 0.15 or x1 > PW - MARGIN + 0.15:
        problems.append(f"第{pno+1}页 内容超出左右页边距 ({x0:.2f}~{x1:.2f})")
    if y0 < MARGIN - 0.15 or y1 > PH - MARGIN + 0.15:
        problems.append(f"第{pno+1}页 内容超出上下页边距 ({y0:.2f}~{y1:.2f})")
    if not np.isnan(cw) and abs(cw - CELL_CM) > 0.08:
        problems.append(f"第{pno+1}页 格宽 {cw:.3f} 偏离 {CELL_CM:.2f}cm")
    if not np.isnan(ch) and abs(ch - CELL_CM) > 0.08:
        problems.append(f"第{pno+1}页 格高 {ch:.3f} 偏离 {CELL_CM:.2f}cm")
    if not np.isnan(cw) and not np.isnan(ch) and abs(cw - ch) > 0.05:
        problems.append(f"第{pno+1}页 格子不是正方形：宽 {cw:.3f} 高 {ch:.3f}")

print()
if STROKE_ORDER:
    print(f"描红统计（笔顺练习模式）：每行只查第 1 格 —— 带图的田字格行 {total_rows} 行"
          f"（实际田字格行 {expect_rows} 行以内）")
    # 笔顺模式下"留白"是设计的一部分（笔顺练 3 遍 + 整字 3 个之后才留白），
    # 所以不能要求"每一行都有图"。这里只守住两头：每个字至少有一行有图、
    # 有图的行不超过实际田字格行数。每一格到底铺了什么，由 verify_stroke_order.py
    # 回到文件里逐格精确核对。
    if total_rows < NCHAR:
        problems.append(f"带描红的田字格行只有 {total_rows} 行，少于字数 {NCHAR}"
                        f"（有整个字没贴上）")
    if total_rows > expect_rows:
        problems.append(f"带描红的田字格行 {total_rows} 超过实际田字格行 {expect_rows}")
    if total_clusters < NCHAR:
        problems.append(f"描红格合计 {total_clusters} 少于字数 {NCHAR}")
else:
    print(f"描红统计：带描红的田字格行 {total_rows} 行（预期 {NCHAR}）"
          f"，描红格合计 {total_clusters} 个（预期 {NCHAR*TRACE_CELLS}）")
    if total_rows != NCHAR:
        problems.append(f"带描红的田字格行 {total_rows} != {NCHAR}")
    if total_clusters != NCHAR * TRACE_CELLS:
        problems.append(f"描红格合计 {total_clusters} != {NCHAR*TRACE_CELLS}")
    if any(p != TRACE_CELLS for p in all_per):
        problems.append(f"有字的描红格数不是 {TRACE_CELLS}: {all_per}")

print()
_pg = min(1, doc.page_count - 1)     # 单页文档就查第 1 页
print(f"颜色检查（第{_pg+1}页）:")
pm = doc[_pg].get_pixmap(dpi=200)
a = np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.width, pm.n)[..., :3].astype(int)
for name, col, lo in [("绿色格线", (0, 176, 80), 500), ("灰色虚线", (153, 153, 153), 300),
                      ("浅粉描红", (232, 192, 192), 300), ("黑色已写笔画", (26, 26, 26), 200),
                      ("红色当前笔画", (223, 49, 48), 200)]:
    n = (np.abs(a - np.array(col)).sum(axis=2) < 30).sum()
    print(f"  {name}: {n} 像素" + ("" if n > lo else "   <-- 过少!"))
    if n <= lo:
        problems.append(f"第2页 {name} 过少 ({n})")

print()
print("全部验收通过 ✓" if not problems else f"发现 {len(problems)} 个问题：")
for p in problems:
    print("  -", p)