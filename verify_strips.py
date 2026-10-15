# -*- coding: utf-8 -*-
"""教学条宽度 / 裁切检查。

专治一类静默 bug：教学条比田字格（16cm）宽，如果打印区域没跟着加宽，
导出 PDF 或打印时**超出的部分会被直接裁掉**，但页面几何检查察觉不到。

本脚本对每一页：
  1. 从 PDF 里**实测**田字格尺寸（格子多大的版式都能用，不写死 2cm）
  2. 从文字里读出该页的字，算出应有最右墨迹 = 左边距 + max(最宽教学条, 田字格总宽)
  3. 与实测最右墨迹比对，差超过 0.15cm 就判为被裁切或溢出

用法： python3 verify_strips.py <pdf> [meta.json]
"""
import json
import os
import re
import sys

import numpy as np
import pymupdf

_POS = [a for a in sys.argv[1:] if not a.startswith("--")]   # 位置参数（跳过 --cols= 之类）
PDF = _POS[0]
# 每行几个田字格（a5 纵向/a4 = 8；a5 横向 = 5），可用 --cols=N 指定
COLS = next((int(a.split("=", 1)[1]) for a in sys.argv[1:]
             if a.startswith("--cols=")), 8)
_mp = _POS[1] if len(_POS) > 1 else None
if _mp is None:                       # 默认找 PDF 同目录下的 meta.json
    _cand = os.path.join(os.path.dirname(os.path.abspath(PDF)), "meta.json")
    _mp = _cand if os.path.exists(_cand) else "stroke_meta_fill.json"
META = json.load(open(_mp, encoding="utf-8"))
print(f"（meta: {_mp}）")
DPI = 300
PX_CM = DPI / 2.54
MARGIN = 0.8
TOL = 0.15

def _bands(mask, gap):
    idx = np.flatnonzero(mask)
    if idx.size == 0:
        return []
    out = [[idx[0]]]
    for v in idx[1:]:
        if v - out[-1][-1] <= gap:
            out[-1].append(v)
        else:
            out.append([v])
    return [(o[0], o[-1]) for o in out]


def detect_grid(pm):
    """从页面实测田字格：返回 (格子边长cm, 一个田字格块宽度cm, 各块左边界[cm])。

    一条横向格线在裁切线处会被间隔列断开，每一段就是一个"田字格块"的宽度，
    段起点就是块起点。纵向版式只有 1 段，横向版式（左右两张 A5）有 2 段。
    这样任何版式、任何格子尺寸都适用。
    """
    a = np.frombuffer(pm.samples, dtype=np.uint8).reshape(
        pm.height, pm.width, pm.n)[..., :3].astype(int)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    green = (g > 110) & (r < 110) & (b < 130)
    # 网格宽度 = 横向绿线里**最长的那一段连续线段**。
    # 横向版式一条横线上有左右两个半页（中间被裁切线隔开），所以必须按"段"取，
    # 不能取整行的最左到最右 —— 那样会把两个半页之间的空白也算进去。
    grid_px, lefts = 0, []
    for i in range(a.shape[0]):
        idx = np.flatnonzero(green[i])
        if idx.size == 0:
            continue
        br = np.flatnonzero(np.diff(idx) > 1)
        runs = [r for r in np.split(idx, br + 1) if (r[-1] - r[0]) > PX_CM * 4]
        if runs and (runs[0][-1] - runs[0][0]) > grid_px:
            grid_px = runs[0][-1] - runs[0][0]
            lefts = [r[0] / PX_CM for r in runs]      # 这一行上的所有块起点
    if grid_px == 0 or not lefts:
        return None
    grid_cm = grid_px / PX_CM
    cell_cm = grid_cm / COLS          # 每块固定 COLS 格，用块宽反推格子边长
    return cell_cm, grid_cm, lefts


doc = pymupdf.open(PDF)
# 纸张尺寸从 PDF 现读（横向 29.7x21，纵向 21x29.7），不能写死
PAGE_W = doc[0].rect.width / 72 * 2.54
PAGE_H = doc[0].rect.height / 72 * 2.54
print(f"文件 {PDF}   共 {doc.page_count} 页")
_g = detect_grid(doc[0].get_pixmap(dpi=DPI))
print(f"页面 {PAGE_W:.1f} x {PAGE_H:.1f} cm"
      f"   可打印区 {MARGIN:.2f} ~ {PAGE_W-MARGIN:.2f} cm")
if _g:
    print(f"实测田字格：格子 {_g[0]:.3f}cm　每块宽 {_g[1]:.2f}cm　"
          f"每页 {len(_g[2])} 块（块起点 {'/'.join(f'{x:.2f}' for x in _g[2])}cm）")
print()
print(f"{'页':<4}{'本页字':<14}{'最宽条cm':<10}{'应有右边界':<12}{'实测右边界':<12}{'判定'}")

problems = []
for pi in range(doc.page_count):
    page = doc[pi]
    chars = re.findall(r"【.+?】\s*(\S+)\s*共\s*\d+\s*画", page.get_text())
    known = [c for c in chars if c in META]
    unknown = [c for c in chars if c not in META]
    pm = page.get_pixmap(dpi=DPI)
    geo = detect_grid(pm)
    if geo is None:
        print(f"{pi+1:<4}{''.join(chars):<14}{'？':<10}（页面上量不到田字格，跳过）")
        continue
    cell_cm, grid_cm, lefts = geo
    widths = [META[c]["w"] / META[c]["frame"] * cell_cm for c in known]
    strip_max = max(widths) if widths else None
    left = lefts[-1]                            # 最后一个田字格块的左边 = 该页最右墨迹的基准
    if strip_max is None:
        print(f"{pi+1:<4}{''.join(chars):<14}{'？':<10}{'（meta 里没有这些字，跳过）'}")
        continue
    want = left + max(strip_max, grid_cm)
    a = np.frombuffer(pm.samples, dtype=np.uint8).reshape(
        pm.height, pm.width, pm.n)[..., :3].astype(int)
    ink = a.sum(axis=2) < 720
    xs = np.flatnonzero(ink.any(axis=0))
    got = xs.max() / PX_CM if xs.size else 0.0

    ok = abs(got - want) <= TOL
    verdict = "✓" if ok else ("✗ 被裁切" if got < want else "✗ 溢出")
    if not ok:
        problems.append(f"第{pi+1}页 实测 {got:.2f}cm，应为 {want:.2f}cm（{verdict[2:]}）")
    if unknown:
        print(f"     （未在 meta 中找到：{''.join(unknown)}，已忽略其宽度）")
    if got > PAGE_W - MARGIN + 0.05:
        problems.append(f"第{pi+1}页 右边界 {got:.2f}cm 超出可打印区 {PAGE_W-MARGIN:.2f}cm")
    print(f"{pi+1:<4}{''.join(chars):<14}{strip_max:<10.2f}{want:<12.2f}{got:<12.2f}{verdict}")

print()
if problems:
    print(f"发现 {len(problems)} 个问题：")
    for p in problems:
        print("  - " + p)
    sys.exit(1)
print("教学条宽度检查全部通过 ✓（没有被裁切，也没有超出纸张）")