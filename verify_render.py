# -*- coding: utf-8 -*-
"""对渲染出的教学条图片做像素级结构校验（替代肉眼检查）。

校验点：
1. 格数 = 笔画数 + 1（末尾那个是描红示范格）
2. 教学格 k 必须有红色像素（当前这一笔），且别格不该有红色
3. 教学格 k>=1 必须有黑色像素（已经写完的笔画）
4. 黑色像素数随 k 单调不减（笔画在累加）
5. 描红格只有粉色，没有红/黑

排版几何**直接从 render_strokes import**（layout / F / LABEL_H / ROW_GAP），
不再自己抄一份常量。抄一份的下场是：改渲染器时这边悄悄失效 —— 这里曾经把格间距
抄成 14（渲染器实际是 12），而且只按单行取格，于是把折行的教学条误判成「缺格」。

用法：
    python3 verify_render.py output/某字帖/strips
    python3 verify_render.py output/横测/strips --strip-max-cells=5    # 横向版式（5 格换行）
"""
import json
import os
import sys

import numpy as np
from PIL import Image

import render_strokes as R

_pos = [a for a in sys.argv[1:] if not a.startswith("--")]
IMG_DIR = _pos[0] if _pos else "stroke_images"
# render_strokes 在 import 时已经按同样的规则读过这些开关了，这里只是为了明确
for _a in sys.argv[1:]:
    if _a.startswith("--strip-max-cells="):
        R.MAX_W = int(float(_a.split("=", 1)[1]) * R.F)
    elif _a.startswith("--strip-max-cm="):
        R.MAX_W = int(float(_a.split("=", 1)[1]) * R.F / 2.0)

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = json.load(open(os.path.join(HERE, "stroke_data_cn.json"), encoding="utf-8"))

F = R.F
PALETTE = {
    "white": (255, 255, 255),
    "green": (0, 176, 80),
    "gray": (190, 190, 190),
    "black": (26, 26, 26),
    "red": (224, 49, 49),
    "pink": (232, 192, 192),
}


def classify(a):
    """把 RGB 数组分类到最近调色板颜色，返回 {name: count}。"""
    cols = np.array(list(PALETTE.values()))
    names = list(PALETTE.keys())
    flat = a.reshape(-1, 1, 3).astype(int)
    d = ((flat - cols.reshape(1, -1, 3)) ** 2).sum(axis=2)
    idx = d.argmin(axis=1)
    return {nm: int((idx == i).sum()) for i, nm in enumerate(names)}


def frame_boxes(n):
    """按渲染器真实排版给出每一格的像素框，返回 [(box, 是否描红格), ...]，按笔顺排列。"""
    rows, _w, _h = R.layout(n)
    out = []
    for ri, row in enumerate(rows):
        y0 = ri * (F + R.LABEL_H + R.ROW_GAP)
        for (x0, is_pink) in row:
            out.append(((int(x0), int(y0), int(x0 + F), int(y0 + F)), is_pink))
    return out


def find_targets():
    """返回 (可校验的字, 因不在对照数据里而跳过的字)。

    笔画数要从 stroke_data_cn.json 取（才能判断"格数 = 笔画数 + 1"），所以只有这份
    参照数据里的字能校验；其它字只报告跳过，不当成错误。
    """
    if not os.path.isdir(IMG_DIR):
        raise SystemExit(f"找不到目录：{IMG_DIR}")
    png = [f[:-4] for f in sorted(os.listdir(IMG_DIR)) if f.endswith(".png")]
    todo = [c for c in png if c in DATA]
    skipped = [c for c in png if c not in DATA]
    if not todo:
        raise SystemExit(f"{IMG_DIR} 里没有可校验的教学条图片"
                         f"（{len(png)} 张都不在参照数据 stroke_data_cn.json 里）")
    return todo, skipped


problems = []
summary = []

TODO, SKIPPED = find_targets()
for ch in TODO:
    n = len(DATA[ch]["strokes"])
    a = np.array(Image.open(os.path.join(IMG_DIR, f"{ch}.png")).convert("RGB"))
    boxes = frame_boxes(n)

    if len(boxes) != n + 1:
        problems.append(f"{ch}: 格数 {len(boxes)} != 笔画数 {n} + 1")
        continue

    counts = []
    for (x0, y0, x1, y1), _pink in boxes:
        if x1 > a.shape[1] or y1 > a.shape[0]:
            counts.append(None)                  # 格框超出图片：排版和校验对不上
            continue
        counts.append(classify(a[y0:y1, x0:x1]))
    if any(c is None for c in counts):
        problems.append(f"{ch}: 格框超出图片范围（图片 {a.shape[1]}x{a.shape[0]}，"
                        f"{len(boxes)} 格）")
        continue

    teach = counts[:n]
    pk = counts[n]

    for k in range(n):
        if teach[k]["red"] < 60:
            problems.append(f"{ch} 教学格{k+1}: 红色像素过少 ({teach[k]['red']})")
    # 第 1 格只有"第一笔"是红的、还没有已写完的笔画，所以黑笔从第 2 格起才要求
    for k in range(1, n):
        if teach[k]["black"] < 60:
            problems.append(f"{ch} 教学格{k+1}: 黑色(已写笔画)像素过少 ({teach[k]['black']})")

    blacks = [c["black"] for c in teach]
    for k in range(1, n):
        if blacks[k] < blacks[k - 1] * 0.85:
            problems.append(f"{ch}: 黑色像素未累加 k{k}->{k+1} ({blacks[k-1]}->{blacks[k]})")

    if pk["pink"] < 60:
        problems.append(f"{ch} 描红格: 粉色像素过少 ({pk['pink']})")
    if pk["red"] > 40:
        problems.append(f"{ch} 描红格: 出现红色点 ({pk['red']})")
    if pk["black"] > 40:
        problems.append(f"{ch} 描红格: 出现黑色点 ({pk['black']})")

    summary.append((ch, n, blacks, pk["pink"]))

print("=== 逐字校验 ===")
print(f"{'字':<4}{'笔':<4}{'各教学格黑像素(应逐格累加)':<44}{'描红格粉像素'}")
for ch, n, blacks, pink in summary:
    print(f"{ch:<4}{n:<4}{str(blacks):<44}{pink}")

print()
if SKIPPED:
    print(f"（另有 {len(SKIPPED)} 个字不在参照数据里，已跳过：{''.join(SKIPPED[:12])}"
          f"{'…' if len(SKIPPED) > 12 else ''}）")
    print()
if problems:
    print(f"共校验 {len(summary)} 字，发现 {len(problems)} 个问题：")
    for p in problems:
        print("  -", p)
    raise SystemExit(1)
print(f"共校验 {len(summary)} 字，全部通过 ✓"
      f"（格数 / 红笔 / 黑笔累加 / 描红粉色 均正确）")