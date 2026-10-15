# -*- coding: utf-8 -*-
"""渲染田字格里的"描红"字（改为与笔顺教学行同一套字形）。

之前这里用霞鹜文楷**字体**渲染，而笔顺教学行用的是 hanzi-writer 的
笔画轮廓填充 —— 两处字形来源不同，所以田字格里的描红和教学行最后那个
粉色描红格看起来不是一个字。

现在统一：直接复用 render_strokes 的轮廓填充管线（同样的减细量、同样的映射），
田字格里的描红与教学行最后一格就是**同一个字形**，只是尺寸不同。

映射说明
--------
教学格里字形盒占 (F-2*PAD)/F = 93.3%（240px 格子里 224px）。
本脚本输出的 SIZE x SIZE 图中字形盒同样占 93.3%，
所以把整张图按"整个田字格"大小放入时，字形与教学行的比例、位置完全一致。
数据盒最小余量 28/1024 = 0.55mm（2cm 格子），透明边缘足够，
无需再额外内缩，不会遮住田字格边框。
"""
import os
import sys
import numpy as np
from PIL import Image

import render_strokes as R

ROOT = os.path.dirname(os.path.abspath(__file__))
SIZE = 300                        # 输出边长
PINK = (232, 192, 192)            # 描红颜色
SS_TRACE = 5                      # 超采样：240*5 = 1200 -> 缩到 300

OUT_DIR = f"{ROOT}/trace_images_fill"
CHARS_ARG = ""
for _a in sys.argv[1:]:
    if _a.startswith("--out="):
        _v = _a.split("=", 1)[1]
        OUT_DIR = _v if _v.startswith("/") else f"{ROOT}/{_v}"
    elif _a.startswith("--chars="):
        CHARS_ARG = _a.split("=", 1)[1]

REQ = []
if CHARS_ARG:
    import stroke_store
    REQ = stroke_store.parse_chars(CHARS_ARG)
    R.DATA.update(stroke_store.load_many(REQ))


def render_trace(ch):
    n = len(R.DATA[ch]["strokes"])
    inner_s = (R.F - 2 * R.PAD) * SS_TRACE / R.BOX
    erode_px = R.THIN_UNITS * inner_s          # 与教学行完全相同的减细量
    m = R.stroke_mask(ch, range(n), SS_TRACE, erode_px)
    m = m.resize((SIZE, SIZE), Image.LANCZOS)

    im = Image.new("RGBA", (SIZE, SIZE), PINK + (255,))
    im.putalpha(m)
    return im


def check(im):
    """自检：字形不越出字形盒（留白足够），并报告墨迹中心偏移。"""
    a = np.array(im)[..., 3] > 20
    ys, xs = np.where(a)
    if len(xs) == 0:
        return ["空图"], 0, 0
    box = (R.F - 2 * R.PAD) / R.F * SIZE       # 字形盒边长
    pad = (SIZE - box) / 2
    probs = []
    if xs.min() < pad - 1 or xs.max() > SIZE - pad + 1:
        probs.append(f"横向超出字形盒 (x {xs.min()}~{xs.max()}，盒内应为 {pad:.0f}~{SIZE-pad:.0f})")
    if ys.min() < pad - 1 or ys.max() > SIZE - pad + 1:
        probs.append(f"纵向超出字形盒 (y {ys.min()}~{ys.max()})")
    dx = (xs.min() + xs.max()) / 2 - SIZE / 2
    dy = (ys.min() + ys.max()) / 2 - SIZE / 2
    return probs, dx, dy


if __name__ == "__main__":
    CHARS = (REQ if CHARS_ARG else list(R.DEFAULT_CHARS))
    os.makedirs(OUT_DIR, exist_ok=True)
    bad, offs = [], []
    for ch in CHARS:
        im = render_trace(ch)
        p, dx, dy = check(im)
        if p:
            bad.append((ch, p))
        offs.append((abs(dx), abs(dy), ch))
        im.save(f"{OUT_DIR}/{ch}.png")
    o = np.array([(a, b) for a, b, _ in offs])
    mm = 20 / SIZE
    print(f"共 {len(CHARS)} 张 -> {OUT_DIR}")
    print(f"字形来源 = 笔画轮廓填充（与教学行同源）；减细量 = {R.THIN_UNITS}")
    print(f"墨迹中心相对格子中心：|dx| 最大 {o[:,0].max():.1f}px ({o[:,0].max()*mm:.2f}mm)，"
          f"|dy| 最大 {o[:,1].max():.1f}px ({o[:,1].max()*mm:.2f}mm)")
    if bad:
        print("警告：", bad)
    else:
        print("字形留白 / 不贴边 自检通过 ✓")