# -*- coding: utf-8 -*-
"""渲染"笔顺教学条"图片（填充真实字形版）。

与旧版的区别
------------
旧版把笔画路径当成"中心线"用粗笔描，但 hanzi-writer-data 里的 `strokes`
其实是笔画的**外轮廓闭合路径**。用粗笔描外轮廓 = 字形本身厚度 + 笔宽，
所以看起来总是偏粗，而且笔宽对粗细的控制很弱。

新版直接**把外轮廓填色**，得到的就是字真实的样子；再用形态学腐蚀
做整体减细（`--thin`，单位是 1024 数据盒的坐标），从而精确控制笔画粗细。
4 倍超采样 + LANCZOS 缩小，边缘平滑无锯齿。

用法：
  python3 render_strokes.py --thin=24 --out=stroke_images_fill --meta=stroke_meta_fill.json
"""
import json
import os
import sys
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from scipy.ndimage import distance_transform_edt
from stroke_lib import stroke_polylines, resolve_font

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = f"{ROOT}/stroke_data_cn.json"
CHARS_ARG = ""
FONT_PATH = None            # 启动时自动探测（见 _pick_font()）

# 尺寸参数（输出像素）
F = 240          # 每格边长
GAP = 12         # 格间距（压紧一点，让 8 画字的 9 格能单行放下）
GAP_BEFORE_PINK = 26
LABEL_H = 48     # 格下方标签高度
PAD = 8          # 字形相对格子的内缩
SS = 4           # 超采样倍数
BOX = 1024       # 数据坐标盒

MAX_W = 1920     # 教学条单行最大宽度（源像素）= 16cm = 田字格宽度，绝不超出格子区
ROW_GAP = 18     # 换行后两行之间的间距
THIN_UNITS = 10  # 减细量（1024 盒坐标）——实测安全上限：>10 会把"右/耳/间"的细横压断
OUT_DIR = f"{ROOT}/stroke_images_fill"
META_PATH = f"{ROOT}/stroke_meta_fill.json"
for _a in sys.argv[1:]:
    if _a.startswith("--thin="):
        THIN_UNITS = float(_a.split("=", 1)[1])
    elif _a.startswith("--out="):
        _v = _a.split("=", 1)[1]
        OUT_DIR = _v if _v.startswith("/") else f"{ROOT}/{_v}"
    elif _a.startswith("--meta="):
        _v = _a.split("=", 1)[1]
        META_PATH = _v if _v.startswith("/") else f"{ROOT}/{_v}"
    elif _a.startswith("--data="):
        DATA_PATH = _a.split("=", 1)[1]
    elif _a.startswith("--chars="):
        CHARS_ARG = _a.split("=", 1)[1]
    elif _a.startswith("--strip-max-cm="):
        MAX_W = int(float(_a.split("=", 1)[1]) * F / 2.0)   # 条内每格 2cm
    elif _a.startswith("--strip-max-cells="):
        # 直接按"最多几个田字格宽"来限制。田字格不是 2cm 时必须用这个，
        # 否则 --strip-max-cm 的 2cm 假设会让条超出格子区。
        MAX_W = int(float(_a.split("=", 1)[1]) * F)

# 单字表文件（每行一个字表）或直接给字串；缺省用项目里那份 34 字数据
if CHARS_ARG:
    import stroke_store
    from stroke_store import parse_chars as _pc
    DATA = stroke_store.load_many(_pc(CHARS_ARG))
else:
    DATA = json.load(open(DATA_PATH, encoding="utf-8"))

DEFAULT_CHARS = "一二三十八人上下左右大小多少日月水火山石田土口手目耳木禾竹门问间中立"

GREEN = (0, 176, 80)
GRAY_CROSS = (190, 190, 190)
BLACK = (26, 26, 26)
RED = (224, 49, 49)
PINK = (232, 192, 192)
WHITE = (255, 255, 255)
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫"

_poly_cache = {}


def polygons_of(ch, i):
    """第 i 笔的填充多边形（1024 盒坐标，y 已翻转到图像方向）。"""
    key = (ch, i)
    if key not in _poly_cache:
        _poly_cache[key] = stroke_polylines(DATA[ch]["strokes"][i])
    return _poly_cache[key]


def dashed_line(dr, p0, p1, fill, dash, gap, width):
    import math
    x0, y0 = p0
    x1, y1 = p1
    dist = math.hypot(x1 - x0, y1 - y0)
    if dist == 0:
        return
    ux, uy = (x1 - x0) / dist, (y1 - y0) / dist
    t = 0.0
    while t < dist:
        e = min(t + dash, dist)
        dr.line([(x0 + ux * t, y0 + uy * t), (x0 + ux * e, y0 + uy * e)], fill=fill, width=width)
        t = e + gap


def draw_frame(dr, x0, y0, size, s):
    dr.rectangle([x0, y0, x0 + size - 1, y0 + size - 1], outline=GREEN, width=max(1, round(3 * s)))
    cx = x0 + size / 2
    cy = y0 + size / 2
    dashed_line(dr, (cx, y0 + 2 * s), (cx, y0 + size - 2 * s), GRAY_CROSS,
                dash=7 * s, gap=6 * s, width=max(1, round(2 * s)))
    dashed_line(dr, (x0 + 2 * s, cy), (x0 + size - 2 * s, cy), GRAY_CROSS,
                dash=7 * s, gap=6 * s, width=max(1, round(2 * s)))


def stroke_mask(ch, stroke_ids, s, erode_px):
    """把若干笔填成一张 F*s 的灰度掩膜，并按 erode_px 做欧氏腐蚀。"""
    n = F * s
    m = Image.new("L", (n, n), 0)
    dm = ImageDraw.Draw(m)
    inner = (F - 2 * PAD) / BOX
    for i in stroke_ids:
        for pl in polygons_of(ch, i):
            pts = [((PAD + x * inner) * s, (PAD + y * inner) * s) for (x, y) in pl]
            if len(pts) >= 3:
                dm.polygon(pts, fill=255)
            elif len(pts) == 2:
                dm.line(pts, fill=255, width=max(2, int(6 * s)))
    if erode_px > 0:
        a = np.array(m) > 127
        if a.any():
            d = distance_transform_edt(a)
            a = d >= erode_px
        m = Image.fromarray((a * 255).astype(np.uint8))
    return m


def _pack(flags):
    """贪心换行：返回每行的「是否描红格」标记序列。"""
    rows, cur, cur_w = [], [], 0
    for is_pink in flags:
        gap = GAP_BEFORE_PINK if (is_pink and cur) else (GAP if cur else 0)
        if cur and cur_w + gap + F > MAX_W:
            rows.append(cur)
            cur, cur_w, gap = [], 0, 0
        cur.append(is_pink)
        cur_w += gap + F
    if cur:
        rows.append(cur)
    return rows


def _positions(rows):
    """给每行的格子算出 x 坐标（描红格前面留大一点的间距）。"""
    out = []
    for row in rows:
        cur_w, r = 0, []
        for is_pink in row:
            gap = GAP_BEFORE_PINK if (is_pink and r) else (GAP if r else 0)
            r.append((cur_w + gap, is_pink))
            cur_w += gap + F
        out.append(r)
    return out


def layout(n):
    """把 n 个教学格 + 1 个描红格排成若干行（超出可打印宽度就换行）。

    返回 (rows, total_w, total_h)，rows 是 [[(x, 是否描红格), ...], ...]。
    """
    rows = _pack([False] * n + [True])
    # 换行后如果最后一行只剩"描红"一个孤格，就从上一行挪一格下来作伴
    if len(rows) > 1 and len(rows[-1]) == 1:
        rows[-1].insert(0, rows[-2].pop())
    rows = _positions(rows)
    total_w = max(r[-1][0] + F for r in rows)
    total_h = len(rows) * (F + LABEL_H) + (len(rows) - 1) * ROW_GAP
    return rows, total_w, total_h


def render_char(ch, s=SS):
    n = len(DATA[ch]["strokes"])
    # 减细量换算到超采样像素
    inner_s = (F - 2 * PAD) * s / BOX
    erode_px = THIN_UNITS * inner_s

    rows, total_w, total_h = layout(n)

    im = Image.new("RGB", (total_w * s, total_h * s), WHITE)
    dr = ImageDraw.Draw(im)
    try:
        f_label = ImageFont.truetype(_pick_font(), 30 * s)
    except Exception:
        f_label = ImageFont.load_default()

    for ri, row in enumerate(rows):
        y0 = ri * (F + LABEL_H + ROW_GAP)
        for x0, is_pink in row:
            draw_frame(dr, x0 * s, y0 * s, F * s, s)
            # 全局笔序：教学行按顺序编号，描红格单独
            k = sum(1 for r_ in rows[:ri] for (_x, p) in r_ if not p) + \
                sum(1 for (_x, p) in row if not p and _x < x0)
            if is_pink:
                m = stroke_mask(ch, range(n), s, erode_px)
                im.paste(Image.new("RGB", (F * s, F * s), PINK), (x0 * s, y0 * s), m)
                label, lc = "描红", (120, 120, 120)
            else:
                if k > 0:
                    m = stroke_mask(ch, range(k), s, erode_px)
                    im.paste(Image.new("RGB", (F * s, F * s), BLACK), (x0 * s, y0 * s), m)
                m = stroke_mask(ch, range(k, k + 1), s, erode_px)
                im.paste(Image.new("RGB", (F * s, F * s), RED), (x0 * s, y0 * s), m)
                label = CIRCLED[k] if k < len(CIRCLED) else str(k + 1)
                lc = (60, 60, 60)
            bb = dr.textbbox((0, 0), label, font=f_label)
            dr.text((x0 * s + F * s / 2 - (bb[0] + bb[2]) / 2,
                     (y0 + F) * s + 8 * s - bb[1]), label, font=f_label, fill=lc)

    im = im.resize((total_w, total_h), Image.LANCZOS)
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, f"{ch}.png")
    im.save(out)
    return out, total_w, total_h, n


# ---------------- 自检：减细后每一笔都还在 ----------------
def _pick_font():
    global FONT_PATH
    if FONT_PATH is None:
        FONT_PATH = resolve_font() or ""
    return FONT_PATH


def check_no_stroke_lost(ch, s=SS):
    """减细后，每一笔的中心线（medians）必须仍落在该笔的填充掩膜里。"""
    inner_s = (F - 2 * PAD) * s / BOX
    erode_px = THIN_UNITS * inner_s
    n = len(DATA[ch]["strokes"])
    bad = []
    for i in range(n):
        m = np.array(stroke_mask(ch, [i], s, erode_px)) > 127
        if not m.any():
            bad.append((i + 1, "整笔消失"))
            continue
        med = DATA[ch]["medians"][i]
        ok = 0
        # 用与掩膜一致的映射（数据 y 轴向上，Y_FLIP=900）检查中心线点
        for (mx, my) in med:
            gx = (PAD + mx * ((F - 2 * PAD) / BOX)) * s
            gy = (PAD + (900 - my) * ((F - 2 * PAD) / BOX)) * s
            xi, yi = int(round(gx)), int(round(gy))
            if 0 <= xi < m.shape[1] and 0 <= yi < m.shape[0] and m[yi, xi]:
                ok += 1
        if ok < max(1, len(med) // 3):
            bad.append((i + 1, f"中心线 {ok}/{len(med)} 点在掩膜内"))
    return bad


if __name__ == "__main__":
    CHARS = list(DATA.keys()) if CHARS_ARG else list(DEFAULT_CHARS)
    meta, problems = {}, []
    for ch in CHARS:
        p, w, h, n = render_char(ch)
        meta[ch] = {"w": w, "h": h, "n": n, "frame": F}
        bad = check_no_stroke_lost(ch)
        if bad:
            problems.append((ch, bad))
    json.dump(meta, open(META_PATH, "w"), ensure_ascii=False, indent=1)
    print(f"共 {len(CHARS)} 张 -> {OUT_DIR}")
    print(f"渲染方式 = 填充字形轮廓；减细 THIN_UNITS = {THIN_UNITS}（1024 盒坐标）")
    print(f"元数据 -> {META_PATH}")
    if problems:
        print("警告：以下字有笔画被减细破坏：")
        for ch, bad in problems:
            print(f"  {ch}: {bad}")
    else:
        print("笔画完整性自检通过 ✓（每一笔的中心线都仍落在笔画内）")