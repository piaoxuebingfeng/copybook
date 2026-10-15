#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 README 里用到的示例图片（docs/*.png）。

图片全部从真实产物 PDF 渲染而来，不画示意图 —— README 上看到的就是打印出来的样子。
几何（格线位置、裁切线）都是从 PDF 现测的，所以换字表 / 换格子尺寸后重跑这个脚本
即可，不用手改坐标。

用法（在仓库根目录）：
    python3 docs/make_docs_images.py
"""
import os
import sys

import numpy as np
import pymupdf
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs")

# 三份成品里都含有「水 / 火」，取这几页做对比最直观
PORTRAIT = ("幼小衔接字帖_细笔画.pdf", 9)      # 第 9 页 = 水 火
LANDSCAPE = ("幼小衔接字帖_横向5格.pdf", 9)    # 第 9 页 = 水 火
COMPACT = ("幼小衔接字帖_紧凑4字.pdf", 5)      # 第 5 页 = 月 水 火 山

INK = (34, 40, 49)
GREEN = (0, 176, 80)
ACCENT = (219, 68, 55)
FONT_CANDIDATES = [
    "/usr/share/fonts/Fonts_windows/NotoSansSC-VF.ttf",
    "/usr/share/fonts/Fonts_windows/msyh.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "C:/Windows/Fonts/msyh.ttc",
]


def font(size, bold=False):
    """找一款中文字体；找不到就退回 PIL 默认字体（英文还行，中文会变方框）。"""
    for p in ([c for c in FONT_CANDIDATES if "Bold" in c or "bold" in c] if bold else []) + FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


def render(pdf, page_no, dpi):
    """把某一页渲染成 PIL 图。page_no 从 1 开始。"""
    doc = pymupdf.open(os.path.join(ROOT, pdf))
    pm = doc[page_no - 1].get_pixmap(dpi=dpi)
    im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
    doc.close()
    return im


def green_lines(im, px_cm):
    """找出横向绿格线的 y 坐标（cm）。用于自动定位教学条和第一行田字格。"""
    a = np.asarray(im).astype(int)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    green = (g > r + 40) & (g > b + 20) & (g > 110)
    rows = green.sum(axis=1)
    idx = np.flatnonzero(rows > px_cm * 4)      # 至少横跨 4cm 才算一条格线
    if idx.size == 0:
        return []
    groups = [[idx[0]]]
    for v in idx[1:]:
        (groups[-1].append(v) if v - groups[-1][-1] <= 3 else groups.append([v]))
    return [float(np.mean(g)) / px_cm for g in groups]


def grid_region(lines, cell_cm):
    """从所有绿线里挑出**真正的田字格区**。

    教学条用的也是田字格框（同一个绿色），所以它的上下边框同样是长绿线，直接用
    第一条线会截到教学条。田字格是一行挨一行铺开的 —— 它的格线必然构成一条
    等间距序列，取最长的那条，返回 (第一条格线 y, 第二条格线 y)。
    """
    best = []
    cur = []
    for a, b in zip(lines, lines[1:]):
        if 0.92 * cell_cm < b - a < 1.10 * cell_cm:
            cur = (cur or [a]) + [b]
        else:
            best = cur if len(cur) > len(best) else best
            cur = []
    best = cur if len(cur) > len(best) else best
    if len(best) < 2:
        raise SystemExit("找不到田字格区，检查 PDF 是否正常")
    return best[0], best[1]


def ink_box(im, px_cm, y0_cm, y1_cm, x0_cm, x1_cm, margin=0.14):
    """给定行范围，返回这一行里实际有墨迹的横向范围（cm），用来把两边空白裁掉。"""
    a = np.asarray(im).astype(int)
    band = a[int(y0_cm * px_cm):int(y1_cm * px_cm), int(x0_cm * px_cm):int(x1_cm * px_cm)]
    ink = (band < 235).any(axis=2)
    xs = np.flatnonzero(ink.any(axis=0))
    if xs.size == 0:
        return x0_cm, x1_cm
    return (x0_cm + xs[0] / px_cm - margin, x0_cm + xs[-1] / px_cm + margin)


def cut_x(im):
    """裁切线的 x 像素。版式上它就落在纸张正中（纵向 21/2、横向 29.7/2），
    所以直接用页宽的一半 —— 比去检测那排很淡的细虚线稳得多。"""
    return im.width // 2


def fit_height(im, h):
    w = max(1, round(im.width * h / im.height))
    return im.resize((w, h), Image.LANCZOS)


def panel(im, caption, sub, band=92, pad=18, bg=(255, 255, 255)):
    """给一张页图加底部说明条，返回拼好的面板。"""
    canvas = Image.new("RGB", (im.width + pad * 2, im.height + band + pad), bg)
    canvas.paste(im, (pad, pad))
    d = ImageDraw.Draw(canvas)
    d.rectangle([pad - 1, pad - 1, pad + im.width, pad + im.height], outline=(226, 232, 240))
    f1, f2 = font(30), font(21)
    d.text((pad + 2, im.height + pad + 16), caption, font=f1, fill=INK)
    d.text((pad + 2, im.height + pad + 52), sub, font=f2, fill=(100, 116, 139))
    return canvas


def hstack(imgs, gap=22, bg=(244, 246, 249)):
    w = sum(i.width for i in imgs) + gap * (len(imgs) + 1)
    h = max(i.height for i in imgs) + gap * 2
    c = Image.new("RGB", (w, h), bg)
    x = gap
    for i in imgs:
        c.paste(i, (x, gap))
        x += i.width + gap
    return c


def with_title(im, title, sub, bg=(244, 246, 249), pad=26, band=104):
    c = Image.new("RGB", (im.width + pad * 2, im.height + band + pad * 2), bg)
    c.paste(im, (pad, pad + band - 40))
    d = ImageDraw.Draw(c)
    d.text((pad + 4, 26), title, font=font(38, bold=True), fill=INK)
    d.text((pad + 4, 74), sub, font=font(22), fill=(100, 116, 139))
    return c


def main():
    os.makedirs(OUT, exist_ok=True)
    made = []

    # ---------- 1. 三版式总览（README 首图）----------
    panels = []
    for (pdf, pg), cap, sub in (
            (PORTRAIT, "A4 纵向 · 8 格/行", "上下两张横版 A5"),
            (LANDSCAPE, "A4 横向 · 5 格/行", "左右两张竖版 A5"),
            (COMPACT, "紧凑版式 · 8 格/行", "一张 A4 放 4 个字")):
        p = panel(fit_height(render(pdf, pg, 150), 760), cap, sub)
        panels.append(p)
    hero = with_title(hstack(panels), "同一份字表，三种打印版式",
                      "沿每页正中的虚线剪开，就是一张张独立的 A5 练习纸")
    p = os.path.join(OUT, "hero.png")
    hero.save(p, optimize=True)
    made.append(p)

    # ---------- 2. 三种版式各一张整页 ----------
    for (pdf, pg), name, cap, sub in (
            (PORTRAIT, "layout-portrait", "A4 纵向 · 2 字/页", "每字一张横版 A5，8 格/行 · 4 行练习"),
            (LANDSCAPE, "layout-landscape", "A4 横向 · 2 字/页", "每字一张竖版 A5，5 格/行 · 6 行练习"),
            (COMPACT, "layout-compact", "紧凑版式 · 4 字/页", "一张 A4 放 4 个字，8 格/行 · 2 行练习")):
        im = render(pdf, pg, 170)
        p = os.path.join(OUT, f"{name}.png")
        panel(im, cap, sub).save(p, optimize=True)
        made.append(p)

    # ---------- 3. 细节放大：笔顺教学条 / 田字格描红 ----------
    im = render(*PORTRAIT, dpi=420)
    px_cm = 420 / 2.54
    CELL = 2.38125
    g_top, g_next = grid_region(green_lines(im, px_cm), CELL)

    # 教学条：从「页边距 + 标题行 + 字头行」的底部，到田字格区第一条格线
    strip_top = (18 + 14.25) / 72 * 2.54 + 0.8
    x0, x1 = ink_box(im, px_cm, strip_top, g_top, 0.5, 20.4)
    crop = im.crop((int(x0 * px_cm), int(strip_top * px_cm),
                    int(x1 * px_cm), int(g_top * px_cm)))
    crop = with_title(crop, "① 笔顺教学条",
                      "逐笔演示：黑=已写完的笔画，红=当前这一笔，末尾 3 格是描红示范")
    p = os.path.join(OUT, "zoom-strip.png")
    crop.save(p, optimize=True)
    made.append(p)

    # 田字格描红行 = 田字格区的前两行格线之间，宽度取整个格区
    x0, x1 = ink_box(im, px_cm, g_top, g_next, 0.5, 20.4)
    box = (int(0.5 * px_cm), int(g_top * px_cm), int(20.4 * px_cm), int(g_next * px_cm))
    crop = with_title(im.crop(box), "② 田字格描红",
                      "前 3 格描红、后面留白自己写；格子是精确 2.38cm 正方形")
    p = os.path.join(OUT, "zoom-grid.png")
    crop.save(p, optimize=True)
    made.append(p)

    # ---------- 4. 打印裁剪示意 ----------
    im = render(*PORTRAIT, dpi=170)
    x = cut_x(im)
    d = ImageDraw.Draw(im)
    # 高亮裁切线（把淡虚线描成醒目的红色虚线）
    for y in range(0, im.height, 30):
        d.line([x, y, x, min(y + 17, im.height)], fill=ACCENT, width=5)
    # 标注气泡：放在下半页偏右，压到的是空白练习格，不会盖住字头
    f = font(30)
    label = "沿这条线剪开 →  上下/左右即两张 A5"
    tw = d.textlength(label, font=f)
    bx = min(x + 30, im.width - tw - 60)
    by = int(im.height * 0.70)
    pad = 18
    d.rounded_rectangle([bx, by, bx + tw + pad * 2, by + 64], radius=12,
                        fill=(255, 255, 255), outline=ACCENT, width=3)
    d.text((bx + pad, by + 15), label, font=f, fill=ACCENT)
    d.line([bx, by + 32, x, by + 32], fill=ACCENT, width=3)
    im = with_title(im, "打印后沿中线剪开",
                    "A4 上上下（或左右）排着两张 A5 —— 一页只练一个字，不剪也完全可以正常用")
    p = os.path.join(OUT, "print-cut.png")
    im.save(p, optimize=True)
    made.append(p)

    for p in made:
        print(f"  {os.path.relpath(p, ROOT)}  {os.path.getsize(p)//1024} KB")
    print(f"共 {len(made)} 张")


if __name__ == "__main__":
    sys.exit(main())