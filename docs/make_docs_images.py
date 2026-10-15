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


def cut_line(im, orient):
    """在页面正中附近找**真正的裁切线**，返回 (方向 'h'/'v', 位置像素)。

    这里绝不能图省事取宽/高的一半：纵向版式的两张 A5 是**上下**堆叠的，裁切线是
    横线；横向版式左右并排，裁切线是竖线。README 里的示意图曾经对纵向版式画了一条
    **竖线**，把"上下裁开"标成了"左右裁开" —— 方向正好说反，就是因为在猜而不是在测。

    判据用虚线的特征：白色/非白交替次数远多于实心格线。实测纵向版式裁切线在
    49.6% 高处、交替 136 段，而普通横格线只有个位数，两者不会混。
    """
    a = np.asarray(im).astype(int)
    nw = (a.sum(axis=2) < 720)
    if orient == "portrait":
        best = bestr = 0
        for y in range(int(im.height * 0.42), int(im.height * 0.58)):
            seg = nw[y]
            if seg.sum() < 300 or abs(y - im.height / 2) > im.height * 0.08:
                continue
            r = int((np.diff(seg.astype(int)) == 1).sum())
            if r > bestr:
                best, bestr = y, r
        return "h", best or im.height // 2
    best = bestr = 0
    for x in range(int(im.width * 0.42), int(im.width * 0.58)):
        seg = nw[:, x]
        if seg.sum() < 300 or abs(x - im.width / 2) > im.width * 0.08:
            continue
        r = int((np.diff(seg.astype(int)) == 1).sum())
        if r > bestr:
            best, bestr = x, r
    return "v", best or im.width // 2


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

    # ---------- 3b. 笔顺练习模式（--stroke-order）：田字格逐格练笔顺 ----------
    SO = "output/示例_笔顺练习/示例_笔顺练习.pdf"
    if os.path.exists(os.path.join(ROOT, SO)):
        im3 = render(SO, 1, 420)
        g3, _g3b = grid_region(green_lines(im3, px_cm), CELL)
        # 取「水」的第一行田字格（8 格）——正好是两遍笔顺的前 4 个阶段，
        # 一眼能看出墨迹一格格变多
        box = (int(0.5 * px_cm), int(g3 * px_cm),
               int(20.4 * px_cm), int((g3 + CELL) * px_cm))
        crop = with_title(im3.crop(box), "③ 笔顺练习模式（--stroke-order）",
                          "每个田字格都按笔顺递进：第 1 格只有第 1 笔、第 2 格前 2 笔……"
                          "练 3 遍、再重复 3 个整字，剩下的格子留白自己写")
        p = os.path.join(OUT, "zoom-stroke-order.png")
        crop.save(p, optimize=True)
        made.append(p)

    # ---------- 3c. 拼音字帖：一方格一个音节（首格示范 + 描红 + 落带） ----------
    # 两个例子**跨页按内容找**：第一页的音节（bā mā bà mǎ dà tǔ mǐ nǚ）碰巧都没有降部，
    # 所以"降部进下格"的例子要到第二页（shàng 的 g）才有。按行号写死就会选错。
    PY = "output/示例_拼音/示例_拼音.pdf"
    if os.path.exists(os.path.join(ROOT, PY)):
        def solid_lines(imp):
            """只认**实线**（整行覆盖率 ≥80%）：虚线有间隙，覆盖率只有 ~60%。

            用旧版 green_lines 的判据在这里恰好踩边缘（浅绿虚线 241 vs 242 差 1，
            时灵时不灵）—— 换成覆盖率判据，实虚之分是结构性的，不靠色值巧合。
            """
            a = np.asarray(imp).astype(int)
            R_, G_, B_ = a[..., 0], a[..., 1], a[..., 2]
            green = (G_ > 180) & (G_ - R_ > 15) & (G_ - B_ > 15)
            rowcov = green.sum(axis=1)
            idx = np.flatnonzero(rowcov > 0.8 * a.shape[1])
            if idx.size == 0:
                return []
            groups = [[idx[0]]]
            for v in idx[1:]:
                (groups[-1].append(v) if v - groups[-1][-1] <= 3
                 else groups.append([v]))
            return [float(np.mean(g)) for g in groups]

        want = ["up", "down"]
        found = {}
        for pg in (1, 2):
            imp = render(PY, pg, 420)
            a = np.asarray(imp).astype(int)
            R_, G_, B_ = a[..., 0], a[..., 1], a[..., 2]
            greenmask = (G_ > 180) & (G_ - R_ > 15) & (G_ - B_ > 15)
            inkmask = (a.sum(axis=2) < 690) & ~greenmask   # 黑示范 + 灰描红都算墨
            lines = solid_lines(imp)
            for i in range(len(lines) - 1):
                if not want:
                    break
                top, bot = lines[i], lines[i + 1]
                h = bot - top
                if h < 2.0 * px_cm or h > 3.0 * px_cm:    # 一个四线格 ≈ 2.38cm
                    continue
                rr = np.flatnonzero(inkmask[int(top):int(bot) + 1].any(axis=1))
                if not rr.size:
                    continue
                if "up" in want and rr.min() < h / 3:
                    want.remove("up"); found["up"] = (imp, top, bot)
                elif "down" in want and rr.max() > 0.75 * h:
                    # 0.75h 才算真降部（~0.82）；基线圆弧超出只到 ~0.69 ——
                    # 用 2h/3+4px 的松阈值会把无降部音节误当成降部例
                    want.remove("down"); found["down"] = (imp, top, bot)

        def annotate_bands(im, top_px, bot_px):
            """右侧加一条白边标三条带各放什么 —— 整行都是格子，没有空位可叠字。"""
            strip = int(3.4 * px_cm)
            wide = Image.new("RGB", (im.width + strip, im.height), (255, 255, 255))
            wide.paste(im, (0, 0))
            d = ImageDraw.Draw(wide)
            h = bot_px - top_px
            f = font(int(h * 0.20))
            for i, txt in enumerate(("上格 · 升部与声调", "中格 · 字母主体", "下格 · 降部")):
                cy = top_px + h * (i + 0.5) / 3
                tw = d.textlength(txt, font=f)
                x = im.width + (strip - tw) / 2
                d.text((x, cy - h * 0.11), txt, font=f, fill=INK)
            return wide

        crops = []
        for key in ("up", "down"):
            if key not in found:
                continue
            imp, top, bot = found[key]
            pad = 0.30 * px_cm
            c = imp.crop((int(0.5 * px_cm), int(top - pad),
                          int(20.4 * px_cm), int(bot + pad)))
            crops.append(annotate_bands(c, int(pad), int(pad + (bot - top))))
        if crops:
            W = max(c.width for c in crops)
            gap = 22
            total = Image.new("RGB", (W, sum(c.height for c in crops) + gap), (244, 246, 249))
            yy = 0
            for c in crops:
                total.paste(c, (0, yy)); yy += c.height + gap
            crop = with_title(total, "④ 拼音字帖：一方格一个音节（--type=pinyin）",
                              "首格黑色示范、其余浅灰描红，字母主体写在中格、升部与声调在上格、"
                              "降部进下格 —— 格宽按最长音节自适应，字号绝不缩"
                              "（两张图是自动挑出来的：一张有声调、一张有降部）")
            p = os.path.join(OUT, "zoom-pinyin.png")
            crop.save(p, optimize=True)
            made.append(p)

    # ---------- 4. 打印裁剪示意 ----------
    # 裁切线是**从 PDF 里探测出来的**，不是想当然取宽/高的一半：
    # 纵向版式的两张 A5 上下堆叠，裁切线是横线；横向版式左右并排，裁切线是竖线。
    # 曾经这里对纵向版式画了一条竖线，把"上下裁开"标成了"左右裁开" —— 方向正好说反。
    def annotate_cut(im, orient, label):
        """把探测到的裁切线描成醒目的红色虚线，并在旁边放一个说明气泡。"""
        axis, pos = cut_line(im, orient)
        d = ImageDraw.Draw(im)
        if axis == "h":                              # 上下两张 A5：横线
            for x in range(0, im.width, 30):
                d.line([x, pos, min(x + 17, im.width), pos], fill=ACCENT, width=5)
        else:                                        # 左右两张 A5：竖线
            for y in range(0, im.height, 30):
                d.line([pos, y, pos, min(y + 17, im.height)], fill=ACCENT, width=5)
        f = font(30)
        tw = d.textlength(label, font=f)
        pad = 18
        if axis == "h":
            bx = int(im.width * 0.30)
            by = int(pos + im.height * 0.16)
            d.rounded_rectangle([bx, by, bx + tw + pad * 2, by + 64], radius=12,
                                fill=(255, 255, 255), outline=ACCENT, width=3)
            d.text((bx + pad, by + 15), label, font=f, fill=ACCENT)
            d.line([bx, by, bx + tw / 2, pos], fill=ACCENT, width=3)
        else:
            bx = int(im.width * 0.14)
            by = int(im.height * 0.70)
            d.rounded_rectangle([bx, by, bx + tw + pad * 2, by + 64], radius=12,
                                fill=(255, 255, 255), outline=ACCENT, width=3)
            d.text((bx + pad, by + 15), label, font=f, fill=ACCENT)
            d.line([bx + tw + pad * 2, by + 32, pos, by + 32], fill=ACCENT, width=3)
        return im

    panels = [
        annotate_cut(render(*PORTRAIT, dpi=170), "portrait", "沿这条横线剪开"),
        annotate_cut(render(*LANDSCAPE, dpi=170), "landscape", "沿这条竖线剪开"),
    ]
    hero2 = hstack([panel(fit_height(im, 900), cap, sub)
                    for im, cap, sub in zip(panels,
                                            ("A4 纵向 · 上下两张 A5", "A4 横向 · 左右两张 A5"),
                                            ("裁切线是横线 —— 沿它剪开，上下各一张",
                                             "裁切线是竖线 —— 沿它剪开，左右各一张"))])
    hero2 = with_title(hero2, "打印后沿正中裁切线剪开",
                       "不剪也完全可以正常用：不剪就是一张 A4 练两个字，剪开就是一人一张 A5")
    p = os.path.join(OUT, "print-cut.png")
    hero2.save(p, optimize=True)
    made.append(p)

    for p in made:
        print(f"  {os.path.relpath(p, ROOT)}  {os.path.getsize(p)//1024} KB")
    print(f"共 {len(made)} 张")


if __name__ == "__main__":
    sys.exit(main())