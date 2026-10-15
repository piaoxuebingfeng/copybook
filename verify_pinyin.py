"""拼音字帖验收 —— 从产物反查，不信生成器的自述。

查四件事，每件都能独立把「看着差不多但其实是错的」抓出来：

1. **三带等高**：每个四线三格必须由 3 个等高的子行组成（四线三格的立身之本）
2. **贴图对齐**：描红图的锚点必须在三条带的第一行、**纵向零偏移**，高度必须
   **正好等于三行子行之和**。差 1px，图里的带界就和 Excel 的虚线错开，
   字母会整条带地偏 —— 而且不报错。
3. **贴图内容**：把 xlsx 里嵌的图按**像素内容**反查是哪个音节，与音节气清单逐块比对
   （不看文件名，重复渲染 / 贴错图都会被抓出来）
4. **PDF 落带**：导出后量真实像素 —— 两条虚虚线必须把上下绿实线**严格三等分**，
   描红墨迹的顶部要进上带、底部要落在中带下沿（基线）

用法：
    python3 verify_pinyin.py 字帖.xlsx --chars="bā 八 mā 妈" [--pdf=字帖.pdf]
"""
import glob
import hashlib
import os
import sys

import numpy as np
from PIL import Image

import pinyin_input as P
import render_pinyin as RP

ROOT = os.path.dirname(os.path.abspath(__file__))

ARG = {}
for _a in sys.argv[1:]:
    if _a.startswith("--"):
        _k, _sep, _v = _a.partition("=")
        ARG[_k] = _v if _sep else "1"

TOL_CM = 0.02          # 三带等高的容差
FAIL = []


def fail(msg):
    FAIL.append(msg)
    print(f"  ✗ {msg}")


def ok(msg):
    print(f"  ✓ {msg}")


def pixhash(img):
    """按**解码后的像素**做指纹，而不是 PNG 字节 —— 编码差异不该算作不同。"""
    return hashlib.md5(np.array(img.convert("RGBA")).tobytes()).hexdigest()


# ---------------------------------------------------------------- xlsx 侧
def find_bands(ws):
    """结构识别四线三格：上边绿实线 + 两条灰虚线 + 下边绿实线，且三行等高。

    认「边框长什么样」而不是认行号 —— 行号会随排版改动而变。
    """
    def bd(cell, which):
        s = getattr(cell.border, which)
        if s is None or not s.style:
            return None
        col = getattr(s.color, "rgb", None)
        return (s.style, (str(col)[2:] if col and len(str(col)) == 8 else str(col)).upper())

    bands, heights = [], {}
    for r in sorted(int(k) for k in ws.row_dimensions):
        heights[r] = ws.row_dimensions[r].height
    for r in sorted(heights):
        try:
            t = bd(ws.cell(row=r, column=1), "top")
            b1 = bd(ws.cell(row=r, column=1), "bottom")
            b2 = bd(ws.cell(row=r + 1, column=1), "bottom")
            b3 = bd(ws.cell(row=r + 2, column=1), "bottom")
        except Exception:
            continue
        if not (t and t[0] == "thin" and b1 and b1[0] == "dashed"
                and b2 and b2[0] == "dashed" and b3 and b3[0] == "thin"):
            continue
        bands.append(r)
    return bands, heights


def images_by_row(ws):
    EMU = 9525
    out = {}
    for im in ws._images:
        a = im.anchor
        out.setdefault(a._from.row, []).append((a, im))
    return out, EMU


def check_xlsx(path, items, geo, n_trace, font):
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb.active
    print(f"\n=== xlsx 几何与贴图 ===  {os.path.basename(path)}")

    bands, heights = find_bands(ws)
    print(f"  识别到 {len(bands)} 个四线三格，清单里有 {len(items)} 个音节")
    if len(bands) != len(items):
        fail(f"四线三格个数 {len(bands)} 与音节个数 {len(items)} 不一致")

    # 1) 三带等高
    bad_eq = []
    for r in bands:
        v = [heights.get(r + i) for i in range(3)]
        if None in v or len(set(round(x, 4) for x in v)) != 1:
            bad_eq.append((r, v))
    if bad_eq:
        fail(f"三带不等高的块：{bad_eq[:3]}")
    else:
        ok(f"全部 {len(bands)} 个块三带等高（{heights[bands[0]]}pt x 3）")

    # 2) 贴图对齐
    ibyr, EMU = images_by_row(ws)
    sub_px = geo["band_px"]
    bad_anchor, missing = [], []
    for r in bands:
        lst = ibyr.get(r - 1, [])
        if len(lst) != 1:
            missing.append((r, len(lst)))
            continue
        a, im = lst[0]
        cy = a.ext.cy / EMU
        cx = a.ext.cx / EMU
        if abs(cy - geo["box_h"]) > 0.5:
            bad_anchor.append((r, "高度", cy, geo["box_h"]))
        if abs(a._from.rowOff / EMU) > 0.5:
            bad_anchor.append((r, "纵向偏移非零", a._from.rowOff / EMU, 0))
        if abs(cx - geo["box_w"]) > 0.5:
            bad_anchor.append((r, "宽度", cx, geo["box_w"]))
    if missing:
        fail(f"这些块没有唯一的贴图：{missing[:3]}")
    if bad_anchor:
        fail(f"贴图锚点不对（差 1px 就整条带偏）：{bad_anchor[:3]}")
    else:
        ok(f"全部贴图纵向零偏移、高度 {geo['box_h']}px = 三行子行之和、宽 {geo['box_w']}px")

    # 3) 贴图内容 -> 反查音节
    lookup = {}
    for it, hz in items:
        img, drawn = RP.render_band(it, font, geo["box_w"], geo["box_h"], n_trace)
        lookup[pixhash(img)] = (it, hz, drawn)
    wrong, unknown = [], 0
    for r in bands:
        lst = ibyr.get(r - 1, [])
        if not lst:
            continue
        img = Image.open(lst[0][1].ref)
        hit = lookup.get(pixhash(img))
        if hit is None:
            unknown += 1
            continue
        it, hz, drawn = hit
        # 墨迹必须落在正确的带上
        a = np.array(img.convert("RGBA"))[..., 3] > 20
        h = img.size[1]
        rows = np.flatnonzero(a.any(axis=1))
        if not rows.size:
            wrong.append((it, "整张图没有墨迹"))
            continue
        top, bot = int(rows.min()), int(rows.max())
        if bot > int(2 * h / 3) + 2:          # 底部不该越到中带以下，除非真有降部
            pass
        if top > int(h / 3):                  # 顶部必须进上带
            wrong.append((it, f"顶部 {top} 没进上带（带界 {h//3}）"))
        if abs(bot - 2 * h / 3) > 0.10 * h:
            wrong.append((it, f"基线 {bot} 偏离 2/3（应为 {2*h/3:.0f}）"))
    if unknown:
        fail(f"{unknown} 张贴图在清单里找不到对应的音节（贴错图或渲染不一致）")
    if wrong:
        fail(f"落带不对：{wrong[:5]}")
    if not unknown and not wrong:
        ok(f"全部 {len(bands)} 张贴图按像素反查命中，且墨迹落带正确")
    return ws


# ---------------------------------------------------------------- PDF 侧
def check_pdf(path, items, geo, n_trace, font):
    """PDF 侧：四线格必须严格三等分；印出来的墨迹位置必须和渲染器的意图一致。

    注意**不能**用「墨迹底部 ≈ 2/3」当判据：`shàng` 的 `g` 有降部，底部本来就该
    伸到下带（78%），拿基线去卡它会把正确的产物判成错的。真正的不变量是
    「PDF 里量到的墨迹范围 == 渲染器为该音节生成的那张图里的墨迹范围」，
    它同时覆盖升部、降部和声调，且对任何音节都成立。
    """
    import pymupdf
    print(f"\n=== PDF 真实像素 ===  {os.path.basename(path)}")
    # 渲染器意图：每个音节的墨迹在 90px 高盒子里占哪几行
    want = []
    for it, hz in items:
        img, _ = RP.render_band(it, font, geo["box_w"], geo["box_h"], n_trace)
        a = np.array(img.convert("RGBA"))[..., 3] > 20
        rr = np.flatnonzero(a.any(axis=1))
        want.append((int(rr.min()), int(rr.max())) if rr.size else None)

    doc = pymupdf.open(path)
    DPI, SCALE = 200, 200 / 96.0        # 图里 1px = PDF 里 2.08px
    TOL = 5                             # 抗锯齿边缘的容差（PDF 侧 ~2px@96dpi）
    bad, total, checked = 0, 0, 0
    for pno in range(doc.page_count):
        pm = doc[pno].get_pixmap(dpi=DPI)
        arr = np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.width, pm.n)
        R, G, B = (arr[..., i].astype(int) for i in range(3))
        pink = (abs(R - 232) < 45) & (abs(G - 192) < 45) & (abs(B - 192) < 45) & (R > G + 10)
        green = (G > 120) & (R < 110) & (B < 130)
        gray = (abs(R - G) < 14) & (abs(G - B) < 14) & (R > 100) & (R < 215)
        W = pm.width

        def rows(frac, mask):
            return [y for y in range(mask.shape[0])
                    if mask[y].any() and mask[y].sum() >= frac * W]

        def collapse(ys, gap=4):
            out = []
            for y in ys:
                if out and y - out[-1][-1] <= gap:
                    out[-1].append(y)
                else:
                    out.append([y])
            return [int(np.mean(g)) for g in out]

        gr, dr = collapse(rows(0.80, green)), collapse(rows(0.20, gray))
        for i in range(0, len(gr) - 1, 2):
            top, bot = gr[i], gr[i + 1]
            h = bot - top
            total += 1
            idx = total - 1                      # 第几块 = 第几个音节（都按文档顺序）
            mid = [y for y in dr if top < y < bot]
            t1, t2 = top + h / 3, top + 2 * h / 3
            if len(mid) != 2 or abs(mid[0] - t1) > 3 or abs(mid[1] - t2) > 3:
                bad += 1
                fail(f"第 {pno+1} 页第 {total} 块三带不等分：虚线 {mid}，应在 {t1:.0f}/{t2:.0f}")
                continue
            rr = np.flatnonzero(pink[top:bot + 1].any(axis=1))
            if not rr.size or idx >= len(want) or want[idx] is None:
                continue
            checked += 1
            got = (float(rr.min()), float(rr.max()))
            exp = (want[idx][0] * SCALE, want[idx][1] * SCALE)
            if abs(got[0] - exp[0]) > TOL or abs(got[1] - exp[1]) > TOL:
                it = items[idx][0]
                bad += 1
                fail(f"第 {pno+1} 页第 {total} 块（{it}）：墨迹 {got[0]:.0f}~{got[1]:.0f}px，"
                     f"渲染意图 {exp[0]:.0f}~{exp[1]:.0f}px（容差 {TOL}）")
    if not bad:
        ok(f"{doc.page_count} 页共 {total} 个四线格：虚线严格三等分（≤3px@200dpi），"
           f"{checked} 个块的墨迹位置与渲染意图一致（≤{TOL}px）")
    return doc.page_count, total


if __name__ == "__main__":
    xlsx = None
    for a in sys.argv[1:]:
        if not a.startswith("--"):
            xlsx = a if a.startswith("/") else os.path.join(ROOT, a)
            break
    if not xlsx:
        raise SystemExit(__doc__)

    if ARG.get("--file"):
        f = ARG["--file"]
        src = open(f if f.startswith("/") else os.path.join(ROOT, f), encoding="utf-8").read()
    else:
        src = ARG.get("--chars", "")
    items = [(i, h) for _, lst in P.parse_items(src) for i, h in lst]
    if not items:
        raise SystemExit("✗ 需要 --chars 或 --file 提供音节清单，否则无法反查")
    bad = P.validate(items)
    if bad:
        raise SystemExit(f"✗ 非法音节：{bad}")

    geo = RP.band_geometry(int(ARG.get("--band-px", "30")), int(ARG.get("--bands", "3")),
                           int(ARG.get("--cols", "8")), int(ARG.get("--inset-x", "4")))
    n_trace = int(ARG.get("--trace-times", "2"))
    font = RP.find_font(P.needed_chars(items), ARG.get("--font"))

    check_xlsx(xlsx, items, geo, n_trace, font)
    pdf = ARG.get("--pdf")
    if pdf:
        check_pdf(pdf if pdf.startswith("/") else os.path.join(ROOT, pdf),
                  items, geo, n_trace, font)

    print()
    if FAIL:
        print(f"❌ 拼音字帖验收未通过：{len(FAIL)} 项")
        for m in FAIL:
            print(f"  · {m}")
        sys.exit(1)
    print("✅ 拼音字帖验收通过（三带等高 / 贴图对齐 / 落带正确 / 严格三等分）")