"""拼音字帖的 Excel 排版 —— A4 纵向，**一行一个音节（四线三格）**。

排版结构（每个音节一块）：
    标签行         【组名】　bā　八
    上带   ┐
    中带   ├ 四线三格（3 个子行，每条带 1 个子行）
    下带   ┘
    间隔行

四线三格用**原生 Excel 边框**画（和汉字字帖的田字格一样，不合并单元格）：
上下两条绿实线是外框，中间两条灰虚线是辅助线。音节的描红图贴在上面，
贴图高度正好等于 3 个子行之和、纵向不内缩 —— 这样图里的带界和 Excel 的虚线
严格重合，字母才会真的坐在带上。

v1 **一行一个音节**而不是「一个方格一个音节」：`zhuāng` 这种 6 个字母的音节在
2.38cm 的方格里放不下，硬塞就得缩小字号，字母一缩就不在带上了。
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

import pinyin_input as P
import render_pinyin as RP
import sheet_layout as S

ROOT = os.path.dirname(os.path.abspath(__file__))

ARG = {}
for _a in sys.argv[1:]:
    if _a.startswith("--"):
        _k, _sep, _v = _a.partition("=")
        ARG[_k] = _v if _sep else "1"

MARGIN_CM = 0.8
# 几何取自 render_pinyin.band_geometry —— 渲染和排版共用同一份，
# 否则图里的带界会和 Excel 画的虚线错开，字母整条带地偏
GEO = RP.band_geometry(int(ARG.get("--band-px", "30")),
                       int(ARG.get("--bands", "3")),
                       int(ARG.get("--cols", "8")),
                       int(ARG.get("--inset-x", "4")))
BANDS = GEO["bands"]
BAND_PX = GEO["band"]
GRID_PX = GEO["grid_px"]
BOX_W, BOX_H, INSET_X = GEO["box_w"], GEO["box_h"], GEO["inset_x"]
BAND_CM = BAND_PX / 96 * 2.54
GRID_CM = GRID_PX / 96 * 2.54    # 一行四线格的总宽（文字自检要用宽度，不是带高）
SUB_ROW_PT = S.snap_up_pt(GEO["band_px"] * 72 / 96)   # 子行高 = 整数像素，过 0.75pt 量化
COL_W = S.col_width_for_px(GRID_PX, ARG.get("--mode", "excel"))
N_TRACE = int(ARG.get("--trace-times", "2"))

TITLE_PT, INFO_PT, HINT_PT, HDR_SPACER_PT = 26, 16, 16, 5
LABEL_PT, ITEM_SPACER_PT = 16, 6
ITEM_PT = LABEL_PT + BANDS * SUB_ROW_PT + ITEM_SPACER_PT

TITLE = ARG.get("--title", "幼小衔接 · 拼音字帖")
SUBTITLE = ARG.get("--subtitle",
                   "四线三格：字母主体写在中格，升部与声调写在上格，降部写在下格")
OUT = ARG.get("--out", "拼音字帖.xlsx")
OUT = OUT if OUT.startswith("/") else f"{ROOT}/{OUT}"
IMG_DIR = ARG.get("--images", "pinyin_images")
IMG_DIR = IMG_DIR if IMG_DIR.startswith("/") else f"{ROOT}/{IMG_DIR}"
FONT = ARG.get("--font") or None
UNIT = ARG.get("--unit", "行")             # 一个音节的练习单位，用于文案

CHARS_ARG = ARG.get("--chars", "")
FILE_ARG = ARG.get("--file", "")

if FILE_ARG:
    _p = FILE_ARG if FILE_ARG.startswith("/") else f"{ROOT}/{FILE_ARG}"
    ITEMS = P.parse_items(open(_p, encoding="utf-8").read())
elif CHARS_ARG:
    ITEMS = P.parse_items(CHARS_ARG)
else:
    ITEMS = P.parse_items("bā 八\nmā 妈\nbà 爸\nmǎ 马")
FLAT = [(it, hz, g) for g, lst in ITEMS for it, hz in lst]
if not FLAT:
    raise SystemExit("✗ 没有解析出任何音节，检查 --chars / --file 的内容")
_bad = P.validate([(i, h) for i, h, _ in FLAT])
if _bad:
    raise SystemExit("✗ 这些不是合法音节："
                     + "、".join(_bad)
                     + "\n  音节只能由字母和带调的元音组成（声调可写 ba1/ma3/ma5）。")

FONT_PATH = RP.find_font(P.needed_chars([(i, h) for i, h, _ in FLAT]), FONT)


def slug(text):
    """音节 -> 文件名。声调字符不进文件名，改用码点，避免各种文件系统编码问题。"""
    return "_".join(f"{ord(c):04x}" for c in text)


def build():
    wb = Workbook()
    ws = wb.active
    ws.title = "拼音字帖"
    ws.sheet_view.showGridLines = False
    S.setup_page(ws, "portrait", MARGIN_CM)
    ws.column_dimensions["A"].width = COL_W

    avail = GRID_CM
    r = 1
    t = ws.cell(row=r, column=1, value=TITLE)
    t.font = Font(name="黑体", size=18, bold=True)
    t.alignment = Alignment(horizontal="center", vertical="center")
    S.check_text_width(TITLE, 18, r, avail)
    ws.row_dimensions[r].height = TITLE_PT
    r += 1
    s = ws.cell(row=r, column=1, value="姓名：____________     日期：______月____日")
    s.font = Font(name="黑体", size=10, color="444444")
    s.alignment = Alignment(horizontal="left", vertical="center")
    S.check_text_width(s.value, 10, r, avail)
    ws.row_dimensions[r].height = INFO_PT
    r += 1
    s = ws.cell(row=r, column=1, value=SUBTITLE)
    s.font = Font(name="黑体", size=10, color="444444")
    s.alignment = Alignment(horizontal="left", vertical="center")
    S.check_text_width(s.value, 10, r, avail)
    ws.row_dimensions[r].height = HINT_PT
    r += 1
    ws.row_dimensions[r].height = HDR_SPACER_PT
    r += 1

    # 可打印高度（pt）。分页按**块高**算：块与块之间分页，块内绝不切开
    usable = (29.7 - 2 * MARGIN_CM) / 2.54 * 72
    used = TITLE_PT + INFO_PT + HINT_PT + HDR_SPACER_PT
    blocks, pages = [], 1
    for idx, (it, hz, g) in enumerate(FLAT):
        if used + ITEM_PT > usable + 1e-6:        # 这一块放不下 -> 先分页
            S.add_page_break(ws, r - 1)
            used, pages = 0.0, pages + 1
        used += ITEM_PT
        lab = ws.cell(row=r, column=1, value=f"【{g}】　{it}　　{hz}" if hz else f"【{g}】　{it}")
        lab.font = Font(name="黑体", size=12, bold=True, color=S.GREEN)
        lab.alignment = Alignment(horizontal="left", vertical="center")
        S.check_text_width(lab.value, 12, r, avail)
        ws.row_dimensions[r].height = LABEL_PT
        lab_row = r
        r += 1

        # 四线三格：三条带 = 三个子行；上下绿实线，中间两条灰虚线
        band_row = r
        for b in range(BANDS):
            rr = band_row + b
            ws.row_dimensions[rr].height = SUB_ROW_PT
            cell = ws.cell(row=rr, column=1)
            top = S.OUTER if b == 0 else None
            bottom = S.OUTER if b == BANDS - 1 else S.DASH
            from openpyxl.styles import Border
            cell.border = Border(top=top, bottom=bottom)
        r += BANDS

        ws.row_dimensions[r].height = ITEM_SPACER_PT
        r += 1

        blocks.append({"item": it, "hanzi": hz, "group": g, "label_row": lab_row,
                       "band_row": band_row, "next_row": r})
    return wb, ws, r, blocks, pages


def load_manifest():
    """读渲染阶段写的 plan.json —— 里面是**实际画出的描红遍数**。

    音节太长时一遍都放不下（或只放得下 1 遍），排版必须照实回报，
    不能想当然地写 N_TRACE。
    """
    import json
    p = os.path.join(IMG_DIR, "plan.json")
    if not os.path.exists(p):
        return {}
    with open(p, encoding="utf-8") as fh:
        return json.load(fh).get("items", {})


def place(ws, blocks):
    """把描红图贴上去，返回每个音节 (音节, 汉字, 实际描红遍数)。"""
    man = load_manifest()
    plans, missing = [], []
    for b in blocks:
        it = b["item"]
        src = os.path.join(IMG_DIR, slug(it) + ".png")
        info = man.get(slug(it), {})
        drawn = int(info.get("drawn", 0))
        if os.path.exists(src):
            # 纵向 0 内缩：贴图高度正好等于三行子行之和，带界才能和虚线重合
            S.place_image(ws, src, 0, b["band_row"] - 1, BOX_W, BOX_H, INSET_X, 0)
        else:
            missing.append(it)
        plans.append((it, b["hanzi"], drawn))
    if missing:
        print(f"  ⚠ 以下音节没有描红图，已留白：{''.join(missing)}")
    return plans


if __name__ == "__main__":
    wb, ws, last_row, blocks, pages = build()
    plans = place(ws, blocks)
    S.set_print_area(ws, 1, last_row)
    wb.save(OUT)
    print(f"已生成 {OUT}")
    print(f"  音节 {len(blocks)} 个　共 {pages} 页　末行 {last_row}"
          f"　每页最多 {int(((29.7-2*MARGIN_CM)/2.54*72) // ITEM_PT)} 个")
    print(f"  四线三格：一条带 {GEO['band_px']}px，共 {BAND_PX}px = {BAND_CM:.4f}cm；"
          f"子行高 {SUB_ROW_PT}pt；总宽 {GRID_PX}px = {GRID_PX/96*2.54:.2f}cm")
    print(f"  字体 {os.path.basename(FONT_PATH)}　每音节描红 {N_TRACE} 遍")
    for it, hz, drawn in plans:
        print(f"    {it}　{hz}　-> 描红 {drawn} 遍 + 留白")
    for w in S.TEXT_WARNINGS:
        print(f"  ⚠ {w}")