"""拼音字帖的 Excel 排版 —— A4 纵向，**一方格一个音节**（参考 Copybook 工程的格式）。

排版结构（每个音节占一行格子）：
    标签行       【组名】　bā　八
    四线三格     [首格·黑色示范] [描红格·浅灰] × N [留白格…]
    间隔行

四线三格用**原生 Excel 边框**画（和汉字字帖的田字格一样，不合并单元格）：
上下两条浅绿实线、中间两条更浅的虚线、每格之间一条竖实线 —— 线色取自参考
工程成品的实测值。首格是黑色示范，其余格浅灰描红（默认整行描红，--trace-times
可减少，留白格给孩子自己写）。

格宽**按最长音节自适应**：先用「x-height = 一条带」反解字号、量出最长音节的
墨迹宽度，格宽 = max(最小格宽, 最长墨迹 + 边距)，列数 = floor(总宽/格宽)。
于是 `zhuāng`（6 字母）会把格加宽、列数变少，但字号**绝不缩** —— 缩字号字母
就不在带上了，这是四线三格的硬约束，比「格子整齐」重要。

贴图仍然遵守 v1 的硬不变量：**纵向零内缩、高度正好等于三行子行之和**，
图里的带界才会和 Excel 画的虚线严格重合。
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

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

# 四线三格线色（参考工程 Copybook 成品实测；比汉字田字格的深绿更浅、更像课本）
LINE_SOLID = "92E892"      # 上下边线 + 竖分隔线
LINE_DASH = "CAF1CA"       # 中间两条辅助虚线
SOLID = Side(style="thin", color=LINE_SOLID)
DASHM = Side(style="dashed", color=LINE_DASH)

TITLE_PT, INFO_PT, HINT_PT, HDR_SPACER_PT = 26, 16, 16, 5
LABEL_PT, ITEM_SPACER_PT = 16, 6

TITLE = ARG.get("--title", "幼小衔接 · 拼音字帖")
SUBTITLE = ARG.get("--subtitle",
                   "四线三格：字母主体写在中格，升部与声调写在上格，降部写在下格")
OUT = ARG.get("--out", "拼音字帖.xlsx")
OUT = OUT if OUT.startswith("/") else f"{ROOT}/{OUT}"
IMG_DIR = ARG.get("--images", "pinyin_images")
IMG_DIR = IMG_DIR if IMG_DIR.startswith("/") else f"{ROOT}/{IMG_DIR}"
FONT = ARG.get("--font") or None
MODE = ARG.get("--mode", "excel")

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

# ── 几何：与渲染阶段同一份 cell_plan（确定性重算），并与 plan.json 互相印证 ──
FONT_PATH = RP.find_font(P.needed_chars([(i, h) for i, h, _ in FLAT]), FONT)
GEO = RP.cell_plan([i for i, _, _ in FLAT], FONT_PATH,
                   int(ARG.get("--band-px", "30")),
                   int(ARG.get("--bands", "3")),
                   int(ARG.get("--grid-px", "720")),
                   int(ARG.get("--inset-x", "4")))
BANDS, BAND_PX, CELL_PX, COLS = GEO["bands"], GEO["band"], GEO["cell_px"], GEO["cols"]
GRID_PX, BOX_W, BOX_H, INSET_X = GEO["grid_px"], GEO["box_w"], GEO["box_h"], GEO["inset_x"]
BAND_CM = BAND_PX / 96 * 2.54
GRID_CM = GRID_PX / 96 * 2.54
SUB_ROW_PT = S.snap_up_pt(GEO["band_px"] * 72 / 96)
ITEM_PT = LABEL_PT + BANDS * SUB_ROW_PT + ITEM_SPACER_PT
COL_W = S.col_width_for_px(CELL_PX, MODE)

# 描红格数：默认整行（首格示范之外全描），--trace-times=N 只描 N 格
_tt = ARG.get("--trace-times", "full")
N_TRACE = COLS - 1 if _tt == "full" else max(0, min(int(_tt), COLS - 1))


def _check_plan_consistency():
    """排版算出的几何必须和渲染阶段写下的完全一致 —— 差一个数带界就错位。"""
    p = os.path.join(IMG_DIR, "plan.json")
    if not os.path.exists(p):
        return
    import json
    with open(p, encoding="utf-8") as fh:
        saved = json.load(fh)
    if saved.get("geometry") != GEO:
        raise SystemExit("✗ 渲染阶段与排版阶段的几何不一致！\n"
                         f"  渲染: {saved.get('geometry')}\n  排版: {GEO}\n"
                         "  通常是两次运行之间字体或参数变了 —— 请重跑 render_pinyin。")


_check_plan_consistency()


def slug(text):
    """音节 -> 文件名。声调字符不进文件名，改用码点，避免文件系统编码问题。"""
    return "_".join(f"{ord(c):04x}" for c in text)


def build():
    wb = Workbook()
    ws = wb.active
    ws.title = "拼音字帖"
    ws.sheet_view.showGridLines = False
    S.setup_page(ws, "portrait", MARGIN_CM)
    for c in range(COLS):
        ws.column_dimensions[get_column_letter(c + 1)].width = COL_W

    avail = GRID_CM
    r = 1
    t = ws.cell(row=r, column=1, value=TITLE)
    t.font = Font(name="黑体", size=18, bold=True)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
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

        # 四线三格：3 个子行 × COLS 列。上下浅绿实线、中间浅虚线、每格之间竖实线
        band_row = r
        for b in range(BANDS):
            rr = band_row + b
            ws.row_dimensions[rr].height = SUB_ROW_PT
            for c in range(COLS):
                ws.cell(row=rr, column=c + 1).border = Border(
                    top=SOLID if b == 0 else None,
                    bottom=SOLID if b == BANDS - 1 else DASHM,
                    left=SOLID if c == 0 else None,
                    right=SOLID)
        r += BANDS

        ws.row_dimensions[r].height = ITEM_SPACER_PT
        r += 1

        blocks.append({"item": it, "hanzi": hz, "group": g, "label_row": lab_row,
                       "band_row": band_row, "next_row": r})
    return wb, ws, r, blocks, pages


def place(ws, blocks):
    """贴图：首格黑示范，随后 N 格浅灰描红。返回 (音节, 汉字, 实际贴了几格)。"""
    placed = []
    missing = []
    for b in blocks:
        it = b["item"]
        model = os.path.join(IMG_DIR, slug(it) + ".png")
        trace = os.path.join(IMG_DIR, slug(it) + "_t.png")
        n = 0
        if os.path.exists(model):
            # 纵向 0 内缩：贴图高度正好等于三行子行之和，带界才能和虚线重合
            S.place_image(ws, model, 0, b["band_row"] - 1, BOX_W, BOX_H, INSET_X, 0)
            n += 1
            for k in range(1, N_TRACE + 1):
                if os.path.exists(trace):
                    S.place_image(ws, trace, k, b["band_row"] - 1, BOX_W, BOX_H,
                                  INSET_X, 0)
                    n += 1
        else:
            missing.append(it)
        placed.append((it, b["hanzi"], n))
    if missing:
        print(f"  ⚠ 以下音节没有贴图，已留白：{''.join(missing)}")
    return placed


if __name__ == "__main__":
    wb, ws, last_row, blocks, pages = build()
    placed = place(ws, blocks)
    S.set_print_area(ws, COLS, last_row)
    wb.save(OUT)
    print(f"已生成 {OUT}")
    print(f"  音节 {len(blocks)} 个　共 {pages} 页　每行 {COLS} 格"
          f"（格宽 {CELL_PX}px = {CELL_PX/96*2.54:.2f}cm，"
          f"每音节 = 首格示范 + {N_TRACE} 格描红 + {COLS-1-N_TRACE} 格留白）")
    print(f"  四线三格：一条带 {GEO['band_px']}px，格高 {BAND_PX}px = {BAND_CM:.4f}cm；"
          f"子行高 {SUB_ROW_PT}pt；总宽 {GRID_PX}px = {GRID_CM:.2f}cm")
    print(f"  字体 {os.path.basename(FONT_PATH)}")
    print(f"  线色 实线#{LINE_SOLID} 虚线#{LINE_DASH}（参考工程实测值）")
    for it, hz, n in placed:
        print(f"    {it}　{hz}　-> {n} 格")
    for w in S.TEXT_WARNINGS:
        print(f"  ⚠ {w}")
