# -*- coding: utf-8 -*-
"""幼小衔接汉字字帖生成器（v3 · 重做版）

设计要点
--------
1. 笔顺教学条（图片）：每个字一条，N 个教学格 + 1 个描红示范格。
   第 k 格画出前 k 笔 —— 第 k 笔【红色】= 当前要写的这一笔，
   前面的笔画【黑色】= 已经写好的部分（与样例的红黑规则一致）。
   这样孩子一眼就能看出"这一笔落在字的哪个位置"。
2. 田字格：Excel 原生单元格，每格由 2x2 子格拼成 ——
   外框绿色实线，内部灰色虚线十字。**全程不合并单元格**，
   因为实测合并后 Excel/LibreOffice 都会丢掉内部十字线。
3. 描红字：透明背景 PNG（楷体·浅粉）叠在田字格上，
   田字格边框不受影响，十字线在任何软件里都完整。
4. 尺寸换算（关键，网传"列宽8.38/行高50=2cm"是错的）：
   Excel 列宽 -> 像素： px = 7 * width + 5      (Calibri 11 / 宋体 11)
   行高 -> 像素：       px = pt * 96 / 72
   要得到 2cm 正方形：半个格子 = 1cm = 37.795px
       子列宽 = (37.795 - 5) / 7 = 4.685
       子行高 = 37.795 * 72 / 96 = 28.35 pt
"""
import json
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

ROOT = os.path.dirname(os.path.abspath(__file__))
from openpyxl.worksheet.pagebreak import Break

# ---------------- 版式参数 ----------------
CELLS_PER_ROW = 8        # 每行田字格数
TRACE_CELLS = 3          # 第 1 行前几格放描红字
LABEL_PT = 12            # 字头标签行高（a4 版式用）
SPACER_PT = 2            # 字块之间空行高（a4 版式用）
MARGIN_CM = 0.8          # 页边距

# ---- 先解析命令行：下面很多常量依赖它 ----
ARG = {}
for _a in __import__("sys").argv[1:]:
    if _a.startswith("--"):
        _k, _sep, _v = _a.partition("=")
        ARG[_k] = _v if _sep else "1"

# 版式：
#   a5 = 一张 A4 纵向放 **2 个字**，上下各占半页（= 一张 A5），可从正中裁成两张 A5
#   a4 = 原来的紧凑版式，一张 A4 放 4 个字
LAYOUT = {"a5": "a5", "2": "a5", "a4": "a4", "4": "a4"}.get(
    ARG.get("--layout", "a5"), "a5")
_CELL_REQ = float(ARG.get("--cell-cm", 2.4 if LAYOUT == "a5" else 2.0))
PER_PAGE = 2 if LAYOUT == "a5" else 4

# 纸张方向（只对 --layout=a5 有意义）：
#   portrait  = A4 纵向，上下两张 A5（每张 21 x 14.85cm，横版 A5）—— 默认
#   landscape = A4 横向，左右两张 A5（每张 14.85 x 21cm，竖版 A5）
ORIENT = {"portrait": "portrait", "p": "portrait", "v": "portrait", "纵向": "portrait",
          "landscape": "landscape", "l": "landscape", "h": "landscape", "横向": "landscape"
          }.get(ARG.get("--orient", "portrait"), "portrait")
if ORIENT == "landscape" and LAYOUT != "a5":
    __import__("sys").stderr.write(
        "错误：--orient=横向 只适用于 --layout=a5（一张 A4 放 2 个字）\n")
    __import__("sys").exit(1)
LANDSCAPE = (ORIENT == "landscape")

# A4 横向时每半页只有 14.85cm 宽，放不下 8 个 2.4cm 的格子 —— 用 5 格（= 11.9cm）。
CELLS_PER_ROW = 5 if LANDSCAPE else 8
PAGE_W_CM = 29.7 if LANDSCAPE else 21.0
PAGE_H_CM = 21.0 if LANDSCAPE else 29.7

PX_PER_CM = 96 / 2.54

# ---- 田字格尺寸**按整数像素定义** ----
# 行高的最小单位是 1px（= 0.75pt @96dpi），列宽也可以精确到 1px。
# 所以格子边长必须取**整数像素**，Excel 与 LibreOffice 才会渲染出**精确的正方形**；
# 若按 cm 取整（如 1.2cm -> 45.35px），行高会被取整成 46px 而列宽仍是 45.35px，
# 结果格子是 1.217 x 1.190cm —— 不是正方形。
SUB_PX = max(4, int(round(_CELL_REQ / 2 / 2.54 * 96)))   # 半格边长（像素）
CELL_PX = SUB_PX * 2                                     # 整格边长（像素）
CELL_CM = CELL_PX / 96 * 2.54                            # 实际厘米边长
SUB_CM = CELL_CM / 2
SUB_ROW_PT = SUB_PX * 72 / 96                            # 子行高(pt) = 整数像素

# 列宽的「像素 -> 列宽」换算，Excel/WPS 与 LibreOffice 的常量不同：
#   Excel / WPS :  px = 7.000 * w + 5.0
#   LibreOffice :  px = 7.374 * w + 0.1   （实测标定）
# 反解出列宽，使两边都恰好得到 SUB_PX 像素。
_MODE = ARG.get("--mode", "excel")
_v = ARG.get("--out", "幼小衔接字帖.xlsx")
_OUT = _v if _v.startswith("/") else f"{ROOT}/{_v}"

_MODELS = {"excel": (7.0, 5.0), "lo": (7.374, 0.1)}
_slope, _offset = _MODELS[_MODE]
SUB_COL_W = round((SUB_PX - _offset) / _slope, 4)        # 子列宽 -> 恰好 SUB_PX 像素

# 横向版式的列布局（每半页固定 14.85cm）：
#   [半页格子][右侧留白][内 0.5cm][内 0.5cm][半页格子][右侧留白]
# 裁切线 = 前一个 0.5cm 间隔列的右边框 = 纸正中 14.85cm。
LAND_HALF_W_CM = 14.85   # 横向时每半页宽度 = 29.7 / 2
# 裁切线左右各留一点白：左边 0.5cm、右边 0.8cm。
# 这样剪开后两张 A5 的左边距**都是 0.8cm**（左边的来自页边距、右边的来自这 0.8cm），
# 右侧留白也都是 2.14cm —— 两张纸完全对称。
LAND_INNER_L_CM = 0.5
LAND_INNER_R_CM = MARGIN_CM
GUTTER_W_CM = LAND_INNER_L_CM
# 半页里格子排不满，剩下的就是"右侧留白"（用户选的方案：格子保持 2.4cm）
LAND_SLACK_CM = (LAND_HALF_W_CM - MARGIN_CM - LAND_INNER_L_CM) - CELLS_PER_ROW * CELL_CM
# 注意：右半页的 14.85cm 里也要扣掉右侧页边距，否则打印区域会超出可打印宽度，
# 触发 fitToWidth 把整页缩小 —— 格子就不再是 2.38cm 了。
LAND_TAIL_CM = ((LAND_HALF_W_CM - MARGIN_CM) - LAND_INNER_R_CM
                - CELLS_PER_ROW * CELL_CM)
# 最外侧那一条留白再收 0.1cm：让打印区域略小于可打印宽度。
# 否则"正好相等"时 fitToWidth 可能把整页缩一丁点，格子就不再是 2.38125cm 了。
LAND_TAIL_TRIM_CM = 0.1
_GUT_PX = round(GUTTER_W_CM / 2.54 * 96)
GUTTER_COL_W = round((_GUT_PX - _offset) / _slope, 4)      # 裁切线间隔列宽
_SLK_PX = round(LAND_SLACK_CM / 2.54 * 96)
_TAIL_PX = round((LAND_SLACK_CM - LAND_TAIL_TRIM_CM) / 2.54 * 96)
SLACK_COL_W = round((_SLK_PX - _offset) / _slope, 4)       # 半页右侧留白列宽
_TAIL2_PX = round((LAND_TAIL_CM - LAND_TAIL_TRIM_CM) / 2.54 * 96)
TAIL_COL_W = round((_TAIL2_PX - _offset) / _slope, 4)      # 最外侧留白列宽
INNER_R_COL_W = round((round(LAND_INNER_R_CM / 2.54 * 96) - _offset) / _slope, 4)

TOTAL_SUBCOLS = CELLS_PER_ROW * 2          # 半页的子列数（合并单元格用）
HALF_SUBCOLS = TOTAL_SUBCOLS
# 横向：左半 + 2 个间隔列（裁切线在它们中间）+ 右半
SHEET_COLS = (2 * HALF_SUBCOLS + 4) if LANDSCAPE else TOTAL_SUBCOLS

# ---- A5 版式（--layout=a5）----
# 一张 A4 纵向 = 上下两张 A5。两半的内容高度必须**严格相等**，从正中裁开才刚好。
# 另外所有行高一律取整到 0.75pt（= 1px@96dpi）的整数倍：Excel 与 LibreOffice 对
# 行高的取整方式不同，用非整倍数时累积十几行后裁切线会明显偏移。
HALF_PT = 396.0          # 每半页内容高度；2 半 = 792pt = 27.94cm < 28.1cm 可用
# 横向版式：两半是**左右并排**的，所以它们共用同一批行 —— 行高必须同步，
# 否则左右两张 A5 会不一样高。每半页可用高度 = 21 - 2*0.8 = 19.4cm。
LAND_HALF_PT = 547.5     # 547.5pt = 19.31cm，留 0.09cm 余量（0.75pt 的整数倍）
H5_HEAD_PT = 18.0        # 半页头（标题 + 姓名/日期）
H5_LABEL_PT = 14.25      # 字头行
H5_CUT_PT = 6.0          # 裁切线行


def snap_up_pt(x):
    """行高向上取整到 0.75pt 的整数倍（保证 Excel 与 LibreOffice 算出同样的高度）。"""
    import math
    return math.ceil(round(x / 0.75, 6)) * 0.75


GREEN = "00B050"
GRAY = "999999"
OUTER = Side(style="thin", color=GREEN)
DASH = Side(style="dashed", color=GRAY)


# ---- 命令行参数 ----
TRACE_INSET_PX = int(ARG.get("--trace-inset", 4))
CHARS_ARG = ARG.get("--chars", "")
DATA_FILE = ARG.get("--data", f"{ROOT}/stroke_data_cn.json")
TITLE = ARG.get("--title", "幼小衔接 · 汉字笔顺字帖")
SUBTITLE = ARG.get("--subtitle",
                   "看图学笔顺：红色 = 正在写的这一笔　黑色 = 已经写好的　最后一格 = 描红示范")
TITLE_PT = 26          # 标题行高（a4 版式用）
INFO_PT = 16           # 姓名/日期行高（a4 版式用）
HINT_PT = 16           # 看图学笔顺说明行高（a4 版式用）
HDR_SPACER_PT = 5      # 表头下方空行（a4 版式用）


def _dir_arg(key, dflt):
    v = ARG.get(key, dflt)
    return v if v.startswith("/") else f"{ROOT}/{v}"


STROKES_DIR = _dir_arg("--strokes", "stroke_images")
META_FILE = _dir_arg("--meta", "stroke_meta.json")
TRACES_DIR = _dir_arg("--traces", "trace_images")

META = json.load(open(META_FILE, encoding="utf-8"))

GROUPS_DEFAULT = [
    ("基础数字", "一二三十八人"),
    ("基础方位", "上下左右大小多少"),
    ("自然事物", "日月水火山石田土口手目耳"),
    ("基础字根", "木禾竹门问间中立"),
]

if CHARS_ARG:
    import stroke_store
    _list = stroke_store.parse_chars(CHARS_ARG)
    DATA = stroke_store.load_many(_list)
    _missing = [c for c in _list if c not in DATA]
    if _missing:
        print(f"⚠ 以下字缺笔画数据，已跳过：{''.join(_missing)}")
    GROUPS = [("生字", "".join(_list))]
    CHARS = [(c, "生字") for c in _list if c in DATA]
else:
    DATA = json.load(open(DATA_FILE, encoding="utf-8"))
    GROUPS = GROUPS_DEFAULT
    CHARS = [(c, g) for g, cs in GROUPS for c in cs]

# 打印区域必须覆盖最宽的教学条，否则导出/打印时教学条右侧会被裁掉。
# 田字格只用前 TOTAL_SUBCOLS 列，教学条可能更宽，故单独算打印区域列数。
STRIP_MAX_CM = max(CELL_CM * META[ch]["w"] / META[ch]["frame"] for ch, _ in CHARS)
PRINT_COLS = max(TOTAL_SUBCOLS,
                 min(int(STRIP_MAX_CM / SUB_CM) + 1, int((21 - 2 * MARGIN_CM) / SUB_CM)))
if LANDSCAPE:
    PRINT_COLS = SHEET_COLS               # 横向：打印区域要覆盖左右两个半页


# 教学条高度**逐字计算**（笔画多的字条会换行，条更高）
def strip_metrics(ch):
    cm = CELL_CM * (META[ch]["h"] / META[ch]["frame"])
    pt = round(cm / 2.54 * 72, 2)
    return cm, pt, pt + 1.2, cm / 2.54 * 96     # 条高cm, 条高pt, 行高pt, 图片高px


TEXT_WARNINGS = []


def text_width_cm(text, pt):
    """估算一段文字的渲染宽度（cm）。全角按 1 em，半角按 0.5 em。"""
    em = sum(1.0 if ord(c) > 0x2000 else 0.5 for c in str(text))
    return em * pt / 72 * 2.54


def check_text_width(text, pt, row, avail=None):
    """合并单元格会**直接裁掉**超出宽度的文字，所以这里必须自检。"""
    if avail is None:
        avail = TOTAL_SUBCOLS * SUB_CM
    w = text_width_cm(text, pt)
    if w > avail:
        TEXT_WARNINGS.append(
            f"第 {row} 行文字约 {w:.2f}cm，超过合并宽度 {avail:.2f}cm，会被裁掉：{str(text)[:40]}…")
    return w


def apply_grid_borders(ws, row_a, col_a):
    """为一个 2x2 田字格设置边框：外框绿实线 + 内部灰虚线十字。"""
    spec = [
        # (行偏移, 列偏移, top,      bottom,   left,     right)
        (0, 0, OUTER, DASH,  OUTER, DASH),
        (0, 1, OUTER, DASH,  DASH,  OUTER),
        (1, 0, DASH,  OUTER, OUTER, DASH),
        (1, 1, DASH,  OUTER, DASH,  OUTER),
    ]
    for dr, dc, top, bottom, left, right in spec:
        ws.cell(row=row_a + dr, column=col_a + dc).border = Border(
            top=top, bottom=bottom, left=left, right=right)


def build():
    wb = Workbook()
    ws = wb.active
    ws.title = "汉字笔顺字帖"
    ws.sheet_view.showGridLines = False

    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=MARGIN_CM / 2.54, right=MARGIN_CM / 2.54,
        top=MARGIN_CM / 2.54, bottom=MARGIN_CM / 2.54, header=0.2, footer=0.2)

    for i in range(1, PRINT_COLS + 1):
        ws.column_dimensions[get_column_letter(i)].width = SUB_COL_W

    r = 1
    # 标题
    t = ws.cell(row=r, column=1, value=TITLE)
    t.font = Font(name="黑体", size=18, bold=True)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_SUBCOLS)
    check_text_width(TITLE, 18, r)
    ws.row_dimensions[r].height = TITLE_PT
    r += 1
    # 姓名 / 日期
    s = ws.cell(row=r, column=1, value="姓名：____________     日期：______月____日")
    s.font = Font(name="黑体", size=10, color="444444")
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_SUBCOLS)
    check_text_width(s.value, 10, r)
    ws.row_dimensions[r].height = INFO_PT
    r += 1
    # 看图学笔顺说明（单独一行，避免和姓名日期挤在一行被裁掉）
    s = ws.cell(row=r, column=1, value=SUBTITLE)
    s.font = Font(name="黑体", size=10, color="444444")
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_SUBCOLS)
    check_text_width(s.value, 10, r)
    ws.row_dimensions[r].height = HINT_PT
    r += 1
    ws.row_dimensions[r].height = HDR_SPACER_PT
    r += 1

    blocks = []
    for idx, (ch, gname) in enumerate(CHARS):
        n = len(DATA[ch]["strokes"])
        # 字头
        lab = ws.cell(row=r, column=1, value=f"【{gname}】　{ch}　　共 {n} 画")
        lab.font = Font(name="黑体", size=10.5, bold=True, color=GREEN)
        lab.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_SUBCOLS)
        ws.row_dimensions[r].height = LABEL_PT
        r += 1
        # 教学条（逐字计算高度）
        strip_cm, strip_pt, strip_row_pt, strip_img_px = strip_metrics(ch)
        strip_row = r
        ws.row_dimensions[r].height = strip_row_pt
        r += 1
        # 两行田字格
        grid_rows = []
        for _ in range(2):
            ra = r
            for k in range(CELLS_PER_ROW):
                apply_grid_borders(ws, ra, 1 + k * 2)
            ws.row_dimensions[ra].height = SUB_ROW_PT
            ws.row_dimensions[ra + 1].height = SUB_ROW_PT
            grid_rows.append(ra)
            r += 2
        # 空行
        ws.row_dimensions[r].height = SPACER_PT
        r += 1
        # 该字块总高（pt）：字头 + 教学条行 + 两行田字格 + 空行
        block_pt = LABEL_PT + strip_row_pt + 2 * (2 * SUB_ROW_PT) + SPACER_PT
        blocks.append({"char": ch, "n": n, "strip_row": strip_row, "grid_rows": grid_rows,
                       "strip_img_px": strip_img_px, "strip_row_pt": strip_row_pt,
                       "block_pt": block_pt})
    return wb, ws, r, blocks


def place_inset_image(ws, path, col0, row0, size_px, inset_px):
    """把图片放进单元格，四周内缩 inset_px —— 这样不会遮住格子边框。"""
    im = XLImage(path)
    inner = size_px - 2 * inset_px
    im.width = inner
    im.height = inner
    anchor = OneCellAnchor()
    anchor._from = AnchorMarker(col=col0, colOff=pixels_to_EMU(inset_px),
                                row=row0, rowOff=pixels_to_EMU(inset_px))
    anchor.ext = XDRPositiveSize2D(pixels_to_EMU(inner), pixels_to_EMU(inner))
    im.anchor = anchor
    ws.add_image(im)


def add_images(ws, blocks):
    for b in blocks:
        ch = b["char"]
        m = META[ch]
        img_h = b["strip_img_px"]
        img_w = img_h * (m["w"] / m["h"])
        im = XLImage(f"{STROKES_DIR}/{ch}.png")
        im.height = img_h
        im.width = img_w
        _c0 = b.get("col0", 1)
        ws.add_image(im, f"{get_column_letter(_c0)}{b['strip_row']}")

        tr = f"{TRACES_DIR}/{ch}.png"
        if os.path.exists(tr):
            r0 = b["grid_rows"][0] - 1        # AnchorMarker 的 row/col 从 0 开始
            for k in range(TRACE_CELLS):
                place_inset_image(ws, tr, _c0 - 1 + 2 * k, r0, CELL_PX, TRACE_INSET_PX)


def add_page_breaks(ws, blocks):
    """按每个字块的**实际高度**分页，保证字块不被切到两页。"""
    usable = (29.7 - 2 * MARGIN_CM) / 2.54 * 72      # 可打印高度(pt)
    used = TITLE_PT + INFO_PT + HINT_PT + HDR_SPACER_PT     # 第 1 页表头占高
    cuts = 0
    for i, b in enumerate(blocks):
        if i > 0 and used + b["block_pt"] > usable:
            ws.row_breaks.append(Break(id=blocks[i - 1]["grid_rows"][1] + 2))
            used = 0
            cuts += 1
        used += b["block_pt"]
    return cuts


def build_half(ws, r, ch, gname, first):
    """构造半个 A4 页（= 一张 A5）。总高严格等于 HALF_PT，这样从正中裁开刚好两张 A5。"""
    n = len(DATA[ch]["strokes"])
    used = 0.0
    hcols = TOTAL_SUBCOLS // 2

    # ① 半页头：左边标题，右边姓名/日期（裁开后每张 A5 都是完整可用的）
    t = ws.cell(row=r, column=1, value=TITLE)
    t.font = Font(name="黑体", size=11, bold=True)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=hcols)
    check_text_width(TITLE, 11, r, avail=hcols * SUB_CM)
    s = ws.cell(row=r, column=hcols + 1, value="姓名：________　日期：___月___日")
    s.font = Font(name="黑体", size=9, color="444444")
    s.alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells(start_row=r, start_column=hcols + 1, end_row=r, end_column=TOTAL_SUBCOLS)
    check_text_width(s.value, 9, r, avail=hcols * SUB_CM)
    ws.row_dimensions[r].height = H5_HEAD_PT
    r += 1
    used += H5_HEAD_PT

    # ② 字头
    lab = ws.cell(row=r, column=1, value=f"【{gname}】　{ch}　　共 {n} 画")
    lab.font = Font(name="黑体", size=11, bold=True, color=GREEN)
    lab.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_SUBCOLS)
    ws.row_dimensions[r].height = H5_LABEL_PT
    r += 1
    used += H5_LABEL_PT

    # ③ 笔顺教学条
    strip_cm, strip_pt, strip_row_pt, strip_img_px = strip_metrics(ch)
    strip_row_pt = snap_up_pt(strip_row_pt)
    strip_row = r
    ws.row_dimensions[r].height = strip_row_pt
    r += 1
    used += strip_row_pt

    # ④ 田字格：把剩下的高度尽量排满（笔画多的字教学条更高，格子相应少排）
    tail = H5_CUT_PT if first else 0.0
    nrows = max(1, int((HALF_PT - used - tail) // (2 * SUB_ROW_PT)))
    while nrows > 1 and used + nrows * 2 * SUB_ROW_PT + tail > HALF_PT + 1e-6:
        nrows -= 1
    grid_rows = []
    for _ in range(nrows):
        ra = r
        for k in range(CELLS_PER_ROW):
            apply_grid_borders(ws, ra, 1 + k * 2)
        ws.row_dimensions[ra].height = SUB_ROW_PT
        ws.row_dimensions[ra + 1].height = SUB_ROW_PT
        grid_rows.append(ra)
        r += 2
        used += 2 * SUB_ROW_PT

    # ⑤ 垫底行：把这一半补齐到恰好 HALF_PT
    pad = HALF_PT - used - tail
    if pad > 0.005:
        ws.row_dimensions[r].height = round(pad, 2)
        r += 1

    # ⑥ 裁切线：只画在第一半末尾（A4 正中），沿着它剪开就是两张 A5
    if first:
        for c in range(1, PRINT_COLS + 1):
            ws.cell(row=r, column=c).border = Border(
                bottom=Side(style="dashed", color="AAAAAA"))
        ws.row_dimensions[r].height = H5_CUT_PT
        r += 1

    return {"char": ch, "n": n, "strip_row": strip_row, "grid_rows": grid_rows,
            "strip_img_px": strip_img_px, "strip_row_pt": strip_row_pt,
            "nrows": nrows, "next_row": r}



# ==================== A4 横向版式（左右两张 A5）====================
def _half_head(ws, r, c0, title_sz=10, info_sz=8.5):
    """半页表头：左边标题，右边姓名/日期。c0 是该半页的第一列。"""
    t = ws.cell(row=r, column=c0, value=TITLE)
    t.font = Font(name="黑体", size=title_sz, bold=True)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 4)
    check_text_width(TITLE, title_sz, r, avail=5 * SUB_CM)

    s = ws.cell(row=r, column=c0 + 5, value="姓名：________　日期：___月___日")
    s.font = Font(name="黑体", size=info_sz, color="444444")
    s.alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells(start_row=r, start_column=c0 + 5, end_row=r,
                   end_column=c0 + HALF_SUBCOLS - 1)
    check_text_width(s.value, info_sz, r, avail=5 * SUB_CM)


def build_a5_landscape():
    """A4 横向：左右并排两张 A5（各 14.85 x 21cm），沿正中竖线裁开。

    和纵向版式最大的不同：两半是**左右并排**的，也就是**共用同一批行**。所以
      * 教学条行高必须取两半中较高的那个（否则一行放不下高的那条）；
      * 田字格行数也必须两半相同 —— 否则左右两张 A5 会不一样高。
    这里按"较高那条教学条"能排下的最大行数来定，两半一致。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "汉字笔顺字帖"
    ws.sheet_view.showGridLines = False

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=MARGIN_CM / 2.54, right=MARGIN_CM / 2.54,
        top=MARGIN_CM / 2.54, bottom=MARGIN_CM / 2.54, header=0.2, footer=0.2)

    # 列布局：[左半][左留白][内0.5][内0.5][右半][右留白]
    C_SLACK_L = HALF_SUBCOLS + 1
    C_GUT1 = HALF_SUBCOLS + 2          # 它的右边框就是裁切线（纸正中 14.85cm）
    C_GUT2 = HALF_SUBCOLS + 3
    C0L, C0R = 1, HALF_SUBCOLS + 4
    C_SLACK_R = SHEET_COLS
    for i in range(1, SHEET_COLS + 1):
        if i == C_SLACK_L:
            _w = SLACK_COL_W
        elif i == C_GUT1:
            _w = GUTTER_COL_W
        elif i == C_GUT2:
            _w = INNER_R_COL_W
        elif i == C_SLACK_R:
            _w = TAIL_COL_W
        else:
            _w = SUB_COL_W
        ws.column_dimensions[get_column_letter(i)].width = _w

    pairs = [CHARS[i:i + 2] for i in range(0, len(CHARS), 2)]
    blocks, r, cuts = [], 1, 0

    for pi, pair in enumerate(pairs):
        halves = [(C0L, pair[0][0], pair[0][1])]
        if len(pair) > 1:
            halves.append((C0R, pair[1][0], pair[1][1]))

        # 教学条行高：两半取较高者（共用同一行）
        srow = max(snap_up_pt(strip_metrics(ch)[2]) for _, ch, _ in halves)
        fixed = H5_HEAD_PT + H5_LABEL_PT + srow
        cap = max(1, int((LAND_HALF_PT - fixed) // (2 * SUB_ROW_PT)))
        while cap > 1 and fixed + cap * 2 * SUB_ROW_PT > LAND_HALF_PT + 1e-6:
            cap -= 1

        page_top = r

        # ① 半页头
        for c0, ch, g in halves:
            _half_head(ws, r, c0)
        ws.row_dimensions[r].height = H5_HEAD_PT
        r += 1

        # ② 字头
        for c0, ch, g in halves:
            n = len(DATA[ch]["strokes"])
            lab = ws.cell(row=r, column=c0, value=f"【{g}】　{ch}　　共 {n} 画")
            lab.font = Font(name="黑体", size=10, bold=True, color=GREEN)
            lab.alignment = Alignment(horizontal="left", vertical="center")
            ws.merge_cells(start_row=r, start_column=c0, end_row=r,
                           end_column=c0 + HALF_SUBCOLS - 1)
            check_text_width(lab.value, 10, r, avail=HALF_SUBCOLS * SUB_CM)
        ws.row_dimensions[r].height = H5_LABEL_PT
        r += 1

        # ③ 教学条行（图片由 add_images 放到各自半页的列上）
        strip_row = r
        ws.row_dimensions[r].height = srow
        r += 1

        # ④ 田字格：两半行数一致
        grid_rows = []
        for _ in range(cap):
            ra = r
            for c0, ch, g in halves:
                for k in range(CELLS_PER_ROW):
                    apply_grid_borders(ws, ra, c0 + k * 2)
            ws.row_dimensions[ra].height = SUB_ROW_PT
            ws.row_dimensions[ra + 1].height = SUB_ROW_PT
            grid_rows.append(ra)
            r += 2

        # ⑤ 垫底行：把这一页补齐到恰好 LAND_HALF_PT
        used = fixed + cap * 2 * SUB_ROW_PT
        pad = LAND_HALF_PT - used
        if pad > 0.005:
            ws.row_dimensions[r].height = round(pad, 2)
            r += 1

        # ⑥ 裁切线：竖虚线，落在第一个间隔列的右边框（= 纸的正中 14.85cm）
        for rr in range(page_top, r):
            ws.cell(row=rr, column=C_GUT1).border = Border(
                right=Side(style="dashed", color="AAAAAA"))

        for c0, ch, g in halves:
            blocks.append({
                "char": ch, "n": len(DATA[ch]["strokes"]), "col0": c0,
                "strip_row": strip_row, "grid_rows": grid_rows,
                "strip_img_px": strip_metrics(ch)[3],
                "strip_row_pt": snap_up_pt(strip_metrics(ch)[2]),
                "nrows": cap})

        if pi < len(pairs) - 1:
            ws.row_breaks.append(Break(id=r - 1))
            cuts += 1

    return wb, ws, r, blocks, cuts


def build_a5():
    """一张 A4 纵向放 2 个字，各占半页。每对字占满一页并强制分页。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "汉字笔顺字帖"
    ws.sheet_view.showGridLines = False

    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=MARGIN_CM / 2.54, right=MARGIN_CM / 2.54,
        top=MARGIN_CM / 2.54, bottom=MARGIN_CM / 2.54, header=0.2, footer=0.2)

    for i in range(1, PRINT_COLS + 1):
        ws.column_dimensions[get_column_letter(i)].width = SUB_COL_W

    pairs = [CHARS[i:i + 2] for i in range(0, len(CHARS), 2)]
    blocks = []
    r = 1
    cuts = 0
    for pi, pair in enumerate(pairs):
        for hi, (ch, gname) in enumerate(pair):
            rec = build_half(ws, r, ch, gname, first=(hi == 0))
            blocks.append(rec)
            r = rec["next_row"]
        if pi < len(pairs) - 1:                 # 每对字一页
            ws.row_breaks.append(Break(id=r - 1))
            cuts += 1
    return wb, ws, r, blocks, cuts


if __name__ == "__main__":
    if LAYOUT == "a5" and LANDSCAPE:
        wb, ws, last_row, blocks, nb = build_a5_landscape()
    elif LAYOUT == "a5":
        wb, ws, last_row, blocks, nb = build_a5()
    else:
        wb, ws, last_row, blocks = build()
        nb = None
    add_images(ws, blocks)
    if nb is None:
        nb = add_page_breaks(ws, blocks)
    ws.print_area = f"A1:{get_column_letter(PRINT_COLS)}{last_row}"

    wb.save(_OUT)

    print(f"已生成 {_OUT}   [模式 {_MODE} / 版式 {LAYOUT}"
          f"{' 横向' if LANDSCAPE else ' 纵向' if LAYOUT == 'a5' else ''}]")
    print(f"  字数 {len(blocks)}　末行 {last_row}　分页符 {nb} 个")
    print(f"  子列宽 {SUB_COL_W}　子行高 {SUB_ROW_PT}pt　"
          f"-> 田字格 {CELL_CM}x{CELL_CM}cm ({CELL_PX:.1f}px)")
    for w in TEXT_WARNINGS:
        print(f"  ⚠ {w}")
    if LANDSCAPE:
        _pw = (2 * TOTAL_SUBCOLS * SUB_CM + LAND_SLACK_CM
               + LAND_INNER_L_CM + LAND_INNER_R_CM
               + LAND_TAIL_CM - LAND_TAIL_TRIM_CM)
    else:
        _pw = PRINT_COLS * SUB_CM
    print(f"  打印区域 {get_column_letter(PRINT_COLS)} 列 = {_pw:.1f}cm"
          f"（田字格 {TOTAL_SUBCOLS*SUB_CM:.1f}cm，最宽教学条 {STRIP_MAX_CM:.2f}cm）")
    hs = sorted({round(b["strip_row_pt"], 1) for b in blocks})
    print(f"  教学条行高 {hs} pt（笔画多的字自动换行，条更高）")
    if LAYOUT == "a5":
        ns = sorted({b["nrows"] for b in blocks})
        if LANDSCAPE:
            print(f"  A5 横向版式：一张 A4 左右 2 个字，每半页 {CELLS_PER_ROW} 格/行，"
                  f"每字 {ns} 行田字格")
            print(f"    每张 A5 = 14.85 x 21cm（竖版）；半页高 {LAND_HALF_PT}pt；"
                  f"裁切线在第 {HALF_SUBCOLS + 2} 列右边框 = 纸正中 14.85cm")
        else:
            print(f"  A5 纵向版式：一张 A4 上下 2 个字，每字 {ns} 行田字格；"
                  f"半个 A4 高 {HALF_PT}pt，裁切线在正中")
