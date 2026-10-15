"""共享版面层 —— Excel/WPS 与 LibreOffice 的纸张几何、列宽校准、贴图、文字自检。

这一层只放**与「格子里画什么」无关**的东西：汉字字帖（generate_zitie.py）和
拼音字帖（generate_pinyin.py）共用。两条生成线都踩过的坑集中在这里，动之前先读注释。

里面的每个常量都是**实测标定**出来的，不是查文档抄来的：
  · 列宽 <-> 像素：Excel/WPS 与 LibreOffice 的换算公式不同（差一个常数项）
  · 行高的最小单位是 1px = 0.75pt，所有行高必须量化到这个粒度
  · 打印区域必须比可打印宽度略小，否则 fitToWidth 会把整页缩小
"""
import math

from openpyxl.styles import Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# ---------------------------------------------------------------- 单位换算
# 列宽和行高的基准是 96dpi 的像素：1px = 0.75pt。格子边长一律按**整数像素**定义，
# 这样 Excel 与 LibreOffice 才会渲染出精确的正方形（按 cm 取整会一边 45px 一边 46px）。
PX_PER_CM = 96 / 2.54
PT_PER_PX = 72 / 96


def px_to_pt(px):
    return px * PT_PER_PX


def cm_to_px(cm):
    return cm / 2.54 * 96


def px_to_cm(px):
    return px / 96 * 2.54


def cm_to_pt(cm):
    return cm / 2.54 * 72


def snap_up_pt(x):
    """行高向上取整到 0.75pt（= 1px@96dpi）的整数倍。

    Excel 与 LibreOffice 对非整倍数行高的取整方式不同，累积十几行之后
    裁切线会明显偏移 —— 所以每个行高都必须先过这里。
    """
    return math.ceil(round(x / 0.75, 6)) * 0.75


# ---------------------------------------------------------------- 列宽标定
# 列宽的「像素 -> 列宽」换算，两家的公式不同（实测标定）：
#   Excel / WPS :  px = 7.000 * w + 5.0
#   LibreOffice :  px = 7.374 * w + 0.1
COL_MODELS = {"excel": (7.0, 5.0), "lo": (7.374, 0.1)}


def col_width_for_px(px, mode="excel"):
    """反解列宽，使 Excel/WPS（或 LibreOffice）恰好渲染成 px 像素。"""
    slope, offset = COL_MODELS[mode]
    return round((px - offset) / slope, 4)


# ---------------------------------------------------------------- 颜色与边框
GREEN = "00B050"        # 田字格/四线格外框
GRAY = "999999"         # 格内辅助线（田字格十字）
CUT_GRAY = "AAAAAA"     # 裁切线（比辅助线更淡）
OUTER = Side(style="thin", color=GREEN)
DASH = Side(style="dashed", color=GRAY)
CUT = Side(style="dashed", color=CUT_GRAY)


# ---------------------------------------------------------------- 页面
def setup_page(ws, orientation, margin_cm=0.8, paper=9):
    """A4（paperSize=9）+ 四边等边距 + fitToWidth。

    fitToWidth=1 时，只要打印区域**超过**可打印宽度，整页就会被缩小、
    格子就不再是正方形了 —— 所以调用方必须把打印区域取得比可打印宽度略小。
    """
    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = paper
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=margin_cm / 2.54, right=margin_cm / 2.54,
        top=margin_cm / 2.54, bottom=margin_cm / 2.54,
        header=0.2, footer=0.2)


def set_print_area(ws, ncols, last_row):
    """打印区域必须覆盖最宽的教学条/四线格，否则导出时右侧会被裁掉。"""
    ws.print_area = f"A1:{get_column_letter(ncols)}{last_row}"


def add_page_break(ws, row):
    """在第 row 行下方插一个分页符。"""
    from openpyxl.worksheet.pagebreak import Break
    ws.row_breaks.append(Break(id=row))


# ---------------------------------------------------------------- 贴图
def place_image(ws, path, col0, row0, w_px, h_px, inset_x=0, inset_y=0):
    """把图片放进单元格并精确控制尺寸与内缩（宽高可以不同）。

    注意：AnchorMarker 的 row/col 从 **0** 开始，而 openpyxl 的单元格行号从 1 开始，
    所以调用方要减 1。用 OneCellAnchor（而不是 ws.add_image(im, "A1")）是为了
    精确控制左上角偏移，否则图片会贴在单元格左上角、盖住边框。

    拼音字帖要用 inset_y=0：贴图的高度必须**正好**等于三行子行的高度之和，
    这样图里的两条带界才和 Excel 画的虚线严格重合（差 1px 字母就不在带上了）。
    """
    im = XLImage(path)
    im.width = w_px
    im.height = h_px
    anchor = OneCellAnchor()
    anchor._from = AnchorMarker(col=col0, colOff=pixels_to_EMU(inset_x),
                                row=row0, rowOff=pixels_to_EMU(inset_y))
    anchor.ext = XDRPositiveSize2D(pixels_to_EMU(w_px), pixels_to_EMU(h_px))
    im.anchor = anchor
    ws.add_image(im)


def place_inset_image(ws, path, col0, row0, size_px, inset_px):
    """正方形图片，四周内缩 inset_px —— 这样不会遮住格子边框（田字格描红用）。"""
    inner = size_px - 2 * inset_px
    place_image(ws, path, col0, row0, inner, inner, inset_px, inset_px)


def box_borders(ws, row, col, outer=OUTER, inner=DASH):
    """在 2x2 子单元格上画「外框 + 内部十字」——这就是田字格的画法。

    田字格**绝不能合并单元格**：合并会把内部十字一起删掉，只剩一个空框。
    所以用 2 个子行 x 2 个子列拼出来，外框绿实线、内部灰虚线。
    """
    spec = [
        # (行偏移, 列偏移, top,    bottom, left,  right)
        (0, 0, outer, inner, outer, inner),
        (0, 1, outer, inner, inner, outer),
        (1, 0, inner, outer, outer, inner),
        (1, 1, inner, outer, inner, outer),
    ]
    for dr, dc, top, bottom, left, right in spec:
        ws.cell(row=row + dr, column=col + dc).border = Border(
            top=top, bottom=bottom, left=left, right=right)


def dash_bottom(ws, row, col):
    """把某一行的下边框画成裁切线（虚线）。"""
    ws.cell(row=row, column=col).border = Border(bottom=CUT)


def dash_right(ws, row, col):
    """把某一列的右边框画成裁切线（虚线）。"""
    ws.cell(row=row, column=col).border = Border(right=CUT)


# ---------------------------------------------------------------- 文字宽度自检
# 合并单元格会**直接裁掉**超出宽度的文字（不报错、不换行），所以每处都要自检。
TEXT_WARNINGS = []


def text_width_cm(text, pt):
    """估算一段文字的渲染宽度（cm）。全角按 1 em，半角按 0.5 em。"""
    em = sum(1.0 if ord(c) > 0x2000 else 0.5 for c in str(text))
    return em * pt / 72 * 2.54


def check_text_width(text, pt, row, avail=None):
    """avail 单位是 cm；不传就用默认（调用方自己传更准）。"""
    if avail is None:
        avail = 20.0
    w = text_width_cm(text, pt)
    if w > avail:
        TEXT_WARNINGS.append(
            f"第 {row} 行文字约 {w:.2f}cm，超过合并宽度 {avail:.2f}cm，会被裁掉："
            f"{str(text)[:40]}…")
    return w