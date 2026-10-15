# -*- coding: utf-8 -*-
"""SVG 笔画路径解析工具。

支持数据源：hanzi-writer-data（MakeMeAHanzi 系，中文笔顺标准）
坐标约定：数据为 1024x1024 视框、y 轴向上；本模块统一转换为
图像坐标（y 轴向下），即 y_img = 900 - y_data。
"""
import re

CMD_RE = re.compile(r"([MLCQVHSTAZmlcqvhstaz])([^MLCQVHSTAZmlcqvhstaz]*)")
NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")
BEZIER_STEPS = 12

VIEWBOX = 1024
Y_FLIP = 900


def _nums(body):
    return [float(m.group()) for m in NUM_RE.finditer(body)]


def parse_svg_path(d):
    """解析 path d，返回折线列表 [[(x,y),...], ...]，坐标已转为图像坐标（y 向下）。"""
    polylines = []
    cur = [0.0, 0.0]
    current = None
    prev_ctrl = None   # 用于 S / T 的平滑控制点反射

    def push(p):
        nonlocal current
        if current is None:
            current = [p]
            polylines.append(current)
        else:
            current.append(p)

    for cmd, body in CMD_RE.findall(d):
        rel = cmd.islower()
        uc = cmd.upper()
        nums = _nums(body)
        j = 0

        if uc == "M":
            while j + 2 <= len(nums):
                x, y = nums[j], nums[j + 1]
                if rel:
                    x += cur[0]; y += cur[1]
                cur = [x, y]
                current = None
                push((x, y))
                j += 2
            prev_ctrl = None
        elif uc == "L":
            while j + 2 <= len(nums):
                x, y = nums[j], nums[j + 1]
                if rel:
                    x += cur[0]; y += cur[1]
                cur = [x, y]
                push((x, y))
                j += 2
            prev_ctrl = None
        elif uc == "H":
            while j + 1 <= len(nums):
                x = nums[j]
                if rel:
                    x += cur[0]
                cur = [x, cur[1]]
                push(tuple(cur))
                j += 1
            prev_ctrl = None
        elif uc == "V":
            while j + 1 <= len(nums):
                y = nums[j]
                if rel:
                    y += cur[1]
                cur = [cur[0], y]
                push(tuple(cur))
                j += 1
            prev_ctrl = None
        elif uc in ("C", "S"):
            while True:
                if uc == "C":
                    if j + 6 > len(nums):
                        break
                    c1x, c1y, c2x, c2y, x1, y1 = nums[j:j + 6]
                    j += 6
                    if rel:
                        c1x += cur[0]; c1y += cur[1]
                        c2x += cur[0]; c2y += cur[1]
                        x1 += cur[0]; y1 += cur[1]
                else:  # S：第一控制点是上一控制点的反射
                    if j + 4 > len(nums):
                        break
                    c2x, c2y, x1, y1 = nums[j:j + 4]
                    j += 4
                    if rel:
                        c2x += cur[0]; c2y += cur[1]
                        x1 += cur[0]; y1 += cur[1]
                    if prev_ctrl is None:
                        c1x, c1y = cur[0], cur[1]
                    else:
                        c1x = 2 * cur[0] - prev_ctrl[0]
                        c1y = 2 * cur[1] - prev_ctrl[1]
                x0, y0 = cur
                for s in range(1, BEZIER_STEPS + 1):
                    t = s / BEZIER_STEPS
                    it = 1 - t
                    bx = it**3*x0 + 3*it*it*t*c1x + 3*it*t*t*c2x + t**3*x1
                    by = it**3*y0 + 3*it*it*t*c1y + 3*it*t*t*c2y + t**3*y1
                    push((bx, by))
                prev_ctrl = (c2x, c2y)
                cur = [x1, y1]
        elif uc in ("Q", "T"):
            while True:
                if uc == "Q":
                    if j + 4 > len(nums):
                        break
                    cx, cy, x1, y1 = nums[j:j + 4]
                    j += 4
                    if rel:
                        cx += cur[0]; cy += cur[1]
                        x1 += cur[0]; y1 += cur[1]
                else:  # T：控制点是上一控制点的反射
                    if j + 2 > len(nums):
                        break
                    x1, y1 = nums[j:j + 2]
                    j += 2
                    if rel:
                        x1 += cur[0]; y1 += cur[1]
                    if prev_ctrl is None:
                        cx, cy = cur[0], cur[1]
                    else:
                        cx = 2 * cur[0] - prev_ctrl[0]
                        cy = 2 * cur[1] - prev_ctrl[1]
                x0, y0 = cur
                for s in range(1, BEZIER_STEPS + 1):
                    t = s / BEZIER_STEPS
                    it = 1 - t
                    bx = it*it*x0 + 2*it*t*cx + t*t*x1
                    by = it*it*y0 + 2*it*t*cy + t*t*y1
                    push((bx, by))
                prev_ctrl = (cx, cy)
                cur = [x1, y1]
        else:
            j = len(nums)

    # 统一转换为图像坐标：y 轴翻转
    out = []
    for pl in polylines:
        if len(pl) < 2:
            continue
        out.append([(x, Y_FLIP - y) for (x, y) in pl])
    return out


def stroke_polylines(d):
    return parse_svg_path(d)


def stroke_bbox(d):
    pts = [p for pl in parse_svg_path(d) for p in pl]
    if not pts:
        return None
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return (min(xs), max(xs), min(ys), max(ys))


def char_bbox(stroke_paths):
    """整字包围盒。"""
    boxes = [stroke_bbox(d) for d in stroke_paths]
    boxes = [b for b in boxes if b]
    if not boxes:
        return None
    return (min(b[0] for b in boxes), max(b[1] for b in boxes),
            min(b[2] for b in boxes), max(b[3] for b in boxes))

# ---------------------------------------------------------------- 字体探测
# 标签文字（①②③…「描红」）需要一款中文字体。不同系统字体位置不同，
# 这里按优先级找一款能用的，找不到就回退到 PIL 内置字体（并给出提示）。
FONT_CANDIDATES = [
    # Linux
    "/usr/share/fonts/Fonts_windows/HarmonyOS_Sans_SC_Bold.ttf",
    "/usr/share/fonts/Fonts_windows/HarmonyOS_Sans_SC_Regular.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
    # macOS
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    # Windows
    "C:/Windows/Fonts/msyhbd.ttc",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simhei.ttf",
    "C:/Windows/Fonts/simsun.ttc",
]


def resolve_font(verbose=True):
    """返回一款可用的中文字体路径；都找不到返回 None（调用方回退）。"""
    import os
    import subprocess
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            return p
    # 再试系统自带的字体匹配工具
    for cmd in (["fc-match", "-f", "%{file}", "sans-serif:lang=zh"],
                ["fc-match", "-f", "%{file}", "sans:lang=zh-cn"]):
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
            p = (r.stdout or "").strip()
            if p and os.path.exists(p):
                return p
        except Exception:
            pass
    if verbose:
        print("⚠ 没找到可用的中文字体，标签将用内置字体（可能显示成方框）。")
        print("  可安装一款中文字体，或在 stroke_lib.FONT_CANDIDATES 里加上你的字体路径。")
    return None
