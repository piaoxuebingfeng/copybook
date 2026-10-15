"""拼音渲染 —— 字体探测 + 四线三格里的音节渲染。

**这个模块的核心不变量**：字母必须坐在正确的带上。
四线三格的三条带分工是固定的：上带放升部字母和声调符号（`b d h k l`、`ā` 的一横）、
中带放字母主体（x-height）、下带放降部（`g p q y` 的尾巴）。

所以这里**不居中、不按框缩放**，而是：
  1. 反解字号，使 **x-height 恰好等于一条带**（实测 'x' 的墨迹高度，不信字体度量）
  2. 用 `anchor="ms"`（水平居中 + 基线对齐）把基线钉在框高的 **2/3** 处

于是中带 = [1/3, 2/3]，正好是字母主体；升部与降部自然伸进上下带。
这与汉字那侧「格子边长必须取整数像素」是同一类精确度问题：做歪了不会报错，
只会「看着像但不像课本」。
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.abspath(__file__))

# 描红用的浅粉，与汉字字帖的田字格描红**同一个色号**（视觉统一）
PINK = (232, 192, 192)
INK = (34, 40, 49)

# 候选字体：按「教学合适度 + 可分发性」排序。
# 楷体最适合儿童识字教学，但它是微软/方正私有字体 —— 本机装了就用，**绝不随仓库分发**。
FONT_CANDIDATES = [
    "fonts/*.ttf", "fonts/*.ttc", "fonts/*.otf",          # 用户自备优先
    "/usr/share/fonts/Fonts_windows/simkai.ttf",          # 楷体 KaiTi（私有）
    "/usr/share/fonts/truetype/arphic/ukai.ttc",          # AR PL UKai（Arphic，可分发，楷体风格）
    "/usr/share/fonts/truetype/arphic/uming.ttc",         # AR PL UMing（Arphic，可分发）
    "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/Fonts_windows/msyh.ttc",            # 微软雅黑（私有）
    "/usr/share/fonts/Fonts_windows/simsun.ttc",          # 宋体（私有）
]

# 运行期实际探测到的候选（含 glob 展开）
_RESOLVED = None


def candidates():
    """展开候选字体列表（glob 优先，保留优先级顺序）。"""
    global _RESOLVED
    if _RESOLVED is None:
        import glob
        out = []
        for pat in FONT_CANDIDATES:
            if any(ch in pat for ch in "*?["):
                out.extend(sorted(glob.glob(os.path.join(ROOT, pat))))
            else:
                out.append(pat)
        _RESOLVED = [p for p in out if os.path.exists(p)]
    return _RESOLVED


# ---------------------------------------------------------------- 字体覆盖校验
def missing_chars(path, chars):
    """返回字体缺哪些字符。

    有 fontTools 就用 cmap 精确判断（**推荐**，也是 requirements 里列它的原因）；
    没有就退回「渲染出来是空白就当作缺字」的粗略判断，并明确告知不可靠。
    """
    chars = {c for c in chars if not c.isspace()}
    if not chars:
        return set()
    try:
        from fontTools.ttLib import TTFont, TTCollection
        if path.lower().endswith(".ttc"):
            cmaps = [set(f.getBestCmap()) for f in TTCollection(path).fonts]
        else:
            cmaps = [set(TTFont(path, fontNumber=0, lazy=True).getBestCmap())]
        best = min(cmaps, key=lambda cm: len(chars - {c for c in chars if ord(c) in cm}))
        return {c for c in chars if ord(c) not in best}
    except ImportError:
        f = ImageFont.truetype(path, 48)
        out = set()
        for c in chars:
            try:
                if f.getmask(c).getbbox() is None:      # 渲染为空白 -> 大概率缺字
                    out.add(c)
            except Exception:
                out.add(c)
        return out


def find_font(chars, explicit=None):
    """挑一个覆盖到 chars 的字体。都不覆盖就抛错并说清缺什么、怎么办。"""
    if explicit:
        p = explicit if explicit.startswith("/") else os.path.join(ROOT, explicit)
        if not os.path.exists(p):
            raise SystemExit(f"✗ 指定的字体不存在：{p}")
        miss = missing_chars(p, chars)
        if miss:
            raise SystemExit(f"✗ 指定的字体缺少字符 {''.join(sorted(miss))}：{p}")
        return p
    tried = []
    for p in candidates():
        miss = missing_chars(p, chars)
        if not miss:
            return p
        tried.append((os.path.basename(p), "".join(sorted(miss))))
    msg = ["✗ 没有找到能完整渲染这些字音的字体：" + "".join(sorted(set(chars) - set(" ")))]
    for name, miss in tried[:8]:
        msg.append(f"    {name}：缺 {miss}")
    msg.append("  解决：把一个覆盖拼音声调的字体放到 fonts/ 目录，或用 --font=路径 指定。")
    msg.append("  可自由分发的推荐：AR PL UKai / Noto Serif CJK / LXGW WenKai。")
    raise SystemExit("\n".join(msg))


# ---------------------------------------------------------------- 字号反解
def measure_x_height(path, size):
    """实测 'x' 的墨迹高度（= x-height，像素）。宁可靠像素也不信字体度量。"""
    f = ImageFont.truetype(path, size)
    img = Image.new("L", (size * 4, size * 4), 0)
    # anchor="ls"：左对齐 + 基线对齐到 y=size*2，于是墨迹底部就是基线
    ImageDraw.Draw(img).text((size, size * 2), "x", font=f, fill=255, anchor="ls")
    box = img.getbbox()
    if not box:
        raise SystemExit(f"✗ 字体渲染 'x' 得到空白：{path}")
    return box[3] - box[1]


def size_for_x_height(path, target_px, ref=200):
    """反解字号，使 x-height = target_px。字号与 x-height 近似线性，迭代两轮足够。"""
    size = ref
    for _ in range(3):
        xh = measure_x_height(path, size)
        if xh <= 0:
            break
        size = max(4, int(round(size * target_px / xh)))
    return size


# ---------------------------------------------------------------- 音节渲染
def render_band(text, path, box_w, box_h, n_trace, color=PINK, ss=4):
    """把一行四线格的描红部分渲染成一张 RGBA 图。

    box_w x box_h 是**最终贴进 Excel 的像素尺寸**，内部按 4 倍超采样再缩回来
    （与汉字描红图同一套做法）。三带的边界固定在 1/3、2/3 处，所以只要贴图时
    **纵向不内缩**，图里的带界就和 Excel 画的虚线严格重合。

    返回 (图, 实际画出的描红遍数)。音节太长时可能画不满 n_trace 遍 —— 实际遍数
    必须回报给调用方，验收脚本拿它逐格核对。
    """
    if box_w <= 0 or box_h <= 0:
        raise ValueError("box 尺寸必须为正")
    band = box_h / 3.0
    size = size_for_x_height(path, band * ss)
    f = ImageFont.truetype(path, size)
    img = Image.new("RGBA", (box_w * ss, box_h * ss), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    baseline = 2.0 * box_h * ss / 3.0
    margin = int(0.20 * box_h * ss)                 # 左侧留一点白，不贴着格线起笔
    try:
        w = d.textlength(text, font=f)
    except AttributeError:                          # 老版本 Pillow
        w = f.getsize(text)[0]
    pitch = w + 0.6 * band * ss                     # 两遍之间留出写字的间隔
    drawn = 0
    x = margin
    while drawn < n_trace and x + w <= box_w * ss:
        d.text((x, baseline), text, font=f, fill=color + (255,), anchor="ls")
        drawn += 1
        x += pitch
    return img.resize((box_w, box_h), Image.LANCZOS), drawn


def band_geometry(band_px=30, bands=3, cols=8, inset_x=4):
    """四线格几何的**唯一来源** —— 渲染和排版都必须用这一份。

    两边各算一套的话，只要有一个数不一致，图里的带界就和 Excel 画的虚线错开，
    字母会整条带地偏 —— 而且不会报错，只会「看着不太对」。
    """
    band = band_px * bands
    grid = band * cols
    return {"band_px": band_px, "bands": bands, "band": band, "grid_px": grid,
            "box_w": grid - 2 * inset_x, "box_h": band, "inset_x": inset_x}


def render_syllable(text, path, box_w, box_h, color=INK, ss=4):
    """渲染单个音节（教学/示范用，非描红）。同样的基线对齐规则。"""
    band = box_h / 3.0
    size = size_for_x_height(path, band * ss)
    f = ImageFont.truetype(path, size)
    img = Image.new("RGBA", (box_w * ss, box_h * ss), (0, 0, 0, 0))
    ImageDraw.Draw(img).text((box_w * ss / 2.0, 2.0 * box_h * ss / 3.0), text,
                             font=f, fill=color + (255,), anchor="ms")
    return img.resize((box_w, box_h), Image.LANCZOS)


# ---------------------------------------------------------------- 批量渲染 CLI
def _slug(text):
    """音节 -> 文件名：声调字符不进文件名，改用码点，避免文件系统编码问题。"""
    return "_".join(f"{ord(c):04x}" for c in text)


def main(argv):
    import json
    import pinyin_input as P
    arg = {}
    for a in argv:
        if a.startswith("--"):
            k, sep, v = a.partition("=")
            arg[k] = v if sep else "1"

    src = ""
    f = arg.get("--file")
    if f:
        p = f if f.startswith("/") else os.path.join(ROOT, f)
        src = open(p, encoding="utf-8").read()
    else:
        src = arg.get("--items", arg.get("--chars", ""))
    items = [(i, h) for _, lst in P.parse_items(src) for i, h in lst]
    if not items:
        raise SystemExit("✗ 没有解析出音节")

    geo = band_geometry(int(arg.get("--band-px", "30")),
                        int(arg.get("--bands", "3")),
                        int(arg.get("--cols", "8")),
                        int(arg.get("--inset-x", "4")))
    n_trace = int(arg.get("--trace-times", "2"))
    font = find_font(P.needed_chars(items), arg.get("--font"))

    outdir = arg.get("--outdir", "pinyin_images")
    outdir = outdir if outdir.startswith("/") else os.path.join(ROOT, outdir)
    os.makedirs(outdir, exist_ok=True)

    plan = {"font": font, "geometry": geo, "n_trace": n_trace, "items": {}}
    print(f"字体 {os.path.basename(font)}　"
          f"贴图 {geo['box_w']}x{geo['box_h']}px　每音节描红 {n_trace} 遍")
    for it, hz in items:
        img, drawn = render_band(it, font, geo["box_w"], geo["box_h"], n_trace)
        img.save(os.path.join(outdir, _slug(it) + ".png"))
        plan["items"][_slug(it)] = {"item": it, "hanzi": hz, "drawn": drawn,
                                    "box_w": geo["box_w"], "box_h": geo["box_h"]}
        flag = "" if drawn >= n_trace else f"　⚠ 只放得下 {drawn} 遍"
        print(f"  {it}　{hz}　-> 描红 {drawn} 遍 + 留白{flag}")
    with open(os.path.join(outdir, "plan.json"), "w", encoding="utf-8") as fh:
        json.dump(plan, fh, ensure_ascii=False, indent=1)
    print(f"  图片目录 {outdir}（{len(items)} 张 + plan.json）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))