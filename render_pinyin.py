"""拼音渲染 —— 字体探测 + 单音节单元格渲染。

**这个模块的核心不变量**：字母必须坐在正确的带上。
四线三格的三条带分工是固定的：上带放升部字母和声调符号（`b d h k l`、`ā` 的一横）、
中带放字母主体（x-height）、下带放降部（`g p q y` 的尾巴）。

所以这里**不居中、不按框缩放**，而是：
  1. 反解字号，使 **x-height 恰好等于一条带**（实测 'x' 的墨迹高度，不信字体度量）
  2. 基线钉在框高的 **2/3** 处（anchor="ls" + 手动对位）

于是中带 = [1/3, 2/3]，正好是字母主体；升部与降部自然伸进上下带。

两个实战教训（都是踩出来的，写在这里防止回退）：

* **cmap 必须合并全部子表再查。** 专用拼音字体（如 汉语拼音.ttf）常把拉丁字母
  放在 (1,0) Mac Roman 子表里，`getBestCmap()` 只挑一个「最佳」子表会误报整页缺字，
  于是这个字体被自己的覆盖检查拒掉 —— 恰恰是最该用的那个。
  TTC 则必须只看第 0 个字体：PIL 渲染固定用 index=0，「挑覆盖最好的那个」
  会造成「检查通过了、渲染出来却是豆腐块」。
* **楷体的 `g` 是双层印刷体。** 教材里孩子学的是单层手写体（`ɡ`）。
  渲染前要实测字形（数墨迹里的封闭空腔：双层 g 有 2 个，单层只有 1 个），
  双层且字体提供单层码点（U+0261）时自动替换。`a` 同理（U+0251）。
  汉语拼音.ttf 这类专用字体本身就是单层，替换自然不发生。

字母间距**不做任何手工调整**：simkai 里带调元音的边距忽宽忽窄（实测 `shàng`
的间隙是 8/21/23/6px），而拼音专用字体的 advance 天然紧凑均匀（5~7px），
和参考工程逐像素一致 —— 间距问题的正解是「用为拼音设计的字体」，不是自己排。
"""
import os
import sys
from collections import deque

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.abspath(__file__))

# 首列示范色（近黑，与汉字教学条的墨色一致）与描红浅灰（参考工程实测色 #CECECE）
INK = (34, 40, 49)
TRACE = (206, 206, 206)

# 候选字体：按「教学合适度 + 可分发性」排序。
# 楷体最适合儿童识字教学，但它是微软/方正私有字体 —— 本机装了就用，**绝不随仓库分发**。
# fonts/ 里的「拼音」字体会排到最前：它们是专为四线三格设计的（单层 a/g、紧凑 advance）。
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
    """展开候选字体列表（glob 优先；名字里带「拼音/pinyin」的排最前）。"""
    global _RESOLVED
    if _RESOLVED is None:
        import glob
        out = []
        for pat in FONT_CANDIDATES:
            if any(ch in pat for ch in "*?["):
                hits = sorted(glob.glob(os.path.join(ROOT, pat)))
                # 拼音专用字体优先：专为四线三格设计（单层字形 + 紧凑间距）
                hits.sort(key=lambda p: 0 if ("拼音" in os.path.basename(p)
                                              or "pinyin" in os.path.basename(p).lower()) else 1)
                out.extend(hits)
            else:
                out.append(pat)
        _RESOLVED = [p for p in out if os.path.exists(p)]
    return _RESOLVED


# ---------------------------------------------------------------- 字体覆盖校验
def _merged_charmap(font):
    """合并一个字体的**全部** cmap 子表 —— 见模块 docstring 的教训。"""
    chars = set()
    for table in font["cmap"].tables:
        chars.update(table.cmap.keys())
    return chars


def missing_chars(path, chars):
    """返回字体缺哪些字符。

    有 fontTools 就用 cmap 精确判断（**合并所有子表**；TTC 只看第 0 个字体，
    因为 PIL 渲染固定用 index=0）；没有就退回「渲染出来是空白就当作缺字」的
    粗略判断，并明确告知不可靠。
    """
    chars = {c for c in chars if not c.isspace()}
    if not chars:
        return set()
    try:
        from fontTools.ttLib import TTFont, TTCollection
        if path.lower().endswith(".ttc"):
            face = TTCollection(path).fonts[0]        # PIL 用的就是第 0 个
        else:
            face = TTFont(path, fontNumber=0, lazy=True)
        have = _merged_charmap(face)
        return {c for c in chars if ord(c) not in have}
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
    msg.append("  推荐（可自由分发）：AR PL UKai / Noto Serif CJK / LXGW WenKai；")
    msg.append("  专为拼音设计的（单层 a/g，间距最匀）参考 Copybook 工程的 汉语拼音.ttf。")
    raise SystemExit("\n".join(msg))


# ---------------------------------------------------------------- 字形归一化（单层 a/g）
_HOLES = {}


def glyph_holes(path, ch):
    """数一个字形里被墨迹完全封闭的空腔数。

    双层 `g` 有 2 个空腔（上下两个圈），单层 `ɡ` 只有 1 个；`a` 同理。
    这是判断「印刷体还是教材手写体」的可靠办法 —— 量像素，不猜字体名。
    """
    key = (path, ch)
    if key in _HOLES:
        return _HOLES[key]
    size = 120
    f = ImageFont.truetype(path, size)
    img = Image.new("L", (size * 2, size * 2), 0)
    ImageDraw.Draw(img).text((size // 2, size), ch, font=f, fill=255, anchor="ls")
    import numpy as np
    a = np.asarray(img) > 60
    rows = np.flatnonzero(a.any(axis=1))
    cols = np.flatnonzero(a.any(axis=0))
    n = 0
    if rows.size and cols.size:
        box = a[rows.min():rows.max() + 1, cols.min():cols.max() + 1]
        H, W = box.shape
        seen = np.zeros_like(box, bool)
        q = deque()
        for x in range(W):
            for y in (0, H - 1):
                if not box[y, x] and not seen[y, x]:
                    seen[y, x] = True
                    q.append((y, x))
        for y in range(H):
            for x in (0, W - 1):
                if not box[y, x] and not seen[y, x]:
                    seen[y, x] = True
                    q.append((y, x))
        while q:
            y, x = q.popleft()
            for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                ny, nx = y + dy, x + dx
                if 0 <= ny < H and 0 <= nx < W and not box[ny, nx] and not seen[ny, nx]:
                    seen[ny, nx] = True
                    q.append((ny, nx))
        cavity = (~box) & (~seen)
        vis = np.zeros_like(cavity, bool)
        for yy in range(H):
            for xx in range(W):
                if cavity[yy, xx] and not vis[yy, xx]:
                    n += 1
                    dq = deque([(yy, xx)])
                    vis[yy, xx] = True
                    while dq:
                        y, x = dq.popleft()
                        for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                            ny, nx = y + dy, x + dx
                            if 0 <= ny < H and 0 <= nx < W and cavity[ny, nx] and not vis[ny, nx]:
                                vis[ny, nx] = True
                                dq.append((ny, nx))
    _HOLES[key] = n
    return n


# 单层替换码点：拉丁扩展 IPA 里的教学手写形
SINGLE = {"a": "ɑ", "g": "ɡ"}     # U+0251 / U+0261


def normalize_text(text, path):
    """把双层的 a/g 换成单层教学形（字体提供时）。

    汉语拼音.ttf 本身就是单层，替换不会发生；楷体/宋体的 g 是双层 ——
    不换的话，孩子在四线格里练的是印刷体写法。
    """
    out = []
    for ch in text:
        repl = SINGLE.get(ch)
        if repl and glyph_holes(path, ch) >= 2 and not missing_chars(path, repl):
            ch = repl
        out.append(ch)
    return "".join(out)


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


# ---------------------------------------------------------------- 单元格几何（唯一来源）
def cell_plan(items, path, band_px=30, bands=3, grid_px=720,
              inset_x=4, pad=8, min_cell=90):
    """计算「一方格一个音节」的整张表几何 —— 渲染和排版**必须**用这一份。

    格宽按**最长音节**自适应：先按 x-height = 一条带 反解字号、量出每个音节的
    墨迹宽度，格宽 = max(最小格宽, 最长墨迹 + 2*pad)，列数 = floor(总宽/格宽)，
    再把格宽放大到恰好铺满总宽（floor 保证不超）。

    于是 `zhuāng`（6 字母）也能放进格子：格变宽、列数变少，但字号**不缩** ——
    缩字号字母就不在带上了，这是四线三格的硬约束。
    """
    worst = 0
    for it in items:
        w = syllable_ink_size(it, path, band_px)[0]
        worst = max(worst, w)
    need = max(min_cell, worst + 2 * pad)
    cols = max(1, grid_px // need)
    cell_px = grid_px // cols
    return {"band_px": band_px, "bands": bands, "band": band_px * bands,
            "cell_px": cell_px, "cols": cols, "grid_px": cell_px * cols,
            "box_w": cell_px - 2 * inset_x, "box_h": band_px * bands,
            "inset_x": inset_x, "ink_max": worst}


def syllable_ink_size(text, path, band_px, ss=4):
    """按「x-height = band_px」的字号渲染 text，返回墨迹 (宽, 高)（最终尺度）。

    用来算格宽 —— 量的是**墨迹**不是 advance：声调符号会超出 em 框，
    只看 advance 会把调号裁掉。
    """
    size = size_for_x_height(path, band_px * ss)
    f = ImageFont.truetype(path, size)
    W = max(int(size * (len(text) + 2)), size)
    img = Image.new("L", (W, band_px * ss * 3), 0)
    d = ImageDraw.Draw(img)
    d.text((0, band_px * ss * 2), text, font=f, fill=255, anchor="ls")
    b = img.getbbox()
    if b is None:
        raise SystemExit(f"✗ 渲染 {text!r} 得到空白（缺字？）")
    return (b[2] - b[0] + ss - 1) // ss, (b[3] - b[1] + ss - 1) // ss   # 向上取整


# ---------------------------------------------------------------- 音节渲染
def render_cell(text, path, box_w, box_h, color=TRACE, ss=4):
    """渲染**一个音节**到一张方格贴图（RGBA）。

    box_w x box_h 是**最终贴进 Excel 的像素尺寸**，内部按 ss 倍超采样再缩回来
    （与汉字描红图同一套做法）。三带的边界固定在 1/3、2/3 处，所以只要贴图时
    **纵向不内缩**，图里的带界就和 Excel 画的虚线严格重合。

    水平方向按**墨迹**居中（不是 advance 居中）：调号和降部会超出 em 框，
    advance 居中的音节在格子里看着会偏。返回 None 表示一个字都放不下。
    """
    if box_w <= 0 or box_h <= 0:
        raise ValueError("box 尺寸必须为正")
    band = box_h / 3.0
    size = size_for_x_height(path, band * ss)
    f = ImageFont.truetype(path, size)
    W, H = box_w * ss, box_h * ss
    baseline = 2.0 * H / 3.0
    scratch = Image.new("L", (W * 2, H), 0)
    d0 = ImageDraw.Draw(scratch)
    d0.text((0, baseline), text, font=f, fill=255, anchor="ls")
    b = scratch.getbbox()
    if b is None:
        return None
    ink_w = b[2] - b[0]
    if ink_w > W:
        return None                             # 墨迹比格子还宽：调用方该把格加宽
    x0 = (W - ink_w) / 2.0 - b[0]               # 墨迹居中
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.text((x0, baseline), text, font=f, fill=color + (255,), anchor="ls")
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

    # 先定字体再算几何：格宽来自「这个字体下的最长音节墨迹」
    font = find_font(P.needed_chars(items), arg.get("--font"))
    geo = cell_plan([i for i, _ in items], font,
                    int(arg.get("--band-px", "30")),
                    int(arg.get("--bands", "3")),
                    int(arg.get("--grid-px", "720")),
                    int(arg.get("--inset-x", "4")))

    outdir = arg.get("--outdir", "pinyin_images")
    outdir = outdir if outdir.startswith("/") else os.path.join(ROOT, outdir)
    os.makedirs(outdir, exist_ok=True)

    plan = {"font": font, "geometry": geo, "colors": {"model": INK, "trace": TRACE},
            "items": {}}
    print(f"字体 {os.path.basename(font)}　格 {geo['cell_px']}px x {geo['box_h']}px"
          f"　每行 {geo['cols']} 格　最长音节墨迹 {geo['ink_max']}px")
    for it, hz in items:
        rt = normalize_text(it, font)          # 双层 g -> 单层 ɡ（字体需要时）
        model = render_cell(rt, font, geo["box_w"], geo["box_h"], INK)
        trace = render_cell(rt, font, geo["box_w"], geo["box_h"], TRACE)
        if model is None:
            print(f"  ⚠ {it}　{hz}　-> 放不下（墨迹超过格宽），已跳过")
            continue
        model.save(os.path.join(outdir, _slug(it) + ".png"))
        trace.save(os.path.join(outdir, _slug(it) + "_t.png"))
        plan["items"][_slug(it)] = {"item": it, "render": rt, "hanzi": hz,
                                    "box_w": geo["box_w"], "box_h": geo["box_h"]}
        print(f"  {it}　{hz}　-> 首格示范 + 描红格")
    with open(os.path.join(outdir, "plan.json"), "w", encoding="utf-8") as fh:
        json.dump(plan, fh, ensure_ascii=False, indent=1)
    print(f"  图片目录 {outdir}（{len(plan['items'])} 音节 x 示范/描红 + plan.json）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
