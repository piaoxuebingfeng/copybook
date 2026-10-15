"""拼音音节的输入解析 —— 把用户写的音节表变成可以直接渲染的预组合字符串。

设计要点：
  · 声调一律转成**预组合字符**（`ā` = U+0101），再由字体渲染。
    绝不用「字母 + 独立声调符号叠加」的方式 —— 那样符号位置会漂，而且不同字体漂得不一样。
  · 支持键盘友好的数字标调（`ba1` -> `bā`），省得用户去找 `ǎ` 这种字。
  · v1 **不引入词典**：音节由用户直接给出，因此不存在多音字判断错误的问题。
"""
import re

# 四个声调 + 轻声。ü 的打字替代是 v / V，两个都收。
_TONE_MAP = {
    1: {"a": "ā", "o": "ō", "e": "ē", "i": "ī", "u": "ū", "ü": "ǖ"},
    2: {"a": "á", "o": "ó", "e": "é", "i": "í", "u": "ú", "ü": "ǘ"},
    3: {"a": "ǎ", "o": "ǒ", "e": "ě", "i": "ǐ", "u": "ǔ", "ü": "ǚ"},
    4: {"a": "à", "o": "ò", "e": "è", "i": "ì", "u": "ù", "ü": "ǜ"},
    5: {},          # 轻声不标
}

MARKED = set("āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜ")
VOWELS = "aoeiuü"


def _mark_position(s):
    """声调标在哪个字母上 —— 按《汉语拼音方案》的规则：

    a、o、e 优先（谁在前面标谁）；iu / ui 标在后一个；其余标在唯一的元音上。
    """
    for v in ("a", "o", "e"):
        if v in s:
            return s.index(v)
    iu = s.find("iu")
    if iu >= 0:
        return iu + 1
    ui = s.find("ui")
    if ui >= 0:
        return ui + 1
    for idx, c in enumerate(s):
        if c in VOWELS:
            return idx
    return -1


def to_marked(text):
    """把数字标调写法转成预组合字符：`ba1`->`bā`、`ma5`->`ma`、`nv3`->`nǚ`。

    已经是预组合写法的原样返回（`bā` -> `bā`）。末尾不是 1~5 的数字时也原样返回。
    """
    t = str(text).strip()
    if not t:
        return t
    t = t.replace("v", "ü").replace("V", "ü")     # 键盘替代
    m = re.fullmatch(r"([A-Za-züÜ]+)([1-5])", t)
    if not m:
        return t
    body, tone = m.group(1), int(m.group(2))
    table = _TONE_MAP[tone]
    if not table:
        return body.lower()
    body = body.lower()
    pos = _mark_position(body)
    if pos < 0:
        return body
    ch = body[pos]
    if ch not in table:
        return body
    return body[:pos] + table[ch] + body[pos + 1:]


def parse_items(text, sep=None):
    """解析音节表，返回 [(组名, [(音节, 汉字标注), ...]), ...]。

    格式与汉字字帖的 `字表示例.txt` 一致：
        · `组名: 音节...` 一行一组
        · 没有冒号的行归到「音节」组
        · 音节之间用空白或 `,` `，` `、` `/` 分隔
        · 可以写成 `bā` 或 `ba1`（数字标调）
        · 一个条目后面可跟一个汉字做标注：`bā 八`，也可以是 `ba1 八`
    """
    groups, order = {}, []
    for raw in str(text).splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        gname = "音节"
        if ":" in line or "：" in line:
            head, line = re.split(r"[:：]", line, maxsplit=1)   # 只切一次，得到 2 段
            if head.strip():
                gname = head.strip()
        line = re.sub(r"[,，、/]+", " ", line)
        toks = line.split()
        items, i = [], 0
        while i < len(toks):
            item = to_marked(toks[i])
            hanzi = ""
            # 后面紧跟的单个汉字视为标注（如 `bā 八`）
            if i + 1 < len(toks) and re.fullmatch(r"[\u4e00-\u9fff]", toks[i + 1]):
                hanzi = toks[i + 1]
                i += 1
            if item:
                items.append((item, hanzi))
            i += 1
        if not items:
            continue
        if gname not in groups:
            groups[gname] = []
            order.append(gname)
        groups[gname].extend(items)
    return [(g, groups[g]) for g in order]


def validate(items):
    """挑出明显不是合法音节的条目。

    数字标调写成 `ba9`、或者混进别的符号时，如果不管，它会被**原样渲染成一个
    带数字的音节**贴进字帖 —— 不报错、但内容是错的。所以这里必须拦下来。
    """
    bad = []
    for it, _hz in items:
        if not re.fullmatch(r"[a-zü" + "".join(MARKED) + r"]+", it or ""):
            bad.append(it)
    return bad


def needed_chars(items):
    """这些条目要用哪些字符来校验**拼音字体**覆盖。

    只算音节字母，不算汉字注记：汉字是 Excel 单元格文字（由黑体渲染），
    拼音专用字体（如 汉语拼音.ttf）只有拉丁与声调 —— 把汉字算进来会
    恰好把最合适的字体拒掉。
    """
    s = set()
    for it, _hz in items:
        s |= set(it)
    return s