"""汉字 → 拼音：离线词典的最大匹配消歧（第三种输入方式）。

用法上的位置：用户可以直接写汉字原文，本模块把它展开成「音节 汉字」的
标准音节表写法，后面的解析/渲染/排版/验收完全不感知这个功能的存在：

    我爱北京天安门   ->   wǒ 我 ài 爱 běi 北 jīng 京 tiān 天 ān 安 mén 门

消歧规则（正着贪心，取词典里能命中的最长词）：
    长大 -> zhǎng dà（词命中）     长城 -> cháng chéng（词命中）
    单字兜底取词典里该字最常见的读音（词典数据已按词频选好）。

数据：data/pinyin-dict.json.gz（34.2 万条，≤4 字词），从 pinyin-pro 的
complete.json 裁剪而来（MIT，github.com/zh-lx/pinyin-pro）——裁剪规则见
docs/PINYIN.md。声调 Sandhi（一下 yí、不是 bú）和轻声（头发 tóu fa）
都在词表里，命中词即得正确变调。

已知边界（诚实记录，不装没有）：
  · 贪心最大匹配不等于概率最优分词：「研究生命」这类两可串可能切错，
    影响的是个别多音字的读音；自己给音节（ba1 八）永远是精确路径
  · 词典没有的字（生僻字、嗯/呣 这类读 ńg 的叹词）会明确报错，不猜
"""
import gzip
import json
import os
import re

ROOT = os.path.dirname(os.path.abspath(__file__))
DICT_PATH = os.path.join(ROOT, "data", "pinyin-dict.json.gz")
MAXLEN = 4                      # 词典里最长 4 字词，匹配窗口同宽
_CJK_RUN = re.compile(r"[\u4e00-\u9fff]+")
_ASCII = re.compile(r"[A-Za-z0-9]")

_DICT = None


def _load():
    global _DICT
    if _DICT is None:
        try:
            with gzip.open(DICT_PATH, "rt", encoding="utf-8") as fh:
                _DICT = json.load(fh)
        except FileNotFoundError:
            raise SystemExit(
                "✗ 汉字转拼音需要词典 data/pinyin-dict.json.gz，但没有找到。\n"
                "  该词典取自 pinyin-pro（MIT）：github.com/zh-lx/pinyin-pro\n"
                "  的 packages/data/json/complete.json（≤4 字词、34 万条）。\n"
                "  没有词典也能用拼音字帖：直接给音节（bā 八 / ba1 八），\n"
                "  或用 拼音音节总表.txt 当底稿。")
    return _DICT


def convert(text):
    """汉字串 -> ([(音节, 汉字), ...], [查不到的字, ...])。

    非汉字字符（标点/空格/数字）当作分隔：断开词界、不产生行。
    """
    d = _load()
    out, errors = [], []
    for run in _CJK_RUN.findall(text):
        i = 0
        while i < len(run):
            hit = None
            for n in range(min(MAXLEN, len(run) - i), 1, -1):
                w = run[i:i + n]
                py = d.get(w)
                if py and len(py.split()) == n:
                    hit = (w, py.split())
                    break
            if hit:
                w, pys = hit
                out.extend(zip(pys, w))
                i += len(w)
            else:
                ch = run[i]
                py = d.get(ch)
                if py:
                    out.append((py, ch))
                else:
                    errors.append(ch)
                i += 1
    return out, errors


def _line_is_hanzi(content):
    """一行的音节部分是「汉字原文」：有汉字、且没有任何 ASCII 字母/数字。

    汉字标注（bā 八 的 八）不会误判 —— 那种行的音节部分有 ASCII 字母。
    """
    if _ASCII.search(content):
        return False
    return bool(_CJK_RUN.search(content))


def expand_spec(spec):
    """把音节表里的汉字原文行展开成「音节 汉字」写法。

    返回 (新音节表, [(组名, [(音节, 汉字), ...]), ...], [查不到的字, ...])。
    只动汉字行，普通音节行原样保留 —— 两种写法可以在同一个文件里混用。
    """
    out_lines, converted, errors = [], [], []
    for raw in str(spec).splitlines():
        line = raw.strip()
        m = re.split(r"[:：]", line, maxsplit=1) if re.search(r"[:：]", line) else None
        gname = m[0].strip() if m else ""
        content = (m[1] if m else line)
        if line and not line.startswith("#") and _line_is_hanzi(content):
            pairs, errs = convert(content)
            errors.extend(errs)
            converted.append((gname or "音节", pairs))
            toks = " ".join(f"{s} {h}" for s, h in pairs)
            out_lines.append(f"{gname}: {toks}" if gname else toks)
        else:
            out_lines.append(raw)
    return "\n".join(out_lines), converted, errors


if __name__ == "__main__":
    import sys
    for t in sys.argv[1:]:
        pairs, errs = convert(t)
        print(f"{t} -> {' '.join(f'{s}({h})' for s, h in pairs)}"
              + (f"  [查不到: {errs}]" if errs else ""))
