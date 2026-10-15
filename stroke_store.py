# -*- coding: utf-8 -*-
"""汉字笔画数据仓库。

数据来源：hanzi-writer-data 2.0.1（源自 Make Me a Hanzi），共 9575 个字。
- 首次运行 `setup_data()` 会把整个字库下载并解压到 data/hanzi-writer-data/
- 之后**完全离线**可用（一年级生字 44/44、常用字全部覆盖）
- 若某个字不在本地字库里，自动从 jsDelivr CDN 取一次并缓存到 data/cache/

用法：
    from stroke_store import load_many, parse_chars, setup_data
    data = load_many("天地人")        # -> {"天": {...}, "地": {...}, ...}
"""
import json
import os
import shutil
import sys
import tarfile
import urllib.parse
import urllib.request

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "data", "hanzi-writer-data")
CACHE_DIR = os.path.join(ROOT, "data", "cache")
PKG_VERSION = "2.0.1"
TARBALL_NAME = f"hanzi-writer-data-{PKG_VERSION}.tgz"
TARBALL_URL = (f"https://registry.npmjs.org/hanzi-writer-data/-/"
               f"hanzi-writer-data-{PKG_VERSION}.tgz")
TARBALL_LOCAL = os.path.join(ROOT, "data", TARBALL_NAME)
CDN_URL = "https://cdn.jsdelivr.net/npm/hanzi-writer-data@{v}/{c}.json"


# ---------------------------------------------------------------- 字符表解析
def parse_chars(spec):
    """把命令行/文本文件里的字符整理成列表。

    支持：直接给一串字（"天地人"）、逗号/空格分隔（"天,地,人"）、或一个 .txt 文件路径
    （文件里可以换行、可以用 # 注释）。
    """
    if not spec:
        return []
    if os.path.exists(spec):
        text = open(spec, encoding="utf-8").read()
        out = []
        for line in text.splitlines():
            line = line.split("#", 1)[0]
            out.extend(_clean(line))
        return out
    return _clean(spec)


def _clean(s):
    """去掉空白、逗号、标点等非汉字内容，保留字序且不重复。"""
    out, seen = [], set()
    for c in s:
        if c.isspace() or c in ",，、;；|/\\-_+":
            continue
        if not _is_han(c):
            continue
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _is_han(c):
    o = ord(c)
    return (0x4E00 <= o <= 0x9FFF or      # CJK 基本区
            0x3400 <= o <= 0x4DBF or      # 扩展 A
            0xF900 <= o <= 0xFAFF)        # 兼容汉字


# ---------------------------------------------------------------- 数据准备
def has_local_data():
    return os.path.isdir(DATA_DIR) and any(f.endswith(".json") for f in os.listdir(DATA_DIR))


def _extract(tgz, quiet=False):
    """把压缩包里的字库解压到 DATA_DIR（文件名 URL 解码成纯汉字）。"""
    if not quiet:
        print("正在解压字库…")
    tmpdir = os.path.join(ROOT, "data", "_extract")
    shutil.rmtree(tmpdir, ignore_errors=True)
    os.makedirs(tmpdir, exist_ok=True)
    with tarfile.open(tgz, "r:gz") as tf:
        tf.extractall(tmpdir)
    src = os.path.join(tmpdir, "package")
    os.makedirs(DATA_DIR, exist_ok=True)
    n = 0
    for f in os.listdir(src):
        if not f.endswith(".json"):
            continue
        ch = urllib.parse.unquote(f[:-5])          # 文件名是 URL 编码的
        shutil.move(os.path.join(src, f), os.path.join(DATA_DIR, ch + ".json"))
        n += 1
    shutil.rmtree(tmpdir, ignore_errors=True)
    if not quiet:
        print(f"完成：{n} 个字 -> {DATA_DIR}")
    return n


def setup_data(quiet=False):
    """下载并解压字库（只需一次）。已存在则跳过。"""
    if has_local_data():
        if not quiet:
            print(f"本地字库已存在：{DATA_DIR}")
        return True
    os.makedirs(os.path.dirname(DATA_DIR), exist_ok=True)
    tgz = TARBALL_LOCAL
    if not os.path.exists(tgz):
        if not quiet:
            print(f"下载字库（约 13MB）: {TARBALL_URL}")
        tmp = tgz + ".part"
        urllib.request.urlretrieve(TARBALL_URL, tmp)
        os.replace(tmp, tgz)
    _extract(tgz, quiet=quiet)
    return True


def ensure_data(quiet=False):
    """保证字库可用：已解压 -> 直接用；只有压缩包 -> 自动解压；都没有 -> 联网下载。

    这样打包带走时只需要带一个 13MB 的压缩包，首次运行自动展开成 9575 个 json。
    """
    if has_local_data():
        return True
    if os.path.exists(TARBALL_LOCAL):
        _extract(TARBALL_LOCAL, quiet=quiet)
        return True
    try:
        setup_data(quiet=quiet)
        return True
    except Exception as e:
        print(f"字库不可用（{e}）")
        print("请联网运行一次： python3 stroke_store.py --setup")
        return False


def _fetch_remote(ch):
    """本地没有时从 CDN 取一次，并写入缓存。"""
    os.makedirs(CACHE_DIR, exist_ok=True)
    url = CDN_URL.format(v=PKG_VERSION, c=urllib.parse.quote(ch))
    try:
        with urllib.request.urlopen(url, timeout=20) as r:
            raw = r.read().decode("utf-8")
        obj = json.loads(raw)
        with open(os.path.join(CACHE_DIR, ch + ".json"), "w", encoding="utf-8") as fh:
            fh.write(raw)
        return obj
    except Exception:
        return None


# ---------------------------------------------------------------- 读取
_cache = {}
_missing = set()


def load_char(ch):
    """取一个字的 {strokes, medians}；取不到返回 None。"""
    if ch in _cache:
        return _cache[ch]
    if ch in _missing:
        return None
    for path in (os.path.join(DATA_DIR, ch + ".json"),
                 os.path.join(CACHE_DIR, ch + ".json")):
        if os.path.exists(path):
            try:
                obj = json.load(open(path, encoding="utf-8"))
                _cache[ch] = obj
                return obj
            except Exception:
                pass
    obj = _fetch_remote(ch)
    if obj is None:
        _missing.add(ch)
        return None
    _cache[ch] = obj
    return obj


def load_many(chars):
    """批量取，返回 {字: 数据}，缺数据的字不出现在结果里。"""
    out = {}
    for c in chars:
        d = load_char(c)
        if d and d.get("strokes"):
            out[c] = d
    return out


def check(chars):
    """返回 (有数据的字, 缺数据的字)，不触发网络请求。"""
    have, miss = [], []
    for c in chars:
        p = os.path.join(DATA_DIR, c + ".json")
        q = os.path.join(CACHE_DIR, c + ".json")
        (have if (os.path.exists(p) or os.path.exists(q)) else miss).append(c)
    return have, miss


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else "--status"
    if arg in ("--setup", "--fetch-data"):
        setup_data()
    elif arg == "--status":
        if has_local_data():
            n = len([f for f in os.listdir(DATA_DIR) if f.endswith(".json")])
            print(f"本地字库已就绪：{n} 个字  ({DATA_DIR})")
        elif os.path.exists(TARBALL_LOCAL):
            print(f"字库压缩包已就位（{TARBALL_NAME}），首次生成字帖时会自动解压，无需操作。")
        else:
            print("本地字库未安装。联网运行一次： python3 stroke_store.py --setup")
    else:
        ensure_data()                      # 只有压缩包时先自动解压
        chars = parse_chars(arg)
        have, miss = check(chars)
        print(f"共 {len(chars)} 字：有数据 {len(have)}，缺数据 {len(miss)}"
              + (f" -> {''.join(miss)}" if miss else ""))