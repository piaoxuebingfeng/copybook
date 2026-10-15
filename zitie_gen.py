#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""字帖生成器 —— 一句话生成任意汉字的 A4 笔顺描红字帖。

用法
----
    # 1) 直接给字（最常用）
    python3 zitie_gen.py 天地花鸟虫

    # 2) 从文本文件读字表（# 开头是注释，可用「组名: 字表」分组）
    python3 zitie_gen.py --file=字表.txt

    # 3) 常用选项
    python3 zitie_gen.py 天地人 --name=一年级上 --pdf
    python3 zitie_gen.py 天地人 --mode=lo          # 用 LibreOffice/WPS 校对列宽
    python3 zitie_gen.py --setup                   # 首次安装字库（约 13MB，只需一次）

输出
----
    output/<名字>/<名字>.xlsx     可编辑、可直接打印
    output/<名字>/<名字>.pdf      用 --pdf 时生成（打印版，几何已校准）
    output/<名字>/strips/         笔顺教学条图片（中间产物，可复用于改写）

字表文件格式
------------
    # 可以是纯字表
    天地人你我他

    # 也可以用「组名: 字表」分组（当前版本会把所有字依次排入，组名暂只作备注）
    基础数字: 一二三十
    自然事物: 日月水火山石
"""
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable


HELP = """
汉字笔顺字帖生成器 —— 给一串汉字，生成可打印的笔顺描红字帖（Excel + PDF）

用法：
  python3 zitie_gen.py 天地花鸟虫 [选项]
  python3 zitie_gen.py --file=字表示例.txt [选项]

要生成的字（二选一，也可以都用，会合并）：
  字串                     直接给，例如：天地花鸟虫
  --file=文件              从文本读，支持「组名: 字表」格式（见 字表示例.txt）

版式（核心选项）：
  --layout=a5              一张 A4 放 **2 个字**，各占半页（一张 A5），
                           正中一条裁切线，剪开即两张 A5。**默认**
  --layout=a4              一张 A4 放 **4 个字**（原来的紧凑版式）
  也可以写 --layout=2 / --layout=4

纸张方向（只对 --layout=a5 有效）：
  --orient=portrait        A4 **纵向**，上下两张 A5（每张 21 x 14.85cm）。
                           每行 8 格。**默认**
  --orient=landscape       A4 **横向**，左右两张 A5（每张 14.85 x 21cm）。
                           每半页只有 14.85cm 宽，所以每行 5 格（格子仍是 2.4cm）
  也可以写 --orient=纵向 / --orient=横向 / --orient=v / --orient=h
  --cell-cm=2.4            田字格边长（cm）。a5 默认 2.4、a4 默认 2.0。
                           实际会取整到整数像素，保证格子是精确正方形

外观：
  --title=标题             字帖标题（默认「幼小衔接 · 汉字笔顺字帖」）
  --subtitle=文字          标题下第 2 行说明（默认「看图学笔顺：…」）
  --thin=N                 笔画减细量 0~20，默认 10（越大笔画越细）
  --strip-max-cm=16        笔顺教学条单行宽度上限，默认 16（= 田字格总宽）

笔顺练习模式：
  --stroke-order           打开后，每个字的所有田字格都用来练笔顺：
                           第 1 格写第 1 笔、第 2 格写前 2 笔……第 N 格写完整字，
                           如此练 3 遍，再重复 3 个整字，剩下的格子留白给学生自己写。
                           格子不够时自动降级（先减整字、再减遍数），
                           笔画比格子还多时把一遍笔顺均匀压缩进格子，顺序一笔不漏。
                           默认关闭（关闭时是原来的排法：每行前 3 格整字描红，其余留白）

输出：
  --name=名字              输出目录 / 文件名（默认由字表自动生成）
  --pdf                    同时导出 PDF（默认：装了 LibreOffice 就导出）
  --no-pdf                 只出 xlsx，不导 PDF
  --mode=excel             列宽按 Excel / WPS 校准（默认）
  --mode=lo                列宽按 LibreOffice 校准（导 PDF 时内部会用）
  --outdir=目录            输出根目录，默认 ./output

其他：
  --setup                  下载字库后退出（只在首次或换设备时需要）
  -h, --help               显示本帮助

例子：
  python3 zitie_gen.py 天地花鸟虫                 # 2 字/页，A5 版式（默认）
  python3 zitie_gen.py 天地花鸟虫 --layout=a4     # 4 字/页，紧凑版式
  python3 zitie_gen.py 天地花鸟虫 --cell-cm=3     # 格子放大到 3cm
  python3 zitie_gen.py --file=字表示例.txt --name=我的字帖 --pdf
"""


def die(msg):
    print(f"错误：{msg}")
    sys.exit(1)


def parse_args(argv):
    o = {
        "spec": "", "file": "", "name": "", "title": "", "subtitle": "",
        "mode": "excel", "pdf": None, "thin": 10, "setup": False,
        "strip_max_cm": 16,
        "layout": "a5",          # a5 = 一张 A4 放 2 个字；a4 = 放 4 个字
        "orient": "portrait",    # portrait=纵向(上下两个A5)；landscape=横向(左右两个A5)
        "cell_cm": "",           # 田字格边长，留空则 a5=2.4cm / a4=2.0cm
        "outdir": os.path.join(ROOT, "output"),
        "stroke_order": False,   # 田字格里逐格练笔顺
    }
    if any(a in ("-h", "--help", "help") for a in argv):
        print(HELP)
        sys.exit(0)
    for a in argv:
        if a.startswith("--"):
            k, _, v = a[2:].partition("=")
            k = k.replace("-", "_")
            if k in ("pdf", "setup", "stroke_order"):
                o[k] = True
            elif k == "no_pdf":
                o["pdf"] = False
            elif k in o:
                o[k] = v
            elif k in ("thin",):
                o[k] = v
            else:
                die(f"未知参数 --{k}")
        else:
            o["spec"] = (o["spec"] + a).strip()
    return o


def _strip_cells(o):
    """这一版式一行最多几个田字格 —— 教学条宽度上限就按它渲染。
    （不能用 --strip-max-cm：那个参数假设"一格 = 2cm"，格子不是 2cm 时会算错。）"""
    return 5 if (o["layout"] == "a5" and o["orient"] == "landscape") else 8


def main():
    o = parse_args(sys.argv[1:])

    import stroke_store
    from stroke_store import parse_chars

    if o["setup"]:
        stroke_store.setup_data()
    elif not stroke_store.ensure_data():
        die("字库不可用，请联网运行一次： python3 stroke_store.py --setup")

    spec = o["spec"]
    if o["file"]:
        if not os.path.exists(o["file"]):
            die(f"找不到字表文件 {o['file']}")
        raw = open(o["file"], encoding="utf-8").read()
        # 支持「组名: 字表」格式
        joined = "".join(line.split("：", 1)[-1].split(":", 1)[-1] if ("：" in line or ":" in line)
                         else line
                         for line in raw.splitlines() if not line.strip().startswith("#"))
        spec = spec + joined
    if not spec.strip():
        die("没有指定要生成的字。用法：python3 zitie_gen.py 天地花鸟虫")

    chars = parse_chars(spec)
    if not chars:
        die("解析不出汉字，请检查输入")

    have, miss = stroke_store.check(chars)
    if miss:
        print(f"⚠ 字库里没有这些字，将被跳过：{''.join(miss)}")
    if not have:
        die("没有任何字可用")
    chars = have
    print(f"共 {len(chars)} 个字：{''.join(chars)}")

    name = o["name"] or ("字帖_" + "".join(chars[:8]) + ("…" if len(chars) > 8 else ""))
    name = "".join(c for c in name if c not in '/\\:*?"<>|')
    work = os.path.join(o["outdir"], name)
    strips = os.path.join(work, "strips")
    traces = os.path.join(work, "traces")
    meta = os.path.join(work, "meta.json")
    xlsx = os.path.join(work, name + ".xlsx")
    os.makedirs(work, exist_ok=True)

    cspec = "".join(chars)
    title = o["title"] or "幼小衔接 · 汉字笔顺字帖"
    try:
        o["thin"] = int(o["thin"])
    except (TypeError, ValueError):
        die("--thin 只能是数字，例如 --thin=8")
    if not 0 <= o["thin"] <= 20:
        die("--thin 只能是 0~20")
    if o["thin"] > 10:
        print(f"⚠ --thin={o['thin']} 超过实测安全上限 10，"
              f"「右 / 耳 / 间」这类字的细横可能被压断（建议先用 10 试）")

    _ORIENT_ALIAS = {"portrait": "portrait", "p": "portrait", "v": "portrait",
                     "纵向": "portrait", "竖": "portrait",
                     "landscape": "landscape", "l": "landscape", "h": "landscape",
                     "横向": "landscape", "横": "landscape"}
    if o["orient"] not in _ORIENT_ALIAS:
        die("--orient 只能是 portrait(纵向) 或 landscape(横向)")
    o["orient"] = _ORIENT_ALIAS[o["orient"]]

    _LAYOUT_ALIAS = {"a5": "a5", "2": "a5", "a4": "a4", "4": "a4"}
    if o["layout"] not in _LAYOUT_ALIAS:
        die("--layout 只能是 a5 / a4（也可以写 2 / 4）")
    o["layout"] = _LAYOUT_ALIAS[o["layout"]]
    if o["orient"] == "landscape" and o["layout"] != "a5":
        die("--orient=横向 只适用于 --layout=a5（一张 A4 放 2 个字）")
    cell_opt = [f"--cell-cm={o['cell_cm']}"] if o["cell_cm"] else []
    # 笔顺练习模式：渲染阶段图（前 k 笔），并让拼装脚本按序列铺满田字格
    so_opt = ["--stroke-order"] if o["stroke_order"] else []
    stages_opt = ["--stages"] if o["stroke_order"] else []

    def run(step, args):
        print(f"  · {step}")
        r = subprocess.run([PY] + args, cwd=ROOT, capture_output=True, text=True)
        if r.returncode != 0:
            print(r.stdout[-2000:])
            print(r.stderr[-2000:])
            die(f"{step} 失败")
        return r.stdout

    print("生成中…")
    _cm = o["cell_cm"] or ("2.4" if o["layout"] == "a5" else "2.0")
    if o["layout"] == "a5":
        _d = "横向（左右两张 A5，每半页 14.85 x 21cm，5 格/行）" if o["orient"] == "landscape" \
             else "纵向（上下两张 A5，每半页 21 x 14.85cm，8 格/行）"
        print(f"  版式：一张 A4 放 2 个字，{_d}（田字格 {_cm}cm）")
    else:
        print(f"  版式：一张 A4 放 4 个字（田字格 {_cm}cm）")
    if o["stroke_order"]:
        print("  笔顺练习模式：每个字的田字格都按笔顺逐格递进（--stroke-order）")
    run("渲染笔顺教学条", ["render_strokes.py", f"--chars={cspec}",
                        f"--out={strips}", f"--meta={meta}", f"--thin={o['thin']}",
                        f"--strip-max-cells={_strip_cells(o)}"])
    run("渲染田字格描红", ["render_trace.py", f"--chars={cspec}", f"--out={traces}",
                         *stages_opt])
    out = run("拼装字帖", ["generate_zitie.py", f"--chars={cspec}",
                         f"--strokes={strips}", f"--meta={meta}", f"--traces={traces}",
                         "--trace-inset=0", f"--mode={o['mode']}",
                         f"--layout={o['layout']}", *cell_opt, *so_opt,
                         f"--title={title}", f"--out={xlsx}",
                         f"--orient={o['orient']}"])
    for line in out.strip().splitlines():
        print("    " + line)

    made = [xlsx]
    want_pdf = o["pdf"] if o["pdf"] is not None else shutil.which("soffice") is not None
    if want_pdf:
        if not shutil.which("soffice"):
            print("  ⚠ 未找到 soffice，跳过 PDF（Excel/WPS 里直接打印 xlsx 即可）")
        else:
            # PDF 要用 LibreOffice 的列宽模型重建一份，才能保证几何精确
            tmp_x = os.path.join(work, "_lo_tmp.xlsx")
            run("按 LibreOffice 列宽重建", ["generate_zitie.py", f"--chars={cspec}",
                                        f"--strokes={strips}", f"--meta={meta}",
                                        f"--traces={traces}", "--trace-inset=0",
                                        "--mode=lo", f"--layout={o['layout']}", *cell_opt,
                                        *so_opt,
                                        f"--title={title}", f"--out={tmp_x}",
                                        f"--orient={o['orient']}"])
            pdfdir = os.path.join(work, "_pdf")
            os.makedirs(pdfdir, exist_ok=True)
            subprocess.run(["soffice", "--headless",
                            f"-env:UserInstallation=file://{os.path.join(tempfile.gettempdir(), 'zitie_lo')}",
                            "--convert-to", "pdf", "--outdir", pdfdir, tmp_x],
                           capture_output=True, text=True, timeout=300)
            pdf = os.path.join(work, name + ".pdf")
            src = os.path.join(pdfdir, "_lo_tmp.pdf")
            if os.path.exists(src):
                shutil.move(src, pdf)
                made.append(pdf)
                os.remove(tmp_x)
                shutil.rmtree(pdfdir, ignore_errors=True)
            else:
                print("  ⚠ PDF 转换未产出文件，请直接用 xlsx 打印")

    print("\n完成！")
    for f in made:
        print(f"  {os.path.relpath(f, ROOT)}")
    print(f"  （中间产物：{os.path.relpath(work, ROOT)}/strips、traces）")


if __name__ == "__main__":
    main()