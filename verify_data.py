# -*- coding: utf-8 -*-
"""校验 stroke_data_cn.json —— 笔画数据本身的完整性（早于渲染）。

检查四件不变量：
1) `strokes` 与 `medians` 数量一致（每个笔画既要有轮廓、也要有骨架中心线）
2) 所有轮廓坐标都落在 0..1024 的坐标盒里
3) 所有轮廓都是**闭合**路径（渲染描红靠填充轮廓，不闭合就会漏色）
4) 轮廓非退化（点够多、包围盒不为零）

关于「这个数据是不是真的属于这个字」：**这里不做**。试过拿笔画骨架/轮廓和系统字体
渲染的同一个字比 IoU，结果 `二` 只有 0.021、`一` 却有 0.65 —— 因为字库是**楷体手写
风格**（`一` 右上斜、`二` 上短下长），而对照字体是印刷黑体，笔画粗细和位置本来就不同。
不同字体之间比 IoU 判定不了字形身份，只会误报。

字形正确性由另外两层保证：
* verify_render.py —— 教学条的像素级结构（第 k 格只有第 k 笔是红的、黑笔在累加）
* verify_pdf.py    —— 最终 PDF 的几何、颜色、描红格数量
"""
import json
import os

from stroke_lib import stroke_polylines, char_bbox

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = json.load(open(os.path.join(HERE, "stroke_data_cn.json"), encoding="utf-8"))

BOX = 1024
TOL = 1.0           # 坐标越界容忍（数据里有 -0.0x 这种浮点尾巴）
CLOSE_TOL = 1.0     # 轮廓首尾点距离容忍


def main():
    print(f"笔画数据：stroke_data_cn.json（本目录），共 {len(DATA)} 字")

    n_strokes = 0
    bad_count, bad_range, bad_open, bad_degen, bad_bbox = [], [], [], [], []

    for ch, d in DATA.items():
        strokes, medians = d.get("strokes"), d.get("medians")
        if not strokes or not medians or len(strokes) != len(medians):
            bad_count.append((ch, len(strokes or []), len(medians or [])))
            continue

        for k, sp in enumerate(strokes):
            n_strokes += 1
            polys = stroke_polylines(sp) or []
            pts = [p for pl in polys for p in pl]
            if len(pts) < 3:
                bad_degen.append((ch, k, len(pts)))
                continue
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            if min(xs) < -TOL or max(xs) > BOX + TOL or min(ys) < -TOL or max(ys) > BOX + TOL:
                bad_range.append((ch, k, round(min(xs), 1), round(max(xs), 1),
                                  round(min(ys), 1), round(max(ys), 1)))
            if max(xs) - min(xs) < 1 or max(ys) - min(ys) < 1:
                bad_degen.append((ch, k, "包围盒退化"))
            for pl in polys:                       # 轮廓必须闭合，否则填充会漏
                if pl and (abs(pl[0][0] - pl[-1][0]) > CLOSE_TOL
                           or abs(pl[0][1] - pl[-1][1]) > CLOSE_TOL):
                    bad_open.append((ch, k))

        bb = char_bbox(strokes)
        if not bb or (bb[1] - bb[0] < 10 and bb[3] - bb[2] < 10):
            bad_bbox.append((ch, bb))

    print(f"共 {n_strokes} 条轮廓")
    print()
    checks = [
        ("strokes / medians 数量一致", bad_count),
        (f"轮廓坐标在 0..{BOX} 内", bad_range),
        ("轮廓闭合", bad_open),
        ("轮廓非退化", bad_degen),
        ("整字包围盒有效", bad_bbox),
    ]
    for name, bad in checks:
        print(f"  {'✓' if not bad else '✗'} {name}" + (f"　异常 {len(bad)} 处：{bad[:5]}" if bad else ""))

    failed = [n for n, b in checks if b]
    print()
    if failed:
        print(f"判定：{'、'.join(failed)} 未通过 ✗")
        return 1
    print("判定：数据完整性全部通过 ✓")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())