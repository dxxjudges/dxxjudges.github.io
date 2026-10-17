# -*- coding: utf-8 -*-
"""PPTX 静态检查：越界、文字溢出、备注缺失。

没有 PowerPoint / LibreOffice 时，用几何与排版估算代替肉眼校对。
运行：python verify_ppt.py [pptx 路径]
"""

import math
import os
import sys
import unicodedata

from pptx import Presentation

EMU_PT = 12700.0
TOL = 2.0          # 允许 2pt 的误差


def is_cjk(ch):
    return unicodedata.east_asian_width(ch) in ("W", "F")


def char_w(ch, size):
    """粗略字宽：CJK 全角约 1em，西文约 0.55em，空格约 0.3em。"""
    if ch == " ":
        return size * 0.3
    if is_cjk(ch):
        return size * 1.0
    if ch.isupper():
        return size * 0.62
    return size * 0.55


def frame_report(tf, box_w_pt, box_h_pt):
    """返回 (需要高度pt, 最长单词超宽pt, 段落数)。"""
    total = 0.0
    widest = 0.0
    paras = 0
    for p in tf.paragraphs:
        runs = p.runs
        if not runs:
            continue
        paras += 1
        size = max((r.font.size.pt if r.font.size else 18) for r in runs)
        spacing = p.line_spacing if isinstance(p.line_spacing, float) else 1.2

        width = 0.0
        longest_token = 0.0
        token = 0.0
        for r in runs:
            for ch in r.text:
                w = char_w(ch, size)
                width += w
                # CJK 可任意断行，不计入「不可断的最长词」；只有连续的西文才算一个词
                if is_cjk(ch):
                    longest_token = max(longest_token, token)
                    token = 0.0
                elif ch.isspace():
                    longest_token = max(longest_token, token)
                    token = 0.0
                else:
                    token += w
        longest_token = max(longest_token, token)
        widest = max(widest, longest_token - box_w_pt)

        lines = max(1, math.ceil(width / max(box_w_pt, 1)))
        total += lines * size * spacing
    return total, widest, paras


def check(path):
    prs = Presentation(path)
    sw, sh = prs.slide_width, prs.slide_height
    problems = []

    for i, slide in enumerate(prs.slides, start=1):
        has_notes = slide.has_notes_slide and bool(
            slide.notes_slide.notes_text_frame.text.strip())

        for shp in slide.shapes:
            name = shp.shape_type
            l, t = shp.left, shp.top
            r, b = l + shp.width, t + shp.height

            if l < -TOL or t < -TOL or r > sw + TOL or b > sh + TOL:
                problems.append(
                    "slide %02d: 形状越界 %s  box=(%.2f, %.2f)-(%.2f, %.2f)in"
                    % (i, name, l / 914400, t / 914400,
                       r / 914400, b / 914400))

            if not shp.has_text_frame:
                continue
            tf = shp.text_frame
            if not tf.text.strip():
                continue

            bw = shp.width / EMU_PT
            bh = shp.height / EMU_PT
            need, over, paras = frame_report(tf, bw, bh)
            if need > bh + TOL:
                problems.append(
                    "slide %02d: 文字溢出  需要 %.0fpt / 框高 %.0fpt  文本=%r"
                    % (i, need, bh, tf.text[:34]))
            if over > TOL:
                problems.append(
                    "slide %02d: 单词超宽 %.0fpt / 框宽 %.0fpt  文本=%r"
                    % (i, over, bw, tf.text[:34]))

        if not has_notes:
            problems.append("slide %02d: 缺少解说备注" % i)

    return prs, problems


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "out", "dxxjudges-intro.pptx")
    prs, issues = check(target)
    print("slides: %d   size: %.1f KB"
          % (len(prs.slides._sldIdLst), os.path.getsize(target) / 1024))
    if not issues:
        print("OK - 未发现越界 / 溢出 / 缺备注")
    else:
        for p in issues:
            print("  ! " + p)
        print("\n%d 项待修" % len(issues))
        sys.exit(1)