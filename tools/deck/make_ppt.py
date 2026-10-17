# -*- coding: utf-8 -*-
"""用 python-pptx 生成 dxxjudges 组织介绍 PPT。

画面配色与 https://dxxj.de5.net 保持一致（深蓝底 + 靛蓝/青色渐变）。
运行：python make_ppt.py [输出路径]
"""

import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Inches, Pt

import content

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
LOGO = os.path.join(ROOT, "assets", "logo", "png", "logo-512.png")
MARK = os.path.join(ROOT, "assets", "logo", "png", "mark-white-256.png")

# ── 主题 ──────────────────────────────────────────────────────────────────
BG = RGBColor(0x0B, 0x10, 0x20)
CARD = RGBColor(0x14, 0x1C, 0x38)
BORDER = RGBColor(0x22, 0x30, 0x5C)
TEXT = RGBColor(0xE8, 0xEC, 0xF7)
DIM = RGBColor(0x9A, 0xA6, 0xC8)
FAINT = RGBColor(0x6B, 0x77, 0x99)
INDIGO = RGBColor(0x63, 0x66, 0xF1)
CYAN = RGBColor(0x22, 0xD3, 0xEE)
VIOLET = RGBColor(0xA7, 0x8B, 0xFA)
GREEN = RGBColor(0x34, 0xD3, 0x99)
AMBER = RGBColor(0xFB, 0xBF, 0x24)

FONT = "Microsoft YaHei"
FONT_MONO = "Consolas"

W, H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.9)
PREVIEW_PX_W = 1920          # PNG 预览导出宽度，preview_px() 据此换算


def preview_px(n):
    """把「预览里的像素」换算成 EMU，便于口头微调位置。"""
    return Emu(int(round(n / float(PREVIEW_PX_W) * W)))


# ── 基础绘制工具 ──────────────────────────────────────────────────────────
def blank(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG
    bg.line.fill.background()
    bg.shadow.inherit = False
    return slide


def rect(slide, x, y, w, h, fill=CARD, line=BORDER, radius=0.045, line_w=1.0):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.adjustments[0] = radius
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_w)
    shape.shadow.inherit = False
    shape.text_frame.word_wrap = True
    return shape


def text(slide, x, y, w, h, runs, size=18, color=TEXT, bold=False, font=FONT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=1.25, space=0):
    """runs 的每一项：

    - ``"文本\\n文本"``（str）    -> 每行一个段落
    - ``("片段", {opts})``（tuple）-> 追加到当前段落的内联 run

    传入的 ``\\n`` 不会被 python-pptx 转成换行，必须走 str 的分段逻辑。
    """
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    # 定版式：不要 python-pptx 默认的 spAutoFit，否则框高会被 PowerPoint 重算
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0

    if isinstance(runs, str):
        runs = [runs]

    para = None
    for chunk in runs:
        if isinstance(chunk, tuple):
            body, opts = chunk
            if para is None:
                para = new_para(tf)
            run = para.add_run()
            run.text = body
            style(para, run, size, color, bold, font, opts, align, line, space)
        else:
            para = None
            for line_text in chunk.split("\n"):
                para = new_para(tf)
                run = para.add_run()
                run.text = line_text or " "
                style(para, run, size, color, bold, font, {}, align, line, space)
    return box


def new_para(tf):
    if len(tf.paragraphs) == 1 and not tf.paragraphs[0].runs:
        return tf.paragraphs[0]
    return tf.add_paragraph()


def style(para, run, size, color, bold, font, opts, align, line, space):
    f = run.font
    f.size = Pt(opts.get("size", size))
    f.bold = opts.get("bold", bold)
    f.color.rgb = opts.get("color", color)
    f.name = opts.get("font", font)
    # 段落属性必须写在段落对象上；早期版本从 run._r.getparent() 取段落，
    # 结果 algn / lnSpc 全部丢失（字号颜色却正常），很难发现。
    para.alignment = opts.get("align", align)
    para.line_spacing = opts.get("line", line)
    if space:
        para.space_after = Pt(space)


def gradient_banner(slide, x, y, w, h, c1=INDIGO, c2=CYAN):
    """用两层矩形近似一条渐变光带（python-pptx 原生渐变 API 在形状上可用，
    这里用带渐变的矩形以获得更柔和的过渡）。"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.line.fill.background()
    shape.shadow.inherit = False
    fill = shape.fill
    fill.gradient()
    fill.gradient_angle = 0.0
    stops = fill.gradient_stops
    stops[0].color.rgb = c1
    stops[0].position = 0.0
    stops[1].color.rgb = c2
    stops[1].position = 1.0
    return shape


def header(slide, tag, title, lead=None):
    """页眉：小标签 + 主标题 + 副标题。"""
    gradient_banner(slide, 0, 0, Inches(0.16), H)
    text(slide, MARGIN, Inches(0.62), Inches(6), Inches(0.3), tag.upper(),
         size=12, color=CYAN, bold=True)
    text(slide, MARGIN, Inches(0.95), Inches(11.5), Inches(0.8), title,
         size=34, bold=True)
    if lead:
        text(slide, MARGIN, Inches(1.72), Inches(11.5), Inches(0.5), lead,
             size=15, color=DIM)


def footer(slide, page, total):
    text(slide, MARGIN, H - Inches(0.62), Inches(9), Inches(0.3),
         content.DECK_FOOT, size=10, color=FAINT)
    text(slide, W - MARGIN - Inches(2), H - Inches(0.62), Inches(2),
         Inches(0.3), "%02d / %02d" % (page, total),
         size=10, color=FAINT, align=PP_ALIGN.RIGHT)


def notes(slide, say):
    """把解说词写进演讲者备注，导出后仍可查看。"""
    slide.notes_slide.notes_text_frame.text = say


# ── 版式 ──────────────────────────────────────────────────────────────────
def cover(prs, s, page=1, total=1):
    slide = blank(prs)
    # 光带默认 2.42in，下移 50px（按 1920px 宽预览折算）
    gradient_banner(slide, 0, Inches(2.42) + preview_px(50), W, Inches(0.05))

    if os.path.exists(LOGO):
        slide.shapes.add_picture(LOGO, W // 2 - Inches(0.72),
                                 Inches(1.15), Inches(1.44), Inches(1.44))

    text(slide, 0, Inches(2.85), W, Inches(1.0), s["title"],
         size=54, bold=True, align=PP_ALIGN.CENTER)
    text(slide, 0, Inches(3.92), W, Inches(0.5), s["subtitle"],
         size=20, color=DIM, align=PP_ALIGN.CENTER)

    for i, (label, url, colour) in enumerate((
            ("NoldOJ", content.SITE_NOLD, INDIGO),
            ("CliJudge", content.SITE_CLI, CYAN))):
        x = W // 2 + Inches(0.2) if i else W // 2 - Inches(3.4)
        chip = rect(slide, x, Inches(4.75), Inches(3.2), Inches(0.52),
                    fill=CARD, line=BORDER, radius=0.5)
        text(slide, x, Inches(4.87), Inches(3.2), Inches(0.3),
             [(label + "  ", {"size": 13, "bold": True, "color": colour}),
              (url.replace("https://", ""), {"size": 12, "color": DIM})],
             align=PP_ALIGN.CENTER)

    text(slide, 0, H - Inches(0.9), W, Inches(0.3), s["note"],
         size=12, color=FAINT, align=PP_ALIGN.CENTER)
    notes(slide, s["say"])


def bullets(prs, s, page, total):
    slide = blank(prs)
    header(slide, s["tag"], s["title"], s.get("lead"))

    top = Inches(2.45)
    n = len(s["items"])
    card_h = Inches(0.96)
    # 按条目数自适应行距，5 条也不会顶到页脚
    avail = (H - Inches(0.95)) - top
    gap = min(Inches(1.14), (avail - card_h) / max(n - 1, 1))
    for i, (head, body) in enumerate(s["items"]):
        y = top + gap * i
        rect(slide, MARGIN, y, Inches(11.53), card_h)
        accent = (INDIGO, CYAN, VIOLET, GREEN, AMBER)[i % 5]
        rect(slide, MARGIN, y, Inches(0.06), card_h,
             fill=accent, line=None, radius=0.5)
        text(slide, MARGIN + Inches(0.34), y + Inches(0.16), Inches(3.0),
             Inches(0.6), head, size=17, bold=True)
        text(slide, MARGIN + Inches(3.4), y + Inches(0.17), Inches(7.7),
             Inches(0.7), body, size=13.5, color=DIM, line=1.35)
    footer(slide, page, total)
    notes(slide, s["say"])


def cards(prs, s, page, total):
    slide = blank(prs)
    header(slide, s["tag"], s["title"], s.get("lead"))

    cw, ch = Inches(5.62), Inches(3.05)
    for i, card in enumerate(s["cards"]):
        x = MARGIN + (cw + Inches(0.29)) * i
        y = Inches(2.45)
        rect(slide, x, y, cw, ch)
        accent = INDIGO if i == 0 else CYAN
        rect(slide, x, y, cw, Inches(0.05), fill=accent, line=None, radius=0.5)

        text(slide, x + Inches(0.36), y + Inches(0.3), cw - Inches(0.72),
             Inches(0.4), card["title"], size=24, bold=True, color=accent)
        text(slide, x + Inches(0.36), y + Inches(0.78), cw - Inches(0.72),
             Inches(0.34), card["sub"], size=13, color=FAINT)
        text(slide, x + Inches(0.36), y + Inches(1.16), cw - Inches(0.72),
             Inches(0.3), card["site"], size=14, bold=True, color=CYAN)

        body = "\n".join("· " + p for p in card["points"])
        text(slide, x + Inches(0.36), y + Inches(1.62), cw - Inches(0.72),
             Inches(1.3), body, size=13.5, color=DIM, line=1.55)

    text(slide, MARGIN, Inches(6.05), Inches(11.53), Inches(0.3),
         [("源码仓库  ", {"size": 12, "color": FAINT}),
          (content.REPO_NOLD.replace("https://", "") + "     ", {"size": 12, "color": DIM}),
          (content.REPO_CLI.replace("https://", ""), {"size": 12, "color": DIM})])
    footer(slide, page, total)
    notes(slide, s["say"])


def spotlight(prs, s, page, total):
    slide = blank(prs)
    header(slide, s["tag"], s["title"], s.get("lead"))

    top = Inches(2.5)
    cw, ch = Inches(5.62), Inches(1.28)
    for i, (head, body) in enumerate(s["items"]):
        col, row = i % 2, i // 2
        x = MARGIN + (cw + Inches(0.29)) * col
        y = top + (ch + Inches(0.24)) * row
        rect(slide, x, y, cw, ch)
        badge = rect(slide, x + Inches(0.3), y + Inches(0.3), Inches(0.62),
                     Inches(0.62), fill=BG, line=None, radius=0.35)
        text(slide, x + Inches(0.3), y + Inches(0.44), Inches(0.62),
             Inches(0.34), "%02d" % (i + 1), size=15, bold=True,
             color=CYAN, align=PP_ALIGN.CENTER)
        text(slide, x + Inches(1.08), y + Inches(0.28), cw - Inches(1.4),
             Inches(0.32), head, size=15.5, bold=True)
        text(slide, x + Inches(1.08), y + Inches(0.66), cw - Inches(1.4),
             Inches(0.6), body, size=12.5, color=DIM, line=1.3)
    footer(slide, page, total)
    notes(slide, s["say"])


def table(prs, s, page, total):
    slide = blank(prs)
    header(slide, s["tag"], s["title"], s.get("lead"))

    rows, cols = len(s["rows"]) + 1, len(s["headers"])
    top, left = Inches(2.35), MARGIN
    width, height = Inches(11.53), Inches(4.3)
    gt = slide.shapes.add_table(rows, cols, left, top, width, height).table
    gt.columns[0].width = Inches(2.13)
    for c in range(1, cols):
        gt.columns[c].width = Inches(4.7)

    for c, head in enumerate(s["headers"]):
        cell = gt.cell(0, c)
        cell.text = head
        cell.fill.solid()
        cell.fill.fore_color.rgb = BORDER
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.name = FONT
        p.font.color.rgb = TEXT if c == 0 else (VIOLET if c == 1 else CYAN)
        cell.margin_left = Inches(0.16)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    for r, row in enumerate(s["rows"], start=1):
        for c, val in enumerate(row):
            cell = gt.cell(r, c)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(12.5)
            p.font.name = FONT
            p.font.bold = (c == 0)
            p.font.color.rgb = TEXT if c == 0 else DIM
            cell.margin_left = Inches(0.16)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    for r in range(rows):
        gt.rows[r].height = Inches(0.52)
    footer(slide, page, total)
    notes(slide, s["say"])


def code(prs, s, page, total):
    slide = blank(prs)
    header(slide, s["tag"], s["title"], s.get("lead"))

    cw = Inches(5.62)
    for i, (title, lines) in enumerate(s["blocks"]):
        x = MARGIN + (cw + Inches(0.29)) * i
        y, h = Inches(2.45), Inches(2.85)
        rect(slide, x, y, cw, h, fill=RGBColor(0x0A, 0x0E, 0x1D))
        rect(slide, x, y, cw, Inches(0.42), fill=BG, line=None, radius=0.12)
        for k, c in enumerate((RGBColor(0xF8, 0x71, 0x71), AMBER, GREEN)):
            dot = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, x + Inches(0.22 + 0.26 * k),
                y + Inches(0.15), Inches(0.13), Inches(0.13))
            dot.fill.solid()
            dot.fill.fore_color.rgb = c
            dot.line.fill.background()
            dot.shadow.inherit = False
        text(slide, x + Inches(1.06), y + Inches(0.1), cw - Inches(1.2),
             Inches(0.28), title, size=12.5, bold=True, color=DIM)

        body = "\n".join(lines) if lines else " "
        text(slide, x + Inches(0.3), y + Inches(0.68), cw - Inches(0.6),
             h - Inches(0.92), body, size=12, color=RGBColor(0xC7, 0xD2, 0xFE),
             font=FONT_MONO, line=1.7)
    footer(slide, page, total)
    notes(slide, s["say"])


def links(prs, s, page, total):
    slide = blank(prs)
    header(slide, s["tag"], s["title"], s.get("lead"))

    cw, ch = Inches(5.62), Inches(3.5)
    for i, (title, items) in enumerate(s["groups"]):
        x = MARGIN + (cw + Inches(0.29)) * i
        y = Inches(2.45)
        rect(slide, x, y, cw, ch)
        accent = INDIGO if i == 0 else CYAN
        rect(slide, x, y, cw, Inches(0.05), fill=accent, line=None, radius=0.5)
        text(slide, x + Inches(0.36), y + Inches(0.34), cw - Inches(0.72),
             Inches(0.4), title, size=22, bold=True, color=accent)
        for k, (label, url) in enumerate(items):
            yy = y + Inches(1.02 + 0.47 * k)
            text(slide, x + Inches(0.36), yy, Inches(1.5), Inches(0.3),
                 label, size=11.5, color=FAINT)
            text(slide, x + Inches(1.88), yy, cw - Inches(2.24), Inches(0.3),
                 url.replace("https://", ""), size=10.5, color=TEXT)
    footer(slide, page, total)
    notes(slide, s["say"])


def cta(prs, s, page=1, total=1):
    slide = blank(prs)
    # 与封面一致：光带下移 50px，并把下方文案整体跟着挪，
    # 只挪光带会正好落到标题上（2.93in 落在标题框 2.92~3.72 内）
    drop = preview_px(50)
    gradient_banner(slide, 0, Inches(2.58) + drop, W, Inches(0.05))

    if os.path.exists(LOGO):
        slide.shapes.add_picture(LOGO, W // 2 - Inches(0.6), Inches(1.5),
                                 Inches(1.2), Inches(1.2))
    text(slide, 0, Inches(2.92) + drop, W, Inches(0.8), s["title"], size=40,
         bold=True, align=PP_ALIGN.CENTER)
    text(slide, 0, Inches(3.85) + drop, W, Inches(0.4), s["subtitle"], size=19,
         color=DIM, align=PP_ALIGN.CENTER)
    text(slide, 0, Inches(4.55) + drop, W, Inches(0.4), s["note"], size=15,
         color=CYAN, align=PP_ALIGN.CENTER)
    footer(slide, page, total)
    notes(slide, s["say"])


LAYOUTS = {
    "cover": cover,
    "bullets": bullets,
    "cards": cards,
    "spotlight": spotlight,
    "table": table,
    "code": code,
    "links": links,
    "cta": cta,
}


def build(out_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H

    total = len(content.SLIDES)
    for i, s in enumerate(content.SLIDES):
        LAYOUTS[s["kind"]](prs, s, i + 1, total)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    prs.save(out_path)
    return out_path, total


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        HERE, "out", "dxxjudges-intro.pptx")
    path, count = build(target)
    print("PPTX: %s  (%d slides, %.1f KB)" %
          (path, count, os.path.getsize(path) / 1024))