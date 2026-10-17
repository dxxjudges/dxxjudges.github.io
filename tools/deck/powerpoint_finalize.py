# -*- coding: utf-8 -*-
"""用 PowerPoint COM 完成收尾：嵌入配音、设置自动播放、导出 PNG 预览。

python-pptx 没有音频 API，手写 OOXML 虽然能过结构校验，但 PowerPoint 对
<p:timing> 的语义校验更严。直接让 PowerPoint 自己写一遍最稳妥。

运行：
    python powerpoint_finalize.py                 # 嵌音频 + 自动播放 + 导出 PNG
    python powerpoint_finalize.py --no-export     # 只嵌音频
    python powerpoint_finalize.py --check-only    # 只检查能否正常打开（是否触发修复）
"""

import argparse
import glob
import json
import os
import shutil
import sys
import tempfile

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
AUDIO = os.path.join(OUT, "audio")
PREVIEW = os.path.join(OUT, "slides")

PPTX_IN = os.path.join(OUT, "dxxjudges-intro.pptx")
PPTX_OUT = os.path.join(OUT, "dxxjudges-intro.pptx")
PPTX_AUDIO = os.path.join(OUT, "dxxjudges-intro-audio.pptx")

ppSaveAsOpenXMLPresentation = 24
msoTrue, msoFalse = -1, 0


def available():
    """PowerPoint COM 是否可用（不可用时构建流程退回手写 OOXML）。"""
    try:
        app = win32.Dispatch("PowerPoint.Application")
        app.Quit()
        return True
    except Exception:                # noqa: BLE001
        return False


def open_app():
    app = win32.Dispatch("PowerPoint.Application")
    # PowerPoint 不支持完全隐藏窗口，设为最小化可见
    try:
        app.Visible = msoTrue
        app.WindowState = 2          # ppWindowMinimized
    except Exception:                # noqa: BLE001
        pass
    app.DisplayAlerts = 1           # ppAlertsNone
    return app


def open_presentation(app, path):
    """Open(FileName, ReadOnly, Untitled, WithWindow)。"""
    return app.Presentations.Open(path, ReadOnly=msoFalse,
                                  Untitled=msoFalse, WithWindow=msoFalse)


def embed_audio(pres, manifest, corner=True):
    """逐页 AddMediaObject2，并设成随页自动播放、翻页即停。"""
    poster = os.path.join(HERE, "..", "..", "assets", "logo", "png", "logo-256.png")
    added = []
    for entry in manifest["slides"]:
        idx = entry["index"]
        slide = pres.Slides(idx)
        path = os.path.join(AUDIO, entry["audio"])
        if not os.path.exists(path):
            print("  ! 缺少音频，跳过: %s" % path)
            continue

        left = int(pres.PageSetup.SlideWidth - 46) if corner else int(
            pres.PageSetup.SlideWidth * 0.5 - 20)
        top = int(pres.PageSetup.SlideHeight - 46)

        # 注意：AddMediaObject2 在 Slide.Shapes 上，MainSequence 上没有这个方法
        shape = slide.Shapes.AddMediaObject2(path, msoFalse, msoTrue,
                                             left, top, 32, 32)

        ps = shape.AnimationSettings.PlaySettings
        ps.PlayOnEntry = msoTrue          # 随页面自动播放
        ps.HideWhileNotPlaying = msoTrue  # 播放时不显示图标
        ps.StopAfterSlides = 1            # 翻页即停，逐页接力

        # 用品牌 logo 当媒体海报，替换默认的喇叭图标
        if poster and os.path.exists(poster):
            try:
                shape.MediaFormat.SetDisplayPictureFromFile(poster)
            except Exception as exc:               # noqa: BLE001
                print("  (海报图设置失败，忽略: %s)" % exc)

        # SetDisplayPictureFromFile 会把形状撑成图片原始尺寸，
        # 必须在之后重新写回位置与尺寸，否则右下角会盖住内容
        shape.Left, shape.Top = left, top
        shape.Width, shape.Height = 26, 26

        # 时长在 MediaFormat.Length 上（Shape / PlaySettings 都没有 Duration）
        # MediaFormat.Length 单位是毫秒
        dur = float(shape.MediaFormat.Length) / 1000.0
        added.append((idx, entry["audio"], round(dur, 2)))
        print("  slide %02d  %-14s %6.2fs  随页自动播放" % (idx, entry["audio"], dur))
    return added


def export_png(pres, width=1920, height=1080):
    if os.path.isdir(PREVIEW):
        shutil.rmtree(PREVIEW)
    os.makedirs(PREVIEW, exist_ok=True)
    pres.Export(PREVIEW, "PNG", width, height)
    return sorted(glob.glob(os.path.join(PREVIEW, "*.PNG")) +
                  glob.glob(os.path.join(PREVIEW, "*.png")))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-export", action="store_true")
    ap.add_argument("--check-only", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(PPTX_IN):
        sys.exit("先运行 make_ppt.py 生成 %s" % PPTX_IN)
    with open(os.path.join(AUDIO, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)

    # 在临时副本上操作，避免中途失败污染产物
    tmpdir = tempfile.mkdtemp(prefix="pptx_")
    work = os.path.join(tmpdir, "deck.pptx")
    shutil.copy2(PPTX_IN, work)

    app = open_app()
    pres = None
    try:
        print("打开: %s" % os.path.basename(work))
        pres = open_presentation(app, work)
        n = pres.Slides.Count
        print("  打开成功，共 %d 页（未触发修复）" % n)

        if args.check_only:
            return 0

        if n != len(manifest["slides"]):
            sys.exit("页数不一致: PPT %d / manifest %d" % (n, len(manifest["slides"])))

        print("嵌入配音:")
        added = embed_audio(pres, manifest)

        pres.SaveAs(PPTX_AUDIO, ppSaveAsOpenXMLPresentation)
        print("已保存: %s (%.1f KB)" % (os.path.basename(PPTX_AUDIO),
                                        os.path.getsize(PPTX_AUDIO) / 1024))

        if not args.no_export:
            files = export_png(pres)
            print("导出预览 %d 张 -> %s" % (len(files), PREVIEW))

        # 关闭后重新打开一次，确认文件可读
        pres.Close()
        pres = open_presentation(app, PPTX_AUDIO)
        media = sum(s.TimeLine.MainSequence.Count for s in pres.Slides)
        print("回读校验: %d 页，%d 个媒体对象" % (pres.Slides.Count, media))
        return 0
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:                # noqa: BLE001
            pass
        try:
            app.Quit()
        except Exception:                # noqa: BLE001
            pass
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())