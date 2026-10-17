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
ppTransitionSpeedMedium = 2
msoMedia = 16
# 只设 PlaySettings.PlayOnEntry 是不够的：PowerPoint 会把媒体写成
# <p:cond delay="indefinite"/>，即必须点击才播。
# TriggerType=WithPrevious 才会加上 <p:cond evt="onBegin" delay="0"><p:tn val="2"/>
# —— 页面一开始就播，这才是真正的自动播放。
msoAnimTriggerOnPageClick = 1
msoAnimTriggerWithPrevious = 2
msoAnimTriggerAfterPrevious = 3

# PpEntryEffect -> 实际写进 OOXML 的 <p:transition> 元素（已实测）
#   257=<p:cut/> 513=<p:random/> 769=<p:blinds/> 1025=<p:checker/>
#   1281=<p:cover/> 1537=<p:dissolve/> 3073=<p:zoom/> 3585=<p:split/>
TRANSITIONS = {
    "dissolve": 1537,      # 柔和叠化，适合整片
    "fade": 1537,
    "cover": 1281,
    "zoom": 3073,
    "split": 3585,
    "cut": 257,
    "checker": 1025,
    "blinds": 769,
    "random": 513,
    "none": 0,
}


def available():
    """PowerPoint COM 是否可用（不可用时构建流程退回手写 OOXML）。"""
    try:
        app = win32.Dispatch("PowerPoint.Application")
        app.Quit()
        return True
    except Exception:                # noqa: BLE001
        return False


def set_transitions(pres, manifest, effect="dissolve", pad=1.2,
                    min_time=4.0, duration=0.7, advance=True,
                    advance_last=True):
    """按每页解说长度设置定时自动切页，并统一入场动效。

    AdvanceOnTime  + AdvanceTime = 定时自动翻页
    AdvanceOnClick 保留，这样演讲时仍可手动接管节奏
    """
    effect_id = TRANSITIONS[effect]
    n = pres.Slides.Count
    total = 0.0
    rows = []
    for i in range(1, n + 1):
        slide = pres.Slides(i)
        tr = slide.SlideShowTransition
        tr.EntryEffect = effect_id
        if effect_id:
            tr.Speed = ppTransitionSpeedMedium
            tr.Duration = duration

        dur = min_time
        if i - 1 < len(manifest["slides"]):
            dur = max(min_time, manifest["slides"][i - 1]["duration"] + pad)

        if advance and (advance_last or i < n):
            tr.AdvanceOnTime = msoTrue
            tr.AdvanceTime = dur
        else:
            tr.AdvanceOnTime = msoFalse
        tr.AdvanceOnClick = msoTrue          # 手动仍可翻页

        rows.append((i, dur))
        total += dur

    for i, dur in rows:
        print("  slide %02d  自动切页 %5.1fs" % (i, dur))
    print("  转场=%s 时长=%.1fs  全片约 %.0fs (%.1f 分钟)"
          % (effect, duration, total, total / 60.0))
    return total


def verify_show(pres, advance_last=False):
    """回读放映设置：定时切页、动效、音频自动播放。

    注意 StopAfterSlides 不做断言——开启定时切页后 PowerPoint 会把它归零，
    而每页切页时间已是「解说长度 + 留白」，音频必然先结束，不会串到下一页。
    """
    problems = []
    n = pres.Slides.Count
    for i in range(1, n + 1):
        slide = pres.Slides(i)
        tr = slide.SlideShowTransition
        want_advance = advance_last or i < n
        is_advance = tr.AdvanceOnTime == msoTrue
        if is_advance != want_advance:
            problems.append("slide %02d: 定时切页=%s，期望 %s"
                            % (i, is_advance, want_advance))
        if tr.EntryEffect == 0:
            problems.append("slide %02d: 无入场动效" % i)
        if is_advance and tr.AdvanceTime < 1.0:
            problems.append("slide %02d: AdvanceTime 过小 %.2f"
                            % (i, tr.AdvanceTime))

        seq = slide.TimeLine.MainSequence
        if seq.Count < 1:
            problems.append("slide %02d: 没有媒体对象" % i)
            continue
        for j in range(1, seq.Count + 1):
            eff = seq.Item(j)
            if eff.Shape.Type != msoMedia:
                continue
            ps = eff.Shape.AnimationSettings.PlaySettings
            if ps.PlayOnEntry != msoTrue:
                problems.append("slide %02d: 音频不是随页自动播放" % i)
            if eff.Timing.TriggerType != msoAnimTriggerWithPrevious:
                problems.append("slide %02d: 触发方式=%s，期望 WithPrevious"
                                % (i, eff.Timing.TriggerType))
    return problems


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


def set_trigger_autoplay(slide, shape):
    """把某个媒体形状所在的动画节点设为 WithPrevious（=页面开始即播）。"""
    seq = slide.TimeLine.MainSequence
    target = shape.Id
    for j in range(1, seq.Count + 1):
        eff = seq.Item(j)
        try:
            if eff.Shape.Id != target:
                continue
        except Exception:                # noqa: BLE001
            continue
        eff.Timing.TriggerType = msoAnimTriggerWithPrevious
        eff.Timing.TriggerDelayTime = 0
        return True
    return False


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

        # 关键：PlayOnEntry 单独设只会得到 clickEffect（需点击）。
        # 把该媒体所在的动画节点改成 WithPrevious，才会写入 onBegin 触发条件。
        set_trigger_autoplay(slide, shape)

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
    ap.add_argument("--no-advance", action="store_true",
                    help="不开启定时自动切页")
    ap.add_argument("--advance-last", action="store_true",
                    help="末页也定时切页（默认停在末页）")
    ap.add_argument("--pad", type=float, default=1.2,
                    help="每页解说结束到翻页的留白秒数")
    ap.add_argument("--transition", default="dissolve",
                    choices=sorted(TRANSITIONS), help="入场动效")
    ap.add_argument("--transition-seconds", type=float, default=0.7)
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

        print("放映设置:")
        set_transitions(pres, manifest, effect=args.transition, pad=args.pad,
                        duration=args.transition_seconds,
                        advance=not args.no_advance,
                        advance_last=args.advance_last)

        pres.SaveAs(PPTX_AUDIO, ppSaveAsOpenXMLPresentation)
        print("已保存: %s (%.1f KB)" % (os.path.basename(PPTX_AUDIO),
                                        os.path.getsize(PPTX_AUDIO) / 1024))

        if not args.no_export:
            files = export_png(pres)
            print("导出预览 %d 张 -> %s" % (len(files), PREVIEW))

        # 关闭后重新打开一次，确认文件可读且放映设置真的写进去了
        pres.Close()
        pres = open_presentation(app, PPTX_AUDIO)
        media = sum(s.TimeLine.MainSequence.Count for s in pres.Slides)
        print("回读校验: %d 页，%d 个媒体对象" % (pres.Slides.Count, media))
        problems = verify_show(pres, advance_last=args.advance_last)
        if problems:
            for p in problems:
                print("  ! " + p)
            raise SystemExit(1)
        print("放映设置回读通过：定时切页 + 入场动效 + 音频随页自动播放")
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