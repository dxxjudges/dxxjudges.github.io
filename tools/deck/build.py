# -*- coding: utf-8 -*-
"""一键构建：PPT → 配音 → 配乐 → 混音 → 嵌入音频。

    python build.py                # 完整流程
    python build.py --skip-audio   # 只出画面（无配音时用）
    python build.py --fast         # 配乐不重新合成，只复用 out/audio 里的配音
"""

import argparse
import json
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
AUDIO = os.path.join(OUT, "audio")
sys.path.insert(0, HERE)


def step(n, total, title):
    print("\n[%d/%d] %s" % (n, total, title), flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--skip-audio", action="store_true")
    ap.add_argument("--no-autoplay", action="store_true")
    ap.add_argument("--voice", default=os.path.join(
        HERE, "voices", "zh_CN-huayan-medium.onnx"))
    args = ap.parse_args()

    total = 3 if args.skip_audio else 4
    t0 = time.time()

    # 1) 画面
    step(1, total, "生成 PPTX")
    import make_ppt
    pptx, count = make_ppt.build(os.path.join(OUT, "dxxjudges-intro.pptx"))
    print("  %d 页  %.1f KB" % (count, os.path.getsize(pptx) / 1024))

    import verify_ppt
    prs, issues = verify_ppt.check(pptx)
    if issues:
        for i in issues:
            print("  ! " + i)
        raise SystemExit("排版自检未通过")
    print("  排版自检通过（无越界 / 无溢出 / 备注齐全）")

    if args.skip_audio:
        print("\n完成，用时 %.1fs" % (time.time() - t0))
        return

    # 2) 配音 + 配乐 + 混音
    step(2, total, "Piper 配音 + numpy 配乐 + 闪避混音")
    import make_audio
    if not os.path.exists(args.voice):
        raise SystemExit(
            "缺少语音模型 %s\n请下载 zh_CN-huayan-medium.onnx{,.json} 到 voices/" % args.voice)
    sys.argv = ["make_audio.py", "--voice", args.voice]
    make_audio.main()
    with open(os.path.join(AUDIO, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    print("  解说 %.1fs，输出 %d 个音频文件"
          % (manifest["total_duration"], len(manifest["slides"]) + 4))

    # 3) 嵌入音频
    step(3, total, "嵌入逐页配音")
    import embed_audio
    embed_audio.embed(pptx, os.path.join(OUT, "dxxjudges-intro-audio.pptx"),
                      manifest, autoplay=False)
    print("  dxxjudges-intro-audio.pptx（点图标播放，最稳妥）")
    if not args.no_autoplay:
        embed_audio.embed(pptx, os.path.join(OUT, "dxxjudges-intro-autoplay.pptx"),
                          manifest, autoplay=True)
        print("  dxxjudges-intro-autoplay.pptx（随页自动播放）")

    # 4) 自检
    step(4, total, "校验")
    for name in ("dxxjudges-intro-audio.pptx", "dxxjudges-intro-autoplay.pptx"):
        path = os.path.join(OUT, name)
        if not os.path.exists(path):
            continue
        problems, n = embed_audio.validate(path)
        status = "OK" if not problems else "%d 项异常" % len(problems)
        print("  %-38s %7.1f KB  %d 页带音频  %s"
              % (name, os.path.getsize(path) / 1024, n, status))
        for p in problems:
            print("    ! " + p)
        if problems:
            raise SystemExit(1)

    print("\n全部完成，用时 %.1fs" % (time.time() - t0))
    print("产物目录: %s" % OUT)


if __name__ == "__main__":
    main()