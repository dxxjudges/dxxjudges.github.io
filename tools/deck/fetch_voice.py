# -*- coding: utf-8 -*-
"""下载 Piper 中文语音模型到 voices/。

huggingface.co 在部分网络不可达，默认走 hf-mirror.com，可 --mirror 切换。

    python fetch_voice.py
    python fetch_voice.py --mirror https://huggingface.co
"""

import argparse
import os
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
DEST = os.path.join(HERE, "voices")

VOICE = "zh_CN-huayan-medium"
BASES = [
    "https://hf-mirror.com/rhasspy/piper-voices/resolve/main/zh/zh_CN/huayan/medium/",
    "https://huggingface.co/rhasspy/piper-voices/resolve/main/zh/zh_CN/huayan/medium/",
]
FILES = [VOICE + ".onnx", VOICE + ".onnx.json"]


def download(url, path):
    print("  -> %s" % url)
    req = urllib.request.Request(url, headers={"User-Agent": "curl/8"})
    with urllib.request.urlopen(req, timeout=600) as r, open(path, "wb") as f:
        total = 0
        while True:
            chunk = r.read(1 << 16)
            if not chunk:
                break
            f.write(chunk)
            total += len(chunk)
    return total


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mirror", default=None)
    args = ap.parse_args()
    bases = [args.mirror] if args.mirror else BASES

    os.makedirs(DEST, exist_ok=True)
    for name in FILES:
        path = os.path.join(DEST, name)
        if os.path.exists(path) and os.path.getsize(path) > 0:
            print("已存在 %s (%.1f MB)" % (name, os.path.getsize(path) / 1048576))
            continue

        last = None
        for base in bases:
            try:
                size = download(base + name, path)
                print("完成 %s (%.1f MB)" % (name, size / 1048576))
                last = None
                break
            except Exception as exc:                     # noqa: BLE001
                last = exc
                print("失败 %s: %s" % (base, exc))
        if last is not None:
            sys.exit("全部镜像都失败，最后一个错误: %s" % last)

    print("\n语音模型就绪: %s" % os.path.join(DEST, VOICE + ".onnx"))


if __name__ == "__main__":
    main()