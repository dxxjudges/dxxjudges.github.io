# -*- coding: utf-8 -*-
"""生成 dxxjudges 介绍片的配音与配乐。

配音：Piper（piper-onnx + zh_CN-huayan-medium），按句合成并插入停顿
配乐：numpy 加性合成（bgm.py），按人声包络做闪避（ducking）
混音：numpy，输出 wav 与 mp3（ffmpeg 编码），并给出 manifest.json

运行：python make_audio.py [--voice voices/zh_CN-huayan-medium.onnx]
"""

import argparse
import builtins
import json
import os
import re
import subprocess
import sys

import numpy as np
import soundfile as sf

import bgm
import content

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out", "audio")

SR = 44100
VOICE_SR = 22050           # huayan-medium 的原生采样率
SENT_GAP = 0.22            # 句间停顿
SLIDE_GAP = 1.10           # 页间停顿
MIN_SLIDE = 5.0            # 每页最少占位，避免翻页太快

# piper_onnx 用无 encoding 的 open() 读配置，Windows 默认 GBK 码页会炸，
# 这里统一把文本模式的 open 换成 UTF-8。
_open = builtins.open


def _open_utf8(file, mode="r", *args, **kwargs):
    if "b" not in mode and "encoding" not in kwargs:
        kwargs["encoding"] = "utf-8"
    return _open(file, mode, *args, **kwargs)


builtins.open = _open_utf8

from piper_onnx import Piper  # noqa: E402  (必须在 open shim 之后导入)


# ── 音频小工具 ────────────────────────────────────────────────────────────
def resample(x, src, dst):
    if src == dst:
        return x
    n = int(round(len(x) * dst / src))
    return np.interp(np.linspace(0, len(x) - 1, n),
                     np.arange(len(x)), x).astype(np.float32)


def trim_silence(x, thresh_ratio=0.006, pad=0.06):
    if len(x) == 0:
        return x
    peak = float(np.abs(x).max()) or 1.0
    win = max(1, int(0.01 * SR))
    frames = np.lib.stride_tricks.sliding_window_view(x, win)[::win]
    rms = np.sqrt((frames.astype(np.float64) ** 2).mean(axis=1))
    loud = np.where(rms > peak * thresh_ratio)[0]
    if not len(loud):
        return x
    p = int(pad * SR)
    a = max(0, loud[0] * win - p)
    b = min(len(x), (loud[-1] + 1) * win + p)
    return x[a:b]


def soft_limit(x, ceiling=0.95):
    return np.tanh(x * 1.05) / 1.05 * ceiling


def rms(x):
    return float(np.sqrt((x.astype(np.float64) ** 2).mean())) if len(x) else 0.0


def envelope(x, hop=0.02):
    """20ms 帧 RMS 包络。"""
    w = max(1, int(hop * SR))
    n = int(np.ceil(len(x) / w))
    padded = np.pad(x, (0, n * w - len(x)))
    frames = padded.reshape(n, w).astype(np.float64)
    return np.sqrt((frames ** 2).mean(axis=1))


def smooth_envelope(env, attack=0.05, release=0.40):
    """非对称单极点平滑：起音快、释放慢。"""
    out = np.empty_like(env)
    acc = 0.0
    a = float(np.exp(-1.0 / max(attack / 0.02, 1e-6)))
    r = float(np.exp(-1.0 / max(release / 0.02, 1e-6)))
    for i, v in enumerate(env):
        acc = (1 - (a if v > acc else r)) * v + (a if v > acc else r) * acc
        out[i] = acc
    return out


# ── 配音 ──────────────────────────────────────────────────────────────────
def split_sentences(text):
    parts = re.split(r"(?<=[。！？!?；;])\s*", text.strip())
    return [p for p in parts if p]


def synth_slide(piper, text, length_scale=1.05):
    chunks = []
    for i, sentence in enumerate(split_sentences(text)):
        audio, sr = piper.create(sentence, length_scale=length_scale)
        chunks.append(np.asarray(audio, dtype=np.float32))
        if i < len(split_sentences(text)) - 1:
            chunks.append(np.zeros(int(SENT_GAP * sr), dtype=np.float32))
    if not chunks:
        return np.zeros(int(0.5 * VOICE_SR), dtype=np.float32)

    voice = resample(np.concatenate(chunks), VOICE_SR, SR)
    voice = trim_silence(voice)
    peak = float(np.abs(voice).max()) or 1.0
    voice = soft_limit(voice / peak * 0.92)
    return voice.astype(np.float32)


# ── 混音 ──────────────────────────────────────────────────────────────────
def duck(bgm_stereo, voice_env, depth_db=13.0, floor_db=-24.0):
    """按人声包络闪避配乐。

    用 95 分位做参考电平并在 dB 域做检测（压缩器式），
    避免用绝对峰值归一化后正常语音帧只落到 0.2 的 level、闪避几乎不生效。
    """
    env = smooth_envelope(voice_env)
    ref = float(np.percentile(env, 95)) or float(env.max()) or 1.0
    db = 20.0 * np.log10(np.maximum(env, 1e-9) / ref)
    level = np.clip((db - floor_db) / (0.0 - floor_db), 0.0, 1.0)
    gain = 10.0 ** (-depth_db * level / 20.0)
    n = bgm_stereo.shape[1]
    # 注意：np.resize 会循环平铺而不是重采样，包络必须用插值拉伸到样本数
    g = np.interp(np.linspace(0, len(gain) - 1, n),
                  np.arange(len(gain)), gain)
    fade = min(n, int(0.05 * SR))
    g[:fade] *= np.linspace(0, 1, fade)              # 起始不突然压低
    return bgm_stereo * g[None, :]


def to_mp3(wav_path, mp3_path, bitrate="128k"):
    subprocess.run(
        ["ffmpeg", "-y", "-loglevel", "error", "-i", wav_path,
         "-c:a", "libmp3lame", "-b:a", bitrate, "-ar", str(SR), "-ac", "1",
         mp3_path],
        check=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--voice", default=os.path.join(
        HERE, "voices", "zh_CN-huayan-medium.onnx"))
    ap.add_argument("--length-scale", type=float, default=1.05)
    args = ap.parse_args()

    cfg = args.voice + ".json"
    if not os.path.exists(args.voice):
        sys.exit("找不到语音模型: %s\n"
                 "请先下载 zh_CN-huayan-medium 的 .onnx 与 .onnx.json 到 voices/" % args.voice)

    os.makedirs(OUT, exist_ok=True)
    print("Piper: %s" % os.path.basename(args.voice))
    piper = Piper(args.voice, cfg)
    print("  sample rate: %d Hz" % piper.sample_rate)

    # 1) 逐页配音
    clips, layout = [], []
    for i, slide in enumerate(content.SLIDES, start=1):
        audio = synth_slide(piper, slide["say"], args.length_scale)
        clips.append(audio)
        layout.append((i, slide.get("title", slide.get("subtitle", "")),
                       len(audio) / SR))
        print("  slide %02d  %5.1fs  %s" % (i, len(audio) / SR, layout[-1][1]))

    # 2) 拼成整条解说时间线
    voice_track = np.zeros(int(1.0 * SR), dtype=np.float32)
    offsets = []
    for clip in clips:
        offsets.append(len(voice_track) / SR)
        voice_track = np.concatenate([voice_track, clip, np.zeros(int(SLIDE_GAP * SR), np.float32)])
    total = len(voice_track) / SR
    print("解说总长 %.1fs" % total)

    # 3) 配乐 + 闪避 + 混音
    print("合成配乐 %.1fs ..." % total)
    music = bgm.build(total + 2.0)
    voice_env = envelope(voice_track)
    ducked = duck(music, voice_env)
    mix_n = music.shape[1]
    voice_full = np.pad(voice_track, (0, max(0, mix_n - len(voice_track))))
    mix = voice_full[None, :] + ducked * 0.42
    mix = np.tanh(mix * 1.02) / 1.02
    peak = float(np.abs(mix).max())
    if peak > 0:
        mix *= 0.95 / peak

    # 4) 落盘
    files = {}
    for i, clip in enumerate(clips, start=1):
        wav = os.path.join(OUT, "slide-%02d.wav" % i)
        sf.write(wav, clip, SR, subtype="PCM_16")
        mp3 = os.path.join(OUT, "slide-%02d.mp3" % i)
        to_mp3(wav, mp3)
        files["slide%02d" % i] = os.path.basename(mp3)

    narration = os.path.join(OUT, "narration.wav")
    sf.write(narration, voice_track, SR, subtype="PCM_16")

    bgm_mono = (music.mean(axis=0) * 0.9).astype(np.float32)
    bgm_wav = os.path.join(OUT, "bgm.wav")
    sf.write(bgm_wav, bgm_mono, SR, subtype="PCM_16")
    to_mp3(bgm_wav, os.path.join(OUT, "bgm.mp3"))

    mix_wav = os.path.join(OUT, "mix.wav")
    sf.write(mix_wav, mix.T, SR, subtype="PCM_16")
    to_mp3(mix_wav, os.path.join(OUT, "mix.mp3"), "192k")

    manifest = dict(
        voice=os.path.basename(args.voice),
        sample_rate=SR,
        total_duration=round(total, 3),
        narration_rms=round(rms(voice_track), 5),
        bgm_rms=round(rms(music.mean(axis=0)), 5),
        slides=[dict(index=i, title=t, duration=round(d, 3), start=round(o, 3),
                     audio=files["slide%02d" % i])
                for (i, t, d), o in zip(layout, offsets)],
    )
    with open(os.path.join(OUT, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print("\n输出目录: %s" % OUT)
    for name in sorted(os.listdir(OUT)):
        print("  %-18s %7.1f KB" % (name, os.path.getsize(os.path.join(OUT, name)) / 1024))


if __name__ == "__main__":
    main()