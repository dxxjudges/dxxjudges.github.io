# -*- coding: utf-8 -*-
"""用 numpy 合成 dxxjudges 介绍片的背景配乐。

纯加性合成 + 包络，不依赖任何采样素材：
    pad    缓慢铺底和声（失谐双振荡器 + 单极点低通）
    bass   根音拨弦
    arp    八分音符琶音
    kick   轻底鼓
    shaker 高频沙锤
和声进行 Am7 - Fmaj7 - Cmaj7 - G6，每小节一个和弦。
"""

import numpy as np

SR = 44100
BPM = 78.0
BEAT = 60.0 / BPM
BAR = 4.0 * BEAT


def midi_hz(m):
    return 440.0 * (2.0 ** ((m - 69) / 12.0))


# (低音 MIDI, 和弦音 MIDI)
PROGRESSION = [
    (45, [57, 60, 64, 67]),   # Am7
    (41, [57, 60, 65, 69]),   # Fmaj7
    (48, [52, 55, 59, 64]),   # Cmaj7
    (43, [55, 59, 62, 64]),   # G6
]


def _t(n):
    return np.arange(n) / SR


def _adsr(n, attack, decay, sustain_level, release):
    """线性包络，返回长度 n 的 [0,1] 曲线。"""
    a = max(1, int(attack * SR))
    d = max(1, int(decay * SR))
    r = max(1, int(release * SR))
    s = max(0, n - a - d - r)
    env = np.concatenate([
        np.linspace(0.0, 1.0, a, endpoint=False),
        np.linspace(1.0, sustain_level, d),
        np.full(s, sustain_level),
        np.linspace(sustain_level, 0.0, r),
    ])
    return env[:n] if len(env) >= n else np.pad(env, (0, n - len(env)))


def _lowpass(x, cutoff):
    """单极点低通。"""
    a = np.exp(-2.0 * np.pi * cutoff / SR)
    # 用指数衰减卷积的近似递推实现
    out = np.empty_like(x)
    acc = 0.0
    for i in range(0, len(x)):
        acc = (1 - a) * x[i] + a * acc
        out[i] = acc
    return out


def _lowpass_fast(x, cutoff, passes=2):
    """FFT 版本的一阶低通，O(n log n)，避免 Python 层的逐样本循环。"""
    spec = np.fft.rfft(x)
    freqs = np.fft.rfftfreq(len(x), 1.0 / SR)
    gain = 1.0 / np.sqrt(1.0 + (freqs / cutoff) ** 4)   # 近似二阶
    for _ in range(passes):
        spec *= gain
    return np.fft.irfft(spec, n=len(x))


def _saw(freq, n, partials=(1.0, 0.45, 0.22, 0.11)):
    t = _t(n)
    out = np.zeros(n)
    for i, amp in enumerate(partials, start=1):
        out += amp * np.sin(2 * np.pi * freq * i * t)
    return out


def _place(track, seg, at):
    """把 seg 叠加到 track 的 at 位置（自动裁剪）。"""
    a = int(at * SR)
    b = min(len(track), a + len(seg))
    if a >= len(track):
        return
    track[a:b] += seg[:b - a]


def _chord(chord_index, dur, rng):
    n = int(dur * SR)
    bass_midi, notes = PROGRESSION[chord_index % len(PROGRESSION)]

    # ── pad：每个和弦音两层失谐，低通后左右略作延时 ──
    pad_l = np.zeros(n)
    pad_r = np.zeros(n)
    env = _adsr(n, dur * 0.30, dur * 0.10, 0.85, dur * 0.35)
    for k, m in enumerate(notes):
        f = midi_hz(m)
        detune = 1.0 + (0.0016 if k % 2 else -0.0016)
        voice_l = _saw(f, n) + _saw(f * detune, n) * 0.8
        voice_r = _saw(f, n, (1.0, -0.45, 0.22, -0.11)) \
            + _saw(f / detune, n) * 0.8
        # 声部错开，避免整体呆板
        off = int(k * 0.045 * SR)
        pad_l[off:] += voice_l[:n - off] * (0.9 - 0.12 * k)
        pad_r[off:] += voice_r[:n - off] * (0.9 - 0.12 * k)
    pad_l *= env / max(len(notes), 1) * 1.9
    pad_r *= env / max(len(notes), 1) * 1.9
    pad_l = _lowpass_fast(pad_l, 2200.0)
    pad_r = _lowpass_fast(pad_r, 2200.0)

    # ── bass：1 拍与 3 拍各一个拨弦根音 ──
    bass = np.zeros(n)
    fb = midi_hz(bass_midi)
    for beat in (0.0, 2.0):
        ln = int(0.62 * SR)
        lt = _t(ln)
        env_b = np.exp(-lt * 7.0)
        tone = (np.sin(2 * np.pi * fb * lt)
                + 0.28 * np.sin(2 * np.pi * fb * 2 * lt)) * env_b
        tone = _lowpass_fast(tone, 900.0) * 0.85
        _place(bass, tone, beat * BEAT)

    # ── arp：八分音符上下行 ──
    arp = np.zeros(n)
    order = notes + [notes[2], notes[1]]
    step = BEAT / 2.0
    steps = int(dur / step)
    for i in range(steps):
        f = midi_hz(order[i % len(order)] + 12)
        ln = int(0.26 * SR)
        lt = _t(ln)
        tone = (np.sin(2 * np.pi * f * lt)
                + 0.20 * np.sin(2 * np.pi * f * 2 * lt)
                + 0.07 * np.sin(2 * np.pi * f * 3 * lt))
        tone *= np.exp(-lt * 15.0) * 0.20
        pan = 0.5 + 0.35 * np.sin(i * 0.9)
        _place(arp, tone * (1.0 - pan), i * step)
        _place(arp, tone * pan, i * step)

    # ── kick：1 拍与 3 拍 ──
    kick = np.zeros(n)
    for beat in (0.0, 2.0):
        ln = int(0.30 * SR)
        lt = _t(ln)
        sweep = 62.0 - 20.0 * (1 - np.exp(-lt * 26.0))
        kick_t = np.sin(2 * np.pi * np.cumsum(sweep) / SR) * np.exp(-lt * 9.0)
        _place(kick, kick_t * 0.55, beat * BEAT)

    # ── shaker：八分音符，前半拍弱、后半拍强 ──
    noise = rng.standard_normal(n)
    shaker = noise - np.concatenate(([0.0], noise[:-1]))   # 一阶高通
    sh = np.zeros(n)
    for i in range(steps):
        ln = int(0.055 * SR)
        burst = shaker[:ln] * np.exp(-_t(ln) * 60.0)
        accent = 1.0 if i % 2 else 0.55
        _place(sh, burst * 0.05 * accent, i * step)

    left = pad_l * 0.55 + bass * 0.62 + arp * 0.5 + kick * 0.5 + sh * 0.7
    right = pad_r * 0.55 + bass * 0.62 + arp * 0.5 + kick * 0.5 + sh * 0.7
    return left, right


def render_loop(dur=BAR * 4, seed=7):
    """生成一个无缝循环（dur 秒）。"""
    rng = np.random.default_rng(seed)
    n = int(dur * SR)
    left = np.zeros(n)
    right = np.zeros(n)

    per = PROGRESSION and BAR
    chords = int(np.ceil(dur / per))
    for c in range(chords):
        cl, cr = _chord(c, per, rng)
        start = int(c * per * SR)
        end = min(n, start + len(cl))
        if start < n:
            left[start:end] += cl[:end - start]
            right[start:end] += cr[:end - start]

    # 首尾交叉淡化，保证循环点无咔哒声
    xf = int(0.35 * SR)
    if len(left) > 2 * xf:
        ramp = np.linspace(0, 1, xf)
        head, tail = left[:xf].copy(), left[-xf:].copy()
        left[:xf] = head * ramp + tail * (1 - ramp)
        left[-xf:] = tail
        head, tail = right[:xf].copy(), right[-xf:].copy()
        right[:xf] = head * ramp + tail * (1 - ramp)
        right[-xf:] = tail
    return left, right


def build(duration, seed=7):
    """生成 length=duration 秒的立体声配乐（numpy float32, 形状 (2, N)）。"""
    loop = render_loop(BAR * 4, seed)
    n = int(duration * SR)
    reps = int(np.ceil(n / len(loop[0]))) + 1
    left = np.tile(loop[0], reps)[:n]
    right = np.tile(loop[1], reps)[:n]

    stereo = np.stack([left, right])
    # 整体淡入淡出
    fi = int(1.5 * SR)
    fo = int(3.0 * SR)
    stereo[:, :fi] *= np.linspace(0, 1, fi)[None, :]
    stereo[:, -fo:] *= np.linspace(1, 0, fo)[None, :]

    # 柔和限幅
    stereo = np.tanh(stereo * 1.15) / 1.15
    peak = float(np.abs(stereo).max())
    if peak > 0:
        stereo *= 0.72 / peak
    return stereo.astype(np.float32)