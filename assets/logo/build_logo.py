#!/usr/bin/env python3
"""dxxjudges logo builder.

Single source of geometry -> renders both the vector (SVG) and raster (PNG/ICO)
assets so the two can never drift apart.

Concept: code brackets wrapping a sandbox shield.
All coordinates live in a 1024x1024 viewBox.

Usage:  python build_logo.py
"""

import math
import os

import numpy as np
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PNG = os.path.join(HERE, "png")

VB = 1024          # viewBox size
SS = 4             # supersampling factor for rasterisation

# Brand ------------------------------------------------------------------
C1 = (0x63, 0x66, 0xF1)      # #6366f1  indigo
C2 = (0x22, 0xD3, 0xEE)      # #22d3ee  cyan
WHITE = (255, 255, 255)

# Geometry ---------------------------------------------------------------
RADIUS = 224                                    # badge corner radius
STROKE = 84                                     # bracket stroke width
LEFT_BRACKET = [(366, 340), (250, 512), (366, 684)]
RIGHT_BRACKET = [(658, 340), (774, 512), (658, 684)]

SH_X0, SH_X1 = 436, 588       # shield left / right
SH_TOP = 380                  # shield top edge
SH_SHOULDER = 502             # where the straight side stops curving
SH_TIP = (512, 646)           # bottom point
SH_TOP_R = 30                 # top corner radius
SH_CTRL_R = (588, 566)        # right side cubic control #1
SH_CTRL_R2 = (556, 610)       # right side cubic control #2
SH_CTRL_L = (468, 610)        # left side cubic control #2
SH_CTRL_L1 = (436, 566)       # left side cubic control #1


def shield_points():
    """Flatten the shield outline into a point list (clockwise)."""
    x0, x1, top = SH_X0, SH_X1, SH_TOP
    r = SH_TOP_R
    pts = [(x0 + r, top), (x1 - r, top)]
    pts += arc(x1 - r, top + r, r, -90, 0)
    pts += [(x1, SH_SHOULDER)]
    pts += cubic((x1, SH_SHOULDER), SH_CTRL_R, SH_CTRL_R2, SH_TIP)
    pts += cubic(SH_TIP, SH_CTRL_L, SH_CTRL_L1, (x0, SH_SHOULDER))
    pts += [(x0, top + r)]
    pts += arc(x0 + r, top + r, r, 180, 270)
    return pts


def arc(cx, cy, r, a0, a1, n=12):
    """Circular arc from angle a0 to a1 (degrees), excluding start point."""
    out = []
    for i in range(1, n + 1):
        a = math.radians(a0 + (a1 - a0) * i / n)
        out.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return out


def cubic(p0, p1, p2, p3, n=40):
    """Cubic bezier flattened to points, excluding the start point."""
    out = []
    for i in range(1, n + 1):
        t = i / n
        u = 1 - t
        out.append(tuple(
            u ** 3 * p0[k] + 3 * u * u * t * p1[k]
            + 3 * u * t * t * p2[k] + t ** 3 * p3[k]
            for k in (0, 1)
        ))
    return out


SHIELD = shield_points()


# Raster helpers ---------------------------------------------------------
def linear_gradient(w, h, c1=C1, c2=C2):
    """135deg linear gradient (top-left -> bottom-right)."""
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    t = (xx + yy) / float(w + h - 2)
    t = t[..., None]
    a = np.array(c1, dtype=np.float32)
    b = np.array(c2, dtype=np.float32)
    return (a + (b - a) * t).astype(np.uint8)


def stroke_path(draw, pts, width, colour, ss=1):
    """Polyline with round joins and round caps."""
    pts = [(x * ss, y * ss) for x, y in pts]
    draw.line(pts, fill=colour, width=int(width * ss), joint="curve")
    r = width * ss / 2
    for x, y in (pts[0], pts[-1]):
        draw.ellipse([x - r, y - r, x + r, y + r], fill=colour)


def render_badge(size, ss=SS):
    """Rounded-square badge: gradient plate + bracket/shield mark."""
    n = VB * ss
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))

    grad = Image.fromarray(linear_gradient(n, n), "RGB").convert("RGBA")

    plate = Image.new("L", (n, n), 0)
    ImageDraw.Draw(plate).rounded_rectangle(
        [0, 0, n - 1, n - 1], radius=RADIUS * ss, fill=255
    )
    img.paste(grad, (0, 0), plate)

    d = ImageDraw.Draw(img)
    stroke_path(d, LEFT_BRACKET, STROKE, WHITE + (255,), ss)
    stroke_path(d, RIGHT_BRACKET, STROKE, WHITE + (255,), ss)
    d.polygon([(x * ss, y * ss) for x, y in SHIELD], fill=WHITE + (255,))

    return img.resize((size, size), Image.LANCZOS)


def render_mark(size, ss=SS, colour=WHITE):
    """Transparent background, mark only."""
    n = VB * ss
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    stroke_path(d, LEFT_BRACKET, STROKE, colour + (255,), ss)
    stroke_path(d, RIGHT_BRACKET, STROKE, colour + (255,), ss)
    d.polygon([(x * ss, y * ss) for x, y in SHIELD], fill=colour + (255,))
    return img.resize((size, size), Image.LANCZOS)


def font(size, bold=True):
    for path in (("C:/Windows/Fonts/segoeuib.ttf", "C:/Windows/Fonts/arialbd.ttf")
                 if bold else
                 ("C:/Windows/Fonts/seguisb.ttf", "C:/Windows/Fonts/arial.ttf")):
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def render_og(w=1200, h=630):
    """Open Graph card: badge + wordmark on the dark site background."""
    ss = 2
    W, H = w * ss, h * ss

    yy, xx = np.mgrid[0:H, 0:W].astype(np.float32)
    r = np.sqrt(((xx - W * 0.26) / (W * 0.52)) ** 2
                + ((yy - H * 0.14) / (H * 0.92)) ** 2)
    glow = np.clip(1.0 - r, 0, 1)[..., None] ** 2
    base = np.zeros((H, W, 3), dtype=np.float32)
    base[:] = np.array((0x0B, 0x10, 0x20), dtype=np.float32) / 255.0
    tint = np.array((0x63, 0x66, 0xF1), dtype=np.float32) / 255.0
    base += (tint - base) * glow * 0.5
    img = Image.fromarray((base * 255).astype(np.uint8), "RGB").convert("RGBA")

    badge = render_badge(300, ss=3)
    bx, by = 104 * ss, (H - 300 * ss) // 2
    img.alpha_composite(badge, (bx, by))

    d = ImageDraw.Draw(img)
    tx = bx + 300 * ss + 62 * ss
    d.text((tx, 289 * ss), "dxxjudges", font=font(80 * ss),
           fill=WHITE + (255,), anchor="lm")
    d.text((tx + 4 * ss, 368 * ss), "NoldOJ  ·  CliJudge",
           font=font(34 * ss, bold=False), fill=(0x9A, 0xA6, 0xC8, 255),
           anchor="lm")

    return img.convert("RGB").resize((w, h), Image.LANCZOS)


def render_lockup(w=1200, h=320):
    """Horizontal lockup on transparent background (for dark surfaces)."""
    ss = 2
    W, H = w * ss, h * ss
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    img.alpha_composite(render_badge(240, ss=3), (0, (H - 240 * ss) // 2))
    d = ImageDraw.Draw(img)
    tx = 240 * ss + 58 * ss
    d.text((tx, 137 * ss), "dxxjudges", font=font(88 * ss),
           fill=WHITE + (255,), anchor="lm")
    d.text((tx + 4 * ss, 216 * ss), "NoldOJ  ·  CliJudge",
           font=font(30 * ss, bold=False), fill=(0x9A, 0xA6, 0xC8, 255),
           anchor="lm")
    return img.resize((w, h), Image.LANCZOS)


# SVG --------------------------------------------------------------------
GRAD_ID = "g"
SVG_BRACKETS = "\n".join(
    "  <path d=\"M%d %d L%d %d L%d %d\" fill=\"none\" stroke=\"#fff\" "
    "stroke-width=\"%d\" stroke-linecap=\"round\" stroke-linejoin=\"round\"/>"
    % (p[0][0], p[0][1], p[1][0], p[1][1], p[2][0], p[2][1], STROKE)
    for p in (LEFT_BRACKET, RIGHT_BRACKET)
)

SHIELD_D = (
    # top edge + right top corner
    "M%d %d H%d a%d %d 0 0 1 %d %d"
    # right side down to the tip, then the left side back up
    " V%d C%d %d %d %d %d %d C%d %d %d %d %d %d"
    # left side back up, then the left top corner
    " V%d a%d %d 0 0 1 %d %d Z"
) % (
    SH_X0 + SH_TOP_R, SH_TOP,
    SH_X1 - SH_TOP_R,
    SH_TOP_R, SH_TOP_R, SH_TOP_R, SH_TOP_R,
    SH_SHOULDER,
    SH_CTRL_R[0], SH_CTRL_R[1], SH_CTRL_R2[0], SH_CTRL_R2[1],
    SH_TIP[0], SH_TIP[1],
    SH_CTRL_L[0], SH_CTRL_L[1], SH_CTRL_L1[0], SH_CTRL_L1[1],
    SH_X0, SH_SHOULDER,
    SH_TOP + SH_TOP_R,
    SH_TOP_R, SH_TOP_R, SH_TOP_R, -SH_TOP_R,
)

SVG_DEFS = (
    '<linearGradient id="%s" x1="0" y1="0" x2="1" y2="1">'
    '<stop offset="0" stop-color="#6366f1"/>'
    '<stop offset="1" stop-color="#22d3ee"/>'
    '</linearGradient>' % GRAD_ID
)

MARK_BODY = SVG_BRACKETS + '\n  <path d="%s" fill="#fff"/>' % SHIELD_D

SVG_BADGE = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %d %d" width="%d" height="%d" role="img" aria-label="dxxjudges">
  <defs>%s</defs>
  <rect width="%d" height="%d" rx="%d" fill="url(#%s)"/>
  <g transform="translate(0 0)">
%s
  </g>
</svg>
""" % (VB, VB, VB, VB, SVG_DEFS, VB, VB, RADIUS, GRAD_ID, MARK_BODY)

SVG_MARK = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %d %d" width="%d" height="%d" role="img" aria-label="dxxjudges">
%s
</svg>
""" % (VB, VB, VB, VB, MARK_BODY)

SVG_LOCKUP = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 320" width="1200" height="320" role="img" aria-label="dxxjudges">
  <defs>%s</defs>
  <rect x="0" y="40" width="240" height="240" rx="60" fill="url(#%s)"/>
  <g transform="translate(0 40) scale(0.234375)">
%s
  </g>
  <text x="298" y="184" font-family="Segoe UI, system-ui, -apple-system, sans-serif" font-size="88" font-weight="700" fill="#e8ecf7">dxxjudges</text>
  <text x="302" y="228" font-family="Segoe UI, system-ui, -apple-system, sans-serif" font-size="30" font-weight="600" fill="#9aa6c8">NoldOJ  &#183;  CliJudge</text>
</svg>
""" % (SVG_DEFS, GRAD_ID, MARK_BODY)


def write(path, data):
    with open(path, "wb") as f:
        f.write(data)
    print("  %-34s %6.1f KB" % (os.path.relpath(path, HERE), len(data) / 1024))


def main():
    os.makedirs(OUT_PNG, exist_ok=True)
    print("vector:")
    write(os.path.join(HERE, "logo.svg"), SVG_BADGE.encode("utf-8"))
    write(os.path.join(HERE, "mark.svg"), SVG_MARK.encode("utf-8"))
    write(os.path.join(HERE, "logo-lockup.svg"), SVG_LOCKUP.encode("utf-8"))

    print("raster:")
    for size in (1024, 512, 256, 128, 64, 48, 32, 16):
        write(os.path.join(OUT_PNG, "logo-%d.png" % size),
              png_bytes(render_badge(size)))

    # favicon.ico (16/32/48) and a dark-bg mark for header use
    ico_src = render_badge(256)
    ico_path = os.path.join(HERE, "favicon.ico")
    ico_src.save(ico_path, sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])
    print("  %-34s %6.1f KB" % ("favicon.ico", os.path.getsize(ico_path) / 1024))

    write(os.path.join(OUT_PNG, "mark-white-256.png"),
          png_bytes(render_mark(256, colour=WHITE)))
    write(os.path.join(OUT_PNG, "mark-indigo-256.png"),
          png_bytes(render_mark(256, colour=C1)))

    write(os.path.join(OUT_PNG, "og-image.png"),
          png_bytes(render_og()))
    write(os.path.join(OUT_PNG, "logo-lockup.png"),
          png_bytes(render_lockup()))


def png_bytes(img):
    import io
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


if __name__ == "__main__":
    main()