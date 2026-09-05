# -*- coding: utf-8 -*-
"""
粒子背景 v3.1 —— 对标参考图三种形态：
  energy_wave : 高密度能量粒子浪（绿色浪脊线 + 前景运动拖尾 + 克制雾光）
  bokeh_wave  : 粒子飘带 + 大光圈散景 + 景深虚化 + 水平镜头光带
  plexus      : 节点互联网络 + 暗绿力场 + 轨道环 + 星芒
全部加法发光合成，深空黑底，荧光绿品牌色 #C8F247 色系。
"""
import os, math, random
from PIL import Image, ImageDraw, ImageFilter, ImageChops

W, H = 2666, 1500
OUT_DIR = r"d:\wwj\_particle_bg"
os.makedirs(OUT_DIR, exist_ok=True)

# 绿 -> 柠 -> 白热 色阶
HOT = (252, 255, 226)
GL  = (220, 248, 130)
G   = (176, 230, 70)
GM  = (112, 182, 54)
GD  = (58, 98, 32)
BG  = (4, 7, 6)


def clamp(v, a=0.0, b=1.0):
    return max(a, min(b, v))


def smooth(t):
    t = clamp(t)
    return t * t * (3 - 2 * t)


def ramp(t):
    """0..1 -> GD -> GM -> G -> GL -> HOT"""
    t = clamp(t)
    stops = [(0.00, GD), (0.30, GM), (0.58, G), (0.80, GL), (1.00, HOT)]
    for i in range(len(stops) - 1):
        t0, c0 = stops[i]
        t1, c1 = stops[i + 1]
        if t <= t1:
            k = smooth((t - t0) / (t1 - t0))
            return tuple(int(c0[j] + (c1[j] - c0[j]) * k) for j in range(3))
    return HOT


def scale_col(c, b):
    return tuple(int(ch * clamp(b)) for ch in c)


def vignette(img, strength=0.86):
    mask = Image.new('L', (W, H), 0)
    md = ImageDraw.Draw(mask)
    md.ellipse([-W * 0.22, -H * 0.28, W * 1.22, H * 1.28], fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(220))
    dark = Image.new('RGB', (W, H), (0, 0, 0))
    return Image.composite(img, dark, mask.point(lambda v: int(255 * (1 - strength) + v * strength)))


def add_glow(layer, rad1=3, rad2=9, mix2=0.55):
    out = layer
    g1 = layer.filter(ImageFilter.GaussianBlur(rad1))
    out = ImageChops.add(out, g1)
    g2 = layer.filter(ImageFilter.GaussianBlur(rad2))
    out = ImageChops.add(out, g2.point(lambda v: int(v * mix2)))
    return out


# ============================================================
# 1) 能量粒子浪
# ============================================================
def energy_wave(seed=1, horizon=0.42, depth=0.60, amp=0.16, density=1.0,
                crest_x=0.62, dim=1.0):
    random.seed(seed * 1000 + 7)
    base = Image.new('RGB', (W, H), BG)
    # 底部环境绿雾
    amb = Image.new('RGB', (W, H), (0, 0, 0))
    ad = ImageDraw.Draw(amb)
    ad.ellipse([W * -0.1, H * (horizon + 0.15), W * 1.1, H * 1.25], fill=(16, 34, 13))
    amb = amb.filter(ImageFilter.GaussianBlur(200))
    base = ImageChops.add(base, amb.point(lambda v: int(v * 0.6)))

    dots = Image.new('RGB', (W, H), (0, 0, 0))
    dr = ImageDraw.Draw(dots)
    streaks = Image.new('RGB', (W, H), (0, 0, 0))
    sd = ImageDraw.Draw(streaks)

    cols = int(260 * density)
    rows = int(110 * density)

    def hfield(nx, ny):
        h = 0.50 * math.sin(nx * 3.0 + seed * 0.9)
        h += 0.25 * math.sin(nx * 6.3 - ny * 2.0 + seed * 1.7)
        h += 0.12 * math.sin(nx * 11.0 + ny * 6.0 + seed * 2.3)
        ridge = math.exp(-((nx - (crest_x + (ny - 0.5) * 0.25)) ** 2) / 0.020)
        h = h * 0.45 + ridge * (0.55 + 0.55 * ny)
        return h

    for r in range(rows):
        ny = r / (rows - 1)                    # 0 远 -> 1 近
        persp = ny ** 1.7
        y_base = H * (horizon + persp * depth)
        scale = 0.30 + ny * 1.9
        row_b = 0.12 + ny * 0.78
        for c in range(cols):
            nx = c / (cols - 1)
            h = hfield(nx, ny)
            x = nx * W + math.sin(ny * 8 + nx * 3 + seed) * 5 * ny
            y = y_base - max(h, -0.2) * H * amp * (0.35 + 0.65 * ny)
            jx = x + random.uniform(-1.8, 1.8) * scale
            jy = y + random.uniform(-1.8, 1.8) * scale
            hh = max(h, 0.0)
            bright = (0.16 + hh * 0.42) * row_b * random.uniform(0.82, 1.12)
            col = ramp(0.22 + hh * 0.42)
            rad = max(0.7, (0.75 + hh * 1.5) * scale * 0.55)
            dr.ellipse([jx - rad, jy - rad, jx + rad, jy + rad],
                       fill=scale_col(col, bright * dim))
            # 仅最近景零星运动拖尾
            if ny > 0.80 and hh > 0.5 and random.random() < 0.12:
                L = random.uniform(14, 46)
                sd.line([(jx, jy + rad), (jx, jy + rad + L)],
                        fill=scale_col(GM, 0.16 * row_b * dim), width=1)

    # 浪脊扫线（招牌亮线）：沿近景地形脊线画发光折线
    ridge_layer = Image.new('RGB', (W, H), (0, 0, 0))
    rd = ImageDraw.Draw(ridge_layer)
    for ny_f, wdt, col, b in [(0.92, 7, G, 0.5), (0.92, 2.5, HOT, 0.7)]:
        pts = []
        for c in range(cols + 1):
            nx = c / cols
            h = hfield(nx, ny_f)
            x = nx * W
            y = H * (horizon + (ny_f ** 1.7) * depth) - max(h, -0.2) * H * amp * (0.35 + 0.65 * ny_f)
            pts.append((x, y))
        rd.line(pts, fill=scale_col(col, b * dim), width=int(wdt), joint='curve')
    ridge_glow = ridge_layer.filter(ImageFilter.GaussianBlur(14))
    ridge_hot = ridge_layer.filter(ImageFilter.GaussianBlur(2))
    dots = ImageChops.add(dots, ridge_glow)
    dots = ImageChops.add(dots, ridge_hot)

    # 远空星点
    for _ in range(140):
        x = random.uniform(0, W); y = random.uniform(0, H * horizon)
        b = random.uniform(0.10, 0.55)
        r = random.uniform(0.6, 1.9)
        dr.ellipse([x - r, y - r, x + r, y + r], fill=scale_col(GL, b))

    streaks = streaks.filter(ImageFilter.GaussianBlur(1.2))
    dots = ImageChops.add(dots, streaks)
    dots = add_glow(dots, 2.5, 8, 0.30)

    img = ImageChops.add(base, dots)
    return vignette(img, 0.90)


# ============================================================
# 2) 散景粒子飘带
# ============================================================
def bokeh_wave(seed=2, band_y=0.62, band_h=0.16, dim=0.9, streaks=False):
    random.seed(seed * 2000 + 13)
    base = Image.new('RGB', (W, H), BG)

    def ribbon_y(nx):
        y = band_y
        y += 0.055 * math.sin(nx * 2.6 + seed * 0.7)
        y += 0.030 * math.sin(nx * 5.7 + seed * 1.5 + 1.2)
        y += 0.014 * math.sin(nx * 11.0 + seed * 2.1)
        return y

    layer = Image.new('RGB', (W, H), (0, 0, 0))
    dr = ImageDraw.Draw(layer)

    # 核心锐粒：沿飘带密集分布（暗绿主体，仅零星高光）
    n_core = 7000
    for _ in range(n_core):
        nx = random.random()
        cy = ribbon_y(nx)
        off = random.gauss(0, band_h * 0.34)
        y = (cy + off) * H
        x = nx * W + random.uniform(-3, 3)
        core = math.exp(-((off / band_h) ** 2) * 3.0)
        hot = 1.0 if random.random() < 0.03 else 0.0   # 零星白热节点
        b = (0.06 + core * 0.34 + hot * 0.45) * random.uniform(0.55, 1.05)
        col = ramp(0.26 + core * 0.45 + hot * 0.4)
        r = random.uniform(0.8, 2.4) * (0.7 + core * 0.7 + hot)
        dr.ellipse([x - r, y - r, x + r, y + r], fill=scale_col(col, b * dim))

    # 中层柔粒
    soft = Image.new('RGB', (W, H), (0, 0, 0))
    sfd = ImageDraw.Draw(soft)
    for _ in range(500):
        nx = random.random()
        cy = ribbon_y(nx)
        off = random.gauss(0, band_h * 0.75)
        y = (cy + off) * H; x = nx * W
        core = math.exp(-((off / band_h) ** 2) * 1.6)
        r = random.uniform(3, 12)
        sfd.ellipse([x - r, y - r, x + r, y + r],
                    fill=scale_col(G, (0.08 + core * 0.28) * dim))
    soft = soft.filter(ImageFilter.GaussianBlur(7))
    layer = ImageChops.add(layer, soft)

    # 飘带雾光（暗）+ 零星亮斑
    haze = Image.new('RGB', (W, H), (0, 0, 0))
    hzd = ImageDraw.Draw(haze)
    for i in range(6):
        cx = (i + 0.5) / 6
        cy = ribbon_y(cx)
        hzd.ellipse([W * (cx - 0.16), H * (cy - band_h * 1.1),
                     W * (cx + 0.16), H * (cy + band_h * 1.1)],
                    fill=scale_col(G, 0.13 * dim))
    for _ in range(5):
        cx = random.uniform(0.1, 0.9); cy = ribbon_y(cx) + random.uniform(-0.03, 0.03)
        hzd.ellipse([W * cx - 90, H * cy - 90, W * cx + 90, H * cy + 90],
                    fill=scale_col(GL, 0.4 * dim))
    haze = haze.filter(ImageFilter.GaussianBlur(90))
    layer = ImageChops.add(layer, haze)

    # 前景大光圈（大而柔，少量）
    bokeh = Image.new('RGB', (W, H), (0, 0, 0))
    bkd = ImageDraw.Draw(bokeh)
    for _ in range(24):
        nx = random.uniform(0.02, 0.98)
        cy = ribbon_y(nx)
        y = (cy + random.uniform(-band_h * 2.0, band_h * 3.0)) * H
        x = nx * W + random.uniform(-50, 50)
        r = random.uniform(30, 130)
        b = random.uniform(0.10, 0.34) * dim
        bkd.ellipse([x - r, y - r, x + r, y + r],
                    fill=scale_col(GL if random.random() < 0.3 else G, b))
    bokeh = bokeh.filter(ImageFilter.GaussianBlur(34))
    layer = ImageChops.add(layer, bokeh)

    # 远空星点
    for _ in range(110):
        x = random.uniform(0, W); y = random.uniform(0, H * (band_y - 0.22))
        b = random.uniform(0.08, 0.5)
        r = random.uniform(0.6, 1.8)
        dr.ellipse([x - r, y - r, x + r, y + r], fill=scale_col(GL, b))

    # 水平镜头光带
    if streaks:
        for sy, wdt, b in [(0.075, 5, 0.30), (0.108, 2, 0.42), (0.93, 4, 0.24)]:
            dr.rectangle([0, H * sy - wdt, W, H * sy + wdt], fill=scale_col(GL, b * dim))

    layer = add_glow(layer, 3, 10, 0.42)
    img = ImageChops.add(base, layer)
    return vignette(img, 0.88)


# ============================================================
# 3) Plexus 节点网络
# ============================================================
def plexus(seed=99, field_x=0.78, field_y=0.45, orbit=True):
    random.seed(seed * 3000 + 31)
    base = Image.new('RGB', (W, H), BG)

    # 右侧暗绿力场（大面积低亮度）
    fld = Image.new('RGB', (W, H), (0, 0, 0))
    fd = ImageDraw.Draw(fld)
    fd.ellipse([W * (field_x - 0.42), H * (field_y - 0.62),
                W * (field_x + 0.42), H * (field_y + 0.62)], fill=(14, 30, 12))
    fld = fld.filter(ImageFilter.GaussianBlur(210))
    base = ImageChops.add(base, fld.point(lambda v: int(v * 0.7)))

    net = Image.new('RGB', (W, H), (0, 0, 0))
    nd = ImageDraw.Draw(net)

    nodes = []
    n = 82
    for i in range(n):
        if random.random() < 0.55:
            x = random.gauss(W * field_x, W * 0.26)
            y = random.gauss(H * field_y, H * 0.28)
        else:
            x = random.uniform(0, W); y = random.uniform(0, H)
        x = clamp(x / W) * W; y = clamp(y / H) * H
        dx = (x / W - field_x); dy = (y / H - field_y)
        glow = clamp(1.0 - math.hypot(dx * 1.2, dy * 1.4) / 0.6)
        nodes.append((x, y, glow))

    # 连线
    maxd = 300
    for i in range(n):
        x1, y1, gl1 = nodes[i]
        for j in range(i + 1, n):
            x2, y2, gl2 = nodes[j]
            d = math.hypot(x1 - x2, y1 - y2)
            if d < maxd:
                a = (1 - d / maxd) * (0.10 + 0.42 * max(gl1, gl2))
                col = ramp(0.32 + 0.4 * max(gl1, gl2))
                nd.line([(x1, y1), (x2, y2)], fill=scale_col(col, a), width=1)

    # 节点
    for x, y, gl in nodes:
        r = 2.0 + gl * 4.5
        nd.ellipse([x - r, y - r, x + r, y + r],
                   fill=scale_col(ramp(0.4 + gl * 0.6), 0.30 + gl * 0.65))

    net = add_glow(net, 4, 12, 0.55)
    img = ImageChops.add(base, net)

    # 轨道环 + 星芒（左上，纤细）
    if orbit:
        deco = Image.new('RGB', (W, H), (0, 0, 0))
        dd = ImageDraw.Draw(deco)
        cx, cy = W * 0.15, H * 0.17
        dd.ellipse([cx - 250, cy - 250 * 0.60, cx + 250, cy + 250 * 0.60],
                   outline=scale_col(GL, 0.32), width=2)
        sx, sy, ss = W * 0.135, H * 0.14, 34
        dd.polygon([(sx, sy - ss), (sx + ss * 0.22, sy - ss * 0.22), (sx + ss, sy),
                    (sx + ss * 0.22, sy + ss * 0.22), (sx, sy + ss),
                    (sx - ss * 0.22, sy + ss * 0.22), (sx - ss, sy),
                    (sx - ss * 0.22, sy - ss * 0.22)], fill=scale_col(HOT, 0.85))
        deco = deco.filter(ImageFilter.GaussianBlur(1.4))
        img = ImageChops.add(img, deco)

    return vignette(img, 0.88)


# ============================================================
# 页面分配
# ============================================================
PAGES = [
    ("bg_01_cover.png",    "energy", dict(seed=1,  horizon=0.42, depth=0.60, amp=0.17, crest_x=0.66, dim=1.0)),
    ("bg_02_profile.png",  "bokeh",  dict(seed=2,  band_y=0.80, band_h=0.13, dim=0.80)),
    ("bg_03_timeline.png", "bokeh",  dict(seed=3,  band_y=0.74, band_h=0.12, dim=0.75)),
    ("bg_04_exp.png",      "bokeh",  dict(seed=4,  band_y=0.82, band_h=0.11, dim=0.70)),
    ("bg_05_certs1.png",   "bokeh",  dict(seed=5,  band_y=0.87, band_h=0.09, dim=0.62)),
    ("bg_06_certs2.png",   "bokeh",  dict(seed=6,  band_y=0.87, band_h=0.09, dim=0.62)),
    ("bg_07_plane.png",    "energy", dict(seed=7,  horizon=0.50, depth=0.56, amp=0.15, crest_x=0.30, dim=0.90)),
    ("bg_08_nebula.png",   "energy", dict(seed=8,  horizon=0.38, depth=0.64, amp=0.20, crest_x=0.55, dim=1.0)),
    ("bg_09_gesture.png",  "bokeh",  dict(seed=9,  band_y=0.80, band_h=0.12, dim=0.75)),
    ("bg_10_website.png",  "bokeh",  dict(seed=10, band_y=0.84, band_h=0.10, dim=0.66)),
    ("bg_11_seedmusic.png","bokeh",  dict(seed=11, band_y=0.52, band_h=0.14, dim=0.85, streaks=True)),
    ("bg_12_roadmap.png",  "energy", dict(seed=12, horizon=0.68, depth=0.40, amp=0.11, crest_x=0.5, dim=0.80)),
    ("bg_13_promise.png",  "bokeh",  dict(seed=13, band_y=0.78, band_h=0.12, dim=0.75)),
    ("bg_14_closing.png",  "energy", dict(seed=14, horizon=0.46, depth=0.60, amp=0.18, crest_x=0.5, dim=0.95)),
    ("bg_15_thanks.png",   "plexus", dict(seed=15)),
]

if __name__ == "__main__":
    print("生成粒子背景 x15 (v3.1) ...")
    for fn, kind, kw in PAGES:
        if kind == "energy":
            img = energy_wave(**kw)
        elif kind == "bokeh":
            img = bokeh_wave(**kw)
        else:
            img = plexus(**kw)
        img.save(os.path.join(OUT_DIR, fn), optimize=True)
        print("  OK", fn)
    print("完成 ->", OUT_DIR)
