# -*- coding: utf-8 -*-
"""
王文杰 PPT 4.0 —— Tech Portfolio（个人科技影响力展示）
设计语言：深空黑 + 荧光绿 #C8F247 + 粒子能量场背景 + Bento 玻璃拟态
对标参考：近黑卡片 / 发丝边框 / 柔光晕 / 克制圆角 / 数据可视化 / 无廉价划线
"""
import os, math, random
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image as PILImage

# === 配色 ===
G   = RGBColor(0xC8, 0xF2, 0x47)   # 荧光绿
GL  = RGBColor(0xE3, 0xF7, 0xA8)   # 亮柠
GD  = RGBColor(0x6E, 0x8F, 0x2E)   # 暗绿
W   = RGBColor(0xFF, 0xFF, 0xFF)
GW  = RGBColor(0x9A, 0xA3, 0x99)   # 正文灰
GS  = RGBColor(0x62, 0x6B, 0x60)   # 弱灰
BG  = RGBColor(0x07, 0x09, 0x08)
CARD  = RGBColor(0x0C, 0x11, 0x0E) # 卡片底
CARD2 = RGBColor(0x14, 0x1B, 0x14) # 浅卡片/磁贴
HAIR  = RGBColor(0x9C, 0xC0, 0x86) # 发丝边框

CN, LAT, MONO, NUM = "微软雅黑", "Segoe UI", "Consolas", "Bahnschrift"
ICON = "Segoe MDL2 Assets"

# MDL2 单色线性图标码位
IC = {"home": "\uE80F", "user": "\uE77B", "people": "\uE716",
      "heart": "\uEB51", "laptop": "\uE7F8", "target": "\uE759",
      "refresh": "\uE777", "bulb": "\uEA80", "bolt": "\uE945",
      "star_o": "\uE734", "star_f": "\uE735", "cap": "\uE7BE",
      "code": "\uE943", "cloud": "\uE753", "terminal": "\uE756",
      "plane": "\uE709", "globe": "\uE774", "eye": "\uE7B3",
      "hand": "\uE815", "music": "\uE8D6", "note": "\uE189",
      "headset": "\uE95B", "link": "\uE71B", "play": "\uE768",
      "check": "\uE73E", "devices": "\uE977"}

IMG = r"d:\wwj\_ppt_extracted\ppt\media"
PBG = r"d:\wwj\_particle_bg"
OUT = r"d:\wwj\王文杰-团支书竞选-3.0科技最终版.pptx"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BL = prs.slide_layouts[6]


# ============================================================
# 底层工具
# ============================================================
def _spPr(sh):
    return sh._element.find(qn('p:spPr'))


def txt(s, l, t, w, h, text, size=14, color=W, bold=False, italic=False,
        font=CN, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=None):
    tb = s.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(0); tf.margin_right = Pt(0)
    tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
        r.font.color.rgb = color
        rPr = r._r.get_or_add_rPr()
        if spacing:
            rPr.set('spc', str(spacing))
        has_cjk = any('一' <= ch <= '鿿' for ch in line)
        lt = rPr.find(qn('a:latin'))
        if lt is None:
            lt = etree.SubElement(rPr, qn('a:latin'))
        lt.set('typeface', font)
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = etree.SubElement(rPr, qn('a:ea'))
        ea.set('typeface', CN if has_cjk else font)
    return tb


def _text_alpha(tb, alpha):
    """设置文本框内所有文字的透明度"""
    for rPr in tb._element.iter(qn('a:rPr')):
        sf = rPr.find(qn('a:solidFill'))
        if sf is None:
            continue
        c = sf.find(qn('a:srgbClr'))
        if c is not None:
            a = c.find(qn('a:alpha'))
            if a is None:
                a = etree.SubElement(c, qn('a:alpha'))
            a.set('val', str(alpha))


def _fill_alpha(sh, alpha):
    sp = _spPr(sh)
    c = sp.find(qn('a:solidFill') + '/' + qn('a:srgbClr'))
    if c is not None:
        a = c.find(qn('a:alpha'))
        if a is None:
            a = etree.SubElement(c, qn('a:alpha'))
        a.set('val', str(alpha))


def _line_alpha(sh, alpha):
    ln = _spPr(sh).find(qn('a:ln'))
    if ln is not None:
        c = ln.find(qn('a:solidFill') + '/' + qn('a:srgbClr'))
        if c is not None:
            etree.SubElement(c, qn('a:alpha')).set('val', str(alpha))


def _round(sh, radius_in, w_in, h_in):
    sp = _spPr(sh)
    pg = sp.find(qn('a:prstGeom'))
    pg.set('prst', 'roundRect')
    av = pg.find(qn('a:avLst'))
    if av is None:
        av = etree.SubElement(pg, qn('a:avLst'))
    for g in list(av):
        av.remove(g)
    adj = min(radius_in / min(w_in, h_in), 0.5)
    gd = etree.SubElement(av, qn('a:gd'))
    gd.set('name', 'adj'); gd.set('fmla', 'val %d' % int(adj * 100000))


def _glow(sh, rad=190500, alpha=12000, color='C8F247'):
    """柔光晕：大半径外阴影当环境光"""
    sp = _spPr(sh)
    eL = sp.find(qn('a:effectLst'))
    if eL is None:
        eL = etree.SubElement(sp, qn('a:effectLst'))
    shd = etree.SubElement(eL, qn('a:outerShdw'))
    shd.set('blurRad', str(rad)); shd.set('dist', '0')
    shd.set('dir', '0'); shd.set('rotWithShape', '0')
    c = etree.SubElement(shd, qn('a:srgbClr')); c.set('val', color)
    etree.SubElement(c, qn('a:alpha')).set('val', str(alpha))


def _gradient(sh, c1, c2, angle=90):
    """竖向渐变填充（c1->c2 hex strings）"""
    sp = _spPr(sh)
    sf = sp.find(qn('a:solidFill'))
    if sf is not None:
        sp.remove(sf)
    gf = etree.Element(qn('a:gradFill'))
    gsLst = etree.SubElement(gf, qn('a:gsLst'))
    for pos, col in ((0, c1), (100000, c2)):
        gs = etree.SubElement(gsLst, qn('a:gs')); gs.set('pos', str(pos))
        c = etree.SubElement(gs, qn('a:srgbClr')); c.set('val', col)
    lin = etree.SubElement(gf, qn('a:lin'))
    lin.set('ang', str(int(angle * 60000))); lin.set('scaled', '1')
    ln = sp.find(qn('a:ln'))
    if ln is not None:
        ln.addprevious(gf)
    else:
        pg = sp.find(qn('a:prstGeom'))
        pg.addnext(gf)


# ============================================================
# 组件库
# ============================================================
def card(s, l, t, w, h, fill=CARD, radius=0.08, border=True,
         b_alpha=9000, glow=11000, fill_alpha=90000):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.shadow.inherit = False
    _round(sh, radius, w / 914400, h / 914400)
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    _fill_alpha(sh, fill_alpha)
    if border:
        sh.line.color.rgb = HAIR; sh.line.width = Pt(0.75)
        _line_alpha(sh, b_alpha)
    else:
        sh.line.fill.background()
    if glow:
        _glow(sh, 200000, glow)
    return sh


def accent_card(s, l, t, w, h, radius=0.08):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.shadow.inherit = False
    _round(sh, radius, w / 914400, h / 914400)
    sh.fill.solid(); sh.fill.fore_color.rgb = G
    sh.line.fill.background()
    _glow(sh, 220000, 26000)
    return sh


def icon_tile(s, x, y, size=0.52, glyph='', gsize=20):
    card(s, Inches(x), Inches(y), Inches(size), Inches(size),
         fill=CARD2, radius=size * 0.22, b_alpha=15000, glow=7000)
    if glyph:
        txt(s, Inches(x), Inches(y - 0.02), Inches(size), Inches(size),
            glyph, size=gsize, color=GL, font=ICON,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def tag(s, x, y, text, w=1.5, h=0.34, color=GL):
    card(s, Inches(x), Inches(y), Inches(w), Inches(h),
         fill=CARD2, radius=h / 2, b_alpha=18000, glow=0, fill_alpha=100000)
    txt(s, Inches(x), Inches(y), Inches(w), Inches(h), text,
        size=9.5, color=color, bold=True, font=MONO,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def stat(s, x, y, w, h, label, value, accent=False, v_size=26, v_color=None):
    if accent:
        accent_card(s, Inches(x), Inches(y), Inches(w), Inches(h), radius=0.08)
        lc = RGBColor(0x33, 0x42, 0x14); vc = BG
    else:
        card(s, Inches(x), Inches(y), Inches(w), Inches(h))
        lc = GW; vc = v_color or W
    txt(s, Inches(x + 0.22), Inches(y + 0.16), Inches(w - 0.4), Inches(0.26),
        label, size=9, color=lc, bold=True, font=MONO)
    has_cjk = any('\u4e00' <= ch <= '\u9fff' for ch in value)
    txt(s, Inches(x + 0.22), Inches(y + 0.40), Inches(w - 0.4), Inches(h - 0.5),
        value, size=v_size, color=vc, bold=True, font=CN if has_cjk else NUM,
        anchor=MSO_ANCHOR.MIDDLE)


def dot(s, x, y, r=0.055, color=G, glow=16000):
    d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - r), Inches(y - r),
                           Inches(r * 2), Inches(r * 2))
    d.fill.solid(); d.fill.fore_color.rgb = color
    d.line.fill.background(); d.shadow.inherit = False
    if glow:
        _glow(d, 60000, glow)
    return d


def bullet(s, x, y, text, size=12.5, color=W, w=5.0, icon='dot', icon_col=GL):
    if icon == 'check':
        txt(s, Inches(x), Inches(y - 0.04), Inches(0.3), Inches(0.35),
            '✓', size=12, color=icon_col, bold=True, font=LAT)
    else:
        dot(s, x + 0.08, y + 0.13, r=0.045, color=icon_col, glow=10000)
    txt(s, Inches(x + 0.32), Inches(y), Inches(w), Inches(0.5),
        text, size=size, color=color)


def hline(s, x1, x2, y, color=GD, dash=True, alpha=38000, width=1.0):
    cn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                Inches(x1), Inches(y), Inches(x2), Inches(y))
    cn.shadow.inherit = False
    cn.line.color.rgb = color; cn.line.width = Pt(width)
    _line_alpha(cn, alpha)
    if dash:
        ln = cn._element.find(qn('p:spPr') + '/' + qn('a:ln'))
        d = etree.SubElement(ln, qn('a:prstDash')); d.set('val', 'sysDash')
    return cn


def page_no(s, page):
    tb = s.shapes.add_textbox(Inches(11.55), Inches(0.5), Inches(1.35), Inches(0.36))
    tf = tb.text_frame
    tf.margin_left = Pt(0); tf.margin_right = Pt(0); tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r1 = p.add_run(); r1.text = '%02d' % page
    r1.font.size = Pt(13); r1.font.bold = True; r1.font.color.rgb = GL; r1.font.name = NUM
    r2 = p.add_run(); r2.text = ' / 15'
    r2.font.size = Pt(10); r2.font.color.rgb = GS; r2.font.name = MONO


def header(s, tag, title, page):
    txt(s, Inches(0.72), Inches(0.5), Inches(9.5), Inches(0.3),
        tag, size=11, color=GL, bold=True, font=MONO, spacing=160)
    if title:
        txt(s, Inches(0.7), Inches(0.84), Inches(10.5), Inches(0.62),
            title, size=29, color=W, bold=True)
    page_no(s, page)


def watermark_num(s, x, y, num, size=66):
    tb = txt(s, Inches(x), Inches(y), Inches(1.6), Inches(1.2),
             '%02d' % num, size=size, color=G, bold=True, font=NUM,
             align=PP_ALIGN.RIGHT)
    _text_alpha(tb, 12000)


def cover_pic(s, path, l, t, w, h, radius=0.08):
    """cover 填充截图（裁切填满），外层为深色框"""
    card(s, l, t, w, h, fill=RGBColor(0x05, 0x08, 0x07), radius=radius,
         b_alpha=12000, glow=14000, fill_alpha=100000)
    if not path or not os.path.exists(path):
        return
    iw, ih = PILImage.open(path).size
    pad = Inches(0.12)
    bx, by = l + pad, t + pad
    bw, bh = w - pad * 2, h - pad * 2
    pic = s.shapes.add_picture(path, bx, by, bw, bh)
    tar = (bw / 914400) / (bh / 914400); src = iw / ih
    if src > tar:
        c = (1 - tar / src) / 2
        pic.crop_left = c; pic.crop_right = c
    else:
        c = (1 - src / tar) / 2
        pic.crop_top = c; pic.crop_bottom = c
    pic.line.fill.background(); pic.shadow.inherit = False
    _round(pic, radius * 0.8, bw / 914400, bh / 914400)


def contain_pic(s, path, l, t, w, h):
    """contain 适配（完整显示证书）"""
    if not path or not os.path.exists(path):
        return
    iw, ih = PILImage.open(path).size
    tar = (w / 914400) / (h / 914400); src = iw / ih
    if src > tar:
        nw = w; nh = int(w / src); nx = l; ny = t + (h - nh) // 2
    else:
        nh = h; nw = int(h * src); ny = t; nx = l + (w - nw) // 2
    pic = s.shapes.add_picture(path, nx, ny, nw, nh)
    pic.line.fill.background(); pic.shadow.inherit = False


def skill_bar(s, x, y, w, label, pct, color=GL):
    txt(s, Inches(x), Inches(y - 0.02), Inches(1.25), Inches(0.3),
        label, size=10, color=GW, font=MONO)
    bx, bw = x + 1.35, w - 1.35
    card(s, Inches(bx), Inches(y + 0.04), Inches(bw), Inches(0.10),
         fill=CARD2, radius=0.05, border=False, glow=0, fill_alpha=100000)
    fillw = max(0.12, bw * pct)
    f = card(s, Inches(bx), Inches(y + 0.04), Inches(fillw), Inches(0.10),
             fill=color, radius=0.05, border=False, glow=9000, fill_alpha=100000)
    _gradient(f, '7AB832', 'C8F247', angle=0)


def waveform(s, x, y, w, h, n=32, seed=11):
    random.seed(seed)
    gap = w / n
    bw = gap * 0.42
    for i in range(n):
        hh = 0.12 + 0.88 * abs(
            0.55 * math.sin(i * 0.55 + seed) +
            0.35 * math.sin(i * 1.3 + 2) + random.uniform(-0.15, 0.15))
        hh = max(0.08, min(1.0, hh))
        bh = Inches(h * hh)
        bar = card(s, Inches(x + i * gap), Inches(y + h - h * hh),
                   Inches(bw), bh,
                   fill=G if i % 3 else GL, radius=bw / 2,
                   border=False, glow=0, fill_alpha=100000)
        if i % 3 == 0:
            _gradient(bar, 'E3F7A8', 'C8F247', 90)


def _set_slide_bg(s, hex_color='070908'):
    """给幻灯片铺近黑实底，防止转场瞬间图片未渲染而闪白"""
    cSld = s._element.find(qn('p:cSld'))
    old = cSld.find(qn('p:bg'))
    if old is not None:
        cSld.remove(old)
    bg = etree.Element(qn('p:bg'))
    bgPr = etree.SubElement(bg, qn('p:bgPr'))
    fill = etree.SubElement(bgPr, qn('a:solidFill'))
    clr = etree.SubElement(fill, qn('a:srgbClr')); clr.set('val', hex_color)
    etree.SubElement(bgPr, qn('a:effectLst'))
    cSld.insert(0, bg)  # p:bg 必须是 cSld 的第一个子元素


def bg_scene(s, bg_name):
    _set_slide_bg(s)
    p = os.path.join(PBG, bg_name)
    if os.path.exists(p):
        s.shapes.add_picture(p, 0, 0, Inches(13.333), Inches(7.5))


# ============================================================
# 15 页
# ============================================================
def build_cover():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_01_cover.png")
    txt(s, Inches(0.72), Inches(0.52), Inches(9), Inches(0.3),
        "CAMPAIGN SPEECH  ·  AI PORTFOLIO  ·  2026.09",
        size=11, color=GL, bold=True, font=MONO, spacing=160)
    page_no(s, 1)
    txt(s, Inches(0.7), Inches(1.42), Inches(12), Inches(2.0),
        "王文杰", size=104, color=W, bold=True)
    txt(s, Inches(0.75), Inches(3.55), Inches(11), Inches(0.6),
        "团支书  ·  以责任之名，赴青春之约", size=24, color=GL, bold=True)
    txt(s, Inches(0.75), Inches(4.25), Inches(11), Inches(0.6),
        '“不是站在台上才发光，而是发光的人才敢站在台上。”',
        size=14, color=GW, italic=True)
    stats = [("PROJECTS", "05", False), ("CERTIFICATES", "04", False),
             ("SKILLS", "AI + WEB", False), ("POSITION", "团支书", True)]
    for i, (lb, vl, ac) in enumerate(stats):
        stat(s, 0.75 + i * 3.02, 5.95, 2.82, 1.02, lb, vl, accent=ac,
             v_size=24 if len(vl) <= 3 else 19)


def build_profile():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_02_profile.png")
    header(s, "PROFILE  ·  个人档案", "个人介绍 · 我是谁", 2)
    # 左：身份卡
    card(s, Inches(0.7), Inches(1.75), Inches(4.55), Inches(5.05))
    accent_card(s, Inches(1.0), Inches(2.05), Inches(0.95), Inches(0.95), radius=0.22)
    txt(s, Inches(1.0), Inches(2.05), Inches(0.95), Inches(0.95), "王",
        size=40, color=BG, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(2.15), Inches(2.12), Inches(2.8), Inches(0.5), "王文杰",
        size=26, color=W, bold=True)
    txt(s, Inches(2.15), Inches(2.66), Inches(2.8), Inches(0.3), "WANG WENJIE",
        size=10, color=GW, font=MONO)
    infos = [(IC["home"], "家乡", "山东潍坊"), (IC["user"], "性格", "温和沉稳，乐于助人"),
             (IC["music"], "爱好", "羽毛球 / AI / 音乐"), (IC["heart"], "初心", "用行动服务集体")]
    for i, (ic, k, v) in enumerate(infos):
        y = 3.35 + i * 0.82
        icon_tile(s, 1.0, y, 0.5, ic, 17)
        txt(s, Inches(1.7), Inches(y + 0.02), Inches(3.2), Inches(0.28),
            k, size=11, color=GW)
        txt(s, Inches(1.7), Inches(y + 0.28), Inches(3.4), Inches(0.34),
            v, size=14.5, color=W, bold=True)
    # 右：核心素养 2x2
    txt(s, Inches(5.55), Inches(1.82), Inches(6), Inches(0.3),
        "核心素养  ·  CORE STRENGTHS", size=12, color=GL, bold=True, font=MONO)
    quals = [(IC["laptop"], "技术驱动", "AI 全栈自学，4 项权威认证\n课程到实战的完整闭环"),
             (IC["people"], "责任担当", "高中团支书实战经验\n组织活动零失误记录"),
             (IC["target"], "执行力", "5 个完整落地项目\n从想法到上线全流程"),
             (IC["refresh"], "成长型", "从内向到破局的蜕变\n敢站上更高的舞台")]
    for i, (ic, k, v) in enumerate(quals):
        cx = 5.55 + (i % 2) * 3.55; cy = 2.3 + (i // 2) * 2.35
        card(s, Inches(cx), Inches(cy), Inches(3.35), Inches(2.15))
        watermark_num(s, cx + 1.62, cy + 0.16, i + 1, size=44)
        icon_tile(s, cx + 0.28, cy + 0.28, 0.55, ic, 20)
        txt(s, Inches(cx + 0.28), Inches(cy + 1.02), Inches(2.8), Inches(0.4),
            k, size=17, color=W, bold=True)
        txt(s, Inches(cx + 0.28), Inches(cy + 1.46), Inches(2.85), Inches(0.65),
            v, size=11, color=GW)


def build_timeline():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_03_timeline.png")
    header(s, "JOURNEY  ·  成长轨迹", "四阶段蜕变 · 从内向到担当", 3)
    stages = [("2021", IC["bulb"], "初中 · 埋下种子", "内向寡言，畏惧公开发言",
               "第一次班会发言，声音都在抖", "萌芽"),
              ("2024", IC["bolt"], "高中 · 勇敢破局", "竞选并担任团支书",
               "组织团日活动，收获同学信任", "破局"),
              ("2025", IC["star_o"], "坚持 · 磨练意志", "羽毛球训练，培养韧性",
               "班级事务零失误，广受好评", "沉淀"),
              ("2026", IC["star_f"], "此刻 · 发光担当", "怀揣 AI 梦想，站上更高舞台",
               "用技术服务班级，继续发光", "发光")]
    y_line = 2.42
    hline(s, 1.15, 12.2, y_line, color=GD, dash=True, alpha=30000)
    for i, (yr, ic, title, desc, detail, kw) in enumerate(stages):
        x = 0.75 + i * 3.08
        cx = x + 1.45
        dot(s, cx, y_line, r=0.075, color=G, glow=22000)
        card(s, Inches(x), Inches(2.85), Inches(2.9), Inches(3.5))
        watermark_num(s, x + 1.18, 5.28, i + 1, size=46)
        txt(s, Inches(x + 0.3), Inches(3.12), Inches(2.3), Inches(0.6),
            yr, size=30, color=GL, bold=True, font=NUM)
        icon_tile(s, x + 0.3, 3.82, 0.55, ic, 20)
        txt(s, Inches(x + 0.3), Inches(4.55), Inches(2.4), Inches(0.4),
            title, size=15.5, color=W, bold=True)
        txt(s, Inches(x + 0.3), Inches(5.06), Inches(2.45), Inches(0.35),
            desc, size=11.5, color=GW)
        txt(s, Inches(x + 0.3), Inches(5.46), Inches(2.45), Inches(0.35),
            detail, size=10.5, color=GS, italic=True)
        tag(s, x + 0.3, 5.88, "· " + kw + " ·", w=0.95, h=0.3)


def build_exp():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_04_exp.png")
    header(s, "EXPERIENCE & LEARNING  ·  经验与学习", "责任担当 × 技术驱动", 4)
    # 左：责任
    card(s, Inches(0.7), Inches(1.75), Inches(5.9), Inches(5.05))
    icon_tile(s, 1.0, 2.08, 0.55, IC["people"], 20)
    txt(s, Inches(1.72), Inches(2.06), Inches(4.5), Inches(0.4),
        "责任担当", size=19, color=W, bold=True)
    txt(s, Inches(1.72), Inches(2.5), Inches(4.5), Inches(0.3),
        "RESPONSIBILITY  ·  高中团支书", size=9.5, color=GL, font=MONO)
    for i, t in enumerate(["组织班级团日活动，协调同学分工",
                           "搭建师生沟通桥梁，及时传达反馈",
                           "管理班级通知体系，响应零延迟",
                           "在压力下保持条理与耐心"]):
        bullet(s, 1.0, 3.15 + i * 0.72, t, size=13, w=5.2, icon='check')
    # 右：技术
    card(s, Inches(6.75), Inches(1.75), Inches(5.9), Inches(5.05))
    icon_tile(s, 7.05, 2.08, 0.55, IC["laptop"], 20)
    txt(s, Inches(7.77), Inches(2.06), Inches(4.5), Inches(0.4),
        "技术驱动", size=19, color=W, bold=True)
    txt(s, Inches(7.77), Inches(2.5), Inches(4.5), Inches(0.3),
        "TECH-DRIVEN  ·  AI 自学之路", size=9.5, color=GL, font=MONO)
    learns = [("01", "AI 大学堂", "四大工程师认证全通关"),
              ("02", "阿里云 Clouder", "大模型应用构建与优化"),
              ("03", "Datawhale", "开源社区协作实战"),
              ("04", "阿里云达摩院", "人工智能训练师双认证"),
              ("05", "Trae AI", "AI 辅助编码实战")]
    for i, (idx, k, v) in enumerate(learns):
        y = 3.0 + i * 0.5
        txt(s, Inches(7.05), Inches(y), Inches(0.5), Inches(0.35),
            idx, size=12, color=GL, bold=True, font=MONO)
        txt(s, Inches(7.55), Inches(y - 0.02), Inches(2.0), Inches(0.35),
            k, size=13, color=W, bold=True)
        txt(s, Inches(9.5), Inches(y), Inches(3.0), Inches(0.35),
            v, size=11, color=GW)
    txt(s, Inches(7.05), Inches(5.66), Inches(4), Inches(0.3),
        "SKILL MATRIX  ·  能力矩阵", size=9.5, color=GL, bold=True, font=MONO)
    for i, (lb, pc) in enumerate([("Python", 0.82), ("Prompt", 0.88),
                                  ("Cloud", 0.62), ("Vision", 0.70)]):
        skill_bar(s, 7.05, 6.02 + i * 0.185, 5.3, lb, pc)


def _cert_page(page, entries, title):
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_%02d_certs%d.png" % (page, 1 if page == 5 else 2))
    header(s, "CERTIFICATION  ·  权威认证", title, page)
    for i, (ic, name, tagv, imgfile, desc) in enumerate(entries):
        x = 0.75 + i * 6.15
        card(s, Inches(x), Inches(1.8), Inches(5.85), Inches(5.0))
        icon_tile(s, x + 0.3, 2.1, 0.55, ic, 20)
        txt(s, Inches(x + 1.0), Inches(2.16), Inches(3.2), Inches(0.45),
            name, size=18, color=W, bold=True)
        tag(s, x + 4.35, 2.18, "✓ 已认证", w=1.2, h=0.36)
        card(s, Inches(x + 0.3), Inches(2.95), Inches(5.25), Inches(2.75),
             fill=RGBColor(0x05, 0x08, 0x07), radius=0.07, b_alpha=10000,
             glow=0, fill_alpha=100000)
        contain_pic(s, os.path.join(IMG, imgfile),
                    Inches(x + 0.45), Inches(3.08), Inches(4.95), Inches(2.5))
        txt(s, Inches(x + 0.3), Inches(5.95), Inches(4), Inches(0.3),
            tagv, size=10, color=GL, bold=True, font=MONO)
        txt(s, Inches(x + 0.3), Inches(6.28), Inches(5.3), Inches(0.45),
            desc, size=11, color=GW)


def build_certs1():
    _cert_page(5, [(IC["cap"], "AI 大学堂", "AI DAXUE · 4 CERTS", "image7.png",
                    "RAG / 微调 / 智能体 / Prompt 工程师，四大工程师认证全通关。"),
                   (IC["cloud"], "阿里云 Clouder", "ALIYUN CLOUD CERT", "image6.png",
                    "大模型 Clouder 认证：RAG 应用构建、VISION 设计、Spring AI 等多项通过。")],
               "AI 认证 · 平台课程")


def build_certs2():
    _cert_page(6, [(IC["terminal"], "Datawhale", "OPEN SOURCE COMMUNITY", "image10.png",
                    "Prompt / Agent / LLM / Fine-tuning 多方向认证，开源协作实战。"),
                   (IC["bulb"], "阿里云达摩院", "DAMO ACADEMY", "image9.png",
                    "人工智能训练师（初级 + 高级）双认证，AI 训练与数据能力证明。")],
               "AI 认证 · 开源 + 产业")


def _project_page(page, pnum, tagv, title, imgfile, tech, info, caption):
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_%02d_%s.png" % (page, tagv.split('_')[1]))
    header(s, "PROJECT %02d / 05  ·  %s" % (pnum, tagv.split('_')[1].upper()),
           title, page)
    cover_pic(s, os.path.join(IMG, imgfile),
              Inches(0.7), Inches(1.75), Inches(7.55), Inches(4.5))
    card(s, Inches(0.78), Inches(6.46), Inches(0.055), Inches(0.24),
         fill=G, radius=0.02, border=False, glow=20000, fill_alpha=100000)
    txt(s, Inches(0.95), Inches(6.42), Inches(7), Inches(0.3),
        caption, size=10, color=GL, bold=True, font=MONO)
    # 右：信息卡
    card(s, Inches(8.5), Inches(1.75), Inches(4.15), Inches(5.05))
    watermark_num(s, 10.95, 1.85, pnum, size=60)
    txt(s, Inches(8.8), Inches(2.05), Inches(3.6), Inches(0.3),
        "PROJECT INFO", size=9.5, color=GW, bold=True, font=MONO)
    txt(s, Inches(8.8), Inches(2.36), Inches(3.6), Inches(0.75),
        title, size=19, color=W, bold=True)
    for i, (k, v) in enumerate(info):
        y = 3.2 + i * 0.56
        txt(s, Inches(8.8), Inches(y), Inches(1.5), Inches(0.3),
            k, size=10.5, color=GW)
        txt(s, Inches(8.8), Inches(y + 0.24), Inches(3.6), Inches(0.34),
            v, size=13, color=W, bold=True)
    txt(s, Inches(8.8), Inches(5.62), Inches(3.5), Inches(0.3),
        "TECH STACK", size=9.5, color=GW, bold=True, font=MONO)
    # 技术栈标签
    for i, t in enumerate(tech):
        tx = 8.8 + (i % 2) * 1.85
        ty = 5.98 + (i // 2) * 0.4
        tag(s, tx, ty, t, w=1.72, h=0.32)


def build_plane():
    _project_page(7, 1, "PROJECT_PLANE", "飞机大战 · Trae AI 辅助开发", "image11.png",
                  ["Python", "Pygame", "AI 辅助", "Trae"],
                  [("开发平台", "Trae AI + Python"), ("游戏模式", "单人 / 双人 / PK"),
                   ("核心特色", "Boss 战 · 波次设计"), ("项目收获", "AI 协作开发全流程实践")],
                  "SKYLINE DEFENDER · PYGAME")


def build_nebula():
    _project_page(8, 2, "PROJECT_NEBULA", "粒子星云 · 物理引擎可视化", "image14.png",
                  ["Python", "Particle", "Physics", "Visual"],
                  [("项目类型", "物理粒子系统"), ("核心能力", "粒子轨迹仿真"),
                   ("项目亮点", "每颗粒子独立轨迹"), ("项目收获", "物理建模与渲染性能调优")],
                  "PARTICLE SYSTEM · PHYSICS ENGINE")


def build_gesture():
    _project_page(9, 3, "PROJECT_GESTURE", "手势切水果 · AI Vision", "image17.png",
                  ["MediaPipe", "OpenCV", "AI Vision", "Real-time"],
                  [("核心技术", "MediaPipe 手部捕捉"), ("交互方式", "实时手势识别"),
                   ("玩法亮点", "碰撞检测切水果"), ("项目收获", "视觉算法实时工程化落地")],
                  "AI VISION PIPELINE · MEDIAPIPE")


def build_website():
    _project_page(10, 4, "PROJECT_WEBSITE", "个人简历网站 · 全流程能力", "image18.png",
                  ["HTML/CSS", "JavaScript", "Deploy", "Responsive"],
                  [("开发能力", "全栈开发与部署"), ("设计细节", "深色主题 · 响应式"),
                   ("性能优化", "高性能加载适配"), ("项目收获", "设计到上线独立闭环")],
                  "FULL-STACK · RESPONSIVE DESIGN")


def build_seedmusic():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_11_seedmusic.png")
    header(s, "PROJECT 05 / 05  ·  AI MUSIC", "AI 制作音乐 · SeedMusic 1.0", 11)
    txt(s, Inches(0.72), Inches(1.95), Inches(5.9), Inches(1.3),
        "用 AI 辅助写歌、生成旋律，\n探索技术与艺术的交汇点，\n让每个人都能成为创作者。",
        size=14, color=GW)
    txt(s, Inches(0.72), Inches(3.08), Inches(5.9), Inches(0.3),
        "从 v1.0 到 v1.2 持续迭代 · 这首歌唱给全班听",
        size=11, color=GL, italic=True, bold=True)
    feats = [(IC["note"], "歌词创作", "AI 辅助写歌"), (IC["headset"], "风格控制", "国风 / 电子 / R&B"),
             (IC["refresh"], "多版本管理", "持续生成歌曲"), (IC["bulb"], "AI 创意实践", "生成式 AI 创作")]
    for i, (ic, t, d) in enumerate(feats):
        cx = 0.72 + (i % 2) * 3.02; cy = 3.55 + (i // 2) * 1.5
        card(s, Inches(cx), Inches(cy), Inches(2.86), Inches(1.32))
        icon_tile(s, cx + 0.22, cy + 0.32, 0.5, ic, 17)
        txt(s, Inches(cx + 0.9), Inches(cy + 0.3), Inches(1.9), Inches(0.4),
            t, size=14.5, color=W, bold=True)
        txt(s, Inches(cx + 0.9), Inches(cy + 0.72), Inches(1.9), Inches(0.4),
            d, size=10.5, color=GW)
    # 播放器卡
    card(s, Inches(7.0), Inches(1.75), Inches(5.65), Inches(4.95))
    txt(s, Inches(7.35), Inches(2.1), Inches(4), Inches(0.3),
        "NOW PLAYING", size=10, color=GW, bold=True, font=MONO)
    txt(s, Inches(7.35), Inches(2.5), Inches(5), Inches(0.6),
        "《青春序曲》  v1.2", size=23, color=W, bold=True)
    txt(s, Inches(7.35), Inches(3.18), Inches(5), Inches(0.4),
        "国风 · 抒情 · 电子融合", size=12, color=GW)
    waveform(s, 7.35, 3.75, 4.95, 1.05, n=30, seed=11)
    # 控制钮
    for i, (lb, ac) in enumerate([("⏮", False), ("▶", True), ("⏭", False)]):
        cx = 9.05 + i * 0.8
        if ac:
            accent_card(s, Inches(cx), Inches(5.35), Inches(0.62), Inches(0.62), radius=0.5)
        else:
            card(s, Inches(cx), Inches(5.35), Inches(0.62), Inches(0.62),
                 fill=CARD2, radius=0.5, b_alpha=16000, glow=8000)
        txt(s, Inches(cx), Inches(5.33), Inches(0.62), Inches(0.62), lb,
            size=15 if ac else 13, color=BG if ac else GW, bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(7.35), Inches(6.25), Inches(5), Inches(0.3),
        "SEEDMUSIC · GENERATIVE AI", size=9, color=GW, font=MONO)


def build_roadmap():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_12_roadmap.png")
    header(s, "MY PLAN  ·  大学四年规划", "Roadmap · 与班级同频共振", 12)
    txt(s, Inches(0.72), Inches(1.62), Inches(12), Inches(0.4),
        "从融入集体 → 深耕方向 → 冲刺沉淀 → 共同上岸，每一步都与大家同行",
        size=13, color=GW)
    phases = [("大一", "融入与建设", ["组织班级活动", "建立学习小组", "破冰团建策划"],
               "让每个人都有归属感"),
              ("大二", "深耕与突破", ["技术分享会", "学科竞赛组队", "AI 工具科普"],
               "让每个人都能成长"),
              ("大三", "冲刺与沉淀", ["考研就业互助", "经验传承", "资源共享库"],
               "不让一个人掉队"),
              ("大四", "收获与上岸", ["毕业纪念", "全员圆满", "班级回忆录"],
               "留下最美的班级记忆")]
    y_line = 2.55
    hline(s, 1.15, 12.2, y_line, color=GD, dash=True, alpha=30000)
    for i, (yr, title, items, goal) in enumerate(phases):
        x = 0.75 + i * 3.08
        cx = x + 1.45
        dot(s, cx, y_line, r=0.075, color=G, glow=22000)
        card(s, Inches(x), Inches(2.95), Inches(2.9), Inches(3.35))
        watermark_num(s, x + 1.18, 5.22, i + 1, size=46)
        txt(s, Inches(x + 0.3), Inches(3.22), Inches(2.3), Inches(0.55),
            yr, size=26, color=GL, bold=True, font=NUM)
        txt(s, Inches(x + 0.3), Inches(3.82), Inches(2.4), Inches(0.4),
            title, size=15.5, color=W, bold=True)
        for j, it in enumerate(items):
            bullet(s, x + 0.3, 4.42 + j * 0.5, it, size=11.5, w=2.4)
        dot(s, x + 0.36, 6.02, r=0.045, color=G, glow=14000)
        txt(s, Inches(x + 0.5), Inches(5.9), Inches(2.3), Inches(0.3),
            goal, size=10.5, color=GL, italic=True, bold=True)


def build_promise():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_13_promise.png")
    header(s, "MY COMMITMENT  ·  三大承诺", "如果我是团支书", 13)
    cards = [(IC["headset"], "服务", "班级服务员",
              "把代码里的严谨，带进班级每一件小事",
              ["班级事务建台账，事事有回应", "重要通知当日达，绝不隔夜",
               "同学困难主动帮，提前搭把手"],
              "服务不是口号，是每天的行动力"),
             (IC["link"], "连接", "师生连接桥",
              "让信息在师生之间零障碍流动",
              ["每周收集建议，整理成清单反馈", "政策通知“翻译”成人话再传达",
               "搭建答疑渠道，问题不过夜"],
              "不让任何一条消息石沉大海"),
             (IC["people"], "成长", "成长同行者",
              "一个人走得快，一群人走得远",
              ["组织 AI / 编程主题分享会", "考前资料共建，笔记共享",
               "组队参加科创赛，一起拿奖"],
              "和你一起，成为更好的自己")]
    for i, (ic, title, sub, idea, acts, pledge) in enumerate(cards):
        x = 0.7 + i * 4.15
        card(s, Inches(x), Inches(2.0), Inches(3.85), Inches(4.55), glow=14000)
        watermark_num(s, x + 2.2, 2.16, i + 1, size=44)
        # 荧光图标（视觉焦点）
        accent_card(s, Inches(x + 0.35), Inches(2.38), Inches(0.72),
                    Inches(0.72), radius=0.16)
        txt(s, Inches(x + 0.35), Inches(2.36), Inches(0.72), Inches(0.72), ic,
            size=26, color=BG, font=ICON,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, Inches(x + 0.35), Inches(3.28), Inches(3.2), Inches(0.55),
            title, size=30, color=W, bold=True)
        txt(s, Inches(x + 0.35), Inches(4.02), Inches(3.2), Inches(0.36),
            sub, size=14, color=GL, bold=True)
        txt(s, Inches(x + 0.35), Inches(4.46), Inches(3.25), Inches(0.32),
            idea, size=11.5, color=GW, italic=True)
        tag(s, x + 0.35, 4.88, "ACTIONS · 行动清单", w=2.0, h=0.32)
        for j, act in enumerate(acts):
            bullet(s, x + 0.35, 5.32 + j * 0.35, act, size=11.5,
                   color=W, w=3.25, icon='check')
        dot(s, x + 0.42, 6.415, r=0.05, color=G, glow=14000)
        txt(s, Inches(x + 0.56), Inches(6.30), Inches(3.1), Inches(0.3),
            pledge, size=11, color=GL, italic=True, bold=True)


def build_closing():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_14_closing.png")
    txt(s, Inches(0), Inches(1.75), Inches(13.333), Inches(0.4),
        "FINAL WORDS  ·  MY COMMITMENT", size=12, color=GL, bold=True,
        font=MONO, align=PP_ALIGN.CENTER, spacing=200)
    tb = s.shapes.add_textbox(Inches(0), Inches(2.5), Inches(13.333), Inches(1.1))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Pt(0); tf.margin_right = Pt(0)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run(); r1.text = "不画大饼，"
    r1.font.size = Pt(52); r1.font.bold = True; r1.font.color.rgb = W; r1.font.name = CN
    r2 = p.add_run(); r2.text = "只做实事"
    r2.font.size = Pt(52); r2.font.bold = True; r2.font.color.rgb = G; r2.font.name = CN
    for r in (r1, r2):
        rPr = r._r.get_or_add_rPr()
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = etree.SubElement(rPr, qn('a:ea'))
        ea.set('typeface', CN)
        lt = rPr.find(qn('a:latin'))
        if lt is None:
            lt = etree.SubElement(rPr, qn('a:latin'))
        lt.set('typeface', LAT)
    _glow(tb, 160000, 20000)
    txt(s, Inches(0), Inches(4.05), Inches(13.333), Inches(1.1),
        "请大家给我一个机会，我会用实际行动，\n把班级变成我们想象中的样子。",
        size=18, color=GW, align=PP_ALIGN.CENTER)
    txt(s, Inches(0), Inches(5.55), Inches(13.333), Inches(0.5),
        "—— 王文杰", size=16, color=GL, bold=True, align=PP_ALIGN.CENTER)
    page_no(s, 14)


def build_thanks():
    s = prs.slides.add_slide(BL)
    bg_scene(s, "bg_15_thanks.png")
    t = txt(s, Inches(0), Inches(2.05), Inches(13.333), Inches(2.2),
            "THANK YOU", size=120, color=G, bold=True, align=PP_ALIGN.CENTER,
            font="Segoe UI Black", anchor=MSO_ANCHOR.MIDDLE)
    _glow(t, 220000, 30000)
    txt(s, Inches(0), Inches(4.55), Inches(13.333), Inches(0.6),
        "感谢聆听 · 以发光之姿，赴青春之约", size=27, color=W, bold=True,
        align=PP_ALIGN.CENTER)
    txt(s, Inches(0), Inches(5.35), Inches(13.333), Inches(0.5),
        "Explore Infinite Possibilities", size=14, color=GW, italic=True,
        font=LAT, align=PP_ALIGN.CENTER)
    page_no(s, 15)


# ============================================================
def main():
    print("生成 PPT 4.0 Tech Portfolio ...")
    pages = [build_cover, build_profile, build_timeline, build_exp,
             build_certs1, build_certs2, build_plane, build_nebula,
             build_gesture, build_website, build_seedmusic, build_roadmap,
             build_promise, build_closing, build_thanks]
    for i, fn in enumerate(pages):
        fn()
        print("  [%d/15] %s" % (i + 1, fn.__name__))
    prs.save(OUT)
    print("完成: %s  (%.1f MB)" % (OUT, os.path.getsize(OUT) / 1024 / 1024))


if __name__ == "__main__":
    main()
