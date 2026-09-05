# -*- coding: utf-8 -*-
"""
PPTX 动画后处理 v2 - 符合 OOXML 标准的 transition + timing
修正:
1. transition 子元素不含 <p:st>，速度用 spd 属性 (slow/med/fast)
2. timing 用标准 par/seq/cTn/animEffect 结构
3. 插入顺序: cSld -> clrMapOvr -> transition -> timing
"""
import os, shutil, zipfile, tempfile
from lxml import etree

NSMAP = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
         'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

PPTX = r"d:\wwj\王文杰-团支书竞选-3.0科技最终版.pptx"


def inject_transition(slide_xml, kind="fade", speed="fast"):
    """标准 transition: <p:transition spd="fast" p14:dur="450"><p:fade/></p:transition>"""
    for t in slide_xml.findall('p:transition', NSMAP):
        slide_xml.remove(t)

    trans = etree.Element(P + 'transition')
    trans.set('spd', speed)
    trans.set('{http://schemas.microsoft.com/office/powerpoint/2010/main}dur', '350')

    if kind == "fade":
        etree.SubElement(trans, P + 'fade')
    elif kind == "push":
        etree.SubElement(trans, P + 'push').set('dir', 'l')
    elif kind == "wipe":
        etree.SubElement(trans, P + 'wipe').set('dir', 'l')
    elif kind == "split":
        sp = etree.SubElement(trans, P + 'split'); sp.set('orient', 'horz'); sp.set('dir', 'out')
    elif kind == "fadeBlack":
        etree.SubElement(trans, P + 'fade').set('thruBlk', '1')
    elif kind == "zoom":
        etree.SubElement(trans, P + 'zoom').set('dir', 'in')

    # transition 必须紧跟 clrMapOvr 之后
    cSld = slide_xml.find('p:cSld', NSMAP)
    clr = slide_xml.find('p:clrMapOvr', NSMAP)
    anchor = clr if clr is not None else cSld
    slide_xml.insert(list(slide_xml).index(anchor) + 1, trans)


def inject_entry_timing(slide_xml, duration_ms=300, mode='row'):
    """标准入场动画: 对页面所有 shape 做 fade 依次出现。
    用动画组序列，引用每个 shape 的 spid。
    """
    for t in slide_xml.findall('p:timing', NSMAP):
        slide_xml.remove(t)

    # 收集 shape 的 id 与位置（y, x），同一排的元素归为同一波齐出
    spTree = slide_xml.find('.//p:cSld/p:spTree', NSMAP)
    items = []  # (y_emu, x_emu, spid, is_pic)
    for tag in ('sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'):
        for el in spTree.findall('p:' + tag, NSMAP):
            cNvPr = el.find('.//p:cNvPr', NSMAP)
            if cNvPr is None or not cNvPr.get('id'):
                continue
            off = el.find('.//a:off', NSMAP)
            ext = el.find('.//a:ext', NSMAP)
            y = int(off.get('y', '0')) if off is not None else 0
            x = int(off.get('x', '0')) if off is not None else 0
            w = int(ext.get('cx', '0')) if ext is not None else 0
            cNvSpPr = el.find('.//p:cNvSpPr', NSMAP)
            is_txbox = cNvSpPr is not None and cNvSpPr.get('txBox') == '1'
            items.append((y, x, cNvPr.get('id'), tag == 'pic', w, is_txbox))
    # 背景大图（位于 0,0 的 pic）不做入场
    items = [it for it in items if not (it[3] and it[0] == 0 and it[1] == 0)]
    if not items:
        return
    plan = []  # (delay_ms, spid)
    if mode == 'col':
        # 以页面大卡（宽>2in 的非文本形状）为列锚，区间间隙 >0.1in 才分列；
        # 元素按 x 落入哪个卡区间就归哪列，列间从左到右齐出
        anchors = sorted((x, x + w) for y, x, _, _, w, tb in items
                         if not tb and 1828800 < w < 5486400 and y >= 1371600)
        cols = []
        for a, b in anchors:
            if not cols or a - cols[-1][1] > 91440:
                cols.append([a, b])
            else:
                cols[-1][1] = max(cols[-1][1], b)

        def col_of(x):
            for i, (a, b) in enumerate(cols):
                if a - 45720 <= x <= b + 45720:
                    return i
            return min(range(len(cols)),
                       key=lambda i: min(abs(x - cols[i][0]), abs(x - cols[i][1])))

        if not cols:
            plan = [(0, it[2]) for it in items]
        else:
            for y, x, spid, *_ in items:
                if y < 1371600:
                    plan.append((0, spid))
                    continue
                plan.append((col_of(x) * 70, spid))
    else:
        # 按 y→x 排序；与所在波基准 y 差 ≤1 英寸视为同一排，同波同时出现
        items.sort(key=lambda t: (t[0], t[1]))
        wave = -1
        base_y = None
        for y, x, spid, *_ in items:
            if base_y is None or y - base_y > 914400:
                wave += 1
                base_y = y
            plan.append((wave * 70, spid))

    timing = etree.Element(P + 'timing')
    tnLst = etree.SubElement(timing, P + 'tnLst')
    par = etree.SubElement(tnLst, P + 'par')
    cTn = etree.SubElement(par, P + 'cTn', id='1', dur='indefinite', restart='never', nodeType='tmRoot')
    childTnLst = etree.SubElement(cTn, P + 'childTnLst')

    seq = etree.SubElement(childTnLst, P + 'seq', concurrent='1', nextAc='seek')
    seqcTn = etree.SubElement(seq, P + 'cTn', id='2', dur='indefinite', nodeType='mainSeq')
    seq_child = etree.SubElement(seqcTn, P + 'childTnLst')

    # 同排同波齐出，波间 70ms 从上往下递进
    nid = 3
    for dly, spid in plan:
        p1 = etree.SubElement(seq_child, P + 'par')
        c1 = etree.SubElement(p1, P + 'cTn', id=str(nid), fill='hold'); nid += 1
        st1 = etree.SubElement(c1, P + 'stCondLst')
        etree.SubElement(st1, P + 'cond', delay=str(dly))
        ch1 = etree.SubElement(c1, P + 'childTnLst')

        p2 = etree.SubElement(ch1, P + 'par')
        c2 = etree.SubElement(p2, P + 'cTn', id=str(nid), fill='hold'); nid += 1
        st2 = etree.SubElement(c2, P + 'stCondLst')
        etree.SubElement(st2, P + 'cond', delay='0')
        ch2 = etree.SubElement(c2, P + 'childTnLst')

        p3 = etree.SubElement(ch2, P + 'par')
        c3 = etree.SubElement(p3, P + 'cTn', id=str(nid), presetID='10', presetClass='entr',
                              presetSubtype='0', fill='hold', grpId='0', nodeType='afterEffect'); nid += 1
        st3 = etree.SubElement(c3, P + 'stCondLst')
        etree.SubElement(st3, P + 'cond', delay='0')
        ch3 = etree.SubElement(c3, P + 'childTnLst')

        # 1) set visibility = visible（标准入场前置节点）
        setel = etree.SubElement(ch3, P + 'set')
        sbhvr = etree.SubElement(setel, P + 'cBhvr')
        sctn = etree.SubElement(sbhvr, P + 'cTn', id=str(nid), dur='1', fill='hold'); nid += 1
        sst = etree.SubElement(sctn, P + 'stCondLst')
        etree.SubElement(sst, P + 'cond', delay='0')
        stgt = etree.SubElement(sbhvr, P + 'tgtEl')
        etree.SubElement(stgt, P + 'spTgt', spid=str(spid))
        attrL = etree.SubElement(sbhvr, P + 'attrNameLst')
        etree.SubElement(attrL, P + 'attrName').text = 'style.visibility'
        to = etree.SubElement(setel, P + 'to')
        etree.SubElement(to, P + 'strVal', val='visible')

        # 2) animEffect: fade
        ae = etree.SubElement(ch3, P + 'animEffect', transition='in', filter='fade')
        cbhvr = etree.SubElement(ae, P + 'cBhvr')
        ctn = etree.SubElement(cbhvr, P + 'cTn', id=str(nid), dur=str(duration_ms), fill='hold'); nid += 1
        cst = etree.SubElement(ctn, P + 'stCondLst')
        etree.SubElement(cst, P + 'cond', delay='0')
        tgt = etree.SubElement(cbhvr, P + 'tgtEl')
        etree.SubElement(tgt, P + 'spTgt', spid=str(spid))

    # timing 必须在 transition 之后
    trans = slide_xml.find('p:transition', NSMAP)
    if trans is not None:
        slide_xml.insert(list(slide_xml).index(trans) + 1, timing)
    else:
        cSld = slide_xml.find('p:cSld', NSMAP)
        clr = slide_xml.find('p:clrMapOvr', NSMAP)
        anchor = clr if clr is not None else cSld
        slide_xml.insert(list(slide_xml).index(anchor) + 1, timing)


def post_process(pptx_path):
    tmpdir = tempfile.mkdtemp()
    with zipfile.ZipFile(pptx_path, 'r') as z:
        names = z.namelist()
        z.extractall(tmpdir)

    slides_dir = os.path.join(tmpdir, 'ppt', 'slides')
    transitions = ["fade", "fade", "push", "fade", "wipe", "wipe",
                   "push", "zoom", "push", "wipe", "fade",
                   "split", "push", "fade", "fadeBlack"]

    # 按列（从左到右分区）入场的页面：2 档案 / 5 证书 / 11 音乐 / 12 规划 / 13 承诺
    col_pages = {2, 5, 11, 12, 13}
    for i in range(1, 16):
        sf = os.path.join(slides_dir, f"slide{i}.xml")
        if not os.path.exists(sf):
            continue
        tree = etree.parse(sf)
        root = tree.getroot()
        inject_transition(root, transitions[i-1] if i <= len(transitions) else "fade")
        inject_entry_timing(root, mode='col' if i in col_pages else 'row')
        tree.write(sf, xml_declaration=True, encoding='UTF-8', standalone=True)
        print(f"  slide{i}: {transitions[i-1]}")

    backup = pptx_path + ".bak"
    shutil.copy2(pptx_path, backup)
    # 先写临时文件，再原子替换（原文件被 WPS/Office 占用时保留临时文件）
    tmp_out = pptx_path + ".new"
    with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.write(os.path.join(tmpdir, name), name)
    shutil.rmtree(tmpdir)
    try:
        os.replace(tmp_out, pptx_path)
        out_path = pptx_path
        os.remove(backup)
    except PermissionError:
        out_path = tmp_out
        print("\n⚠️ 原文件被占用（WPS/Office 未关闭），已输出到: %s" % tmp_out)
        print("   关闭后把 .new 删掉改名即可，或重新运行本脚本。")
    print(f"\n✅ 动画层注入完成: {out_path}  ({os.path.getsize(out_path)/1024/1024:.1f} MB)")


if __name__ == "__main__":
    post_process(PPTX)
