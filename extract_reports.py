# -*- coding: utf-8 -*-
"""提取 Gemini 四份报告的文字内容"""
import zipfile, re, os, sys
from xml.etree import ElementTree as ET

def extract_pptx_text(path):
    """从 pptx 提取所有文字"""
    texts = []
    with zipfile.ZipFile(path) as z:
        slide_files = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                             key=lambda n: int(re.search(r'\d+', n).group()))
        for sf in slide_files:
            xml = z.read(sf).decode('utf-8', errors='ignore')
            # 提取所有 <a:t> 文本
            runs = re.findall(r'<a:t[^>]*>([^<]*)</a:t>', xml)
            page_text = ' | '.join(t for t in runs if t.strip())
            texts.append(f"--- {os.path.basename(sf)} ---\n{page_text}")
    return '\n'.join(texts)

def extract_pdf_text(path):
    from pypdf import PdfReader
    r = PdfReader(path)
    out = []
    for i, p in enumerate(r.pages):
        t = p.extract_text() or ''
        out.append(f"--- page {i+1} ---\n{t}")
    return '\n'.join(out)

files = [
    r"C:\Users\WR\Desktop\王文杰团支书竞选PPT 3.0 科技最终版设计规划.pptx",
    r"C:\Users\WR\Desktop\王文杰竞选PPT_3.0版本升级指南.pptx",
    r"C:\Users\WR\Desktop\王文杰竞选演讲与作品集演示手册.pdf",
    r"C:\Users\WR\Desktop\王文杰-团支书竞选报告.pdf",
]

for f in files:
    print("="*80)
    print("FILE:", f)
    print("="*80)
    if not os.path.exists(f):
        print("!! NOT FOUND")
        continue
    try:
        if f.endswith('.pptx'):
            print(extract_pptx_text(f))
        else:
            print(extract_pdf_text(f))
    except Exception as e:
        print("!! ERROR:", e)
