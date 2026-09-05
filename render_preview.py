# -*- coding: utf-8 -*-
"""用 PowerPoint COM 导出每页 PNG 预览"""
import os, sys
import win32com.client

PPTX = r"d:\wwj\王文杰-团支书竞选-3.0科技最终版.pptx"
OUT = r"d:\wwj\_preview"
os.makedirs(OUT, exist_ok=True)

only = [int(x) for x in sys.argv[1:]] if len(sys.argv) > 1 else None

app = win32com.client.Dispatch("PowerPoint.Application")
try:
    app.DisplayAlerts = 1
except Exception:
    pass
pres = app.Presentations.Open(PPTX, -1, 0, 0)
try:
    n = pres.Slides.Count
    for i in range(1, n + 1):
        if only and i not in only:
            continue
        p = os.path.join(OUT, "slide_%d.png" % i)
        pres.Slides(i).Export(p, "PNG", 1600, 900)
        print("exported", p)
finally:
    pres.Close()
    app.Quit()
print("done")
