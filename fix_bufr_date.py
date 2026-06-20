#!/usr/bin/env python3
"""修复报告中的 BUFR 成立日期错误: 2023年5月 → 2020年8月"""
from docx import Document
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from pathlib import Path

BASE = Path("/Users/castle/Desktop/space for claude")
DOC_PATH = BASE / "海外ETF期权及衍生策略产品研究0618_更新版.docx"

doc = Document(str(DOC_PATH))
body = doc.element.body
children = list(body.iterchildren())

# Find "BUFR" mentions and fix dates
fixes = 0
for child in children:
    if not isinstance(child, CT_P):
        continue
    for r in child.findall(qn("w:r")):
        for t_node in r.findall(qn("w:t")):
            if t_node.text and "BUFR" in t_node.text and "2023年5月" in t_node.text:
                print(f"BEFORE: {t_node.text[:120]}")
                t_node.text = t_node.text.replace("2023年5月", "2020年8月")
                print(f"AFTER:  {t_node.text[:120]}")
                fixes += 1

if fixes:
    doc.save(str(DOC_PATH))
    print(f"\n修复了 {fixes} 处 BUFR 日期错误")
else:
    print("未找到需要修复的内容（可能已被修复）")

# Verify
doc2 = Document(str(DOC_PATH))
body2 = doc2.element.body
for child in body2:
    if isinstance(child, CT_P):
        t = "".join(child.itertext()).strip()
        if "BUFR" in t and ("2020" in t or "2023" in t):
            # extract just the relevant part
            for r in child.findall(qn("w:r")):
                for tn in r.findall(qn("w:t")):
                    if tn.text and "BUFR" in tn.text:
                        if "2020年8月" in tn.text:
                            print(f"✅ 验证: 日期已修正为2020年8月")
                        elif "2023年5月" in tn.text:
                            print(f"❌ 仍有错误日期!")
                        break
