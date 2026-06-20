#!/usr/bin/env python3
"""
修复第三章格式问题：
1. rFonts 缺少 w:eastAsia="KaiTi" — 导致中文字体不是楷体
2. font_run 的 lxml truth-testing bug — r_pr.rFonts or OxmlElement() 不可靠
"""
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.shared import Pt, RGBColor
from pathlib import Path
from lxml import etree

BASE = Path("/Users/castle/Desktop/space for claude")
DOC_PATH = BASE / "海外ETF期权及衍生策略产品研究0618_更新版.docx"
OUT_PATH = BASE / "海外ETF期权及衍生策略产品研究0618_更新版.docx"

CN_FONT = "KaiTi"
EN_FONT = "Times New Roman"

doc = Document(str(DOC_PATH))
body = doc.element.body
children = list(body.iterchildren())

# 找到第三章
si = ei = None
for idx, child in enumerate(children):
    if isinstance(child, CT_P):
        t = "".join(child.itertext()).strip()
        if si is None and t.startswith("三、"):
            si = idx
        elif si is not None and t.startswith("四、"):
            ei = idx
            break

if si is None or ei is None:
    print("ERROR: Cannot find chapter 3 boundaries")
    exit(1)

print(f"第三章: 元素 {si} 到 {ei}, 共 {ei-si} 个元素")

# 统计并修复
fixed_heading = 0
fixed_body = 0
fixed_table = 0

def fix_paragraph(p_element, style_id=""):
    global fixed_heading, fixed_body, fixed_table
    for r in p_element.findall(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            continue
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            continue
        ascii_val = rf.get(qn("w:ascii"))
        hAnsi_val = rf.get(qn("w:hAnsi"))
        ea_val = rf.get(qn("w:eastAsia"))
        if (ascii_val or hAnsi_val) and not ea_val:
            rf.set(qn("w:eastAsia"), CN_FONT)
            if style_id in ('1', '2', '3', '21', '31'):
                fixed_heading += 1
            else:
                fixed_body += 1

def fix_table(table_element):
    global fixed_table
    for row in table_element.findall(qn("w:tr")):
        for cell in row.findall(qn("w:tc")):
            for p in cell.findall(qn("w:p")):
                for r in p.findall(qn("w:r")):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is None:
                        continue
                    rf = rpr.find(qn("w:rFonts"))
                    if rf is None:
                        continue
                    if rf.get(qn("w:ascii")) and not rf.get(qn("w:eastAsia")):
                        rf.set(qn("w:eastAsia"), CN_FONT)
                        fixed_table += 1

for idx in range(si, ei):
    child = children[idx]
    if isinstance(child, CT_P):
        t = "".join(child.itertext()).strip()
        if not t:
            continue
        ppr = child.find(qn("w:pPr"))
        style_id = ""
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None:
                style_id = ps.get(qn("w:val"), "")
        fix_paragraph(child, style_id)
    else:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'tbl':
            fix_table(child)

print(f"修复: {fixed_heading} 个标题 run, {fixed_body} 个正文 run, {fixed_table} 个表格 run")
print(f"为 rFonts 添加 w:eastAsia=\"{CN_FONT}\"")

doc.save(str(OUT_PATH))
print(f"已保存: {OUT_PATH}")

# 验证
doc2 = Document(str(OUT_PATH))
body2 = doc2.element.body
children2 = list(body2.iterchildren())
si2 = ei2 = None
for idx, child in enumerate(children2):
    if isinstance(child, CT_P):
        t = "".join(child.itertext()).strip()
        if si2 is None and t.startswith("三、"):
            si2 = idx
        elif si2 is not None and t.startswith("四、"):
            ei2 = idx
            break

missing_para = 0
missing_table = 0
for idx in range(si2, ei2):
    child = children2[idx]
    if isinstance(child, CT_P):
        for r in child.findall(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None: continue
            rf = rpr.find(qn("w:rFonts"))
            if rf is None: continue
            if rf.get(qn("w:ascii")) and not rf.get(qn("w:eastAsia")):
                missing_para += 1
    else:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'tbl':
            for row in child.findall(qn("w:tr")):
                for cell in row.findall(qn("w:tc")):
                    for p in cell.findall(qn("w:p")):
                        for r in p.findall(qn("w:r")):
                            rpr = r.find(qn("w:rPr"))
                            if rpr is None: continue
                            rf = rpr.find(qn("w:rFonts"))
                            if rf is None: continue
                            if rf.get(qn("w:ascii")) and not rf.get(qn("w:eastAsia")):
                                missing_table += 1

if missing_para == 0 and missing_table == 0:
    print("\n✅ 验证通过：第三章所有 rFonts 均包含 eastAsia 属性（含表格）")
else:
    print(f"\n⚠️ 仍有 {missing_para} 个段落 run, {missing_table} 个表格 run 缺少 eastAsia")
