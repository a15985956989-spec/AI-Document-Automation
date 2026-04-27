# -*- coding: utf-8 -*-
import zipfile, shutil, os, io, random

src = r'~/projects/附件2 班会新闻稿.docx'
dst = r'~/projects/附件2_已完成.docx'
photo_dir = r'~/projects/照片'

with zipfile.ZipFile(src) as z:
    with z.open('word/document.xml') as f:
        xml = f.read().decode('utf-8')
    if 'word/_rels/document.xml.rels' in z.namelist():
        with z.open('word/_rels/document.xml.rels') as f:
            rels_xml = f.read().decode('utf-8')
    else:
        rels_xml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

import re
existing_rids = re.findall(r'Id="rId(\d+)"', rels_xml)
max_rid = max(int(x) for x in existing_rids) if existing_rids else 0

# ===== Photos =====
photos = sorted([f for f in os.listdir(photo_dir) if f.lower().endswith(('.jpg', '.jpeg', '.png'))])
print(f"Photos: {len(photos)}")

# Select 4 photos for the news article
if len(photos) >= 10:
    selected_photos = [
        photos[-3],  # MVIMG_20260422_160023 叶老师讲课
        photos[0],   # IMG_20260422_160745
        photos[2],   # IMG_20260422_160906
        photos[5],   # IMG_20260422_161404
    ]
else:
    selected_photos = photos[:4]

print("Selected:", selected_photos)

def random_hex():
    return ''.join(random.choices('0123456789ABCDEF', k=8))

# Add images to docx
new_rels = []
image_entries = []
for i, photo_name in enumerate(selected_photos):
    rId = f'rId{max_rid + 1 + i}'
    image_path_in_docx = f'word/media/image_news_{i+1}.jpg'
    image_entries.append((image_path_in_docx, os.path.join(photo_dir, photo_name), rId))
    rel_entry = f'<Relationship Id="{rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image_news_{i+1}.jpg"/>'
    new_rels.append(rel_entry)

# ===== Text content =====
paragraphs = [
    ("normal", "为深入贯彻落实高校就业创业工作部署，加强大学生国防教育与安全防范意识，XXXX年X月XX日，XX学院2024级XX班在明志楼教室召开\u201c就业创业教育、征兵宣讲、防诈宣传\u201d主题班会。本次班会由X老师主持，组织委员XXX、团支书XXX协助组织，全班XX名同学全员参加。"),
    ("normal", "班会第一环节聚焦就业创业教育，由组织委员XXX同学主讲。XXX同学围绕\u201c精准赋能促就业，多元协同启新程\u201d主题，从生涯发展规划、职业精准定位、核心能力储备、校园到职场的角色转换、就业心理疏导五个维度进行了系统讲解。通过SWOT自我分析法帮助同学们客观认识自身优势与不足，结合当前数字化、智能化浪潮下新兴职业涌现的行业趋势，引导同学们科学规划职业发展路径。讲解中还穿插了MBTI人格探索互动环节，现场气氛活跃，同学们积极参与分享讨论。"),
    ("image", 0),  # image index
    ("caption", "\u56fe1 \u7ec4\u7ec7\u59d4\u5458\u9648\u4f73\u840d\u540c\u5b66\u8fdb\u884c\u5c31\u4e1a\u521b\u4e1a\u6559\u80b2\u4e3b\u9898\u5206\u4eab"),  # 图1 组织委员XXX同学进行就业创业教育主题分享
    ("normal", "班会第二环节为征兵宣讲，由XXX同学主讲。XXX同学以\u201c以笔为戎守家国，青春建功强军梦\u201d为题，从新时代强军目标、2026年最新征兵政策详解、军营风采展示、参军成长蜕变四个方面进行了详细介绍。重点解读了征兵年龄条件放宽、学历优待、经济待遇保障等最新政策利好，并分享了退役军人在就业安置、创业扶持、复学升学等方面的优待政策，鼓励同学们携笔从戎、报效祖国。"),
    ("image", 1),
    ("caption", "\u56fe2 \u9648\u660a\u540c\u5b66\u8fdb\u884c\u5f81\u5175\u5ba3\u8bb2\u4e3b\u9898\u5206\u4eab"),  # 图2 XXX同学进行征兵宣讲主题分享
    ("normal", "班会第三环节为防诈宣传，由X老师主讲。X老师结合近年来高校高发的电信网络诈骗案例，重点讲解了刷单返利诈骗、冒充公检法诈骗、网络贷款诈骗、虚假投资理财诈骗等常见诈骗类型的作案手法和防范技巧。X老师强调，同学们要时刻保持警惕，做到\u201c不轻信、不转账、不泄露个人信息\u201d，遇到可疑情况第一时间向辅导员或公安机关求助，切实筑牢防诈安全防线。"),
    ("image", 2),
    ("caption", "\u56fe3 \u8f85\u5bfc\u5458\u53f6\u60e0\u73b2\u8001\u5e08\u8fdb\u884c\u9632\u8bc8\u5ba3\u4f20\u4e3b\u9898\u8bb2\u89e3"),  # 图3 X老师进行防诈宣传主题讲解
    ("normal", "在互动讨论环节中，同学们踊跃发言，就自身关心的就业方向选择、征兵报名流程、防诈实用技巧等问题展开了热烈交流。X老师耐心解答同学们的疑问，并结合实际案例给出了针对性的建议。同学们纷纷表示，本次班会内容充实、贴近实际，对未来的职业规划和安全防范都有了更清晰的认识。"),
    ("image", 3),
    ("caption", "\u56fe4 \u540c\u5b66\u4eec\u79ef\u6781\u4e92\u52a8\u8ba8\u8bba"),  # 图4 同学们积极互动讨论
    ("normal", "本次主题班会内容丰富、形式多样，涵盖了就业创业、国防教育、安全防范三大重要主题，取得了良好的教育效果。通过本次班会，同学们进一步明确了职业发展方向，增强了创新创业意识和国防观念，提升了防诈识骗能力，为今后的学习生活和未来发展奠定了坚实基础。"),
]

def esc(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

def random_int():
    return random.randint(100, 999)

# Build body XML - find where <w:body> starts and </w:body> ends in original
body_start_tag = xml.find('<w:body>') + len('<w:body>')
body_end_tag = xml.find('</w:body>')

# Find the end of title paragraphs (the original 3 title lines + empty para after them)
# The original has 3 title paragraphs and then an empty paragraph at index 3
# We want to keep everything up to and including that empty paragraph
# Then insert our body content before sectPr

# The sectPr is at the end of body - extract it
sect_start = xml.rfind('<w:sectPr')
sect_end = xml.rfind('</w:sectPr>') + len('</w:sectPr>')
sect_pr = xml[sect_start:sect_end]

# Keep the title section from original (everything from body start to sectPr)
title_section = xml[body_start_tag:sect_start]

# Now build new body content = title_section + new paragraphs + sectPr
body_parts = []

# Add empty paragraph after body tag
body_parts.append('<w:p w14:paraId="AA000001"><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:p>')

# Image dimensions: 480x270 px = 4572000x2571750 EMU
CX = 4572000
CY = 2571750

for item in paragraphs:
    ptype = item[0]
    pid = random_hex()
    
    if ptype == "normal":
        text = esc(item[1])
        body_parts.append(
            f'<w:p w14:paraId="{pid}">'
            f'<w:pPr>'
            f'<w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/>'
            f'<w:widowControl/><w:kinsoku/><w:wordWrap/><w:overflowPunct/>'
            f'<w:topLinePunct w:val="0"/><w:autoSpaceDE/><w:autoSpaceDN/>'
            f'<w:bidi w:val="0"/><w:adjustRightInd/><w:snapToGrid/>'
            f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
            f'<w:ind w:firstLine="480" w:firstLineChars="200"/>'
            f'<w:jc w:val="left"/>'
            f'<w:rPr>'
            f'<w:rFonts w:hint="default" w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/>'
            f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
            f'<w:lang w:val="en-US" w:eastAsia="zh-CN"/>'
            f'</w:rPr></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:hint="default" w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/>'
            f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
            f'<w:lang w:val="en-US" w:eastAsia="zh-CN"/>'
            f'</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
            f'</w:p>'
        )
    elif ptype == "caption":
        text = esc(item[1])
        body_parts.append(
            f'<w:p w14:paraId="{pid}">'
            f'<w:pPr>'
            f'<w:keepNext w:val="0"/><w:keepLines w:val="0"/>'
            f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
            f'<w:ind w:firstLine="0" w:firstLineChars="0"/>'
            f'<w:jc w:val="center"/>'
            f'<w:rPr>'
            f'<w:b/><w:bCs/>'
            f'<w:rFonts w:hint="default" w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/>'
            f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
            f'<w:lang w:val="en-US" w:eastAsia="zh-CN"/>'
            f'</w:rPr></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:b/><w:bCs/>'
            f'<w:rFonts w:hint="default" w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/>'
            f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
            f'<w:lang w:val="en-US" w:eastAsia="zh-CN"/>'
            f'</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
            f'</w:p>'
        )
    elif ptype == "image":
        img_idx = item[1]
        _, _, img_rId = image_entries[img_idx]
        docPr_id = random_int()
        cNvPr_id = random_int()
        body_parts.append(
            f'<w:p w14:paraId="{pid}">'
            f'<w:pPr>'
            f'<w:keepNext w:val="0"/><w:keepLines w:val="0"/>'
            f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
            f'<w:ind w:firstLine="0" w:firstLineChars="0"/>'
            f'<w:jc w:val="center"/>'
            f'</w:pPr>'
            f'<w:r>'
            f'<w:rPr><w:noProof/></w:rPr>'
            f'<w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            f'<wp:extent cx="{CX}" cy="{CY}"/>'
            f'<wp:docPr id="{docPr_id}" name="Picture {docPr_id}"/>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{cNvPr_id}" name="Picture {cNvPr_id}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill>'
            f'<a:blip r:embed="{img_rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
            f'<a:stretch><a:fillRect/></a:stretch>'
            f'</pic:blipFill>'
            f'<pic:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{CX}" cy="{CY}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'</pic:spPr>'
            f'</pic:pic>'
            f'</a:graphicData>'
            f'</a:graphic>'
            f'</wp:inline>'
            f'</w:drawing>'
            f'</w:r>'
            f'</w:p>'
        )
        # Empty line after image
        body_parts.append(
            f'<w:p w14:paraId="{random_hex()}">'
            f'<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
            f'</w:p>'
        )

# Add trailing empty para
body_parts.append(
    f'<w:p w14:paraId="{random_hex()}">'
    f'<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
    f'</w:p>'
)

# Assemble new XML: keep title from original, add our content
new_body = title_section + ''.join(body_parts)
new_xml = xml[:body_start_tag] + new_body + sect_pr + '</w:body></w:document>'

# Verify XML can be parsed
import xml.etree.ElementTree as ET
try:
    ET.fromstring(new_xml)
    print("XML parse OK!")
except ET.ParseError as e:
    print(f"XML parse ERROR: {e}")

# Update rels
if new_rels:
    insert_pos = rels_xml.rfind('</Relationships>')
    for rel in new_rels:
        rels_xml = rels_xml[:insert_pos] + rel + rels_xml[insert_pos:]

# Write docx
with zipfile.ZipFile(src, 'r') as zin:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                zout.writestr(item, new_xml.encode('utf-8'))
            elif item.filename == 'word/_rels/document.xml.rels':
                zout.writestr(item, rels_xml.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))
        for img_path_in_docx, img_local_path, _ in image_entries:
            with open(img_local_path, 'rb') as img_f:
                img_data = img_f.read()
            zout.writestr(img_path_in_docx, img_data)
            print(f"  Added: {img_path_in_docx} ({len(img_data)//1024}KB)")

with open(dst, 'wb') as f:
    f.write(buf.getvalue())

print(f"\nDone! -> {dst}")
