# -*- coding: utf-8 -*-
import zipfile, shutil, os, re

src = r'~/projects/附件1 主题班会授课登记表.docx'
dst = r'~/projects/附件1_已完成.docx'
shutil.copy2(src, dst)

with zipfile.ZipFile(src) as z:
    with z.open('word/document.xml') as f:
        xml = f.read().decode('utf-8')

# 授课过程内容
guocheng = "本次班会围绕就业创业教育、征兵宣讲与防诈宣传三大主题展开，由X老师主持，全体同学积极参与。班会伊始，X老师就当前就业形势进行深入分析，强调大学生应积极提升自身综合素质，做好职业生涯规划，培养创新创业意识，鼓励同学们勇于尝试、脚踏实地。随后，X老师向同学们详细介绍了征兵政策与相关待遇，讲解了当代大学生参军入伍的重要意义，鼓励有志青年积极响应国家号召，报名参军，在军旅生涯中磨砺意志、锻炼自我。最后，X老师重点讲解了当前高发的电信网络诈骗手段，包括刷单返利、冒充公检法、网络贷款等常见诈骗类型，提醒同学们时刻保持警惕，不轻信陌生人，不随意转账，遇到可疑情况及时报警，切实筑牢防诈安全防线。"

# 授课效果内容
xiaoguo = "本次主题班会内容丰富、重点突出，取得了良好的教育效果。通过就业创业教育环节，同学们进一步明确了自身的职业发展方向，增强了创新创业意识，对未来的求职规划有了更清晰的认识；通过征兵宣讲环节，同学们对征兵政策有了全面了解，爱国热情高涨，部分同学表示有意愿积极考虑参军入伍；通过防诈宣传环节，同学们对各类新型电信网络诈骗手段有了深刻认识，防诈意识得到显著增强，能够自觉做到不轻信、不转账、不泄露。整体而言，本次班会有效拓宽了同学们的视野，激发了责任担当意识，为今后的学习生活和职业发展奠定了坚实基础。"

def make_text_run(text):
    # 转义XML特殊字符
    text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return ('<w:r>'
            '<w:rPr>'
            '<w:rFonts w:hint="eastAsia" w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/>'
            '<w:kern w:val="0"/><w:sz w:val="28"/><w:szCs w:val="28"/>'
            '</w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t>'
            '</w:r>')

def make_para(text, para_id, with_indent=True):
    indent_part = '<w:ind w:firstLine="560" w:firstLineChars="200"/>' if with_indent else ''
    return (f'<w:p w14:paraId="{para_id}">'
            '<w:pPr><w:pStyle w:val="7"/>'
            '<w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:widowControl/>'
            '<w:suppressLineNumbers w:val="0"/>'
            + indent_part +
            '<w:jc w:val="both"/>'
            '<w:rPr>'
            '<w:rFonts w:hint="eastAsia" w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/>'
            '<w:kern w:val="0"/><w:sz w:val="28"/><w:szCs w:val="28"/>'
            '</w:rPr></w:pPr>'
            + make_text_run(text) +
            '</w:p>')

# ---- 找到授课过程右侧单元格中的空段落，用内容替换 ----
# 授课过程右侧单元格 paraId=4CEB04C8
idx1 = xml.find('授课过程')
after1 = xml[idx1:]
tc1_offset = after1.find('</w:tc><w:tc>')
cell1_start = idx1 + tc1_offset
# 找到这个单元格的结束 </w:tc>
cell1_end = xml.find('</w:tc>', cell1_start + 13) + len('</w:tc>')
cell1_xml = xml[cell1_start: cell1_end]
print("授课过程单元格内容:", repr(cell1_xml[:200]))

# ---- 找到授课效果右侧单元格 ----
idx2 = xml.find('授课效果及总结')
after2 = xml[idx2:]
tc2_offset = after2.find('</w:tc><w:tc>')
cell2_start = idx2 + tc2_offset
cell2_end = xml.find('</w:tc>', cell2_start + 13) + len('</w:tc>')
cell2_xml = xml[cell2_start: cell2_end]
print("授课效果单元格内容:", repr(cell2_xml[:200]))

# 新的授课过程单元格
new_cell1 = ('</w:tc><w:tc>'
             '<w:tcPr><w:tcW w:w="9138" w:type="dxa"/><w:gridSpan w:val="7"/><w:vAlign w:val="center"/></w:tcPr>'
             + make_para(guocheng, 'AA11BB22', with_indent=True)
             + '</w:tc>')

# 新的授课效果单元格
new_cell2 = ('</w:tc><w:tc>'
             '<w:tcPr><w:tcW w:w="9138" w:type="dxa"/><w:gridSpan w:val="7"/><w:vAlign w:val="center"/></w:tcPr>'
             + make_para(xiaoguo, 'CC33DD44', with_indent=True)
             + '</w:tc>')

# 替换
new_xml = xml[:cell1_start] + new_cell1 + xml[cell1_end:]

# 重新定位cell2（xml已变化）
idx2b = new_xml.find('授课效果及总结')
after2b = new_xml[idx2b:]
tc2b_offset = after2b.find('</w:tc><w:tc>')
cell2b_start = idx2b + tc2b_offset
cell2b_end = new_xml.find('</w:tc>', cell2b_start + 13) + len('</w:tc>')
new_xml = new_xml[:cell2b_start] + new_cell2 + new_xml[cell2b_end:]

print("替换完成，写入文件...")

# 写回docx
import io
with zipfile.ZipFile(src, 'r') as zin:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                zout.writestr(item, new_xml.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))

with open(dst, 'wb') as f:
    f.write(buf.getvalue())

print("Done! 输出:", dst)
