# AI-Document-Automation

> 用AI把2小时的公文工作压缩到25分钟——一个团支书的效率自救指南

## 📌 项目背景

作为班级团支书，每学期需要完成 **8场** 班会/团日活动的固定公文材料：
- 📄 附件1：主题班会授课登记表
- 📰 附件2：班会新闻稿

**传统痛点：**
- 手工填写模板，单篇耗时 **2-3小时**
- 格式要求高（字体/缩进/页边距），容易出错
- 需嵌入活动照片，手动调整尺寸费时
- 时间紧（要求当天或次日提交）

**解决方案：** 基于Python + 大模型API，实现"一句话指令 → 成品文档"的全自动化链路

---

## 🚀 核心功能

| 功能 | 说明 | 效率提升 |
|------|------|---------|
| **模板格式复刻** | 直接操作Word底层XML，100%保留原模板样式 | 零格式错误 |
| **AI内容生成** | 根据PPT素材自动撰写规范公文 | 质量稳定 |
| **照片智能嵌入** | 自动筛选代表性照片，统一尺寸插入 | 无需手动调整 |
| **脚本复用** | 一次开发，改参数即可批量生成 | 永久复用 |

**实测效果：** 单篇文档制作时间从 **2-3小时 → 25分钟**，效率提升 **80%**

---

## 🛠️ 技术栈

- **Python 3.x**
- **zipfile** - 解析.docx本质（ZIP压缩包）
- **xml.etree** - 操作Word XML结构
- **OpenXML标准** - wp:inline + a:blip 图片嵌入
- **大模型API** - 内容生成与润色

---

## 📂 项目结构

```
AI-Document-Automation/
├── fill_bjh.py          # 授课登记表填充脚本
├── fill_news2.py        # 新闻稿生成脚本（含图片嵌入）
├── templates/           # Word模板文件
│   ├── 附件1_模板.docx
│   └── 附件2_模板.docx
├── assets/              # 活动照片
├── output/              # 生成结果
└── docs/
    └── AI协作案例报告.pdf   # 完整技术报告
```

---

## 🎯 快速开始

### 1. 环境准备

```bash
pip install -r requirements.txt
```

### 2. 准备素材

将以下文件放入对应目录：
- 📎 PPT文件（班会内容素材）→ `assets/班会主题.pptx`
- 📷 活动照片 → `assets/photos/`
- 📋 Word模板 → `templates/`

### 3. 一句话生成

```bash
python fill_news2.py \
  --date "2026年4月22日" \
  --theme "就业创业教育、征兵宣讲、防诈宣传" \
  --ppt "assets/班会主题.pptx" \
  --photos "assets/photos/" \
  --output "output/附件2_已完成.docx"
```

### 4. 获取结果

```
output/
└── 附件2_已完成.docx    ✅ 可直接提交
```

---

## 🔧 技术亮点

### 1. Word XML底层操作

```python
# .docx本质上是ZIP包，直接解析XML
with zipfile.ZipFile(template_path) as zf:
    xml_content = zf.read('word/document.xml')

# 保留标题段落，替换正文内容
title_section = xml[body_start:sect_start]
new_body = title_section + generated_content
new_xml = xml[:body_start] + new_body + sect_pr
```

### 2. OpenXML标准图片嵌入

```python
# 构建符合OOXML规范的图片节点
def make_image_para(rId, cx, cy):
    return f'''<w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="{cx}" cy="{cy}"/>
            <a:graphic>
              <a:graphicData>
                <pic:blipFill>
                  <a:blip r:embed="{rId}"/>
                </pic:blipFill>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>'''
```

### 3. 踩坑复盘

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| Python SyntaxError | 中文引号被识别为字符串边界 | 改用 \u201c/\u201d Unicode转义 |
| XML解析失败 | 拼接时缺少`</w:document>`闭合标签 | 修改组装逻辑，显式追加闭合标签 |
| 标题消失 | 替换body时覆盖了原始标题段落 | 先提取title_section，再追加body_parts |
| PowerShell编码乱码 | 终端默认非UTF-8 | 使用`python -X utf8`参数启动 |

---

## 📊 效率对比

| 维度 | 传统手工 | AI协作方式 |
|------|---------|-----------|
| ⏱️ 耗时 | 2-3小时 | ~25分钟 |
| ✍️ 文字质量 | 依赖个人能力，不稳定 | 规范、准确、得体 |
| 📐 格式合规 | 容易不一致 | 与往期100%一致 |
| 🖼️ 图片处理 | 手动调整，费时 | 自动嵌入，尺寸统一 |
| 🔄 可复用性 | 每次从零开始 | 改参数即可复用 |

---

## 📝 最佳实践：三要素一次性给齐

```
高效协作公式 = 主题说明 + 结构化素材 + 模板/照片

要素1: 主题说明
  → "就业创业教育、征兵宣讲、防诈宣传（三合一主题班会）"

要素2: 结构化素材  
  → @PPT文件（内含各环节主题词和内容框架）

要素3: 模板+照片
  → @附件1/2模板文件 + @照片文件夹
```

---

## 🗺️ 未来扩展

- [ ] 批量生成一学期所有班会文档
- [ ] 一键归档为学期汇报PDF
- [ ] 从照片自动识别活动场景（CV）
- [ ] 团日活动总结同步生成
- [ ] 接入飞书/钉钉机器人，自动提交

---

## 📄 License

MIT License - 欢迎自由使用和二次开发

---

> 💡 **适用场景**：高校团支书、辅导员、行政人员、任何需要批量生成规范Word文档的场景
> 
> 🔗 **关联项目**：[货憨憨标题提取器](https://github.com/yourname/huohanhan-title-extractor) - 我的另一个运营提效工具
