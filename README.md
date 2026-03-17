# Claude Skills 中文增强版

这是 Claude Code 的 DOCX、PPTX 和 PDF skill 的中文增强版本，修复了中文编码问题。

## 修复内容

### 1. DOCX Skill
- **修复 `unpack.py`**: 将编码从 ASCII 改为 UTF-8，正确显示中文字符
- **新增 `extract_text.py`**: 自动检测 UTF-8/GBK/GB18030 编码的 Word 文档提取工具
- **新增 `extract_pptx_text.py`**: PowerPoint 文本提取工具

### 2. PPTX Skill
- **新增 `extract_text.py`**: Word 文档提取工具
- **新增 `extract_pptx_text.py`**: 自动检测编码的 PPT 提取工具
- **更新文档**: 添加中文支持说明

### 3. PDF Skill
- **更新文档**: 添加中文 PDF 处理章节，包括中文文本提取和中文 PDF 创建

## 使用方法

### 提取 Word 文档文本（支持中文）
```bash
python docx/scripts/extract_text.py document.docx
python docx/scripts/extract_text.py document.docx -o output.txt
python docx/scripts/extract_text.py document.docx --json
```

### 提取 PowerPoint 文本（支持中文）
```bash
python pptx/scripts/extract_pptx_text.py presentation.pptx
python pptx/scripts/extract_pptx_text.py presentation.pptx -o output.txt
python pptx/scripts/extract_pptx_text.py presentation.pptx --json
python pptx/scripts/extract_pptx_text.py presentation.pptx --markdown
```

### 解压 Office 文档（修复后的 unpack.py）
```bash
python docx/ooxml/scripts/unpack.py document.docx unpacked/
```

## 与原版的区别

| 功能 | 原版 | 此版本 |
|------|------|--------|
| 中文文档解压 | 乱码 (ASCII) | 正常 (UTF-8) |
| 中文文本提取 | 需外部工具 | 内置支持 |
| 编码检测 | 不支持 | 自动检测 UTF-8/GBK/GB18030 |
| 依赖 | pandoc/markitdown | 纯 Python 标准库 |

## 技术细节

### 编码修复
原版的 `unpack.py` 使用 `encoding="ascii"`，导致中文字符被转换为 XML 实体：
```python
# 原版
xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="ascii"))

# 修复后
xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
```

### 编码检测
新增的提取脚本会自动尝试多种编码：
```python
encodings_to_try = ['utf-8', 'gb18030', 'gbk', 'gb2312', 'utf-16']
```

## 目录结构

```
.
├── docx/           # Word 文档处理 skill
│   ├── ooxml/scripts/unpack.py      # 修复编码的解压脚本
│   ├── scripts/extract_text.py      # 中文 Word 提取
│   ├── scripts/extract_pptx_text.py # PPT 提取
│   └── SKILL.md                     # 更新后的文档
├── pptx/           # PowerPoint 处理 skill
│   ├── scripts/extract_text.py
│   ├── scripts/extract_pptx_text.py
│   └── SKILL.md
└── pdf/            # PDF 处理 skill
    └── SKILL.md                     # 中文 PDF 章节
```

## 许可证

与原 skill 相同（LICENSE.txt）
