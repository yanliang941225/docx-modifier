# DOCX Modifier

**修改 Word 文档文字内容，严格保留所有原始格式。**

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## 这是什么

一个 Python 库，用于修改 `.docx`（Word）文件中的文字内容，同时**100% 保留**原有格式——字体、字号、加粗/斜体、颜色、对齐方式、段落间距、表格边框、页眉页脚、页面布局等一切样式。

常见的应用场景：
- **模板填充**：把合同/报告模板中的占位符替换为实际内容
- **文本批量替换**：全文查找替换公司名称、日期、金额等
- **按段落修改**：精确替换某个段落的文字
- **格式无损编辑**：任何需要改文字但绝不能动格式的场景

## 核心原理

`.docx` 文件本质是一个 ZIP 压缩包，里面包含多个 XML 文件。文字和格式是分离存储的：

```
<w:p>                          <!-- 段落容器 -->
  <w:pPr>                      <!-- 段落格式：对齐/间距 → 不碰 -->
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>                        <!-- 文本运行(run) -->
    <w:rPr>                    <!-- 运行格式：字体/字号/粗体 → 不碰 -->
      <w:rFonts w:ascii="仿宋_GB2312"/>
      <w:sz w:val="30"/>
      <w:b/>
    </w:rPr>
    <w:t>要替换的文字</w:t>     <!-- 只改这里 -->
  </w:r>
</w:p>
```

**只修改 `<w:t>` 节点的文字，绝不触碰 `<w:rPr>` 和 `<w:pPr>` 格式节点。** 这就是输出文档与原始文档格式完全一致的原因。

## 安装

```bash
pip install lxml
```

然后将 `scripts/docx_modifier.py` 复制到你的项目中即可。仅依赖 Python 标准库 + `lxml`。

## 快速开始

### Python API

```python
from docx_modifier import DocxModifier

# 1. 查找替换（最常用）
with DocxModifier('合同.docx') as doc:
    doc.replace('甲方公司', '北京科技有限公司')
    doc.save('合同_已修改.docx')

# 2. 批量替换（模板填充）
mapping = {
    '{{公司名称}}': '北京科技有限公司',
    '{{日期}}': '2026年5月18日',
    '{{金额}}': '¥500,000',
}
with DocxModifier('模板.docx') as doc:
    doc.batch_replace(mapping, scope='all')
    doc.save('已填充.docx')

# 3. 按段落索引替换
with DocxModifier('报告.docx') as doc:
    doc.replace_paragraph(0, '2026年度财务报告')  # 替换标题（第0段）
    doc.replace_paragraph(3, '新的第三段内容')
    doc.save('报告_新标题.docx')

# 4. 查看文档结构
with DocxModifier('文档.docx') as doc:
    for p in doc.list_paragraphs('all'):
        print(f"[{p['index']}] ({p['container']}) {p['text'][:60]}")

# 5. 查找文本位置
with DocxModifier('文档.docx') as doc:
    hits = doc.find('关键词')
    for h in hits:
        print(f"找到于 {h['xml_file']}: {h['match_context'][:50]}")

# 6. 修改页眉页脚
with DocxModifier('文档.docx') as doc:
    doc.replace_header_footer_text('旧页眉文字', '新页眉文字')
    doc.save('文档_新页眉.docx')
```

### 命令行

```bash
# 列出所有段落
python docx_modifier.py list 文档.docx --scope all

# 查找文本
python docx_modifier.py find 文档.docx "关键词"

# 全文替换
python docx_modifier.py replace 文档.docx "旧文字" "新文字" 输出.docx --scope all

# 按段落索引替换
python docx_modifier.py replace-para 文档.docx 0 "新标题" 输出.docx

# 批量替换（JSON 映射）
python docx_modifier.py batch-replace 文档.docx '{"旧1":"新1","旧2":"新2"}' 输出.docx

# 整文替换（从文件读取）
python docx_modifier.py replace-all 文档.docx 新内容.txt 输出.docx

# 查看页眉页脚
python docx_modifier.py headers 文档.docx
```

## API 参考

### `DocxModifier(docx_path)`

打开一个 `.docx` 文件。支持上下文管理器：

```python
with DocxModifier('input.docx') as doc:
    # ... 修改操作 ...
    doc.save('output.docx')
# 退出时自动清理临时文件
```

### 方法

| 方法 | 说明 |
|------|------|
| `list_paragraphs(scope='all')` | 列出所有段落，返回 `[{index, text, container}]` |
| `find(search_text, scope='all')` | 查找文本出现的位置 |
| `replace(old, new, scope='all')` | 全文查找替换 |
| `replace_paragraph(index, new_text, scope='body')` | 按索引替换段落文字 |
| `batch_replace(mapping, scope='all')` | 批量查找替换，`mapping` 为 `{旧: 新}` 字典 |
| `replace_all_body_text(paragraphs, scope='body')` | 用字符串列表覆盖所有段落 |
| `replace_header_footer_text(old, new)` | 替换页眉页脚中的文字 |
| `get_header_footer_texts()` | 获取所有页眉页脚文字 |
| `save(output_path)` | 保存修改到新的 `.docx` 文件 |

### `scope` 参数

控制修改范围：

| 值 | 范围 |
|----|------|
| `'all'` | 全文——正文 + 表格 + 页眉页脚 + 文本框（默认） |
| `'body'` | 仅正文段落（不含表格内的文字） |
| `'tables'` | 仅表格内的文字 |
| `'headers'` | 仅页眉 |
| `'footers'` | 仅页脚 |

## 技术细节

### 处理范围

修改操作覆盖以下 XML 文件：

| 文件 | 内容 |
|------|------|
| `word/document.xml` | 正文、表格、文本框 |
| `word/header*.xml` | 页眉 |
| `word/footer*.xml` | 页脚 |
| `word/footnotes.xml` | 脚注 |
| `word/endnotes.xml` | 尾注 |

### 格式保护机制

- **段落格式** `<w:pPr>`：对齐方式、缩进、间距、行距 → 原封不动
- **运行格式** `<w:rPr>`：字体、字号、粗体、斜体、下划线、颜色 → 原封不动
- **表格格式** `<w:tblPr>`, `<w:tcPr>`：边框、列宽、单元格样式 → 原封不动
- **页面设置**：页边距、纸张大小、分栏 → 原封不动（不在修改范围内）

### 参考文献处理

当需要修改参考文献条目时，使用 `copy.deepcopy` 复制第一条参考文献的格式作为模板，然后逐条填充新内容。参考 SKILL.md 中的示例。

## 依赖

- **Python** ≥ 3.8
- **lxml** — XML 解析（`pip install lxml`）
- **zipfile** — ZIP 处理（Python 标准库，无需安装）

## 限制

- 不支持跨 `<w:t>` 元素的文字替换（如"Hel"在第一个 run、"lo"在第二个 run 的"Hello"）。此场景在实际文档中极少出现。如遇到，调用两次 `replace()` 即可。
- 不修改图片、图表、公式等非文字内容。
- 每次保存生成新文件，不会覆盖原始文件（这是设计意图）。

## 对比其他方案

| 方案 | 格式保留 | 表格支持 | 页眉页脚 | 依赖 |
|------|---------|---------|---------|------|
| **docx-modifier** | 100% | 是 | 是 | lxml |
| `python-docx` | 部分（会丢失部分格式） | 支持 | 支持 | python-docx |
| 手动 `zipfile` + 字符串替换 | 100% | 是 | 需手写 | 无 |

`python-docx` 的问题在于它会重新序列化 XML，导致某些格式属性（尤其是中文排版相关的属性）丢失。而直接操作 XML 字符串或使用 `lxml` 精准修改 `<w:t>` 节点可以完全避免此问题。

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request。请确保：
- 新增功能保持格式保护的核心理念
- 通过现有测试用例
- 为重要功能添加测试
