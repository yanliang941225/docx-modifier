---
name: docx-modifier
description: >
  Modify text content in .docx (Word) files while strictly preserving ALL original
  formatting — fonts, sizes, colors, bold/italic, alignment, spacing, table borders,
  headers, footers, page layout, and styles. Use this skill whenever the user wants to
  edit, change, replace, add to, or update text in a Word document without altering its
  appearance. Triggers on: editing .docx files, Word documents, replacing text in Word,
  updating document content, filling in templates, batch replacing, 修改Word文档,
  替换文档内容, 保留格式, 修改docx文件, 填写Word模板. Always invoke before attempting
  to read or edit any .docx file where content changes are needed — even if the user
  doesn't explicitly say "preserve formatting."
---

# DOCX Content Modifier

Modify text in .docx files while keeping every aspect of the original formatting intact.

## Core principle

A .docx file is a ZIP archive of XML files. Text content lives in `<w:t>` elements, while formatting lives in separate elements (`<w:rPr>` for run properties like font/size/color, `<w:pPr>` for paragraph properties like alignment/spacing). **Only modify `<w:t>` text nodes — never touch formatting nodes.** This is the single rule that guarantees format preservation.

```
<w:p>                       <!-- paragraph container -->
  <w:pPr>                   <!-- PARAGRAPH FORMAT: alignment, spacing → DO NOT TOUCH -->
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>                     <!-- run -->
    <w:rPr>                 <!-- RUN FORMAT: font, size, bold, color → DO NOT TOUCH -->
      <w:rFonts w:ascii="仿宋_GB2312"/>
      <w:sz w:val="30"/>
      <w:b/>
    </w:rPr>
    <w:t>text to replace</w:t>  <!-- ONLY CHANGE THIS -->
  </w:r>
</w:p>
```

## Bundled script

Use `scripts/docx_modifier.py` for all operations. It encapsulates the XML manipulation so you don't need to write it from scratch each time.

```python
from docx_modifier import DocxModifier
```

### Quick reference

| Operation | Code |
|-----------|------|
| List paragraphs | `doc.list_paragraphs(scope='all')` |
| Find text | `doc.find('search text', scope='all')` |
| Replace all occurrences | `doc.replace('old', 'new', scope='all')` |
| Replace by paragraph index | `doc.replace_paragraph(3, 'new text', scope='body')` |
| Batch replace from mapping | `doc.batch_replace({'a':'x', 'b':'y'}, scope='all')` |
| Replace all body paragraphs | `doc.replace_all_body_text([...])` |
| Get header/footer text | `doc.get_header_footer_texts()` |
| Replace in headers/footers | `doc.replace_header_footer_text('old', 'new')` |
| Save | `doc.save('output.docx')` |

### Scope parameter

Controls which parts of the document are modified:
- `'all'` — everything (body, tables, headers, footers, text boxes)
- `'body'` — main body paragraphs only (not inside tables)
- `'tables'` — only text inside table cells
- `'headers'` — only headers
- `'footers'` — only footers

## Workflow by use case

### 1. Find and replace text (most common)

User says: "Replace 'ABC Corp' with 'XYZ Ltd' in this contract."

```python
from scripts.docx_modifier import DocxModifier

with DocxModifier('contract.docx') as doc:
    doc.replace('ABC Corp', 'XYZ Ltd', scope='all')
    doc.save('contract_updated.docx')
```

The `replace()` method searches all `<w:t>` elements across the entire document (body, tables, headers, footers) and does `str.replace()` on each element that contains the search text. Each `<w:t>` keeps its original formatting because only the text string changes.

### 2. Replace a specific paragraph by index

User says: "Change the third paragraph to say..."

```python
with DocxModifier('document.docx') as doc:
    doc.replace_paragraph(2, 'This is the new third paragraph text.', scope='body')
    doc.save('output.docx')
```

Paragraph index is 0-based. The first `<w:t>` in the paragraph gets the new text; all subsequent `<w:t>` elements in that paragraph are cleared. The formatting from the first run (font, size, bold, etc.) is preserved because only `<w:t>.text` changes.

### 3. Batch replace with a mapping

User says: "Replace the company name, date, and amount throughout the document."

```python
mapping = {
    '{{COMPANY}}': 'XYZ Ltd',
    '{{DATE}}': '2026年5月18日',
    '{{AMOUNT}}': '¥50,000',
}

with DocxModifier('template.docx') as doc:
    doc.batch_replace(mapping, scope='all')
    doc.save('filled_template.docx')
```

### 4. Replace all body text

User says: "Overwrite the entire document content with this new text."

```python
new_content = [
    '第一章 总则',
    '第一条 本条例适用于...',
    '第二条 ...',
]

with DocxModifier('template.docx') as doc:
    doc.replace_all_body_text(new_content, scope='body')
    doc.save('new_document.docx')
```

Each string in the list replaces one paragraph. If there are more paragraphs in the document than items in the list, extra paragraphs are left unchanged. If there are more items than paragraphs, extra items are ignored.

### 5. Modify headers and footers

```python
with DocxModifier('document.docx') as doc:
    # See what's in headers/footers
    texts = doc.get_header_footer_texts()
    print(texts)
    
    # Replace text in headers and footers
    doc.replace_header_footer_text('旧页眉文本', '新页眉文本')
    doc.save('output.docx')
```

### 6. Handle references / bibliography

When modifying reference entries, copy the formatting from the first reference entry:

```python
import copy
from lxml import etree

NSMAP = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

with DocxModifier('paper.docx') as doc:
    tree = doc.xml_trees['document']
    body = tree.find('.//w:body', NSMAP)
    all_p = body.findall('.//w:p', NSMAP)
    
    # Find reference paragraphs (usually after "参考文献" heading)
    ref_start = None
    for i, p in enumerate(all_p):
        texts = [t.text or '' for t in p.findall('.//w:t', NSMAP)]
        if '参考文献' in ''.join(texts):
            ref_start = i + 1
            break
    
    if ref_start:
        # Clone the first reference entry as a format template
        template_p = copy.deepcopy(all_p[ref_start])
        
        new_refs = [
            '张三, 李四. 基于AI的文档处理研究[J]. 计算机学报, 2025.',
            '王五, 赵六. 自然语言处理进展综述[J]. 软件学报, 2024.',
        ]
        
        for j, ref_text in enumerate(new_refs):
            if ref_start + j < len(all_p):
                # Replace text in existing paragraph (keeps its own formatting)
                doc.replace_paragraph(ref_start + j, ref_text, scope='body')
    
    doc.save('paper_updated.docx')
```

## Common XML patterns in DOCX

These are reference patterns for when you need to do something the script doesn't handle directly:

### Table structure
```xml
<w:tbl>
  <w:tblPr><!-- table properties (borders, width) --></w:tblPr>
  <w:tblGrid><!-- column widths --></w:tblGrid>
  <w:tr>  <!-- row -->
    <w:tc>  <!-- cell -->
      <w:tcPr><!-- cell properties --></w:tcPr>
      <w:p>  <!-- paragraph inside cell -->
        <w:r><w:t>cell text</w:t></w:r>
      </w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

### Text with mixed formatting (multiple runs)
```xml
<w:p>
  <w:r><w:rPr><w:b/></w:rPr><w:t>Bold text </w:t></w:r>
  <w:r><w:rPr><w:i/></w:rPr><w:t>italic text</w:t></w:r>
</w:p>
```
When replacing text in such paragraphs, the new text inherits the formatting of the first run. If you need to preserve mixed formatting, manually split new text across runs.

### Empty paragraphs
Some paragraphs contain no `<w:t>` elements (just images, fields, or are truly empty). `replace_paragraph` silently does nothing for these — this is intentional to avoid corrupting non-text content.

## Rules

1. **Never modify `<w:rPr>` elements** — they control font, size, bold, italic, color, underline
2. **Never modify `<w:pPr>` elements** — they control alignment, spacing, indentation
3. **Never modify `<w:tblPr>` or `<w:tcPr>`** — they control table formatting
4. **Only set `.text` on `<w:t>` elements** — don't add/remove attributes
5. **Always save to a new file** — never overwrite the original unless the user explicitly asks
6. **When in doubt, use `doc.list_paragraphs()` first** to see what's in the document before modifying
7. **For references/bibliography**, clone the format of the first entry rather than constructing XML manually
8. **Use `scope` parameter** to limit changes to the right part of the document

## Dependencies

- Python 3
- `lxml` (`pip install lxml`)
- `zipfile` (stdlib)
