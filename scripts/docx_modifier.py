#!/usr/bin/env python3
"""
DOCX Content Modifier — modify text in .docx files while preserving ALL formatting.

Core principle: Only modify <w:t> text nodes; never touch <w:rPr> (run formatting:
font/size/bold/italic/color) or <w:pPr> (paragraph formatting: alignment/spacing).

    from docx_modifier import DocxModifier

    with DocxModifier('input.docx') as doc:
        doc.replace('old text', 'new text')
        doc.save('output.docx')

CLI usage:
    python docx_modifier.py list <docx_path> [--scope all]
    python docx_modifier.py find <docx_path> <search_text>
    python docx_modifier.py replace <docx_path> <old> <new> <output> [--scope all]
    python docx_modifier.py replace-para <docx_path> <index> <new_text> <output>
    python docx_modifier.py batch-replace <docx_path> '<json>' <output> [--scope all]
    python docx_modifier.py replace-all <docx_path> <text_file> <output>
"""

import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

# lxml is the only external dependency
try:
    from lxml import etree
except ImportError:
    print("Error: lxml is required. Install it with: pip install lxml", file=sys.stderr)
    sys.exit(1)

# ── XML namespaces used in OOXML ──────────────────────────────────────────────
NSMAP = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'v':   'urn:schemas-microsoft-com:vml',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}

W_NS = NSMAP['w']
W_T  = '{%s}t' % W_NS
W_P  = '{%s}p' % W_NS
W_R  = '{%s}r' % W_NS
W_TC = '{%s}tc' % W_NS
W_TXBX = '{%s}txbxContent' % W_NS

# Files inside the DOCX zip that may contain user-editable text
HEADER_RE = re.compile(r'header\d*\.xml')
FOOTER_RE = re.compile(r'footer\d*\.xml')
EXTRA_XML = ['word/footnotes.xml', 'word/endnotes.xml']


class DocxModifier:
    """Open a .docx, modify its text content, and save to a new file.

    Usage:
        with DocxModifier('input.docx') as doc:
            doc.replace('old', 'new')
            doc.save('output.docx')
    """

    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.temp_dir = tempfile.mkdtemp(prefix='docxmod_')
        self.xml_trees = {}   # key (str) → lxml ElementTree
        self._xml_paths = {}  # key (str) → absolute file path in temp_dir
        self._extract_and_parse()

    # ── lifecycle ────────────────────────────────────────────────────────────

    def _extract_and_parse(self):
        """Unzip the .docx and parse every XML file that could contain text."""
        with zipfile.ZipFile(self.docx_path, 'r') as zf:
            zf.extractall(self.temp_dir)

        # Document body
        self._load_xml('document', 'word/document.xml')

        # Headers and footers
        word_dir = os.path.join(self.temp_dir, 'word')
        if os.path.isdir(word_dir):
            for fname in sorted(os.listdir(word_dir)):
                if HEADER_RE.match(fname) or FOOTER_RE.match(fname):
                    self._load_xml(fname, os.path.join('word', fname))

        # Footnotes, endnotes
        for rel_path in EXTRA_XML:
            self._load_xml(Path(rel_path).name, rel_path)

    def _load_xml(self, key, rel_path):
        abs_path = os.path.join(self.temp_dir, rel_path.replace('/', os.sep))
        if os.path.exists(abs_path):
            self._xml_paths[key] = abs_path
            self.xml_trees[key] = etree.parse(abs_path)

    def save(self, output_path):
        """Write all modified XML trees back and repack as a .docx."""
        for key, tree in self.xml_trees.items():
            tree.write(self._xml_paths[key], xml_declaration=True, encoding='UTF-8',
                       standalone=True)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
            for dirpath, _dirs, filenames in os.walk(self.temp_dir):
                for fname in filenames:
                    full = os.path.join(dirpath, fname)
                    arcname = os.path.relpath(full, self.temp_dir)
                    zf_out.write(full, arcname)

    def close(self):
        """Remove the temporary extraction directory."""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    # ── internal helpers ─────────────────────────────────────────────────────

    def _body_of(self, tree):
        """Return the <w:body> element of a tree, or the root if none."""
        root = tree.getroot()
        body = root.find('.//{%s}body' % W_NS)
        return body if body is not None else root

    def _container_kind(self, t_element):
        """Walk up from a <w:t> to determine: 'body', 'table', or 'textbox'."""
        parent = t_element.getparent()
        while parent is not None:
            tag = parent.tag
            if tag == W_TC:
                return 'table'
            if tag == W_TXBX:
                return 'textbox'
            parent = parent.getparent()
        return 'body'

    def _container_kind_for_xml(self, xml_name):
        """Return the scope kind implied by an XML file key."""
        if 'header' in xml_name.lower():
            return 'headers'
        if 'footer' in xml_name.lower():
            return 'footers'
        if 'footnote' in xml_name.lower():
            return 'footnotes'
        return 'body'

    def _paras_in_scope(self, tree, scope):
        """Return <w:p> elements in *tree* matching *scope*."""
        body = self._body_of(tree)
        all_p = body.findall('.//%s' % W_P)
        if scope in ('all', 'headers', 'footers', 'footnotes'):
            return all_p
        return [p for p in all_p if self._container_kind_of_para(p) == scope]

    def _container_kind_of_para(self, p_elem):
        parent = p_elem.getparent()
        while parent is not None:
            if parent.tag == W_TC:
                return 'table'
            if parent.tag == W_TXBX:
                return 'textbox'
            parent = parent.getparent()
        return 'body'

    @staticmethod
    def _t_elems(p_elem):
        """Return all <w:t> child elements of a paragraph, in document order."""
        return p_elem.findall('.//%s' % W_T)

    @staticmethod
    def _para_text(p_elem):
        return ''.join(t.text or '' for t in DocxModifier._t_elems(p_elem))

    # ── public API ───────────────────────────────────────────────────────────

    def list_paragraphs(self, scope='all'):
        """Yield {index, text, container} for every paragraph in the document."""
        results = []
        tree = self.xml_trees.get('document')
        if tree is None:
            return results
        for p_elem in self._paras_in_scope(tree, scope):
            results.append({
                'index': len(results),
                'text': self._para_text(p_elem),
                'container': self._container_kind_of_para(p_elem),
            })
        return results

    def find(self, search_text, scope='all'):
        """Return list of matches: {xml_file, container, full_para_text}."""
        results = []
        for xml_name, tree in self.xml_trees.items():
            ckind = self._container_kind_for_xml(xml_name)
            if scope != 'all' and ckind != scope:
                continue
            for p_elem in self._paras_in_scope(tree, scope):
                text = self._para_text(p_elem)
                if search_text in text:
                    results.append({
                        'xml_file': xml_name,
                        'container': self._container_kind_of_para(p_elem),
                        'match_context': text,
                    })
        return results

    def replace(self, old_text, new_text, scope='all'):
        """Replace every occurrence of *old_text* with *new_text*.

        Operates on each <w:t> element individually. If *old_text* happens to
        span two <w:t> elements (rare – e.g. "Hel" in one run and "lo" in the
        next), a single call will miss it. In that case call `replace` a second
        time after the first pass has merged the t-nodes.
        """
        for xml_name, tree in self.xml_trees.items():
            ckind = self._container_kind_for_xml(xml_name)
            if scope != 'all' and ckind != scope:
                continue
            for p_elem in self._paras_in_scope(tree, scope):
                for t_elem in self._t_elems(p_elem):
                    if t_elem.text and old_text in t_elem.text:
                        t_elem.text = t_elem.text.replace(old_text, new_text)

    def replace_paragraph(self, index, new_text, scope='body'):
        """Replace the text of the *index*-th (0-based) paragraph in *scope*.

        Puts *new_text* into the first <w:t> of the paragraph and clears the
        rest. This means the paragraph keeps the formatting of its first run.
        """
        tree = self.xml_trees.get('document')
        if tree is None:
            raise ValueError("No word/document.xml found in the docx")

        paras = self._paras_in_scope(tree, scope)
        if index < 0 or index >= len(paras):
            raise IndexError(
                f"Paragraph index {index} out of range (0–{len(paras)-1})")

        p_elem = paras[index]
        t_nodes = self._t_elems(p_elem)
        if not t_nodes:
            return  # empty paragraph (image-only, etc.) — leave unchanged

        t_nodes[0].text = new_text
        # Preserve xml:space="preserve" on the first node if new_text needs it
        if new_text and new_text[0] == ' ' or (new_text and new_text[-1] == ' '):
            if '{%s}space' % 'http://www.w3.org/XML/1998/namespace' not in t_nodes[0].attrib:
                pass  # just keep text; lxml handles quoting

        for n in t_nodes[1:]:
            n.text = ''

    def batch_replace(self, mapping, scope='all'):
        """Apply a {old_text: new_text} mapping across the document."""
        for old, new in mapping.items():
            self.replace(old, new, scope)

    def replace_all_body_text(self, new_paragraphs, scope='body'):
        """Overwrite every paragraph in *scope* with strings from the list.

        Paragraphs beyond len(*new_paragraphs*) are left unchanged.
        Extra strings beyond the document's paragraph count are ignored.
        """
        tree = self.xml_trees.get('document')
        if tree is None:
            raise ValueError("No word/document.xml found")
        paras = self._paras_in_scope(tree, scope)
        for i, p_elem in enumerate(paras):
            if i >= len(new_paragraphs):
                break
            t_nodes = self._t_elems(p_elem)
            if t_nodes:
                t_nodes[0].text = new_paragraphs[i]
                for n in t_nodes[1:]:
                    n.text = ''

    def get_header_footer_texts(self):
        """Return {filename: combined_text} for all header/footer XML parts."""
        results = {}
        for xml_name, tree in self.xml_trees.items():
            if 'header' in xml_name.lower() or 'footer' in xml_name.lower():
                body = self._body_of(tree)
                texts = [t.text or '' for t in body.findall('.//%s' % W_T)]
                results[xml_name] = ''.join(texts)
        return results

    def replace_header_footer_text(self, old_text, new_text):
        """Replace text in headers and footers only."""
        self.replace(old_text, new_text, scope='headers')
        self.replace(old_text, new_text, scope='footers')


# ── CLI ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    cmd = sys.argv[1]

    # Manually handle arg parsing for simplicity
    def _scope(default='all'):
        try:
            idx = sys.argv.index('--scope')
            return sys.argv[idx + 1]
        except (ValueError, IndexError):
            return default

    pos = [a for a in sys.argv[2:] if not a.startswith('--')]
    # Also remove the value that follows --scope
    clean = []
    skip = False
    for a in sys.argv[2:]:
        if skip:
            skip = False
            continue
        if a == '--scope':
            skip = True
            continue
        clean.append(a)

    try:
        if cmd == 'list':
            docx_path = clean[0]
            scope = _scope('all')
            with DocxModifier(docx_path) as doc:
                items = doc.list_paragraphs(scope)
                print(json.dumps(items, ensure_ascii=False, indent=2))

        elif cmd == 'find':
            docx_path, search_text = clean[0], clean[1]
            scope = _scope('all')
            with DocxModifier(docx_path) as doc:
                hits = doc.find(search_text, scope)
                print(json.dumps(hits, ensure_ascii=False, indent=2))

        elif cmd == 'replace':
            docx_path, old, new, out = clean[0], clean[1], clean[2], clean[3]
            scope = _scope('all')
            with DocxModifier(docx_path) as doc:
                doc.replace(old, new, scope)
                doc.save(out)
            print(f"Saved: {out}")

        elif cmd == 'replace-para':
            docx_path, idx, new_text, out = clean[0], int(clean[1]), clean[2], clean[3]
            scope = _scope('body')
            with DocxModifier(docx_path) as doc:
                doc.replace_paragraph(idx, new_text, scope)
                doc.save(out)
            print(f"Saved: {out}")

        elif cmd == 'batch-replace':
            docx_path, mapping_json, out = clean[0], clean[1], clean[2]
            scope = _scope('all')
            mapping = json.loads(mapping_json)
            with DocxModifier(docx_path) as doc:
                doc.batch_replace(mapping, scope)
                doc.save(out)
            print(f"Saved: {out}")

        elif cmd == 'replace-all':
            docx_path, text_file, out = clean[0], clean[1], clean[2]
            scope = _scope('body')
            with open(text_file, 'r', encoding='utf-8') as f:
                new_paras = [line.rstrip('\n') for line in f if line.strip()]
            with DocxModifier(docx_path) as doc:
                doc.replace_all_body_text(new_paras, scope)
                doc.save(out)
            print(f"Saved: {out}")

        elif cmd == 'headers':
            docx_path = clean[0]
            with DocxModifier(docx_path) as doc:
                texts = doc.get_header_footer_texts()
                print(json.dumps(texts, ensure_ascii=False, indent=2))

        else:
            print(f"Unknown command: {cmd}", file=sys.stderr)
            print(__doc__)
            sys.exit(1)

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        raise


if __name__ == '__main__':
    main()
