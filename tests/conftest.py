"""Shared pytest fixtures for chinese-paper-format-skill tests."""

import io
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Pt
from lxml import etree

# OOXML namespaces
WML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
FOOTNOTE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'


def _add_footnote_reference(para_element, footnote_id: int):
    """Add a footnote reference run to a paragraph XML element."""
    r_elem = etree.SubElement(para_element, f'{{{WML}}}r')
    rPr = etree.SubElement(r_elem, f'{{{WML}}}rPr')
    va = etree.SubElement(rPr, f'{{{WML}}}vertAlign')
    va.set(f'{{{WML}}}val', 'superscript')
    fn_ref = etree.SubElement(r_elem, f'{{{WML}}}footnoteReference')
    fn_ref.set(f'{{{WML}}}id', str(footnote_id))


def _build_footnote(fn_id: int, text: str):
    """Build a footnote XML element with given id and text."""
    fn = etree.Element(f'{{{WML}}}footnote')
    fn.set(f'{{{WML}}}id', str(fn_id))
    p = etree.SubElement(fn, f'{{{WML}}}p')

    # Left bracket
    r0 = etree.SubElement(p, f'{{{WML}}}r')
    t0 = etree.SubElement(r0, f'{{{WML}}}t')
    t0.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t0.text = '['

    # Auto-number
    r1 = etree.SubElement(p, f'{{{WML}}}r')
    etree.SubElement(r1, f'{{{WML}}}footnoteRef')

    # Right bracket + citation text
    r2 = etree.SubElement(p, f'{{{WML}}}r')
    t2 = etree.SubElement(r2, f'{{{WML}}}t')
    t2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t2.text = ']' + text

    return fn


def _inject_footnotes(docx_bytes: bytes, footnote_xml_elements: list) -> bytes:
    """Inject footnote XML elements into a .docx ZIP archive."""
    with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as zin:
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == 'word/document.xml':
                    doc_xml = etree.fromstring(data)
                    body = doc_xml.find(f'{{{WML}}}body')
                    first_p = body.find(f'{{{WML}}}p')
                    if first_p is not None:
                        for fid in range(1, len(footnote_xml_elements) + 1):
                            _add_footnote_reference(first_p, fid)
                    data = etree.tostring(doc_xml, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)

                elif item.filename == 'word/_rels/document.xml.rels':
                    rels_xml = etree.fromstring(data)
                    fn_rel = etree.SubElement(rels_xml, f'{{{REL}}}Relationship')
                    fn_rel.set('Id', 'rIdFootnotes')
                    fn_rel.set('Type', FOOTNOTE_REL)
                    fn_rel.set('Target', 'footnotes.xml')
                    data = etree.tostring(rels_xml, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)

                zout.writestr(item, data)

            # Build footnotes.xml
            fns_root = etree.Element(f'{{{WML}}}footnotes',
                                      nsmap={'w': WML})
            # Separator footnotes
            for sep_id, sep_tag in [('-1', 'separator'), ('0', 'continuationSeparator')]:
                sep = etree.SubElement(fns_root, f'{{{WML}}}footnote')
                sep.set(f'{{{WML}}}id', sep_id)
                sp = etree.SubElement(sep, f'{{{WML}}}p')
                sr = etree.SubElement(sp, f'{{{WML}}}r')
                etree.SubElement(sr, f'{{{WML}}}{sep_tag}')

            for elem in footnote_xml_elements:
                fns_root.append(elem)

            zout.writestr('word/footnotes.xml',
                          etree.tostring(fns_root, xml_declaration=True,
                                         encoding='UTF-8', standalone=True))

            # Update Content_Types
            ct_xml = etree.fromstring(zin.read('[Content_Types].xml'))
            override = etree.SubElement(ct_xml, f'{{{CT}}}Override')
            override.set('PartName', '/word/footnotes.xml')
            override.set('ContentType',
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml')
            zout.writestr('[Content_Types].xml',
                          etree.tostring(ct_xml, xml_declaration=True,
                                         encoding='UTF-8', standalone=True))

    return out_buf.getvalue()


# ── Fixtures ──────────────────────────────────────────────────────────


@pytest.fixture
def minimal_docx_path(tmp_path):
    """Create a minimal .docx: title + level-1 heading + body text."""
    doc = Document()
    doc.add_paragraph('测试论文题目')
    doc.add_paragraph('一、引言')
    doc.add_paragraph('这是正文内容，包含一些中文文本和数字123。')

    path = str(tmp_path / 'minimal.docx')
    doc.save(path)
    return path


@pytest.fixture
def docx_with_abstract(tmp_path):
    """Create a .docx with abstract and keywords."""
    doc = Document()
    doc.add_paragraph('担保型以物抵债规则适用研究')
    doc.add_paragraph('【摘要】本文研究了以物抵债的法律适用问题。')
    doc.add_paragraph('【关键词】以物抵债；让与担保；强制清算')
    doc.add_paragraph('一、绪论')
    doc.add_paragraph('这是正文内容。')

    path = str(tmp_path / 'with_abstract.docx')
    doc.save(path)
    return path


@pytest.fixture
def docx_with_footnotes(tmp_path):
    """Create a .docx with 3 footnotes containing citation errors."""
    doc = Document()
    doc.add_paragraph('试论人工智能的侵权责任')
    doc.add_paragraph('一、问题的提出')
    p1 = doc.add_paragraph('人工智能侵权责任问题日益突出。')
    p1.runs[0].font.size = Pt(10.5)

    # Save and inject footnotes
    buf = io.BytesIO()
    doc.save(buf)
    raw = buf.getvalue()

    fn1 = _build_footnote(1, '王利明， 《侵权责任法研究》, 中国人民大学出版社 2016 年, 第 125 页.')
    fn2 = _build_footnote(2, '张新宝：《侵权法》，中国人民大学出版社 2010 年版，p.89.')
    fn3 = _build_footnote(3, '《民法典》第68条。')

    modified = _inject_footnotes(raw, [fn1, fn2, fn3])

    path = str(tmp_path / 'with_footnotes.docx')
    with open(path, 'wb') as f:
        f.write(modified)
    return path


@pytest.fixture
def docx_with_legal_citations(tmp_path):
    """Create a .docx with legal document and case citations."""
    doc = Document()
    doc.add_paragraph('法律文件与案例引用测试')
    doc.add_paragraph('一、法律文件引用')
    doc.add_paragraph('测试规范性文件引用。')

    buf = io.BytesIO()
    doc.save(buf)
    raw = buf.getvalue()

    fn1 = _build_footnote(1, '《国务院关于建立农村最低生活保障制度的通知》，国发[2007]19号。')
    fn2 = _build_footnote(2, '田永诉北京科技大学案，北京市海淀区人民法院[1998]海行初字第142号行政判决书。')
    fn3 = _build_footnote(3, '《最高人民法院关于适用〈行政诉讼法〉的解释》，法释[2018]1号，第100条。')

    modified = _inject_footnotes(raw, [fn1, fn2, fn3])

    path = str(tmp_path / 'with_legal.docx')
    with open(path, 'wb') as f:
        f.write(modified)
    return path
