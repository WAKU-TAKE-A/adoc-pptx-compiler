import importlib.util
import os
import sys

import pytest
from pptx import Presentation
from pptx.util import Inches

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SPEC = importlib.util.spec_from_file_location(
    "adoc_pptx_compiler", os.path.join(ROOT, "adoc-pptx-compiler.py"))
mod = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(mod)

EMU_IN = 914400

DOCBOOK = """<?xml version="1.0" encoding="UTF-8"?>
<article>
  <articleinfo>
    <title>テスト文書</title>
    <author><firstname>花子</firstname><surname>佐藤</surname></author>
    <date>2026-01-01</date>
  </articleinfo>
  <section>
    <title>最初の節</title>
    <simpara>本文の段落です。<emphasis role="strong">太字</emphasis>を含みます。
    <footnote><simpara>脚注テキスト</simpara></footnote></simpara>
    <itemizedlist>
      <listitem><simpara>項目A</simpara></listitem>
      <listitem><simpara>項目B</simpara></listitem>
    </itemizedlist>
  </section>
  <section>
    <title>表の節</title>
    <informaltable>
      <tgroup cols="2">
        <colspec colwidth="1*"/><colspec colwidth="2*"/>
        <thead><row><entry>見出し1</entry><entry>見出し2</entry></row></thead>
        <tbody><row><entry>セル1</entry><entry>セル2</entry></row></tbody>
      </tgroup>
    </informaltable>
    <programlisting language="python">print("hi")</programlisting>
  </section>
</article>
"""


def all_texts(prs):
    texts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        texts.append(cell.text)
    return texts


def test_version():
    assert mod.__version__ == "0.9.2.0"


def test_app_base_dir_script_mode():
    assert mod.app_base_dir() == ROOT


def test_app_base_dir_frozen_mode(tmp_path, monkeypatch):
    exe = tmp_path / "adoc-pptx-compiler.exe"
    exe.write_bytes(b"")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe))
    assert mod.app_base_dir() == str(tmp_path)


def test_compile_from_docbook(tmp_path):
    xml = tmp_path / "doc.xml"
    xml.write_text(DOCBOOK, encoding="utf-8")
    out = tmp_path / "out.pptx"
    report = mod.compile_adoc(str(xml), str(out))
    assert out.is_file()
    assert report["slides"] == 3  # title + 2 sections
    prs = Presentation(str(out))
    texts = all_texts(prs)
    joined = "\n".join(texts)
    assert "テスト文書" in joined
    assert "太字" in joined
    assert "項目A" in joined
    assert "見出し1" in joined and "セル2" in joined
    assert 'print("hi")' in joined
    assert "脚注テキスト" in joined


def test_shapes_within_slide(tmp_path):
    xml = tmp_path / "doc.xml"
    xml.write_text(DOCBOOK, encoding="utf-8")
    out = tmp_path / "out.pptx"
    mod.compile_adoc(str(xml), str(out))
    prs = Presentation(str(out))
    for slide in prs.slides:
        for shape in slide.shapes:
            assert shape.left >= -1000
            assert shape.top >= -1000
            assert shape.left + shape.width <= prs.slide_width + 1000
            assert shape.top + shape.height <= prs.slide_height + 1000


def test_overflow_makes_continuation(tmp_path):
    items = "".join(
        f"<listitem><simpara>長い項目 {i:02d} のテキストです</simpara></listitem>"
        for i in range(40))
    xml_text = f"""<?xml version="1.0" encoding="UTF-8"?>
<article><articleinfo><title>長文</title></articleinfo>
<section><title>長い節</title>
<itemizedlist>{items}</itemizedlist></section></article>"""
    xml = tmp_path / "long.xml"
    xml.write_text(xml_text, encoding="utf-8")
    out = tmp_path / "long.pptx"
    report = mod.compile_adoc(str(xml), str(out))
    assert report["slides"] >= 3  # title + 2+ pages for the section
    prs = Presentation(str(out))
    joined = "\n".join(all_texts(prs))
    assert "長い節（続き）" in joined
    for i in range(40):
        assert f"長い項目 {i:02d}" in joined


def test_overflow_error_mode(tmp_path):
    items = "".join(
        f"<listitem><simpara>項目 {i}</simpara></listitem>" for i in range(40))
    xml_text = f"""<?xml version="1.0" encoding="UTF-8"?>
<article><articleinfo><title>長文</title></articleinfo>
<section><title>長い節</title>
<itemizedlist>{items}</itemizedlist></section></article>"""
    xml = tmp_path / "long.xml"
    xml.write_text(xml_text, encoding="utf-8")
    with pytest.raises(OverflowError):
        mod.compile_adoc(str(xml), str(tmp_path / "x.pptx"), overflow="error")


def test_missing_image_warns(tmp_path):
    xml_text = """<?xml version="1.0" encoding="UTF-8"?>
<article><articleinfo><title>画像</title></articleinfo>
<section><title>画像の節</title>
<mediaobject><imageobject><imagedata fileref="no_such.png"/></imageobject>
</mediaobject></section></article>"""
    xml = tmp_path / "img.xml"
    xml.write_text(xml_text, encoding="utf-8")
    report = mod.compile_adoc(str(xml), str(tmp_path / "img.pptx"))
    assert any("image not found" in w for w in report["warnings"])


def test_frontend_asciidoc_py(tmp_path):
    src = os.path.join(ROOT, "samples", "sample.adoc")
    out = tmp_path / "sample.pptx"
    report = mod.compile_adoc(src, str(out), frontend="asciidoc-py")
    assert out.is_file()
    assert report["slides"] >= 5
    prs = Presentation(str(out))
    joined = "\n".join(all_texts(prs))
    assert "サンプル文書" in joined
    assert "長い節の項目 20" in joined


BOOK = """<?xml version="1.0" encoding="UTF-8"?>
<book>
  <bookinfo><title>本の文書</title></bookinfo>
  <preface><title>まえがき</title><simpara>前文のテキスト</simpara></preface>
  <chapter>
    <title>第一章</title>
    <simpara>章の本文</simpara>
    <section><title>小節</title><simpara>小節の本文</simpara></section>
    <literallayout>リテラル行1
リテラル行2</literallayout>
    <variablelist>
      <varlistentry><term>用語A</term>
        <listitem><simpara>用語Aの定義</simpara></listitem></varlistentry>
    </variablelist>
  </chapter>
</book>
"""


def test_book_chapters_and_blocks(tmp_path):
    xml = tmp_path / "book.xml"
    xml.write_text(BOOK, encoding="utf-8")
    out = tmp_path / "book.pptx"
    report = mod.compile_adoc(str(xml), str(out))
    assert report["slides"] == 4  # title + TOC + preface + chapter
    assert report["orientation"] == "portrait"
    assert report["toc"] is True
    assert not any(u.startswith("block:") for u in report["unsupported"])
    prs = Presentation(str(out))
    joined = "\n".join(all_texts(prs))
    assert "前文のテキスト" in joined
    assert "章の本文" in joined and "小節の本文" in joined
    assert "リテラル行1" in joined and "リテラル行2" in joined
    assert "用語A" in joined and "用語Aの定義" in joined


def test_book_is_portrait_article_is_landscape(tmp_path):
    xml = tmp_path / "book.xml"
    xml.write_text(BOOK, encoding="utf-8")
    out = tmp_path / "book.pptx"
    mod.compile_adoc(str(xml), str(out))
    prs = Presentation(str(out))
    assert prs.slide_height > prs.slide_width
    for slide in prs.slides:
        for shape in slide.shapes:
            assert shape.left + shape.width <= prs.slide_width + 1000
            assert shape.top + shape.height <= prs.slide_height + 1000
    xml2 = tmp_path / "doc.xml"
    xml2.write_text(DOCBOOK, encoding="utf-8")
    out2 = tmp_path / "doc.pptx"
    report = mod.compile_adoc(str(xml2), str(out2))
    assert report["orientation"] == "landscape"
    assert report["toc"] is False
    prs2 = Presentation(str(out2))
    assert prs2.slide_width > prs2.slide_height


def test_orientation_override(tmp_path):
    xml = tmp_path / "book.xml"
    xml.write_text(BOOK, encoding="utf-8")
    out = tmp_path / "book.pptx"
    report = mod.compile_adoc(str(xml), str(out),
                              orientation="landscape", toc="never")
    assert report["orientation"] == "landscape"
    assert report["slides"] == 3


def test_toc_page_numbers(tmp_path):
    xml = tmp_path / "doc.xml"
    xml.write_text(DOCBOOK, encoding="utf-8")
    out = tmp_path / "toc.pptx"
    report = mod.compile_adoc(str(xml), str(out), toc="always")
    assert report["slides"] == 4  # title + TOC + 2 sections
    prs = Presentation(str(out))
    toc_slide = list(prs.slides)[1]
    texts = [s.text_frame.text for s in toc_slide.shapes
             if s.has_text_frame]
    joined = "\n".join(texts)
    assert "目次" in joined
    assert "最初の節" in joined and "表の節" in joined
    assert "3" in texts and "4" in texts  # slide numbers of the sections


def test_design_override(tmp_path):
    toml = tmp_path / "design.toml"
    toml.write_text('[colors]\nband = "FF0000"\n[sizes]\nbody = 20\n',
                    encoding="utf-8")
    design = mod.load_design(str(toml))
    assert design.color_band == "FF0000"
    assert design.size_body == 20
    assert design.unknown_keys == []
    xml = tmp_path / "doc.xml"
    xml.write_text(DOCBOOK, encoding="utf-8")
    out = tmp_path / "designed.pptx"
    mod.compile_adoc(str(xml), str(out), design=design)
    prs = Presentation(str(out))
    content = list(prs.slides)[1]
    band = content.shapes[0]
    assert str(band.fill.fore_color.rgb) == "FF0000"


def test_measurer_fullwidth_rules():
    m = mod.Measurer(mod.Design())
    # monospace: full-width glyphs measure 1em, latin 0.6em
    assert m.char_width_pt("あ", "Consolas", 10.0, mono=True) == 10.0
    assert m.char_width_pt("a", "Consolas", 10.0, mono=True) == 6.0
