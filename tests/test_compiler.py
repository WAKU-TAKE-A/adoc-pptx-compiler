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
    assert mod.__version__ == "0.9.5.0"


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


def test_nested_list_box_width_matches_wrap(tmp_path):
    xml_text = """<?xml version="1.0" encoding="UTF-8"?>
<article><articleinfo><title>入れ子</title></articleinfo>
<section><title>節</title>
<itemizedlist>
  <listitem><simpara>親項目</simpara>
    <itemizedlist>
      <listitem><simpara>これは入れ子になった長い項目のテキストです。折り返しの幅が描画の箱幅と一致しなければなりません。</simpara></listitem>
    </itemizedlist>
  </listitem>
</itemizedlist></section></article>"""
    xml = tmp_path / "nested.xml"
    xml.write_text(xml_text, encoding="utf-8")
    out = tmp_path / "nested.pptx"
    mod.compile_adoc(str(xml), str(out))
    prs = Presentation(str(out))
    slide = list(prs.slides)[1]
    target = None
    for shape in slide.shapes:
        if shape.has_text_frame and "入れ子になった長い項目" in shape.text_frame.text:
            target = shape
    assert target is not None
    # wrapped at content_width - level indent (0.32) - prefix zone (0.3)
    expected = (13.333 - 1.4 - 0.32 - 0.3) * 914400
    assert abs(target.width - expected) < 2000


def test_line_chars_narrows_text_column(tmp_path):
    xml = tmp_path / "doc.xml"
    xml.write_text(DOCBOOK, encoding="utf-8")
    design = mod.Design()
    design.line_chars = 20  # 20 chars * 18pt = 5.0in column
    out = tmp_path / "narrow.pptx"
    mod.compile_adoc(str(xml), str(out), design=design)
    prs = Presentation(str(out))
    slide = list(prs.slides)[1]
    widths = [s.width / 914400 for s in slide.shapes
              if s.has_text_frame and "本文の段落" in s.text_frame.text]
    assert widths and abs(widths[0] - 5.0) < 0.05
    # and the default (line_chars=0) uses the full content width
    out2 = tmp_path / "wide.pptx"
    mod.compile_adoc(str(xml), str(out2))
    prs2 = Presentation(str(out2))
    slide2 = list(prs2.slides)[1]
    widths2 = [s.width / 914400 for s in slide2.shapes
               if s.has_text_frame and "本文の段落" in s.text_frame.text]
    assert widths2 and widths2[0] > 11.0


INLINE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<article><articleinfo><title>文中の画像</title></articleinfo>
<section><title>節</title>
<simpara>文の途中に <inlinemediaobject><imageobject>
<imagedata fileref="lena.jpg"/></imageobject><textobject>
<phrase>lena.jpg</phrase></textobject></inlinemediaobject> があります。</simpara>
</section></article>
"""


def test_inline_image_message_is_explained(tmp_path, capsys):
    xml = tmp_path / "inline.xml"
    xml.write_text(INLINE_XML, encoding="utf-8")
    out = tmp_path / "inline.pptx"
    assert mod.main([str(xml), str(out)]) == 0
    text = capsys.readouterr().out
    assert "inlinemediaobject" in text
    assert "代替テキスト" in text
    # the alt text itself is in the deck, not silently dropped
    prs = Presentation(str(out))
    joined = "\n".join(all_texts(prs))
    assert "lena.jpg" in joined


DB5_SUBTITLE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<book xmlns="http://docbook.org/ns/docbook" version="5.0">
<info><title>メインタイトル</title><subtitle>サブタイトルです</subtitle>
<author><personname>著者 太郎</personname></author></info>
<chapter><title>第一章</title><simpara>本文です。</simpara></chapter>
</book>
"""

COMBINED_TITLE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<article>
  <articleinfo>
    <title>メインタイトル: サブタイトルです</title>
  </articleinfo>
  <section><title>節</title><simpara>本文です。</simpara></section>
</article>
"""


def test_subtitle_from_docbook5(tmp_path):
    xml = tmp_path / "sub.xml"
    xml.write_text(DB5_SUBTITLE_XML, encoding="utf-8")
    doc = mod.parse_docbook(str(xml))
    assert doc.title == "メインタイトル"
    assert doc.subtitle == "サブタイトルです"


def test_subtitle_partitioned_from_combined_title(tmp_path):
    xml = tmp_path / "combined.xml"
    xml.write_text(COMBINED_TITLE_XML, encoding="utf-8")
    doc = mod.parse_docbook(str(xml))
    assert doc.title == "メインタイトル"
    assert doc.subtitle == "サブタイトルです"


def test_subtitle_only_partitioned_once_on_last_colon(tmp_path):
    xml = tmp_path / "multi.xml"
    xml.write_text(COMBINED_TITLE_XML.replace(
        "メインタイトル: サブタイトルです", "A: B: C"), encoding="utf-8")
    doc = mod.parse_docbook(str(xml))
    assert doc.title == "A: B"
    assert doc.subtitle == "C"


def test_subtitle_on_title_slide_only(tmp_path):
    xml = tmp_path / "sub.xml"
    xml.write_text(DB5_SUBTITLE_XML, encoding="utf-8")
    out = tmp_path / "sub.pptx"
    assert mod.main([str(xml), str(out)]) == 0
    prs = Presentation(str(out))
    slides_text = ["\n".join(
        s.text_frame.text for s in slide.shapes if s.has_text_frame)
        for slide in prs.slides]
    assert "サブタイトルです" in slides_text[0]
    # subtitle appears on the title slide only; footers carry the main title
    assert all("サブタイトルです" not in t for t in slides_text[1:])
    assert any("メインタイトル" in t for t in slides_text[1:])


def test_subtitle_size_in_design_toml(tmp_path):
    toml = tmp_path / "d.toml"
    toml.write_text('[sizes]\ntitle = 44\nsubtitle = 30\n', encoding="utf-8")
    design = mod.load_design(str(toml))
    assert design.size_title == 44
    assert design.size_subtitle == 30
    assert design.unknown_keys == []


REV_ADOC = """= テスト文書
:author:
:revdate: 2024年06月15日
:revnumber: 1.0版
:version-label:

== 節

本文です。
"""


def test_scan_revision_attributes(tmp_path):
    src = tmp_path / "rev.adoc"
    src.write_text(REV_ADOC, encoding="utf-8")
    options = mod.scan_adoc_options(str(src))
    assert options["author"] == ""
    assert options["revnumber"] == "1.0版"
    assert options["revdate"] == "2024年06月15日"
    assert options["version_label"] == ""


def test_revision_line_suppresses_label(tmp_path):
    src = tmp_path / "rev.adoc"
    src.write_text(REV_ADOC, encoding="utf-8")
    out = tmp_path / "rev.pptx"
    assert mod.main([str(src), str(out)]) == 0
    prs = Presentation(str(out))
    title_texts = [s.text_frame.text for s in prs.slides[0].shapes
                   if s.has_text_frame]
    joined = "\n".join(title_texts)
    assert "1.0版, 2024年06月15日" in joined
    assert "Version" not in joined


def test_revision_line_default_label(tmp_path):
    src = tmp_path / "rev2.adoc"
    src.write_text(REV_ADOC.replace(":version-label:\n", "")
                   .replace("1.0版", "2.0"), encoding="utf-8")
    out = tmp_path / "rev2.pptx"
    assert mod.main([str(src), str(out)]) == 0
    prs = Presentation(str(out))
    joined = "\n".join(all_texts(prs))
    assert "Version 2.0" in joined


def test_title_background_image(tmp_path):
    from PIL import Image
    Image.new("RGB", (800, 600), (30, 80, 160)).save(
        str(tmp_path / "bg.png"))
    src = tmp_path / "bg.adoc"
    src.write_text("= 背景つき文書\n"
                   ":title-page-background-image: image:bg.png[]\n\n"
                   "== 節\n\n本文です。\n", encoding="utf-8")
    out = tmp_path / "bg.pptx"
    assert mod.main([str(src), str(out)]) == 0
    prs = Presentation(str(out))
    pics = [s for s in prs.slides[0].shapes
            if s.shape_type == 13]  # PICTURE
    assert len(pics) == 1
    assert abs(pics[0].width - prs.slide_width) < 1000
    assert abs(pics[0].height - prs.slide_height) < 1000


def test_title_background_missing_warns(tmp_path, capsys):
    src = tmp_path / "bg2.adoc"
    src.write_text("= 背景なし文書\n"
                   ":title-page-background-image: image:nope.jpg[]\n\n"
                   "== 節\n\n本文です。\n", encoding="utf-8")
    out = tmp_path / "bg2.pptx"
    assert mod.main([str(src), str(out)]) == 0
    assert "title-page-background-image not found" in \
        capsys.readouterr().out


def test_toc_title_attribute(tmp_path):
    src = tmp_path / "toc.adoc"
    src.write_text("= 目次テスト\n:toc:\n:toc-title: もくじ\n\n"
                   "== 節\n\n本文です。\n", encoding="utf-8")
    out = tmp_path / "toc.pptx"
    assert mod.main([str(src), str(out)]) == 0
    prs = Presentation(str(out))
    assert "もくじ" in "\n".join(all_texts(prs))
