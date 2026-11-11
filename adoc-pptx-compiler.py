#!/usr/bin/env python3
"""adoc-pptx-compiler: compile AsciiDoc documents into editable .pptx decks.

Pipeline: AsciiDoc --(Asciidoctor / asciidoc-py)--> DocBook XML
          --> internal model --> slide mapping --> text-measurement layout
          --> python-pptx drawing (plain shapes only, self-contained output).

The output deck does not depend on masters/layouts/themes for its content:
every text is a real text run inside a plain shape, so the file stays
directly editable in PowerPoint, LibreOffice Impress and Google Slides.
"""
import argparse
import json
import re as _re
import os
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

__version__ = "0.9.2.0"

_WS = _re.compile(r"\s+")

# --------------------------------------------------------------------------
# Base directory / path resolution
#
# Rule: wherever this tool needs "the folder the program lives in", that is
# the script folder for a .py run and the .exe folder for a frozen (e.g.
# PyInstaller) run.  app_base_dir() is the single place that knows this.
# --------------------------------------------------------------------------


def app_base_dir():
    """Folder that acts as this program's home base.

    Frozen (exe) builds: the folder containing the executable.
    Plain script runs:   the folder containing this .py file.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def resolve_input_path(path, extra_dirs=()):
    """Resolve an input file: absolute as-is, then CWD, then the document
    folder (extra_dirs), then app_base_dir()."""
    if os.path.isabs(path):
        return path
    candidates = [os.path.join(os.getcwd(), path)]
    candidates += [os.path.join(d, path) for d in extra_dirs]
    candidates.append(os.path.join(app_base_dir(), path))
    for cand in candidates:
        if os.path.isfile(cand):
            return cand
    return candidates[0]


def resolve_output_path(path):
    """Output files are relative to the current working directory."""
    if os.path.isabs(path):
        return path
    return os.path.join(os.getcwd(), path)


# --------------------------------------------------------------------------
# Design (built-in defaults, optionally overridden by a TOML file)
# --------------------------------------------------------------------------


@dataclass
class Design:
    slide_w: float = 13.333
    slide_h: float = 7.5
    margin_x: float = 0.7
    band_h: float = 1.0
    content_top: float = 1.45
    content_bottom: float = 6.9
    color_band: str = "1F3A5F"
    color_accent: str = "E8A33D"
    color_text: str = "222222"
    color_sub: str = "777777"
    color_light: str = "DCE6F2"
    color_code_bg: str = "F2F2F2"
    font_heading: str = "Meiryo"
    font_body: str = "Meiryo"
    font_mono: str = "Consolas"
    size_title_band: float = 24.0
    size_body: float = 18.0
    size_subheading: float = 20.0
    size_table: float = 14.0
    size_code: float = 12.5
    size_caption: float = 12.0
    size_footer: float = 10.5
    line_spacing: float = 1.3
    line_factor: float = 1.2  # font natural line-height allowance
    _cb_explicit: bool = False

    def set_orientation(self, portrait):
        """Force portrait (True) or landscape (False) by swapping the slide
        dimensions when needed, and re-derive the content area."""
        if portrait and self.slide_w > self.slide_h:
            self.slide_w, self.slide_h = self.slide_h, self.slide_w
        if not portrait and self.slide_h > self.slide_w:
            self.slide_w, self.slide_h = self.slide_h, self.slide_w
        if not self._cb_explicit:
            self.content_bottom = self.slide_h - 0.6

    @property
    def content_width(self):
        return self.slide_w - 2 * self.margin_x

    @property
    def content_height(self):
        return self.content_bottom - self.content_top


def load_design(path):
    """Load a design TOML file. Recognised keys:

    [colors] band, accent, text, sub, light, code_bg   (hex strings)
    [fonts]  heading, body, mono                      (font family names)
    [sizes]  title_band, body, subheading, table, code, caption, footer
    [page]   slide_w, slide_h, margin_x, band_h, content_top, content_bottom,
             line_spacing, line_factor                (inches / points)
    """
    design = Design()
    if not path:
        return design
    import tomllib

    keymap = {
        "colors": {"band": "color_band", "accent": "color_accent",
                   "text": "color_text", "sub": "color_sub",
                   "light": "color_light", "code_bg": "color_code_bg"},
        "fonts": {"heading": "font_heading", "body": "font_body",
                  "mono": "font_mono"},
        "sizes": {"title_band": "size_title_band", "body": "size_body",
                  "subheading": "size_subheading", "table": "size_table",
                  "code": "size_code", "caption": "size_caption",
                  "footer": "size_footer"},
        "page": {"slide_w": "slide_w", "slide_h": "slide_h",
                 "margin_x": "margin_x", "band_h": "band_h",
                 "content_top": "content_top",
                 "content_bottom": "content_bottom",
                 "line_spacing": "line_spacing",
                 "line_factor": "line_factor"},
    }
    with open(resolve_input_path(path), "rb") as fh:
        data = tomllib.load(fh)
    unknown = []
    for section, keys in keymap.items():
        for key, value in data.get(section, {}).items():
            attr = keys.get(key)
            if attr is None:
                unknown.append(f"{section}.{key}")
            else:
                setattr(design, attr, value)
                if attr == "content_bottom":
                    design._cb_explicit = True
    for section in data:
        if section not in keymap:
            unknown.append(section)
    design.unknown_keys = unknown
    return design


# --------------------------------------------------------------------------
# Frontend: AsciiDoc -> DocBook XML
# --------------------------------------------------------------------------


class FrontendError(Exception):
    pass


def obtain_docbook(src, frontend="auto", keep_ir=None, workdir=None):
    """Return the path of a DocBook XML file for src (.adoc or .xml)."""
    if src.lower().endswith(".xml"):
        return src
    out_dir = workdir or tempfile.mkdtemp(prefix="adoc_pptx_")
    out = os.path.join(out_dir, "docbook.xml")
    errors = []

    def try_asciidoctor():
        exe = shutil.which("asciidoctor")
        if not exe:
            errors.append("asciidoctor: not found")
            return False
        cmd = [exe, "-b", "docbook5", "-o", out, src]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode != 0 or not os.path.isfile(out):
            errors.append("asciidoctor: " + (proc.stderr.strip()[:300] or "failed"))
            return False
        return True

    def try_asciidoc_py():
        try:
            import asciidoc  # noqa: F401
        except ImportError:
            errors.append("asciidoc-py: not installed (pip install asciidoc)")
            return False
        cmd = [sys.executable, "-m", "asciidoc", "-b", "docbook", "-o", out, src]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode != 0 or not os.path.isfile(out):
            errors.append("asciidoc-py: " + (proc.stderr.strip()[:300] or "failed"))
            return False
        return True

    order = {"asciidoctor": [try_asciidoctor],
             "asciidoc-py": [try_asciidoc_py],
             "auto": [try_asciidoctor, try_asciidoc_py]}[frontend]
    for fn in order:
        if fn():
            if keep_ir:
                shutil.copyfile(out, resolve_output_path(keep_ir))
            return out
    raise FrontendError(
        "Could not convert AsciiDoc to DocBook. " + "; ".join(errors) +
        " Install Asciidoctor (recommended, same toolchain as asciidoctor-pdf)"
        " or run: pip install asciidoc")


# --------------------------------------------------------------------------
# Internal model
# --------------------------------------------------------------------------


@dataclass
class Run:
    text: str
    bold: bool = False
    italic: bool = False
    mono: bool = False
    href: str = None
    note_index: int = None  # footnote marker


@dataclass
class Para:
    runs: list = field(default_factory=list)


@dataclass
class ListItem:
    runs: list = field(default_factory=list)
    children: list = field(default_factory=list)  # nested ListBlock


@dataclass
class ListBlock:
    ordered: bool = False
    items: list = field(default_factory=list)


@dataclass
class TableBlock:
    header: list = field(default_factory=list)  # list of rows; row = list of str
    rows: list = field(default_factory=list)
    weights: list = None


@dataclass
class ImageBlock:
    path: str = ""
    alt: str = ""
    caption: str = ""


@dataclass
class ListingBlock:
    text: str = ""
    language: str = ""


@dataclass
class NoteBlock:
    kind: str = "note"
    runs: list = field(default_factory=list)


@dataclass
class SubHeading:
    text: str = ""
    level: int = 2


@dataclass
class Section:
    title: str = ""
    level: int = 1
    blocks: list = field(default_factory=list)


@dataclass
class Document:
    kind: str = ""  # DocBook root element: article / book / ...
    title: str = ""
    meta: list = field(default_factory=list)  # author / date strings
    sections: list = field(default_factory=list)
    footnotes: list = field(default_factory=list)  # global, in encounter order
    unsupported: list = field(default_factory=list)


# --------------------------------------------------------------------------
# DocBook parsing (namespace-agnostic; accepts DocBook 4.5 and 5)
# --------------------------------------------------------------------------


def _ln(tag):
    return tag.rsplit("}", 1)[-1]


def _kids(elem, name=None):
    return [c for c in elem if name is None or _ln(c.tag) == name]


def _first(elem, name):
    for c in elem:
        if _ln(c.tag) == name:
            return c
    return None


def _text(elem):
    return "".join(elem.itertext()) if elem is not None else ""


class DocBookParser:
    def __init__(self):
        self.doc = Document()

    # -- inline ---------------------------------------------------------
    def inline_runs(self, elem):
        runs = []

        def add(text, **kw):
            text = _WS.sub(" ", text)
            if text:
                runs.append(Run(text=text, **kw))

        def walk(node, bold=False, italic=False, mono=False, href=None):
            add(node.text or "", bold=bold, italic=italic, mono=mono, href=href)
            for child in node:
                name = _ln(child.tag)
                if name == "emphasis":
                    if child.get("role") == "strong":
                        walk(child, bold=True, italic=italic, mono=mono, href=href)
                    else:
                        walk(child, bold=bold, italic=True, mono=mono, href=href)
                elif name in ("literal", "code", "computeroutput", "filename"):
                    walk(child, bold=bold, italic=italic, mono=True, href=href)
                elif name == "ulink":
                    walk(child, bold=bold, italic=italic, mono=mono,
                         href=child.get("url"))
                elif name == "link":
                    url = None
                    for key, val in child.attrib.items():
                        if key.endswith("href"):
                            url = val
                    walk(child, bold=bold, italic=italic, mono=mono,
                         href=url or href)
                elif name == "footnote":
                    note_text = _text(child).strip()
                    self.doc.footnotes.append(note_text)
                    idx = len(self.doc.footnotes)
                    add(f"[{idx}]", bold=bold, italic=italic, mono=mono, href=href)
                    runs[-1].note_index = idx
                elif name in ("subscript", "superscript", "phrase"):
                    walk(child, bold=bold, italic=italic, mono=mono, href=href)
                elif name == "inlinemediaobject":
                    # Inline images degrade to their alt text in this version.
                    self.doc.unsupported.append(
                        "inline:inlinemediaobject(alt text only)")
                    alt = ""
                    for node in child.iter():
                        if _ln(node.tag) == "phrase":
                            alt = _text(node).strip()
                            break
                    add(alt, bold=bold, italic=True, mono=mono, href=href)
                    add(child.tail or "", bold=bold, italic=italic, mono=mono,
                        href=href)
                    continue
                else:
                    self.doc.unsupported.append("inline:" + name)
                    walk(child, bold=bold, italic=italic, mono=mono, href=href)
                add(child.tail or "", bold=bold, italic=italic, mono=mono, href=href)

        walk(elem)
        return runs

    @staticmethod
    def _trim_runs(runs):
        while runs and not runs[0].text.strip():
            runs.pop(0)
        while runs and not runs[-1].text.strip():
            runs.pop()
        if runs:
            runs[0].text = runs[0].text.lstrip()
            runs[-1].text = runs[-1].text.rstrip()
        return runs

    # -- blocks ---------------------------------------------------------
    def parse_list(self, elem):
        block = ListBlock(ordered=(_ln(elem.tag) == "orderedlist"))
        for item in _kids(elem, "listitem"):
            li = ListItem()
            for child in item:
                name = _ln(child.tag)
                if name in ("simpara", "para"):
                    li.runs.extend(self.inline_runs(child))
                    li.runs = self._trim_runs(li.runs)
                elif name in ("itemizedlist", "orderedlist"):
                    li.children.append(self.parse_list(child))
                elif name == "programlisting":
                    li.children.append(ListingBlock(text=_text(child),
                                                    language=child.get("language", "")))
                else:
                    self.doc.unsupported.append("listitem:" + name)
            block.items.append(li)
        return block

    def parse_table(self, elem):
        table = TableBlock()
        tgroup = _first(elem, "tgroup")
        if tgroup is None:
            tgroup = elem
        weights = []
        for cs in _kids(tgroup, "colspec"):
            raw = (cs.get("colwidth") or "1*").rstrip("*")
            try:
                weights.append(float(raw))
            except ValueError:
                weights.append(1.0)
        if weights:
            table.weights = weights
        for part_name, dest in (("thead", table.header), ("tbody", table.rows)):
            part = _first(tgroup, part_name)
            if part is None:
                continue
            for row in _kids(part, "row"):
                cells = []
                for entry in _kids(row, "entry"):
                    paras = [self.inline_runs(p) for p in entry
                             if _ln(p.tag) in ("simpara", "para")]
                    if paras:
                        text = "".join(r.text for runs in paras for r in runs)
                    else:
                        text = _text(entry).strip()
                    cells.append(text)
                dest.append(cells)
        if not table.header and table.rows:
            pass  # headerless tables stay headerless
        return table

    def parse_blocks(self, section_elem, level):
        section = Section(title=_text(_first(section_elem, "title")).strip(),
                          level=level)
        for child in section_elem:
            name = _ln(child.tag)
            if name == "title":
                continue
            if name in ("section", "chapter"):
                sub = self.parse_blocks(child, level + 1)
                section.blocks.append(SubHeading(text=sub.title, level=level + 1))
                section.blocks.extend(sub.blocks)
            elif name in ("simpara", "para"):
                section.blocks.append(
                    Para(runs=self._trim_runs(self.inline_runs(child))))
            elif name in ("itemizedlist", "orderedlist"):
                section.blocks.append(self.parse_list(child))
            elif name in ("table", "informaltable"):
                section.blocks.append(self.parse_table(child))
            elif name in ("mediaobject", "figure", "informalfigure"):
                target = child
                if name == "figure":
                    cap = _text(_first(child, "title")).strip()
                    target = _first(child, "mediaobject")
                    if target is None:
                        target = child
                else:
                    cap = ""
                img = None
                for node in target.iter():
                    if _ln(node.tag) == "imagedata":
                        img = node
                        break
                alt = ""
                for node in target.iter():
                    if _ln(node.tag) == "phrase":
                        alt = _text(node).strip()
                        break
                if img is not None:
                    section.blocks.append(ImageBlock(
                        path=img.get("fileref", ""), alt=alt, caption=cap))
                else:
                    self.doc.unsupported.append(name + ":no-image")
            elif name in ("programlisting", "screen"):
                section.blocks.append(ListingBlock(
                    text=_text(child), language=child.get("language", "")))
            elif name == "literallayout":
                section.blocks.append(ListingBlock(text=_text(child)))
            elif name == "variablelist":
                block = ListBlock(ordered=False)
                for entry in _kids(child, "varlistentry"):
                    term = _text(_first(entry, "term")).strip()
                    item = ListItem()
                    if term:
                        item.runs.append(Run(text=term + ": ", bold=True))
                    for li in _kids(entry, "listitem"):
                        for p in li:
                            if _ln(p.tag) in ("simpara", "para"):
                                item.runs.extend(self.inline_runs(p))
                    item.runs = self._trim_runs(item.runs)
                    block.items.append(item)
                section.blocks.append(block)
            elif name in ("note", "tip", "warning", "caution", "important"):
                runs = []
                for p in child:
                    if _ln(p.tag) in ("simpara", "para"):
                        runs.extend(self.inline_runs(p))
                section.blocks.append(NoteBlock(kind=name, runs=runs))
            elif name == "blockquote":
                runs = []
                for p in child:
                    if _ln(p.tag) in ("simpara", "para"):
                        runs.extend(self.inline_runs(p))
                if runs:
                    section.blocks.append(Para(runs=runs))
            elif name in ("bridgehead",):
                section.blocks.append(SubHeading(
                    text=_text(child).strip(), level=level + 1))
            else:
                self.doc.unsupported.append("block:" + name)
        return section

    def parse(self, xml_path):
        tree = ET.parse(xml_path)
        root = tree.getroot()
        self.doc.kind = _ln(root.tag)
        info = _first(root, "articleinfo")
        if info is None:
            info = _first(root, "info")
        if info is None:
            info = _first(root, "bookinfo")
        title_elem = _first(root, "title")
        if title_elem is not None:
            self.doc.title = _text(title_elem).strip()
        if info is not None:
            if not self.doc.title:
                self.doc.title = _text(_first(info, "title")).strip()
            author = _first(info, "author")
            if author is not None:
                name_parts = []
                for part in ("firstname", "surname", "othername"):
                    value = _text(_first(author, part)).strip()
                    if value:
                        name_parts.append(value)
                if name_parts:
                    self.doc.meta.append(" ".join(name_parts))
                else:
                    text = _text(author).strip()
                    if text:
                        self.doc.meta.append(text)
            date = _text(_first(info, "date")).strip()
            if date:
                self.doc.meta.append(date)
        front = Section(title=self.doc.title, level=1)
        has_front = False
        for child in root:
            name = _ln(child.tag)
            if name in ("section", "chapter", "preface"):
                self.doc.sections.append(self.parse_blocks(child, 1))
            elif name in ("articleinfo", "info", "bookinfo", "title", "subtitle"):
                continue
            else:
                # front matter blocks before the first section
                wrapper = ET.Element("section")
                wrapper.append(child)
                parsed = self.parse_blocks(wrapper, 1)
                front.blocks.extend(parsed.blocks)
                has_front = True
        if has_front and front.blocks:
            self.doc.sections.insert(0, front)
        if not self.doc.sections:
            self.doc.sections.append(Section(title=self.doc.title, level=1))
        return self.doc


def parse_docbook(xml_path):
    return DocBookParser().parse(xml_path)


# --------------------------------------------------------------------------
# Text measurement and wrapping
# --------------------------------------------------------------------------

_WIDE = ("W", "F")
_NO_BREAK_BEFORE = "、。，．・：；！？）」』】〉》ー…‥"


class Measurer:
    """Measures text with real font files when available (via fc-match +
    Pillow), and falls back to per-character width heuristics otherwise."""

    def __init__(self, design):
        self.design = design
        self._font_cache = {}
        self._file_cache = {}
        self.used_fallback = False

    def _font_file(self, family):
        if family in self._file_cache:
            return self._file_cache[family]
        path = None
        exe = shutil.which("fc-match")
        if exe:
            try:
                proc = subprocess.run([exe, "-f", "%{file}", family],
                                      capture_output=True, text=True, timeout=10)
                cand = proc.stdout.strip()
                if cand and os.path.isfile(cand):
                    path = cand
            except Exception:
                path = None
        self._file_cache[family] = path
        return path

    def _pil_font(self, family, size_pt, bold):
        if getattr(self, "_broken", False):
            return None
        key = (family, round(size_pt), bold)
        if key in self._font_cache:
            return self._font_cache[key]
        font = None
        path = self._font_file(family + (":bold" if bold else ""))
        if path:
            try:
                from PIL import ImageFont
                font = ImageFont.truetype(path, int(size_pt * 4))
            except Exception:
                font = None
        if font is None:
            self.used_fallback = True
        self._font_cache[key] = font
        return font

    def char_width_pt(self, ch, family, size_pt, bold=False, mono=False):
        wide = unicodedata.east_asian_width(ch) in _WIDE
        if mono:
            # Monospace latin is 0.6em; full-width glyphs stay 1em even in a
            # monospace stack (and render so in fallback fonts).
            return size_pt * (1.0 if wide else 0.6)
        font = self._pil_font(family, size_pt, bold)
        if font is not None:
            try:
                width = font.getlength(ch) / 4.0
                if wide and width < size_pt * 0.9:
                    # The matched font almost certainly lacks this glyph
                    # (a .notdef box measures ~0.6em). Full-width glyphs are
                    # 1em in every CJK-capable font, so use that instead of
                    # under-measuring and overlapping the rendered text.
                    self.used_fallback = True
                    return size_pt * 1.0
                return width
            except RuntimeError:
                # Pillow built without a working raqm layout for this font:
                # disable real-font measurement for the rest of the run.
                self._font_cache.clear()
                self._broken = True
                self.used_fallback = True
        if ch == " ":
            return size_pt * 0.3
        if unicodedata.east_asian_width(ch) in _WIDE:
            return size_pt * 1.0
        return size_pt * (0.55 if bold else 0.5)

    def wrap(self, runs, family, size_pt, width_in, mono=False):
        """Wrap styled runs into lines. Returns a list of lines; each line is
        a list of (text, style-dict) segments."""
        width_pt = width_in * 72.0
        # Flatten into styled characters.
        chars = []
        for run in runs:
            style = (run.bold, run.italic, run.mono, run.href, run.note_index)
            for ch in run.text:
                chars.append((ch, style))
        lines = []
        cur = []
        cur_w = 0.0
        last_space = -1  # index in cur where a break is allowed (at a space)
        i = 0
        while i < len(chars):
            ch, style = chars[i]
            bold = style[0]
            w = self.char_width_pt(ch, family, size_pt, bold=bold,
                                   mono=mono or style[2])
            if ch == "\n":
                lines.append(cur)
                cur, cur_w, last_space = [], 0.0, -1
                i += 1
                continue
            if cur_w + w > width_pt and cur:
                if ch in _NO_BREAK_BEFORE:
                    cur.append((ch, style))
                    cur_w += w
                    i += 1
                    continue
                if last_space > 0:
                    lines.append(cur[:last_space])
                    rest = cur[last_space + 1:]
                    cur = rest
                    cur_w = sum(self.char_width_pt(
                        c, family, size_pt, bold=s[0], mono=mono or s[2])
                        for c, s in cur)
                    last_space = -1
                else:
                    lines.append(cur)
                    cur, cur_w, last_space = [], 0.0, -1
                continue
            cur.append((ch, style))
            cur_w += w
            if ch == " ":
                last_space = len(cur) - 1
            i += 1
        if cur or not lines:
            lines.append(cur)
        # Merge characters back into segments.
        merged = []
        for line in lines:
            segs = []
            for ch, style in line:
                if segs and segs[-1][1] == style:
                    segs[-1] = (segs[-1][0] + ch, style)
                else:
                    segs.append((ch, style))
            merged.append(segs)
        return merged

    def line_height_in(self, size_pt):
        d = self.design
        return size_pt * d.line_spacing * d.line_factor / 72.0


# --------------------------------------------------------------------------
# Layout: blocks -> atoms -> pages
# --------------------------------------------------------------------------


@dataclass
class Atom:
    kind: str  # lines | subheading | table | image | listing | note | missing
    data: object = None
    height: float = 0.0
    notes: list = field(default_factory=list)  # footnote indices


@dataclass
class Page:
    title: str = ""
    cont: int = 0
    atoms: list = field(default_factory=list)
    footnotes: list = field(default_factory=list)


class LayoutEngine:
    def __init__(self, doc, design, measurer, src_dir):
        self.doc = doc
        self.design = design
        self.m = measurer
        self.src_dir = src_dir
        self.warnings = []
        self.toc_entries = []  # (level, title, content page index)

    # -- atomizers ------------------------------------------------------
    def _lines_atoms(self, runs, size_pt, width_in, prefix_runs_list=None,
                     kind="lines"):
        """Wrap runs and chunk into atoms that fit one page at most."""
        d = self.design
        family = d.font_body
        lines = self.m.wrap(runs, family, size_pt, width_in)
        line_h = self.m.line_height_in(size_pt)
        max_lines = max(1, int(d.content_height / line_h))
        notes = [r.note_index for r in runs if r.note_index]
        atoms = []
        for start in range(0, len(lines), max_lines):
            chunk = lines[start:start + max_lines]
            atoms.append(Atom(kind=kind, data={"lines": chunk, "size": size_pt,
                                              "width": width_in},
                              height=len(chunk) * line_h + 0.06, notes=notes))
        return atoms

    def _list_atoms(self, block, level=0, counter=None):
        d = self.design
        atoms = []
        if counter is None:
            counter = [0]
        for item in block.items:
            counter[0] += 1
            if block.ordered:
                prefix = f"{counter[0]}. "
            else:
                prefix = "• " if level == 0 else "– "
            indent = 0.32 * level
            width = d.content_width - indent - 0.3
            runs = [Run(text=prefix)] + item.runs
            sub = self._lines_atoms(runs, d.size_body, width)
            for atom in sub:
                atom.data["indent"] = indent + 0.3
                atoms.append(atom)
            for child in item.children:
                if isinstance(child, ListBlock):
                    atoms.extend(self._list_atoms(child, level + 1, [0]))
                elif isinstance(child, ListingBlock):
                    atoms.extend(self._block_atoms(child))
        return atoms

    def _table_atoms(self, block):
        d = self.design
        weights = block.weights or [1.0] * max(
            1, max((len(r) for r in block.header + block.rows), default=1))
        total = sum(weights)
        col_ws = [d.content_width * w / total for w in weights]
        line_h = self.m.line_height_in(d.size_table)

        def row_height(row):
            worst = 1
            for ci, cell in enumerate(row):
                if ci >= len(col_ws):
                    break
                lines = self.m.wrap([Run(text=cell)], d.font_body,
                                    d.size_table, col_ws[ci] - 0.14)
                worst = max(worst, len(lines))
            return worst * line_h + 0.1

        header_h = sum(row_height(r) for r in block.header)
        atoms = []
        cur_rows, cur_h = [], header_h
        for row in block.rows:
            rh = row_height(row)
            if cur_rows and cur_h + rh > d.content_height:
                atoms.append(Atom(kind="table",
                                  data={"header": block.header, "rows": cur_rows,
                                        "weights": weights},
                                  height=cur_h))
                cur_rows, cur_h = [], header_h
            cur_rows.append(row)
            cur_h += rh
        if cur_rows or not atoms:
            atoms.append(Atom(kind="table",
                              data={"header": block.header, "rows": cur_rows,
                                    "weights": weights},
                              height=cur_h))
        return atoms

    def _image_atom(self, block):
        d = self.design
        path = resolve_input_path(block.path, extra_dirs=(self.src_dir,))
        if not os.path.isfile(path):
            self.warnings.append(f"image not found: {block.path}")
            return [Atom(kind="missing",
                         data={"text": f"[image not found: {block.path}]"},
                         height=0.4)]
        try:
            from PIL import Image
            with Image.open(path) as im:
                px_w, px_h = im.size
        except Exception:
            self.warnings.append(f"image unreadable: {block.path}")
            return [Atom(kind="missing",
                         data={"text": f"[image unreadable: {block.path}]"},
                         height=0.4)]
        max_w, max_h = d.content_width, d.content_height * 0.62
        w_in = px_w / 96.0
        h_in = px_h / 96.0
        scale = min(max_w / w_in, max_h / h_in, 1.0) if w_in and h_in else 1.0
        w_in, h_in = w_in * scale, h_in * scale
        caption = block.caption or block.alt
        height = h_in + (0.32 if caption else 0.05)
        return [Atom(kind="image",
                     data={"path": path, "w": w_in, "h": h_in,
                           "caption": caption},
                     height=height)]

    def _listing_atoms(self, block):
        d = self.design
        raw_lines = block.text.strip("\n").split("\n") or [""]
        runs_lines = []
        for raw in raw_lines:
            # Wrap listings a little narrower than the box: monospace
            # fallback metrics differ between renderers, and an extra
            # wrapped line would spill out of the background box.
            wrapped = self.m.wrap([Run(text=raw)], d.font_mono, d.size_code,
                                  (d.content_width - 0.24) * 0.9, mono=True)
            runs_lines.extend(wrapped)
        line_h = self.m.line_height_in(d.size_code)
        max_lines = max(1, int((d.content_height - 0.2) / line_h))
        atoms = []
        for start in range(0, len(runs_lines), max_lines):
            chunk = runs_lines[start:start + max_lines]
            atoms.append(Atom(kind="listing",
                              data={"lines": chunk, "size": d.size_code},
                              height=len(chunk) * line_h + 0.2))
        return atoms

    def _note_atoms(self, block):
        d = self.design
        label = block.kind.upper() + ": "
        runs = [Run(text=label, bold=True)] + block.runs
        atoms = self._lines_atoms(runs, d.size_body - 1.5,
                                  d.content_width - 0.55, kind="note")
        for atom in atoms:
            atom.height += 0.14
        return atoms

    def _block_atoms(self, block):
        d = self.design
        if isinstance(block, Para):
            if not block.runs:
                return []
            return self._lines_atoms(block.runs, d.size_body, d.content_width)
        if isinstance(block, SubHeading):
            return [Atom(kind="subheading", data={"text": block.text},
                         height=0.55)]
        if isinstance(block, ListBlock):
            return self._list_atoms(block)
        if isinstance(block, TableBlock):
            return self._table_atoms(block)
        if isinstance(block, ImageBlock):
            return self._image_atom(block)
        if isinstance(block, ListingBlock):
            return self._listing_atoms(block)
        if isinstance(block, NoteBlock):
            return self._note_atoms(block)
        self.warnings.append(f"unhandled block: {type(block).__name__}")
        return []

    # -- pagination -----------------------------------------------------
    def paginate(self, overflow="continue"):
        d = self.design
        pages = []
        for section in self.doc.sections:
            atoms = []
            for block in section.blocks:
                atoms.extend(self._block_atoms(block))
            self.toc_entries.append((1, section.title, len(pages)))
            page = Page(title=section.title, cont=0)
            y = 0.0
            cont = 0
            for atom in atoms:
                notes_h = 0.0
                if atom.notes:
                    notes_h = 0.3 + 0.17 * len(set(atom.notes))
                if page.atoms and y + atom.height + notes_h > d.content_height:
                    if overflow == "error":
                        raise OverflowError(
                            f"section '{section.title}' exceeds one page")
                    pages.append(page)
                    cont += 1
                    page = Page(title=section.title, cont=cont)
                    y = 0.0
                if atom.height > d.content_height:
                    self.warnings.append(
                        f"oversize block on page '{section.title}': "
                        f"{atom.kind} ({atom.height:.2f}in)")
                page.atoms.append(atom)
                if atom.kind == "subheading":
                    self.toc_entries.append(
                        (2, atom.data["text"], len(pages)))
                y += atom.height + 0.08
                for idx in atom.notes:
                    text = self.doc.footnotes[idx - 1] \
                        if 0 < idx <= len(self.doc.footnotes) else ""
                    entry = f"[{idx}] {text}"
                    if entry not in page.footnotes:
                        page.footnotes.append(entry)
            pages.append(page)
        return pages


# --------------------------------------------------------------------------
# Rendering (python-pptx; plain shapes only)
# --------------------------------------------------------------------------

from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Emu, Inches, Pt  # noqa: E402


def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


def _set_run_font(run, design, size, bold=False, italic=False, mono=False,
                  color="color_text", href=None, family=None):
    font = run.font
    font.size = Pt(size)
    font.bold = bold
    font.italic = italic
    if family is None:
        family = design.font_mono if mono else design.font_body
    font.name = family
    font.color.rgb = _rgb(getattr(design, color))
    # East Asian font (python-pptx sets the latin typeface only)
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = rpr.makeelement(qn("a:ea"), {})
        rpr.append(ea)
    ea.set("typeface", family)
    if href:
        run.hyperlink.address = href


def _rect(slide, x, y, w, h, hexcolor):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(hexcolor)
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def _textbox(slide, x, y, w, h):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return box, tf


def _fill_lines(tf, lines, design, size, mono=False, default_color="color_text"):
    first = True
    for line in lines:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.line_spacing = design.line_spacing
        for text, (bold, italic, is_mono, href, _note) in line:
            run = para.add_run()
            run.text = text
            _set_run_font(run, design, size, bold=bold, italic=italic,
                          mono=mono or is_mono, color=default_color, href=href)
        if not line:
            run = para.add_run()
            run.text = ""
            _set_run_font(run, design, size, mono=mono, color=default_color)


class Renderer:
    def __init__(self, doc, design):
        self.doc = doc
        self.design = design
        self.prs = Presentation()
        self.prs.slide_width = Inches(design.slide_w)
        self.prs.slide_height = Inches(design.slide_h)
        self.blank = self.prs.slide_layouts[6]
        self.page_no = 0

    def _new_slide(self, bg=None):
        slide = self.prs.slides.add_slide(self.blank)
        self.page_no += 1
        if bg:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = _rgb(bg)
        return slide

    def _footer(self, slide):
        d = self.design
        fy = d.slide_h - 0.44
        _, tf = _textbox(slide, d.margin_x, fy,
                         d.slide_w - 2 * d.margin_x - 1.1, 0.3)
        run = tf.paragraphs[0].add_run()
        run.text = self.doc.title
        _set_run_font(run, d, d.size_footer, color="color_sub")
        _, tf = _textbox(slide, d.slide_w - d.margin_x - 1.0, fy, 1.0, 0.3)
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.RIGHT
        run = para.add_run()
        run.text = str(self.page_no)
        _set_run_font(run, d, d.size_footer, color="color_sub")

    def title_slide(self):
        d = self.design
        slide = self._new_slide(bg=d.color_band)
        title_y = d.slide_h * 0.29
        _rect(slide, d.margin_x, d.slide_h * 0.44, 2.2, 0.06, d.color_accent)
        _, tf = _textbox(slide, d.margin_x, title_y, d.content_width, 1.6)
        tf.word_wrap = True
        run = tf.paragraphs[0].add_run()
        run.text = self.doc.title
        _set_run_font(run, d, 40, bold=True, color="color_light",
                      family=d.font_heading)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        if self.doc.meta:
            _, tf = _textbox(slide, d.margin_x, d.slide_h * 0.485,
                             d.content_width, 0.8)
            run = tf.paragraphs[0].add_run()
            run.text = "　/　".join(self.doc.meta)
            _set_run_font(run, d, 16, color="color_light")

    def content_slide(self, page):
        d = self.design
        slide = self._new_slide()
        _rect(slide, 0, 0, d.slide_w, d.band_h, d.color_band)
        _rect(slide, 0, d.band_h + 0.1, d.slide_w, 0.045, d.color_accent)
        _, tf = _textbox(slide, d.margin_x, 0.18, d.content_width, 0.65)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        run = tf.paragraphs[0].add_run()
        title = page.title or self.doc.title
        if page.cont > 0:
            title += "（続き）"
        run.text = title
        _set_run_font(run, d, d.size_title_band, bold=True,
                      family=d.font_heading)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        y = d.content_top
        for atom in page.atoms:
            y = self._draw_atom(slide, atom, y) + 0.08
        if page.footnotes:
            note_h = 0.22 * len(page.footnotes) + 0.1
            _, tf = _textbox(slide, d.margin_x, d.content_bottom - note_h + 0.28,
                             d.content_width, note_h)
            for i, entry in enumerate(page.footnotes):
                para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                run = para.add_run()
                run.text = entry
                _set_run_font(run, d, 9.5, color="color_sub")
        self._footer(slide)

    def _draw_atom(self, slide, atom, y):
        d = self.design
        x = d.margin_x
        data = atom.data
        if atom.kind in ("lines", "note"):
            if atom.kind == "note":
                _rect(slide, x, y + 0.04, 0.07, atom.height - 0.12,
                      d.color_accent)
                x += 0.3
            indent = data.get("indent", 0.0)
            _, tf = _textbox(slide, x + indent, y,
                             data["width"] - indent + 0.3, atom.height)
            _fill_lines(tf, data["lines"], d, data["size"])
            return y + atom.height
        if atom.kind == "subheading":
            _, tf = _textbox(slide, x, y + 0.1, d.content_width, 0.45)
            run = tf.paragraphs[0].add_run()
            run.text = data["text"]
            _set_run_font(run, d, d.size_subheading, bold=True,
                          color="color_band", family=d.font_heading)
            return y + atom.height
        if atom.kind == "listing":
            _rect(slide, x, y, d.content_width, atom.height, d.color_code_bg)
            _, tf = _textbox(slide, x + 0.12, y + 0.08,
                             d.content_width - 0.24, atom.height - 0.16)
            _fill_lines(tf, data["lines"], d, data["size"], mono=True)
            return y + atom.height
        if atom.kind == "image":
            left = x + (d.content_width - data["w"]) / 2
            slide.shapes.add_picture(data["path"], Inches(left), Inches(y),
                                     Inches(data["w"]), Inches(data["h"]))
            bottom = y + data["h"]
            if data["caption"]:
                _, tf = _textbox(slide, x, bottom + 0.04, d.content_width, 0.3)
                para = tf.paragraphs[0]
                para.alignment = PP_ALIGN.CENTER
                run = para.add_run()
                run.text = data["caption"]
                _set_run_font(run, d, d.size_caption, color="color_sub")
                bottom += 0.32
            return bottom
        if atom.kind == "missing":
            _, tf = _textbox(slide, x, y + 0.06, d.content_width, 0.3)
            run = tf.paragraphs[0].add_run()
            run.text = data["text"]
            _set_run_font(run, d, d.size_body - 2, italic=True,
                          color="color_sub")
            return y + atom.height
        if atom.kind == "table":
            return self._draw_table(slide, atom, y)
        return y + atom.height

    def _draw_table(self, slide, atom, y):
        d = self.design
        data = atom.data
        header, rows = data["header"], data["rows"]
        weights = data["weights"] or [1.0] * max(
            1, max((len(r) for r in header + rows), default=1))
        ncols = max(len(weights),
                    max((len(r) for r in header + rows), default=1))
        while len(weights) < ncols:
            weights.append(1.0)
        total = sum(weights)
        nrows = len(header) + len(rows)
        if nrows == 0:
            return y
        frame = slide.shapes.add_table(nrows, ncols, Inches(d.margin_x),
                                       Inches(y), Inches(d.content_width),
                                       Inches(atom.height))
        table = frame.table
        for ci in range(ncols):
            table.columns[ci].width = Emu(
                int(Inches(d.content_width) * weights[ci] / total))
        all_rows = [(r, True) for r in header] + [(r, False) for r in rows]
        for ri, (row, is_header) in enumerate(all_rows):
            for ci in range(ncols):
                cell = table.cell(ri, ci)
                cell.margin_left = cell.margin_right = Inches(0.07)
                cell.margin_top = cell.margin_bottom = Inches(0.03)
                cell.text = ""
                tf = cell.text_frame
                tf.word_wrap = True
                run = tf.paragraphs[0].add_run()
                run.text = row[ci] if ci < len(row) else ""
                _set_run_font(run, d, d.size_table, bold=is_header,
                              color="color_light" if is_header else "color_text")
                if is_header:
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                fill = cell.fill
                fill.solid()
                fill.fore_color.rgb = _rgb(
                    d.color_band if is_header else
                    ("FFFFFF" if ri % 2 == 0 else "EFF3F9"))
        return y + atom.height

    def toc_slide(self, entries, toc_title):
        d = self.design
        slide = self._new_slide()
        _rect(slide, 0, 0, d.slide_w, d.band_h, d.color_band)
        _rect(slide, 0, d.band_h + 0.1, d.slide_w, 0.045, d.color_accent)
        _, tf = _textbox(slide, d.margin_x, 0.18, d.content_width, 0.65)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        run = tf.paragraphs[0].add_run()
        run.text = toc_title
        _set_run_font(run, d, d.size_title_band, bold=True,
                      family=d.font_heading)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        y = d.content_top
        for level, title, slide_no in entries:
            indent = 0.35 * (level - 1)
            size = 15.0 if level == 1 else 13.0
            _, tf = _textbox(slide, d.margin_x + indent, y,
                             d.content_width - indent - 0.85, 0.34)
            run = tf.paragraphs[0].add_run()
            run.text = title
            _set_run_font(run, d, size, bold=(level == 1))
            _, tf = _textbox(slide, d.slide_w - d.margin_x - 0.8, y, 0.8, 0.34)
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.RIGHT
            run = para.add_run()
            run.text = str(slide_no)
            _set_run_font(run, d, size, color="color_sub")
            y += 0.44
        self._footer(slide)

    def render(self, pages, out_path, toc=None):
        if self.doc.title:
            self.title_slide()
        if toc:
            for chunk in toc["chunks"]:
                self.toc_slide(chunk, toc["title"])
        for page in pages:
            self.content_slide(page)
        parent = os.path.dirname(os.path.abspath(out_path))
        os.makedirs(parent, exist_ok=True)
        self.prs.save(out_path)
        return self.page_no


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------


def scan_adoc_options(src_path):
    """Read document attributes that DocBook does not carry: :toc: and
    :toc-title: from an .adoc source. Returns {} for non-AsciiDoc input."""
    options = {"toc": False, "toc_title": "目次"}
    if not src_path.lower().endswith((".adoc", ".asciidoc", ".txt")):
        return options
    try:
        with open(src_path, encoding="utf-8", errors="replace") as fh:
            head = fh.read(20000)
    except OSError:
        return options
    if _re.search(r"^:toc:\s*$", head, _re.M):
        options["toc"] = True
    m = _re.search(r"^:toc-title:\s*(.+)$", head, _re.M)
    if m:
        options["toc_title"] = m.group(1).strip()
    return options


def compile_adoc(src, dst, design=None, frontend="auto", overflow="continue",
                 keep_ir=None, report_path=None, orientation="auto",
                 toc="auto"):
    src_abs = resolve_input_path(src)
    if not os.path.isfile(src_abs):
        raise FileNotFoundError(f"input not found: {src}")
    design = design or Design()
    options = scan_adoc_options(src_abs)
    xml_path = obtain_docbook(src_abs, frontend=frontend, keep_ir=keep_ir)
    doc = parse_docbook(xml_path)
    # Orientation: books are portrait, everything else landscape, unless
    # the caller or the design file says otherwise.
    portrait = orientation == "portrait" or \
        (orientation == "auto" and doc.kind == "book")
    design.set_orientation(portrait)
    src_dir = os.path.dirname(src_abs)
    measurer = Measurer(design)
    engine = LayoutEngine(doc, design, measurer, src_dir)
    pages = engine.paginate(overflow=overflow)
    # Table of contents
    want_toc = toc == "always" or \
        (toc == "auto" and (options["toc"] or doc.kind == "book"))
    toc_info = None
    toc_pages = 0
    if want_toc and engine.toc_entries:
        capacity = max(1, int(design.content_height / 0.44))
        toc_pages = -(-len(engine.toc_entries) // capacity)
        first_content_slide = (1 if doc.title else 0) + toc_pages + 1
        numbered = [(level, title, first_content_slide + page_index)
                    for level, title, page_index in engine.toc_entries]
        chunks = [numbered[i:i + capacity]
                  for i in range(0, len(numbered), capacity)]
        toc_info = {"title": options["toc_title"], "chunks": chunks}
    renderer = Renderer(doc, design)
    out_abs = resolve_output_path(dst)
    n_slides = renderer.render(pages, out_abs, toc=toc_info)
    report = {
        "tool": "adoc-pptx-compiler",
        "version": __version__,
        "input": src_abs,
        "output": out_abs,
        "doctitle": doc.title,
        "kind": doc.kind,
        "orientation": "portrait" if design.slide_h > design.slide_w
                       else "landscape",
        "toc": bool(toc_info),
        "slides": n_slides,
        "measurement_fallback": measurer.used_fallback,
        "warnings": engine.warnings,
        "unsupported": sorted(set(doc.unsupported)),
        "design_unknown_keys": getattr(design, "unknown_keys", []),
    }
    if report_path:
        with open(resolve_output_path(report_path), "w",
                  encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="adoc-pptx-compiler",
        description="Compile an AsciiDoc document into an editable .pptx deck.")
    parser.add_argument("src", help="input .adoc (or DocBook .xml)")
    parser.add_argument("dst", help="output .pptx")
    parser.add_argument("--design", help="design TOML file")
    parser.add_argument("--frontend", choices=["auto", "asciidoctor", "asciidoc-py"],
                        default="auto", help="AsciiDoc frontend (default: auto)")
    parser.add_argument("--overflow", choices=["continue", "error"],
                        default="continue",
                        help="what to do when a section exceeds one page")
    parser.add_argument("--orientation",
                        choices=["auto", "landscape", "portrait"],
                        default="auto",
                        help="slide orientation (default: auto = book: "
                             "portrait, article: landscape)")
    parser.add_argument("--toc", choices=["auto", "always", "never"],
                        default="auto",
                        help="table of contents page (default: auto = "
                             "when :toc: is set or the document is a book)")
    parser.add_argument("--keep-ir", help="also save the DocBook XML here")
    parser.add_argument("--report", help="write a JSON report here")
    parser.add_argument("--version", action="version",
                        version=f"adoc-pptx-compiler {__version__}")
    args = parser.parse_args(argv)
    try:
        design = load_design(args.design) if args.design else Design()
        report = compile_adoc(args.src, args.dst, design=design,
                              frontend=args.frontend, overflow=args.overflow,
                              keep_ir=args.keep_ir, report_path=args.report,
                              orientation=args.orientation, toc=args.toc)
    except FileNotFoundError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    except FrontendError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    except OverflowError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    print(f"wrote {report['output']} ({report['slides']} slides)")
    for warning in report["warnings"]:
        print(f"warning: {warning}")
    for item in report["unsupported"]:
        print(f"unsupported: {item}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
