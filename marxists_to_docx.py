"""
Marxists.org -> .docx Konverter
================================
Lädt eine oder mehrere Marxists.org-Seiten, erkennt die Fußnoten-Verlinkung
(<a href="#noteN">) und speichert Text + echte Word-Fußnoten in einer .docx-Datei.

Benutzung in VS Code:
    1. Abhängigkeiten installieren:
         pip install requests beautifulsoup4 python-docx
    2. Datei mit F5 / "Run Python File" starten.
    3. URLs nacheinander eingeben, leere Zeile beendet die Eingabe.
"""

from __future__ import annotations

import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.opc.constants import CONTENT_TYPE, RELATIONSHIP_TYPE
from docx.opc.part import Part
from docx.opc.packuri import PackURI

# ----------------------------- Options -----------------------------

DEFAULT_FONT = "Times New Roman"


@dataclass
class Options:
    keep_colors: bool = False
    italic_quotes: bool = False
    body_font_size: float = 11.0
    line_spacing: float = 1.15
    justify: bool = True
    footnote_font_size: float = 9.0
    footnote_italic: bool = False


COLOR_STYLE_RE = re.compile(r"color\s*:\s*([^;]+)", re.IGNORECASE)
HEX_COLOR_RE = re.compile(r"^#?([0-9a-f]{6})$", re.IGNORECASE)
NAMED_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
    "blue": "0000FF", "yellow": "FFFF00", "gray": "808080", "grey": "808080",
    "navy": "000080", "purple": "800080", "maroon": "800000",
    "darkred": "8B0000", "darkblue": "00008B", "darkgreen": "006400",
}


def parse_color(value: str) -> str | None:
    """Wandelt einen CSS-/HTML-Farbwert in einen 6-stelligen Hex-Wert um."""
    if not value:
        return None
    v = value.strip().lower()
    if v in NAMED_COLORS:
        return NAMED_COLORS[v]
    m = HEX_COLOR_RE.match(v)
    if m:
        return m.group(1).upper()
    return None


def color_from_tag(tag: Tag) -> str | None:
    """Liest eine Farbe aus style="color:..." oder <font color="...">."""
    style = tag.get("style", "")
    if style:
        m = COLOR_STYLE_RE.search(style)
        if m:
            c = parse_color(m.group(1))
            if c:
                return c
    if tag.name == "font":
        c = parse_color(tag.get("color", ""))
        if c:
            return c
    return None


def ask(prompt: str, default: str) -> str:
    raw = input(f"{prompt} [{default}]: ").strip()
    return raw or default


def ask_bool(prompt: str, default: bool) -> bool:
    dflt = "j" if default else "n"
    raw = input(f"{prompt} (j/n) [{dflt}]: ").strip().lower()
    if not raw:
        return default
    return raw.startswith(("j", "y", "1", "t"))


def ask_float(prompt: str, default: float) -> float:
    raw = input(f"{prompt} [{default}]: ").strip().replace(",", ".")
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        print(f"  -> Ungültig, nehme {default}.")
        return default


def ask_options() -> Options:
    print("\nFormatierung (Enter = Default):")
    opts = Options()
    opts.keep_colors = ask_bool("  Textfarben aus dem Original übernehmen?", False)
    opts.italic_quotes = ask_bool('  Zitate in „..." kursiv setzen?', False)
    opts.body_font_size = ask_float("  Fließtext-Schriftgröße (pt)", 11.0)
    opts.line_spacing = ask_float("  Zeilenabstand (1.0 / 1.15 / 1.5 / 2.0)", 1.15)
    opts.justify = ask_bool("  Blocksatz?", True)
    opts.footnote_font_size = ask_float("  Fußnoten-Schriftgröße (pt)", 9.0)
    opts.footnote_italic = ask_bool("  Fußnoten kursiv?", False)
    return opts


# ----------------------------- Footnote support -----------------------------

FOOTNOTES_XML_TEMPLATE = (
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:footnote w:type="separator" w:id="-1">'
    '<w:p><w:r><w:separator/></w:r></w:p>'
    '</w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:r><w:continuationSeparator/></w:r></w:p>'
    '</w:footnote>'
    '</w:footnotes>'
)


class FootnotesPart(Part):
    """Eigene Part-Klasse, die ihr lxml-Element live serialisiert."""

    def __init__(self, partname, content_type, element, package):
        super().__init__(partname, content_type, b"", package)
        self._element = element

    @property
    def element(self):
        return self._element

    @property
    def blob(self):
        from lxml import etree
        return etree.tostring(
            self._element, xml_declaration=True, encoding="UTF-8", standalone=True
        )


class FootnoteManager:
    """Hängt eine footnotes.xml an das Dokument und liefert IDs für neue Fußnoten."""

    def __init__(self, document: Document):
        self.document = document
        self._next_id = 1
        self._footnotes_part = self._ensure_footnotes_part()
        self._footnotes_root = self._footnotes_part.element

    def _ensure_footnotes_part(self) -> Part:
        package = self.document.part.package
        partname = PackURI("/word/footnotes.xml")
        try:
            existing = self.document.part.part_related_by(RELATIONSHIP_TYPE.FOOTNOTES)
            return existing
        except KeyError:
            pass

        from docx.oxml import parse_xml

        element = parse_xml(FOOTNOTES_XML_TEMPLATE)
        part = FootnotesPart(
            partname, CONTENT_TYPE.WML_FOOTNOTES, element, package
        )
        self.document.part.relate_to(part, RELATIONSHIP_TYPE.FOOTNOTES)
        return part

    def add_footnote(self, text: str, size_pt: float | None = None,
                     italic: bool = False) -> int:
        """Erstellt eine neue Fußnote mit dem gegebenen Text, gibt deren ID zurück."""
        fid = self._next_id
        self._next_id += 1

        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xml_ns = "http://www.w3.org/XML/1998/namespace"
        from lxml import etree

        # Word benutzt halbe Punkte für w:sz
        half_pt = str(int(round(size_pt * 2))) if size_pt else None

        def _rpr(parent):
            rPr = etree.SubElement(parent, f"{{{w}}}rPr")
            if italic:
                etree.SubElement(rPr, f"{{{w}}}i")
            if half_pt:
                sz = etree.SubElement(rPr, f"{{{w}}}sz")
                sz.set(f"{{{w}}}val", half_pt)
                szCs = etree.SubElement(rPr, f"{{{w}}}szCs")
                szCs.set(f"{{{w}}}val", half_pt)
            return rPr

        fn = etree.SubElement(self._footnotes_root, f"{{{w}}}footnote")
        fn.set(f"{{{w}}}id", str(fid))

        p = etree.SubElement(fn, f"{{{w}}}p")
        pPr = etree.SubElement(p, f"{{{w}}}pPr")
        pStyle = etree.SubElement(pPr, f"{{{w}}}pStyle")
        pStyle.set(f"{{{w}}}val", "FootnoteText")

        # Reference mark
        ref_run = etree.SubElement(p, f"{{{w}}}r")
        ref_rPr = etree.SubElement(ref_run, f"{{{w}}}rPr")
        ref_rStyle = etree.SubElement(ref_rPr, f"{{{w}}}rStyle")
        ref_rStyle.set(f"{{{w}}}val", "FootnoteReference")
        if half_pt:
            sz = etree.SubElement(ref_rPr, f"{{{w}}}sz")
            sz.set(f"{{{w}}}val", half_pt)
        etree.SubElement(ref_run, f"{{{w}}}footnoteRef")

        # Space
        space_run = etree.SubElement(p, f"{{{w}}}r")
        _rpr(space_run)
        space_t = etree.SubElement(space_run, f"{{{w}}}t")
        space_t.text = " "
        space_t.set(f"{{{xml_ns}}}space", "preserve")

        # Footnote text
        text_run = etree.SubElement(p, f"{{{w}}}r")
        _rpr(text_run)
        text_t = etree.SubElement(text_run, f"{{{w}}}t")
        text_t.text = text
        text_t.set(f"{{{xml_ns}}}space", "preserve")

        return fid


def add_footnote_reference(paragraph, footnote_id: int) -> None:
    """Fügt am Ende des Paragraphen eine Fußnoten-Referenz (hochgestellte Zahl) ein."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    from lxml import etree

    r = etree.SubElement(paragraph._p, f"{{{w}}}r")
    rPr = etree.SubElement(r, f"{{{w}}}rPr")
    rStyle = etree.SubElement(rPr, f"{{{w}}}rStyle")
    rStyle.set(f"{{{w}}}val", "FootnoteReference")
    ref = etree.SubElement(r, f"{{{w}}}footnoteReference")
    ref.set(f"{{{w}}}id", str(footnote_id))


# ----------------------------- HTML parsing -----------------------------

FN_HREF_RE = re.compile(r"^#(.+)$")
URL_ONLY_RE = re.compile(r"^\s*(https?://|www\.)\S+\s*$", re.IGNORECASE)
PDF_MARKER_RE = re.compile(r"^\s*\[?\s*(pdf|download)\s*\]?\s*$", re.IGNORECASE)


def fetch(url: str) -> BeautifulSoup:
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    # Marxists.org liefert oft iso-8859-1; requests rät meist richtig, aber sicherstellen:
    if resp.encoding is None or resp.encoding.lower() == "iso-8859-1":
        resp.encoding = resp.apparent_encoding or "iso-8859-1"
    return BeautifulSoup(resp.text, "html.parser")


def extract_footnotes(soup: BeautifulSoup) -> dict[str, str]:
    """Sammelt aus den <p class="note"> die Fußnoten-Texte.
    Schlüssel = Anker-ID (z.B. 'note12', 'na' für '[1*]'-Übersetzerfußnoten)."""
    notes: dict[str, list[str]] = {}
    current_id: str | None = None
    for p in soup.find_all("p", class_="note"):
        # Erster <a> mit id und href (zurück in den Text) markiert eine neue Fußnote
        anchor = p.find("a", id=True, href=re.compile(r"^#"))
        if anchor:
            current_id = anchor["id"]
            anchor_text = anchor.get_text()
            full = p.get_text()
            text = full[len(anchor_text):].lstrip(" \xa0\t")
            notes.setdefault(current_id, []).append(text.strip())
        else:
            if current_id is not None:
                notes[current_id].append(p.get_text().strip())
    return {k: "\n".join(v).strip() for k, v in notes.items()}


SKIP_CLASSES = {"toplink", "link", "note"}
SKIP_TEXT_PATTERNS = (
    "anfang der seite",
    "zum anfang",
    "nach oben",
)


def find_content_paragraphs(soup: BeautifulSoup) -> list[Tag]:
    """
    Gibt die inhaltlichen Block-Elemente (h1–h4, p) vor dem 'Anmerkungen'-Bereich zurück.
    Lässt Navigation (p.toplink, p.link mit 'Anfang der Seite' etc.) und die
    Fußnoten-Liste (p.note) aus.
    """
    body = soup.body or soup
    result: list[Tag] = []
    for el in body.find_all(["h1", "h2", "h3", "h4", "p"]):
        classes = el.get("class") or []
        if any(c in SKIP_CLASSES for c in classes):
            continue
        text_stripped = el.get_text(strip=True)
        text_lower = text_stripped.lower()
        # Sobald die "Anmerkungen"-Überschrift kommt, abbrechen
        if el.name in ("h2", "h3", "h4") and text_lower.startswith("anmerkung"):
            break
        if any(p in text_lower for p in SKIP_TEXT_PATTERNS):
            continue
        # Leere Absätze (oft nur &#160;) überspringen
        if not text_stripped:
            continue
        result.append(el)
    return result


# ----------------------------- DOCX rendering -----------------------------

def render_element(
    document: Document,
    element: Tag,
    fn_mgr: FootnoteManager,
    footnotes: dict[str, str],
    used_ids: dict[str, int],
    opts: Options,
    quote_state: list[bool],
):
    """
    Rendert ein Block-Element (h1/h2/h3/h4/p) als Paragraph im docx.
    Erkennt Fußnoten-Anker und ersetzt sie durch echte Word-Fußnoten.
    """
    name = element.name
    is_heading = name in ("h1", "h2", "h3", "h4")
    if is_heading:
        level = {"h1": 0, "h2": 1, "h3": 2, "h4": 3}[name]
        paragraph = document.add_heading(level=level)
    else:
        paragraph = document.add_paragraph()
        if opts.justify:
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.line_spacing = opts.line_spacing

    _emit_inline(paragraph, element, fn_mgr, footnotes, used_ids,
                 opts, quote_state, current_color=None)


INLINE_BOLD = {"b", "strong"}
INLINE_ITALIC = {"i", "em"}
INLINE_UNDERLINE = {"u"}


QUOTE_OPEN = "„"   # „
QUOTE_CLOSE_DE = "“"  # "
QUOTE_CLOSE_ALT = "”"  # "
QUOTE_CHARS = (QUOTE_OPEN, QUOTE_CLOSE_DE, QUOTE_CLOSE_ALT)


def _emit_text(paragraph, text: str, bold: bool, italic: bool,
               underline: bool, color_hex: str | None,
               opts: Options, quote_state: list[bool]):
    """Schreibt einen Textabschnitt als Run(s). Splittet bei Anführungszeichen,
    wenn opts.italic_quotes aktiv ist, und toggelt den kursiv-Zustand."""
    def emit(segment: str, force_italic: bool):
        if not segment:
            return
        run = paragraph.add_run(segment)
        run.bold = bold or None
        run.italic = (italic or force_italic) or None
        run.underline = underline or None
        if color_hex:
            run.font.color.rgb = RGBColor.from_string(color_hex)

    if not opts.italic_quotes:
        emit(text, False)
        return

    buf = ""
    for ch in text:
        if ch == QUOTE_OPEN:
            emit(buf, quote_state[0]); buf = ""
            emit(ch, quote_state[0])
            quote_state[0] = True
        elif ch in (QUOTE_CLOSE_DE, QUOTE_CLOSE_ALT) and quote_state[0]:
            emit(buf, True); buf = ""
            emit(ch, True)
            quote_state[0] = False
        else:
            buf += ch
    emit(buf, quote_state[0])


def _emit_inline(
    paragraph,
    element: Tag,
    fn_mgr: FootnoteManager,
    footnotes: dict[str, str],
    used_ids: dict[str, int],
    opts: Options,
    quote_state: list[bool],
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    current_color: str | None = None,
):
    """Geht die Kinder rekursiv durch, schreibt Runs mit Formatierung
    und ersetzt Fußnoten-Anker durch echte Word-Fußnoten."""
    for child in element.children:
        if isinstance(child, NavigableString):
            text = str(child)
            if text:
                _emit_text(paragraph, text, bold, italic, underline,
                           current_color if opts.keep_colors else None,
                           opts, quote_state)
            continue
        if not isinstance(child, Tag):
            continue

        if child.name == "a":
            href = child.get("href", "")
            link_text = child.get_text()
            m = FN_HREF_RE.match(href)
            if m and m.group(1) in footnotes:
                note_key = m.group(1)
                fn_text = footnotes[note_key]
                fn_id = fn_mgr.add_footnote(
                    fn_text,
                    size_pt=opts.footnote_font_size,
                    italic=opts.footnote_italic,
                )
                used_ids[note_key] = fn_id
                add_footnote_reference(paragraph, fn_id)
                continue
            if href.startswith("#") and not link_text.strip():
                continue
            if href in ("#topp", "#top"):
                continue
            if URL_ONLY_RE.match(link_text) or PDF_MARKER_RE.match(link_text):
                continue
            _emit_inline(paragraph, child, fn_mgr, footnotes, used_ids,
                         opts, quote_state, bold, italic, underline,
                         current_color)
        elif child.name == "br":
            paragraph.add_run("\n")
        else:
            new_bold = bold or child.name in INLINE_BOLD
            new_italic = italic or child.name in INLINE_ITALIC
            new_underline = underline or child.name in INLINE_UNDERLINE
            new_color = color_from_tag(child) or current_color
            _emit_inline(paragraph, child, fn_mgr, footnotes, used_ids,
                         opts, quote_state, new_bold, new_italic, new_underline,
                         new_color)


# ----------------------------- Main -----------------------------

def read_urls() -> list[str]:
    print("Marxists.org -> DOCX")
    print("=" * 40)
    print("URLs nacheinander eingeben (eine pro Zeile).")
    print("Leere Zeile = fertig.\n")
    urls: list[str] = []
    while True:
        try:
            line = input(f"URL {len(urls)+1}: ").strip()
        except EOFError:
            break
        if not line:
            break
        if not line.startswith(("http://", "https://")):
            print("  -> Bitte vollständige URL angeben.")
            continue
        urls.append(line)
    return urls


def main():
    urls = read_urls()
    if not urls:
        print("Keine URLs angegeben, Abbruch.")
        sys.exit(0)

    output = input("\nAusgabe-Dateiname [marxists.docx]: ").strip() or "marxists.docx"
    output_path = Path(output).with_suffix(".docx")

    opts = ask_options()

    document = Document()
    # Globalen "Normal"-Stil setzen
    normal = document.styles["Normal"]
    normal.font.name = DEFAULT_FONT
    normal.font.size = Pt(opts.body_font_size)
    # Auch für ostasiatische/komplexe Schriften das gleiche Font setzen
    rpr = normal.element.get_or_add_rPr()
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement("w:rFonts")
        rpr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), DEFAULT_FONT)
    rFonts.set(qn("w:hAnsi"), DEFAULT_FONT)
    rFonts.set(qn("w:cs"), DEFAULT_FONT)

    fn_mgr = FootnoteManager(document)
    seen_top_headings: set[str] = set()
    quote_state = [False]

    for idx, url in enumerate(urls):
        print(f"\n[{idx+1}/{len(urls)}] Lade {url} ...")
        try:
            soup = fetch(url)
        except Exception as e:
            print(f"  Fehler beim Laden: {e}")
            continue

        footnotes = extract_footnotes(soup)
        blocks = find_content_paragraphs(soup)
        print(f"  {len(blocks)} Block-Elemente, {len(footnotes)} Fußnoten gefunden.")

        # Seitenumbruch zwischen Kapiteln
        if idx > 0:
            document.add_page_break()

        used_ids: dict[str, int] = {}
        for block in blocks:
            # Wiederholte Top-Überschriften (Autor, Buchtitel) je Kapitel überspringen
            if block.name in ("h1", "h2"):
                key = block.get_text(strip=True).lower()
                if key in seen_top_headings:
                    continue
                seen_top_headings.add(key)
            render_element(document, block, fn_mgr, footnotes, used_ids,
                           opts, quote_state)
        # Zitat-Zustand pro Kapitel zurücksetzen, falls jemand vergessen hat
        # die Anführungszeichen zu schließen
        quote_state[0] = False

    document.save(output_path)
    print(f"\nFertig: {output_path.resolve()}")


if __name__ == "__main__":
    main()
