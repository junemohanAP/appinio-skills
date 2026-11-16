#!/usr/bin/env python3
"""Appinio deck builder: turns a JSON outline into an on-brand .pptx.

Usage:
  python build_deck.py outline.json -o deck.pptx
  python build_deck.py --list-assets          # icons and photos you can name

Every slide is created from the real slide layouts of the Appinio Slide Deck
Template 2026 (assets/template.pptx): backgrounds, logos, gradients and type
styles come from the template's own masters. Charts, cards and tables are drawn
on the template's "Title only" layout using the template's grid (1920 x 1080 px
canvas, 10 x 5.625 in page). Requires: python-pptx, Pillow. Viewers need Inter.
"""
import argparse, json, os, re, sys, copy
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.normpath(os.path.join(HERE, "..", "assets"))

# ---------------------------------------------------------------- tokens
INK, G700, G500, G300, LINE = "181D24", "39404C", "5F6675", "89909E", "D2D6DB"
G10, BLUE, BLUE10, WHITE, HL_DARK = "F7F7F7", "0038FF", "E5EBFF", "FFFFFF", "BFCDFF"
CHART = ["779DFF", "FFAF75", "8152E3", "88D298", "CD4E5F", "FFCD4C", "ECA1B8", "5E9A8F",
         "B997EF", "FF7D34", "6E7AA5", "FF6C63", "6285E1", "8BD9D0", "AEBBFF", "D1519E",
         "AACBB7", "FF9882", "F4B13B", "67AC65"]
CHART_OTHER = "C3C9D6"
FONT = {300: "Inter Light", 400: "Inter", 500: "Inter Medium", 600: "Inter SemiBold", 700: "Inter", 800: "Inter ExtraBold"}
TAGLINE = ["Rigorous, rapid research.", "Refreshingly human."]
OFFICES = [("Appinio GmbH", "Ballindamm 3\n20095 Hamburg\nGermany"),
           ("Appinio Inc.", "1651 N Broad St, Suite 206\nMiddletown 19709 Delaware\nUSA"),
           ("Appinio Limited", "7 Bell Yard\nWC2A 2JR London\nUK"),
           ("Appinio SLU", "Calle Víctor, 5, 1º, 3I\n48005 Bilbao\nSpain")]

EMU_PX = 9144000 / 1920          # 1 px of the 1920 canvas in EMU
PT_PX = 720 / 1920               # 1 px of the 1920 canvas in points
BODY_TOP, BODY_BOTTOM = 227, 900  # content area of the template (below title + subtitle, above logo)
TEMPLATE = os.path.join(ASSETS, "template.pptx")
WARN = []


def px(v):
    return Emu(int(round(v * EMU_PX)))


def warn(msg):
    WARN.append(msg)


# ---------------------------------------------------------------- assets
def asset(kind, name, ext):
    path = os.path.join(ASSETS, kind, f"{name}.{ext}")
    if not os.path.exists(path):
        raise SystemExit(f"Unknown {kind[:-1]} '{name}'. Run --list-assets.")
    return path


def photo_path(name):
    if os.path.isabs(name) or os.path.exists(name):
        return name
    return asset("photos", name.replace(".jpg", ""), "jpg")


def list_assets(kind=None):
    """Print asset names. Photos are grouped by topic (references/photo-catalog.json)."""
    cat_file = os.path.join(HERE, "..", "references", "photo-catalog.json")
    for k in ("photos", "icons", "gradients", "flags", "maps"):
        if kind and k != kind:
            continue
        names = sorted(os.path.splitext(f)[0] for f in os.listdir(os.path.join(ASSETS, k))
                       if not f.startswith("."))
        if k == "photos" and os.path.exists(cat_file):
            cat = json.load(open(cat_file, encoding="utf-8"))
            groups = {}
            for c in cat:
                groups.setdefault(c["category"], []).append(f'{c["name"]} ({c["orientation"][0]})')
            listed = {c["name"] for c in cat}
            groups["template"] = [n for n in names if n not in listed]
            print("photos (l = landscape, p = portrait, s = square):")
            for g in sorted(groups):
                print(f"  {g}: {', '.join(groups[g])}")
        else:
            print(f"{k}: {', '.join(names)}")


_SCALES = None


def scale(palette, n):
    """Exact n-step row of a data-viz palette (sequential, diverging, trafficLight, categorical)."""
    global _SCALES
    if _SCALES is None:
        _SCALES = json.load(open(os.path.join(HERE, "..", "references", "dataviz-scales.json"), encoding="utf-8"))
    row = _SCALES["palettes"][palette]["steps"].get(str(n))
    if not row:
        raise SystemExit(f"No {n}-step row in the {palette} palette.")
    return [c.lstrip("#").upper() for c in row]


DARK_FILLS = {"3463E0", "2D56C2", "2648A3", "1E3B84", "172D66", "8152E3", "926BE4",
              "6B032E", "A6293E", "D0434A", "4C795B", "2C4B3C", "0038FF", "5E9A8F", "6E7AA5", "67A176"}


def on_fill(fill):
    """Text colour for a data fill: white on the dark steps, ink on the rest."""
    return WHITE if fill.upper() in DARK_FILLS else INK


def fmt_num(v, fmt="0"):
    """Format a number like the Excel codes the outline uses: 0 · 0.0 · 0% (0–1) · 0"%" (whole numbers)."""
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    pct = fmt.endswith("%") and not fmt.endswith('"%"')
    suffix = "%" if "%" in fmt else ""
    dec = len(fmt.split(".")[1].rstrip('%"')) if "." in fmt else 0
    x = v * 100 if pct else v
    return f"{x:,.{dec}f}{suffix}"


def parse_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"[-+−]?\d[\d,]*\.?\d*", str(v))
    return float(m.group(0).replace(",", "").replace("−", "-")) if m else None


# ---------------------------------------------------------------- primitives
def est_lines(text, size, width, weight=400):
    """Rough line count for Inter at a px size inside a px width."""
    cw = size * (0.56 if weight >= 600 else 0.52)
    per = max(1, int(width / cw))
    lines = 0
    for para in str(text).split("\n"):
        words, cur = para.split(), 0
        n = 1
        for w in words:
            L = len(w) + (1 if cur else 0)
            if cur + L > per:
                n += 1; cur = len(w)
            else:
                cur += L
        lines += n
    return lines


def check_fit(label, text, size, width, height, weight=400, lh=1.2):
    need = est_lines(text, size, width, weight) * size * lh
    if need > height * 1.02:
        warn(f"{label}: text likely too long ({need:.0f}px needed, {height:.0f}px available). Shorten it.")


def _set_spc(run, spc_px):
    run._r.get_or_add_rPr().set("spc", str(int(round(spc_px * PT_PX * 100))))


def _alpha(run, alpha):
    clr = run._r.get_or_add_rPr().find(qn("a:solidFill")).find(qn("a:srgbClr"))
    a = etree.SubElement(clr, qn("a:alpha")); a.set("val", str(int(alpha * 100000)))


def text(slide, x, y, w, h, content, size=24, weight=400, color=G700, align="l",
         lh=1.2, spc=None, caps=False, anchor="t", alpha=None, hl=BLUE, bullets=False,
         para_gap=0, name=None):
    """content: str (\\n = new paragraph) or list of paragraphs. [[phrase]] = highlight colour."""
    tb = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    paras = content if isinstance(content, list) else str(content).split("\n")
    for i, ptxt in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = lh
        if para_gap and i:
            p.space_before = Pt(para_gap * PT_PX)
        if bullets:
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(28 * EMU_PX))); pPr.set("indent", str(int(-28 * EMU_PX)))
            bu = etree.SubElement(pPr, qn("a:buChar")); bu.set("char", "•")
        for j, part in enumerate(re.split(r"(\[\[.*?\]\])", str(ptxt))):
            if not part:
                continue
            is_hl = part.startswith("[[")
            r = p.add_run()
            r.text = part[2:-2] if is_hl else part
            f = r.font
            f.size = Pt(size * PT_PX)
            f.name = FONT.get(weight, "Inter")
            f.bold = weight == 700
            f.color.rgb = RGBColor.from_string(hl if is_hl else color)
            if spc is not None:
                _set_spc(r, spc)
            if caps:
                r._r.get_or_add_rPr().set("cap", "all")
            if alpha is not None:
                _alpha(r, alpha)
    return tb


def rect(slide, x, y, w, h, fill, radius=0, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                 px(x), px(y), px(w), px(h))
    if radius:
        shp.adjustments[0] = min(0.5, radius / min(w, h))
    shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line:
        shp.line.color.rgb = RGBColor.from_string(line); shp.line.width = Emu(int(EMU_PX))
    else:
        shp.line.fill.background()
    _nostyle(shp)
    return shp


def _nostyle(shp):
    """Drop the theme style reference (it adds shadows/outlines in some renderers)."""
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)


def hline(slide, x, y, w, color=LINE, weight=1):
    return rect(slide, x, y, w, weight, color)


def vline(slide, x, y, h, color=LINE, weight=1):
    return rect(slide, x, y, weight, h, color)


def picture(slide, path, x, y, w, h, radius=0, fit="cover"):
    iw, ih = Image.open(path).size
    if fit == "contain":
        s = min(w / iw, h / ih); nw, nh = iw * s, ih * s
        return slide.shapes.add_picture(path, px(x + (w - nw) / 2), px(y + (h - nh) / 2), px(nw), px(nh))
    pic = slide.shapes.add_picture(path, px(x), px(y), px(w), px(h))
    box, img = w / h, iw / ih
    if img > box:
        c = (1 - box / img) / 2; pic.crop_left = pic.crop_right = c
    else:
        c = (1 - img / box) / 2; pic.crop_top = pic.crop_bottom = c
    if radius:
        geom = pic._element.spPr.find(qn("a:prstGeom"))
        geom.set("prst", "roundRect")
        av = geom.find(qn("a:avLst"))
        if av is None:
            av = etree.SubElement(geom, qn("a:avLst"))
        gd = etree.SubElement(av, qn("a:gd")); gd.set("name", "adj")
        gd.set("fmla", f"val {int(radius / min(w, h) * 100000)}")
    return pic



BLUE50, CHAP_NUM, PANEL_GREY = "809BFF", "CCCCCC", "F5F6F8"
STYLE = {"name": "master"}


def is_results():
    return STYLE["name"] == "results"


def logo(slide, dark=True, x=66, y=982, w=143, h=37):
    picture(slide, asset("logos", "logo-dark" if dark else "logo-white", "png"), x, y, w, h, fit="contain")


def gradient_path(name):
    return asset("gradients", name, "jpg")


def notes(slide, s):
    if s.get("notes"):
        slide.notes_slide.notes_text_frame.text = s["notes"]


def to_back(slide, shape, after=2):
    """Move a shape to just above the first `after` elements of the spTree (behind later shapes)."""
    tree = slide.shapes._spTree
    tree.remove(shape._element)
    tree.insert(after, shape._element)


def scrim(slide, alpha, x=0, y=0, w=1920, h=1080, behind_index=3):
    r = rect(slide, x, y, w, h, INK)
    sf = r.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    a = etree.SubElement(sf, qn("a:alpha")); a.set("val", str(int(alpha * 100000)))
    to_back(slide, r, behind_index)
    return r


# ---------------------------------------------------------------- template plumbing
class Slide:
    """A slide made from a named template layout; fills placeholders by idx."""

    def __init__(self, prs, layout, dark=False):
        lay = next((l for l in prs.slide_layouts if l.name == "Appinio – " + layout), None)
        if lay is None:
            raise SystemExit(f"Template layout 'Appinio – {layout}' not found in template.pptx")
        self.s = prs.slides.add_slide(lay)
        self.used = set()
        self.dark = dark

    def ph(self, idx):
        for p in self.s.placeholders:
            if p.placeholder_format.idx == idx:
                return p
        return None

    def put(self, idx, content, bullets=False, check=None, align=None, indent=28):
        """Write into placeholder idx, inheriting the template's styling.
        [[phrase]] = highlight colour; a paragraph starting **Headline** gets a bold lead-in."""
        p = self.ph(idx)
        if p is None or content in (None, "", []):
            return
        self.used.add(idx)
        tf = p.text_frame
        paras = content if isinstance(content, list) else str(content).split("\n")
        tf.text = ""
        for i, ptxt in enumerate(paras):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            pPr = para._p.get_or_add_pPr()
            if align:
                pPr.set("algn", {"l": "l", "c": "ctr", "r": "r"}[align])
            if bullets:
                pPr.set("marL", str(int(indent * EMU_PX))); pPr.set("indent", str(int(-indent * EMU_PX)))
                bu = etree.SubElement(pPr, qn("a:buChar")); bu.set("char", "•")
            elif align == "l":
                pPr.set("marL", "0"); pPr.set("indent", "0")
                etree.SubElement(pPr, qn("a:buNone"))
            for part in re.split(r"(\[\[.*?\]\]|\*\*.*?\*\*)", str(ptxt)):
                if not part:
                    continue
                r = para.add_run()
                if part.startswith("[["):
                    r.text = part[2:-2]
                    r.font.color.rgb = RGBColor.from_string(HL_DARK if self.dark else BLUE)
                elif part.startswith("**"):
                    r.text = part[2:-2]; r.font.bold = True
                else:
                    r.text = part
        if check:
            label, size, weight, lh = check
            w, h = p.width / EMU_PX, p.height / EMU_PX
            check_fit(label, re.sub(r"\[\[|\]\]|\*\*", "", "\n".join(paras)), size, w, h, weight, lh)

    def pic(self, idx, path):
        p = self.ph(idx)
        if p is None or not path:
            return None
        self.used.add(idx)
        return p.insert_picture(photo_path(path))

    def finish(self, s):
        for p in list(self.s.placeholders):
            if p.placeholder_format.idx not in self.used:
                p._element.getparent().remove(p._element)
        notes(self.s, s)
        return self.s


# ---------------------------------------------------------------- style chrome
def title_block(sl, s, title_idx=0, sub_idx=3, width=1786):
    """Master style: title + grey subtitle in the template's placeholders."""
    t, sub = s.get("title", ""), s.get("subtitle") or s.get("eyebrow")
    sl.put(title_idx, t)
    sl.put(sub_idx, sub)
    if sub and est_lines(t, 45, width, 800) > 1:
        warn("Title wraps to two lines and collides with the subtitle: shorten it to one line (~60 chars) or drop the subtitle.")
    elif est_lines(t, 45, width, 800) > 2:
        warn("Title longer than two lines: shorten it.")


def results_header(slide, s, width=1786, x=66):
    """Results style: dark grey uppercase eyebrow above a 45px title (drawn)."""
    eyebrow = s.get("eyebrow") or s.get("subtitle")
    y = 67
    if eyebrow:
        text(slide, x, y, width, 32, eyebrow, size=24, weight=700, color=G700, spc=2, caps=True, lh=1.3)
        y += 42
    text(slide, x, y, width, 105, s.get("title", ""), size=45, weight=700, color=INK, spc=-0.5, lh=1.08)
    check_fit("Title (max two lines)", s.get("title", ""), 45, width, 105, 700, 1.12)


def note(slide, s, width=1786, x=66):
    """Footer note. Results: divider at y 900 + note under it. Master: note at x 685 on the logo row."""
    n = s.get("note")
    if is_results():
        hline(slide, x, 900, width)
        if n:
            text(slide, x, 914, width, 40, n, size=24, color=G500)
    elif n:
        text(slide, 685, 944, 1167, 72, n, size=24, color=G500, anchor="b", lh=1.2)


def content_slide(prs, s, width=1786):
    """The base for every drawn content slide, in the deck's style. Returns (Slide, body_top, body_bottom)."""
    if is_results():
        sl = Slide(prs, "Blank")
        results_header(sl.s, s, width)
        return sl, 270, 868
    sl = Slide(prs, "Title only")
    title_block(sl, s, 0, 3)
    return sl, 227, 940


def boxes(slide, top, bottom, x=66, w=1786):
    """Card colour for the style. Master draws a Blue 10 panel first and puts white cards on it."""
    if is_results():
        return G10, (x, top, w, bottom - top)
    rect(slide, x, top, w, bottom - top, BLUE10, radius=16)
    return WHITE, (x + 32, top + 32, w - 64, bottom - top - 64)


def _grid(n, area, cols=None):
    x0, y0, W, H = area
    cols = cols or (3 if n in (3, 5, 6) else 2 if n in (2, 4) else max(1, min(n, 3)))
    rows = (n + cols - 1) // cols
    w = (W - (cols - 1) * 32) / cols
    h = (H - (rows - 1) * 32) / rows
    return [(x0 + (i % cols) * (w + 32), y0 + (i // cols) * (h + 32), w, h) for i in range(n)]


# ---------------------------------------------------------------- shared layouts
def L_cover(prs, s, deck):
    photo = s.get("photo")
    sl = Slide(prs, "Cover photo" if photo else "Cover", dark=True)
    if photo:
        pic = sl.pic(2, photo)
        scrim(sl.s, 0.45)
    sl.put(0, s.get("title", "Presentation title"), check=("Cover title (max 2 short lines)", 171, 700, 0.96))
    sl.put(1, s.get("subtitle"))
    if s.get("partner_logo"):
        rect(sl.s, 385, 67, 2, 73, WHITE)
        picture(sl.s, photo_path(s["partner_logo"]), 427, 67, 279, 73, fit="contain")
    return sl.finish(s)


def L_agenda(prs, s, deck):
    ch = [c if isinstance(c, dict) else {"title": c} for c in s["chapters"]]
    if is_results() and len(ch) <= 6:
        # SlideAgenda: "Contents", blue numbers, page numbers, hairlines, two columns
        sl = Slide(prs, "Blank")
        text(sl.s, 66, 67, 1786, 80, s.get("title", "Contents"), size=61, weight=800, color=INK, spc=-0.5)
        for i, c in enumerate(ch):
            col, row = divmod(i, 3)
            x, y = 66 + col * 927, 230 + row * 200
            text(sl.s, x, y + 40, 110, 70, f"{i + 1:02d}", size=45, weight=800, color=BLUE)
            text(sl.s, x + 120, y + 40, 600, 120, c["title"], size=45, weight=500, color=INK, lh=1.1)
            check_fit(f"Agenda row {i + 1}", c["title"], 45, 600, 120, 500, 1.1)
            if c.get("page"):
                text(sl.s, x + 720, y + 44, 140, 60, str(c["page"]), size=40, color=G300, align="r")
            hline(sl.s, x, y + 180, 860)
        note(sl.s, s)
        return sl.finish(s)
    if len(ch) <= 6:
        sl = Slide(prs, "Agenda")
        nums, titles = [7, 8, 9, 13, 14, 15], [0, 2, 3, 4, 5, 6]
        for i, c in enumerate(ch):
            sl.put(nums[i], f"{i + 1:02d}")
            sl.put(titles[i], c["title"], check=(f"Agenda row {i + 1}", 107, 700, 1.0))
    else:
        sl = Slide(prs, "Agenda long")
        sl.put(0, s.get("title", "Agenda"))
        left_t = [2, 3, 4, 5, 6, 7, 8, 9, 13, 14]; right_t = list(range(15, 25))
        left_n = list(range(25, 44, 2)); right_n = list(range(26, 45, 2))
        for i, c in enumerate(ch[:20]):
            t, n = (left_t[i], left_n[i]) if i < 10 else (right_t[i - 10], right_n[i - 10])
            sl.put(n, f"{i + 1:02d}"); sl.put(t, c["title"])
        if len(ch) > 20:
            warn("Agenda: more than 20 chapters; only the first 20 are shown.")
    return sl.finish(s)


def L_chapter(prs, s, deck):
    sl = Slide(prs, "Chapter", dark=True)
    deck["_chapter"] = s.get("number", deck.get("_chapter", 0) + 1)
    num = f"{int(deck['_chapter']):02d}"
    if is_results():
        # SlideChapter: number, title and one line of context, flush left at mid-height
        text(sl.s, 66, 260, 1500, 220, num, size=200, weight=700, color=WHITE, alpha=0.55, lh=1.0)
        text(sl.s, 66, 500, 1500, 240, s["title"], size=107, weight=700, color=WHITE, lh=1.0, spc=-1)
        check_fit("Chapter title (max ~24 chars)", s["title"], 107, 1500, 240, 700, 1.05)
        if s.get("description"):
            text(sl.s, 66, 760, 1400, 110, s["description"], size=40, color=WHITE, lh=1.25)
    else:
        sl.put(2, num)
        sl.put(0, s["title"], check=("Chapter title (max ~30 chars)", 107, 700, 1.0))
        if s.get("description"):
            warn("Chapter: master-style chapters have no description line; it was left out.")
    return sl.finish(s)


def L_section_intro(prs, s, deck):
    if is_results():
        sl = Slide(prs, "Blank")
        results_header(sl.s, s, 860)
        text(sl.s, 993, 67, 860, 800, s["intro"], size=32, weight=600, color=INK, lh=1.35, para_gap=24)
        check_fit("Section intro", s["intro"], 32, 860, 800, 600, 1.35)
        note(sl.s, s)
        return sl.finish(s)
    sl = Slide(prs, "Section intro")
    sl.put(0, s["title"]); sl.put(2, s.get("subtitle"))
    sl.put(1, s["intro"], check=("Section intro", 32, 600, 1.3), align="l")
    return sl.finish(s)


def L_content(prs, s, deck):
    if is_results():
        sl, top, bot = content_slide(prs, s)
        h = bot - top
        if s.get("body"):
            text(sl.s, 66, top, 860, h, s["body"], size=24, color=G700, lh=1.4, para_gap=16)
            check_fit("Body (left)", s["body"], 24, 860, h, 400, 1.4)
        if s.get("bullets"):
            rect(sl.s, 993, top, 860, h, G10, radius=16)
            text(sl.s, 1033, top + 40, 780, h - 80, s["bullets"][:5], size=24, color=G700, lh=1.5, bullets=True, para_gap=12)
        elif s.get("body_right"):
            text(sl.s, 993, top, 860, h, s["body_right"], size=24, color=G700, lh=1.4, para_gap=16)
        note(sl.s, s)
        return sl.finish(s)
    sl = Slide(prs, "Two columns")
    title_block(sl, s, 0, 3)
    sl.put(1, s.get("body"), check=("Body (left)", 24, 400, 1.4), align="l")
    if s.get("bullets"):
        sl.put(4, s["bullets"][:6], bullets=True, check=("Bullets (right)", 24, 400, 1.5), align="l")
    else:
        sl.put(4, s.get("body_right"), check=("Body (right)", 24, 400, 1.4), align="l")
    note(sl.s, s)
    return sl.finish(s)


def L_image(prs, s, deck):
    if is_results():
        sl = Slide(prs, "Blank")
        picture(sl.s, photo_path(s["photo"]), 993, 0, 927, 1080)
        results_header(sl.s, s, 860)
        if s.get("body"):
            text(sl.s, 66, 270, 860, 598, s["body"], size=24, color=G700, lh=1.4, para_gap=16)
            check_fit("Image slide body (~60 words)", s["body"], 24, 860, 598, 400, 1.4)
        note(sl.s, s, width=860)
        return sl.finish(s)
    sl = Slide(prs, "Text + photo right")
    sl.put(0, s["title"]); sl.put(4, s.get("subtitle"))
    sl.put(1, s.get("body"), check=("Image slide body (~80 words)", 24, 400, 1.4), align="l")
    sl.pic(3, s["photo"])
    return sl.finish(s)


def L_stat(prs, s, deck):
    hl = s.get("highlight")
    if is_results():
        sl, top, bot = content_slide(prs, s)
        st = s["stats"][:3]
        n = len(st); w = (1786 - (n - 1) * 32) / n; h = bot - top
        for i, k in enumerate(st):
            x = 66 + i * (w + 32)
            rect(sl.s, x, top, w, h, G10, radius=16)
            cy = top + (h - 330) / 2
            text(sl.s, x + 40, cy, w - 80, 32, k.get("label", ""), size=24, weight=700, color=G700, spc=2, caps=True, align="c", lh=1.3)
            text(sl.s, x + 20, cy + 56, w - 40, 150, k["value"], size=133, weight=700,
                 color=BLUE if hl == i else INK, align="c", lh=1.0, spc=-1)
            check_fit(f"Stat {i + 1} value", k["value"], 133, w - 40, 150, 700, 1.0)
            text(sl.s, x + 40, cy + 230, w - 80, 120, k.get("meaning", ""), size=32, weight=600, color=INK, align="c", lh=1.25)
        note(sl.s, s)
        return sl.finish(s)
    sl = Slide(prs, "Big numbers")
    title_block(sl, s, 0, 3)
    for i, k in enumerate(s["stats"][:3]):
        sl.put([4, 7, 13][i], k.get("label"))
        v = k["value"]
        sl.put([2, 5, 8][i], f"[[{v}]]" if hl == i else v, check=(f"Stat {i + 1} value", 133, 700, 0.85))
        sl.put([1, 6, 9][i], k.get("meaning"))
    note(sl.s, s)
    return sl.finish(s)


def L_statement(prs, s, deck):
    lines = s.get("lines") or [TAGLINE[0], f"[[{TAGLINE[1]}]]"]
    if s.get("photo"):
        sl = Slide(prs, "Statement photo", dark=True)
        sl.pic(2, s["photo"]); scrim(sl.s, 0.35)
        sl.put(0, [re.sub(r"\[\[|\]\]", "", l) for l in lines], align="c")
        logo(sl.s, dark=False)
        return sl.finish(s)
    grad = s.get("gradient", False)
    sl = Slide(prs, "Statement gradient" if grad else "Statement", dark=bool(grad))
    if isinstance(grad, str):
        to_back(sl.s, picture(sl.s, gradient_path(grad), 0, 0, 1920, 1080))
    sl.put(0, lines)
    for i, l in enumerate(lines):
        if len(re.sub(r"\[\[|\]\]", "", l)) > 40:
            warn(f"Statement line {i + 1}: over ~40 characters, it will wrap.")
    return sl.finish(s)


def L_quote(prs, s, deck):
    q = "“" + s["quote"].strip("“”\"") + "”"
    who = [x for x in (s.get("name"), s.get("role")) if x]
    if s.get("photo"):
        sl = Slide(prs, "Quote + photo")
        sl.put(0, q, check=("Quote (max ~25 words)", 93, 700, 1.15)); sl.put(1, who); sl.pic(2, s["photo"])
    else:  # MasterTestimonial
        sl = Slide(prs, "Quote")
        sl.put(0, q, check=("Quote (max ~35 words)", 93, 700, 1.15)); sl.put(1, who)
        sl.pic(2, s.get("avatar"))
        if s.get("client_logo"):
            picture(sl.s, photo_path(s["client_logo"]), 1593, 900, 260, 110, fit="contain")
    return sl.finish(s)


def L_key_findings(prs, s, deck):
    sl = Slide(prs, "Blank")
    text(sl.s, 66, 67, 600, 300, s.get("title", "Key\nFindings"), size=133, weight=800, color=INK, lh=1.0, spc=-1)
    f = s["findings"][:4]
    bottom = 868 if is_results() else 920
    slot = (bottom - 67) / len(f)
    for i, t in enumerate(f):
        y = 67 + i * slot
        if i:
            hline(sl.s, 685, y, 1167)
        text(sl.s, 685, y + 32, 1167, slot - 64, t, size=29, weight=500, color=INK, lh=1.35)
        check_fit(f"Finding {i + 1}", t, 29, 1167, slot - 64, 500, 1.35)
    note(sl.s, s)
    return sl.finish(s)


def L_fact_cards(prs, s, deck):
    sl, top, bot = content_slide(prs, s)
    fill, area = boxes(sl.s, top, bot)
    cards = s["cards"][:6]
    for i, ((x, y, w, h), c) in enumerate(zip(_grid(len(cards), area), cards)):
        rect(sl.s, x, y, w, h, fill, radius=16)
        cy = y + 40
        if c.get("icon") and h < 360:
            picture(sl.s, asset("icons", c["icon"], "png"), x + 40, cy, 64, 64, fit="contain")
            text(sl.s, x + 124, cy, w - 164, 64, c["label"], size=24, weight=700, color=G700, spc=2, caps=True, lh=1.3, anchor="m")
            cy += 84
        else:
            if c.get("icon"):
                picture(sl.s, asset("icons", c["icon"], "png"), x + 40, cy, 80, 80, fit="contain"); cy += 104
            text(sl.s, x + 40, cy, w - 80, 32, c["label"], size=24, weight=700, color=G700, spc=2, caps=True, lh=1.3); cy += 44
        if c.get("value"):
            text(sl.s, x + 40, cy, w - 80, 60, c["value"], size=45, weight=800, color=INK, lh=1.1); cy += 66
        if c.get("detail"):
            dh = y + h - 40 - cy
            text(sl.s, x + 40, cy, w - 80, dh, c["detail"], size=24, color=G700, lh=1.4)
            check_fit(f"Fact card {i + 1} detail", c["detail"], 24, w - 80, dh, 400, 1.4)
    note(sl.s, s)
    return sl.finish(s)


def L_icon_points(prs, s, deck):
    sl, top, bot = content_slide(prs, s)
    fill, (ax, ay, aw, ah) = boxes(sl.s, top, bot)
    pts = s["points"][:4]
    n = len(pts); w = (aw - (n - 1) * 32) / n
    use_icons = all(p.get("icon") for p in pts)
    for i, p in enumerate(pts):
        x = ax + i * (w + 32)
        rect(sl.s, x, ay, w, ah, fill, radius=16)
        cy = ay + 40
        if use_icons:
            picture(sl.s, asset("icons", p["icon"], "png"), x + 40, cy, 96, 96, fit="contain"); cy += 128
        text(sl.s, x + 40, cy, w - 80, 100, p["headline"], size=35, weight=700, color=INK, lh=1.15)
        cy += est_lines(p["headline"], 35, w - 80, 700) * 35 * 1.2 + 20
        text(sl.s, x + 40, cy, w - 80, ay + ah - 40 - cy, p["text"], size=24, color=G700, lh=1.4)
        check_fit(f"Point {i + 1} text", p["text"], 24, w - 80, ay + ah - 40 - cy, 400, 1.4)
    note(sl.s, s)
    return sl.finish(s)


LETTERS = "ABCDEFGHIJKL"
STACKED = {"stacked_bar": XL_CHART_TYPE.BAR_STACKED_100, "stacked_column": XL_CHART_TYPE.COLUMN_STACKED_100}


def _rich_label(point, parts, size=24, position=None):
    """Custom data label: parts = [(text, colour, bold)]."""
    dl = point.data_label
    tf = dl.text_frame
    para = tf.paragraphs[0]
    for t, col, bold in parts:
        r = para.add_run(); r.text = t
        r.font.size = Pt(size * PT_PX); r.font.bold = bold
        r.font.name = "Inter"
        r.font.color.rgb = RGBColor.from_string(col)
    if position is not None:
        dl.position = position


def _style_chart(chart, kind, n_series, fmt, colors, series=None):
    chart.has_title = False
    chart.font.name = "Inter"; chart.font.size = Pt(24 * PT_PX); chart.font.color.rgb = RGBColor.from_string(INK)
    cs = chart._chartSpace
    spPr = cs.find(qn("c:spPr"))
    if spPr is None:
        spPr = etree.SubElement(cs, qn("c:spPr"))
    etree.SubElement(spPr, qn("a:noFill"))
    plot = chart.plots[0]
    stacked = kind in STACKED
    if kind != "line":
        plot.gap_width = 40 if stacked else 60
        if stacked:
            plot.overlap = 100
        elif n_series > 1:
            plot.overlap = -10
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = fmt; dl.number_format_is_linked = False
    dl.font.size = Pt(24 * PT_PX); dl.font.name = "Inter SemiBold" if n_series > 1 else "Inter"
    if stacked:
        dl.position = XL_LABEL_POSITION.CENTER
    else:
        dl.position = XL_LABEL_POSITION.OUTSIDE_END if kind != "line" else XL_LABEL_POSITION.ABOVE
    va = chart.value_axis
    va.visible = False; va.has_major_gridlines = False
    if kind != "line" and not stacked:
        va.minimum_scale = 0
    ca = chart.category_axis
    ca.has_major_gridlines = False
    ca.format.line.color.rgb = RGBColor.from_string(LINE)
    ca.tick_labels.font.size = Pt(24 * PT_PX)
    if kind in ("bar", "stacked_bar"):
        ca.reverse_order = True
    sig_any = series and any(se.get("sig") for se in series)
    for i, ser in enumerate(plot.series):
        c = colors[i % len(colors)]
        rgb = RGBColor.from_string(c)
        if kind == "line":
            ser.format.line.color.rgb = rgb; ser.format.line.width = Emu(int(6 * EMU_PX))
            ser.smooth = False
            ser.marker.style = 8; ser.marker.size = 9
            ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = rgb
            ser.marker.format.line.color.rgb = rgb
            if n_series > 1:
                ser.data_labels.position = XL_LABEL_POSITION.ABOVE if i % 2 == 0 else XL_LABEL_POSITION.BELOW
                ser.data_labels.number_format = fmt; ser.data_labels.number_format_is_linked = False
                ser.data_labels.font.size = Pt(24 * PT_PX)
                ser.data_labels.show_value = True
        else:
            ser.format.fill.solid(); ser.format.fill.fore_color.rgb = rgb
            if stacked:
                sdl = ser.data_labels
                sdl.show_value = True
                sdl.font.color.rgb = RGBColor.from_string(on_fill(c))
                sdl.font.size = Pt(24 * PT_PX)
                sdl.number_format = fmt; sdl.number_format_is_linked = False
                sdl.position = XL_LABEL_POSITION.CENTER
                small = 0.05 if fmt.endswith("%") and not fmt.endswith('"%"') else 5
                for j, v in enumerate(series[i]["values"] if series else []):
                    if v is not None and v < small:      # too thin to label
                        _rich_label(ser.points[j], [("", on_fill(c), False)])
                    elif v is not None:
                        _rich_label(ser.points[j], [(fmt_num(v, fmt), on_fill(c), False)],
                                    position=XL_LABEL_POSITION.CENTER)
        if series and not stacked:
            sd = series[i]
            sig = sd.get("sig") or []
            for j, v in enumerate(sd["values"]):
                if v is None:
                    continue
                letters = sig[j] if j < len(sig) else ""
                if not sig_any and not sd.get("labels"):
                    continue
                lab = (sd.get("labels") or [None] * len(sd["values"]))[j]
                parts = [(lab if lab is not None else fmt_num(v, fmt), INK, False)]
                for L in str(letters or "").replace(",", " ").split():
                    k = LETTERS.find(L.upper())
                    parts.append((" " + L.upper(), colors[k % len(colors)] if k >= 0 else INK, True))
                pos = None
                if kind == "line":
                    pos = XL_LABEL_POSITION.ABOVE if i % 2 == 0 else XL_LABEL_POSITION.BELOW
                else:
                    pos = XL_LABEL_POSITION.OUTSIDE_END
                _rich_label(ser.points[j], parts, position=pos)
    chart.has_legend = n_series > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(24 * PT_PX)


def L_chart(prs, s, deck):
    """Bar / column / line (SlideChart, SlideSplitChart) and 100 % stacked bars for scales (SlideScaleLogic).
    Series may carry `sig` (significance letters per value, e.g. ["B", "", "A C"]); cuts are lettered A, B, C."""
    sl, top, bot = content_slide(prs, s)
    kind = s.get("kind", "column")
    if kind not in ("bar", "column", "line") and kind not in STACKED:
        raise SystemExit(f"chart kind '{kind}' unknown: bar, column, line, stacked_bar, stacked_column")
    take = s.get("takeaways")
    pw = 1786 - (574 + 32 if take else 0)
    panel = G10 if is_results() else BLUE10
    rect(sl.s, 66, top, pw, bot - top, panel, radius=16)
    series = s["series"]
    sig_any = any(se.get("sig") for se in series)
    if sig_any and not is_results():
        warn("Significance letters belong in results decks; master-style charts stay simple.")
    if kind in STACKED and len(series) < 2:
        warn("Stacked charts need one series per scale point (e.g. 5 series for a 5-point scale).")
    cd = CategoryChartData(); cd.categories = s["categories"]
    for i, ser in enumerate(series):
        name = ser["name"]
        if sig_any and kind not in STACKED and not re.match(rf"{LETTERS[i]}\s", name):
            name = f"{LETTERS[i]}  {name}"
        cd.add_series(name, ser["values"])
    ct = STACKED.get(kind) or {"bar": XL_CHART_TYPE.BAR_CLUSTERED, "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
                               "line": XL_CHART_TYPE.LINE_MARKERS}[kind]
    gf = sl.s.shapes.add_chart(ct, px(66 + 32), px(top + 24), px(pw - 64), px(bot - top - 48), cd)
    if s.get("colors"):
        colors = [c.lstrip("#").upper() for c in s["colors"]]
    elif kind in STACKED:
        colors = scale("sequential", len(series)) if len(series) >= 3 else CHART
        warn("Stacked chart without `colors`: pass the diverging / traffic-light / sequential n-step row "
             "for the scale (references/dataviz-scales.json). Sequential used as a fallback.")
    else:
        colors = ["779DFF"] if len(series) == 1 else CHART
    if len(series) > 1 and len(colors) < len(series):
        warn("Fewer colours than series: colours repeat.")
    _style_chart(gf.chart, kind, len(series), s.get("number_format", "0"), colors, series)
    if sig_any and not s.get("note"):
        warn("Significance letters used but no note: name the cuts in `note` (e.g. 'A 18–29, B 30–44 · Base: …').")
    if take:
        x = 66 + pw + 32
        rect(sl.s, x, top, 574, bot - top, panel, radius=16)
        text(sl.s, x + 40, top + 40, 494, 32, s.get("takeaways_label", "Takeaways"), size=24, weight=700,
             color=G700, spc=2, caps=True, lh=1.3)
        slot = (bot - top - 120) / len(take[:3])
        for i, t in enumerate(take[:3]):
            y = top + 96 + i * slot
            if i:
                hline(sl.s, x + 40, y, 494)
            text(sl.s, x + 40, y + 20, 494, slot - 40, t, size=24, color=G700, lh=1.4)
            check_fit(f"Takeaway {i + 1}", t, 24, 494, slot - 40, 400, 1.4)
    note(sl.s, s)
    return sl.finish(s)


def L_profile(prs, s, deck):
    """SlideProfileChart: semantic differential, connected lines per brand on a grey panel."""
    sl, top, bot = content_slide(prs, s)
    pairs = s["pairs"]; series = s["series"]
    take = s.get("takeaways")
    lw = 300                                     # pole label width
    total = 1786 - (574 + 32 if take else 0)
    px0 = 66 + lw + 24; pw = total - 2 * (lw + 24)
    # legend directly above the plot
    lx = px0
    for i, ser in enumerate(series):
        c = (s.get("colors") or CHART)[i]
        hline(sl.s, lx, top + 16, 36, color=c, weight=6)
        circ = sl.s.shapes.add_shape(MSO_SHAPE.OVAL, px(lx + 11), px(top + 8), px(14), px(14))
        circ.fill.solid(); circ.fill.fore_color.rgb = RGBColor.from_string(c); circ.line.fill.background(); _nostyle(circ)
        text(sl.s, lx + 48, top, 300, 32, ser["name"], size=24, color=INK)
        lx += 60 + est_lines(ser["name"], 24, 9999) * 0 + len(ser["name"]) * 13 + 40
    ptop, pbot = top + 56, bot - 48
    rect(sl.s, px0, ptop, pw, pbot - ptop, PANEL_GREY, radius=12)
    for g in (2, 3, 4):
        vline(sl.s, px0 + (g - 1) / 4 * pw, ptop, pbot - ptop, color=LINE, weight=1)
    n = len(pairs); row = (pbot - ptop) / n
    ys = [ptop + row * (i + 0.5) for i in range(n)]
    for i, (lft, rgt) in enumerate(pairs):
        text(sl.s, 66, ys[i] - 20, lw, 40, lft, size=28, weight=600, color=INK, align="r", anchor="m")
        text(sl.s, px0 + pw + 24, ys[i] - 20, lw, 40, rgt, size=28, weight=600, color=INK, anchor="m")
    for g in range(1, 6):
        text(sl.s, px0 + (g - 1) / 4 * pw - 20, pbot + 8, 40, 32, str(g), size=24, color=G500, align="c")
    for si in reversed(range(len(series))):      # lead brand drawn last, on top
        c = (s.get("colors") or CHART)[si]
        pts = [(px0 + (v - 1) / 4 * pw, ys[i]) for i, v in enumerate(series[si]["values"])]
        fb = sl.s.shapes.build_freeform(px(pts[0][0]), px(pts[0][1]), scale=1.0)
        fb.add_line_segments([(px(x), px(y)) for x, y in pts[1:]], close=False)
        ln = fb.convert_to_shape()
        ln.fill.background(); ln.line.color.rgb = RGBColor.from_string(c); ln.line.width = Emu(int(6 * EMU_PX))
        _nostyle(ln)
        for x, y in pts:
            d = sl.s.shapes.add_shape(MSO_SHAPE.OVAL, px(x - 14), px(y - 14), px(28), px(28))
            d.fill.solid(); d.fill.fore_color.rgb = RGBColor.from_string(c); d.line.fill.background(); _nostyle(d)
    if take:
        x = 66 + total + 32
        rect(sl.s, x, top, 574, bot - top, G10, radius=16)
        text(sl.s, x + 40, top + 40, 494, 32, "Takeaways", size=24, weight=700, color=G700, spc=2, caps=True)
        slot = (bot - top - 120) / len(take[:3])
        for i, t in enumerate(take[:3]):
            y = top + 96 + i * slot
            if i:
                hline(sl.s, x + 40, y, 494)
            text(sl.s, x + 40, y + 20, 494, slot - 40, t, size=24, color=G700, lh=1.4)
    note(sl.s, s)
    return sl.finish(s)


def _heat_bands(rows, cols_idx, n):
    vals = [parse_num(r[c]) for r in rows for c in cols_idx if c < len(r)]
    vals = [v for v in vals if v is not None]
    if not vals:
        return None
    lo, hi = min(vals), max(vals)
    step = (hi - lo) / n or 1
    return lambda v: min(n - 1, int((v - lo) / step)) if v is not None else None


def L_table(prs, s, deck):
    """Tables. Results extras (SlideSplitChart rules):
    heatmap: true | {"bands": 3–5, "columns": [indexes]}  – sequential n-step fills, white text on dark steps
    sig: matrix like rows (letters per cell, "" for none) – data columns are lettered A, B, C
    delta_columns: [indexes] – values starting with + / − in Bold green / red"""
    sl, top, bot = content_slide(prs, s)
    cols, rows = s["columns"], s["rows"]
    hlc = s.get("highlight_column")
    sig = s.get("sig")
    heat = s.get("heatmap")
    deltas = set(s.get("delta_columns", []))
    data_cols = [c for c in range(1, len(cols)) if c not in deltas]
    band_of, heat_cols, heat_row = None, set(), []
    if heat:
        hc = heat.get("columns") if isinstance(heat, dict) else None
        heat_cols = set(hc or data_cols)
        n = (heat.get("bands") if isinstance(heat, dict) else None) or 4
        heat_row = scale("sequential", n)
        band_of = _heat_bands(rows, sorted(heat_cols), n)
        if not is_results():
            warn("Heatmap tables belong in results decks.")
    if len(rows) > 8:
        warn(f"{len(rows)} rows: tables take at most 8 rows; split the table or cut rows.")
    head = list(cols)
    if sig:
        for k, c in enumerate(data_cols):
            if not re.match(rf"{LETTERS[k]}\s", str(cols[c])):
                head[c] = f"{LETTERS[k]}  {cols[c]}"
        if not s.get("note"):
            warn("Significance letters used but no note: name the lettered columns and the test in `note`.")
    nr = len(rows) + 1
    rh = min(90, (bot - top) / nr)
    gf = sl.s.shapes.add_table(nr, len(cols), px(66), px(top), px(1786), px(rh * nr))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for a in ("bandRow", "firstRow"):
        tblPr.set(a, "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    hsize = 24 if is_results() else 32
    bsize = 24 if is_results() else 28
    if s.get("first_column_width"):
        fw = s["first_column_width"]
        tbl.columns[0].width = px(fw)
        for c in range(1, len(cols)):
            tbl.columns[c].width = px((1786 - fw) / (len(cols) - 1))
    for r in range(nr):
        tbl.rows[r].height = px(rh)
        for c in range(len(cols)):
            cell = tbl.cell(r, c)
            raw = head[c] if r == 0 else rows[r - 1][c]
            val = str(raw)
            fill = WHITE
            col = INK if (r == 0 or c == 0) else G700
            if r == 0:
                fill = BLUE if hlc == c else BLUE10
                col = WHITE if hlc == c else INK
            elif band_of and c in heat_cols:
                b = band_of(parse_num(raw))
                if b is not None:
                    fill = heat_row[b]; col = on_fill(fill)
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor.from_string(fill)
            cell.margin_left = cell.margin_right = px(20); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame; tf.text = ""
            para = tf.paragraphs[0]
            if r > 0 and c > 0 and (heat or sig):
                para.alignment = PP_ALIGN.CENTER
            run = para.add_run(); run.text = val
            run.font.size = Pt((hsize if r == 0 else bsize) * PT_PX)
            run.font.name = "Inter SemiBold" if (c == 0 and r > 0 and not is_results()) else "Inter"
            run.font.bold = r == 0 or (c == 0 and is_results())
            if r > 0 and c in deltas and val[:1] in "+-−":
                col = "4C795B" if val[:1] == "+" else "D0434A"
                run.font.bold = True
            run.font.color.rgb = RGBColor.from_string(col)
            if r > 0 and sig and c > 0:
                letters = sig[r - 1][c] if c < len(sig[r - 1]) else ""
                for L in str(letters or "").replace(",", " ").split():
                    k = LETTERS.find(L.upper())
                    lr = para.add_run(); lr.text = " " + L.upper()
                    lr.font.size = Pt(bsize * PT_PX); lr.font.bold = True; lr.font.name = "Inter"
                    lc = col if fill != WHITE else (CHART[k % len(CHART)] if k >= 0 else INK)
                    lr.font.color.rgb = RGBColor.from_string(lc)
            tcPr = cell._tc.get_or_add_tcPr()
            for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcPr, qn(side))
                if side == "a:lnB" or (heat and fill != WHITE):
                    ln.set("w", str(int(EMU_PX * (2 if fill != WHITE and heat else 1))))
                    sf = etree.SubElement(ln, qn("a:solidFill")); cl = etree.SubElement(sf, qn("a:srgbClr"))
                    cl.set("val", WHITE if (heat and fill != WHITE) else LINE)
                else:
                    ln.set("w", "0"); etree.SubElement(ln, qn("a:noFill"))
            f = tcPr.find(qn("a:solidFill"))
            if f is not None:
                tcPr.remove(f); tcPr.append(f)
    note(sl.s, s)
    return sl.finish(s)


def L_steps(prs, s, deck):
    """Results: numbered circles on a line. Master: MasterProcess (blue titles, stems and dots on a baseline)."""
    sl, top, bot = content_slide(prs, s)
    st = s["steps"][:5]
    n = len(st); w = (1786 - (n - 1) * 66) / n if not is_results() else (1786 - (n - 1) * 32) / n
    if is_results():
        y0 = top + 60
        hline(sl.s, 66 + 36, y0 + 36, (n - 1) * (w + 32), color=LINE, weight=2)
        for i, k in enumerate(st):
            x = 66 + i * (w + 32)
            circ = sl.s.shapes.add_shape(MSO_SHAPE.OVAL, px(x), px(y0), px(72), px(72))
            circ.fill.solid(); circ.fill.fore_color.rgb = RGBColor.from_string(BLUE10); circ.line.fill.background(); _nostyle(circ)
            text(sl.s, x, y0, 72, 72, str(i + 1), size=32, weight=700, color=BLUE, align="c", anchor="m")
            text(sl.s, x, y0 + 112, w - 16, 100, k["title"], size=32, weight=700, color=INK, lh=1.15)
            ty = y0 + 112 + est_lines(k["title"], 32, w - 16, 700) * 32 * 1.2 + 16
            text(sl.s, x, ty, w - 16, bot - ty, k.get("text", ""), size=24, color=G700, lh=1.4)
            check_fit(f"Step {i + 1}", k.get("text", ""), 24, w - 16, bot - ty, 400, 1.4)
    else:
        base = bot - 60
        grow = s.get("growing", False)
        for i, k in enumerate(st):
            x = 66 + i * (w + 66)
            text(sl.s, x, top + 40, w, 110, k["title"], size=40, weight=600, color=BLUE, lh=1.1)
            ty = top + 40 + est_lines(k["title"], 40, w, 600) * 40 * 1.15 + 16
            text(sl.s, x, ty, w, base - ty - 80, k.get("text", ""), size=26, color=G700, lh=1.4)
            check_fit(f"Step {i + 1}", k.get("text", ""), 26, w, base - ty - 80, 400, 1.4)
            d = 28 + (i * 40 if grow else 0)
            vline(sl.s, x + 14, ty + 10, base - ty - 10, color=BLUE50, weight=2)
            dot = sl.s.shapes.add_shape(MSO_SHAPE.OVAL, px(x + 14 - d / 2), px(base - d / 2), px(d), px(d))
            dot.fill.solid(); dot.fill.fore_color.rgb = RGBColor.from_string(BLUE); dot.line.fill.background(); _nostyle(dot)
        hline(sl.s, 66, base, 1786, color=BLUE50, weight=2)
        to_back(sl.s, sl.s.shapes[-1], 3)
    note(sl.s, s)
    return sl.finish(s)


def L_contact(prs, s, deck):
    sl = Slide(prs, "Contact")
    sl.put(0, s.get("title", "Contact us")); sl.put(2, s.get("subtitle"))
    for i, p in enumerate(s.get("people", [])[:2]):
        if p.get("photo"):
            sl.pic([3, 5][i], p["photo"])
        else:
            warn(f"Contact {i + 1}: no photo given; add the person's real photo (photo: path).")
        sl.put([1, 6][i], p["name"]); sl.put([4, 7][i], [x for x in (p.get("role"), p.get("email"), p.get("phone")) if x])
    return sl.finish(s)


def L_back_cover(prs, s, deck):
    sl = Slide(prs, "Back cover lilac" if s.get("variant") == "lilac" else "Back cover", dark=True)
    sl.put(0, TAGLINE)
    for i, p in enumerate(s.get("people", [])[:2]):
        text(sl.s, 375 + i * 420, 700, 400, 90, [p["name"], p.get("role", "")], size=24, color=WHITE, lh=1.3)
    return sl.finish(s)


# ---------------------------------------------------------------- master-template components
def L_columns(prs, s, deck):
    n = len(s["columns"])
    if n <= 3:
        lay, heads, bodies = "Three columns", [1, 4, 6], [3, 5, 7]
    elif s.get("grid") == "2x2":
        lay, heads, bodies = "Two by two", [1, 5, 7, 9], [4, 6, 8, 13]
    else:
        lay, heads, bodies = "Four columns", [1, 4, 6, 8], [3, 5, 7, 9]
    sl = Slide(prs, lay)
    title_block(sl, s, 0, 2)
    for i, c in enumerate(s["columns"][:len(heads)]):
        sl.put(heads[i], c["headline"], align="l")
        sl.put(bodies[i], c.get("text"), check=(f"Column {i + 1}", 24, 400, 1.4), align="l")
    note(sl.s, s)
    return sl.finish(s)


def L_ideas(prs, s, deck):
    sl = Slide(prs, "Six ideas")
    sl.put(0, s["title"])
    heads, bodies = [2, 3, 4, 5, 6, 7], [1, 9, 14, 8, 13, 15]
    for i, c in enumerate(s["ideas"][:6]):
        sl.put(heads[i], c["headline"], align="l"); sl.put(bodies[i], c.get("text"), check=(f"Idea {i + 1}", 24, 400, 1.4), align="l")
    return sl.finish(s)


def L_photos(prs, s, deck):
    sl = Slide(prs, "Three photos")
    title_block(sl, s, 0, 2)
    for i, (pic, head, body) in enumerate(zip([3, 5, 6], [1, 7, 9], [4, 8, 13])):
        if i < len(s["items"]):
            it = s["items"][i]
            sl.pic(pic, it["photo"]); sl.put(head, it.get("headline"), align="l"); sl.put(body, it.get("text"), align="l")
    return sl.finish(s)


def L_text_images(prs, s, deck):
    """MasterTextImages: bold-led points left, three-photo collage right."""
    sl = Slide(prs, "Text + photo collage")
    title_block(sl, s, 0, 3, width=859)
    paras = []
    for p in s["points"][:4]:
        paras += [f"**{p['headline']}**", p.get("text", ""), ""]
    sl.put(1, paras[:-1], align="l", check=("Points", 24, 400, 1.4))
    for idx, ph in zip([4, 5, 6], s["photos"][:3]):
        sl.pic(idx, ph)
    return sl.finish(s)


def L_panel(prs, s, deck):
    sl = Slide(prs, "Two-liner + panel")
    sl.put(0, s["title"], check=("Two-liner title", 93, 700, 1.1))
    items = s.get("bullets")
    sl.put(1, items[:6] if items else s.get("body"), bullets=bool(items), check=("Panel text", 53, 600, 1.3), align="l", indent=48)
    return sl.finish(s)


def ring(slide, cx, cy, d, share, stroke=60, track=BLUE10, arc=BLUE):
    t = slide.shapes.add_shape(MSO_SHAPE.DONUT, px(cx - d / 2), px(cy - d / 2), px(d), px(d))
    t.adjustments[0] = stroke / d
    t.fill.solid(); t.fill.fore_color.rgb = RGBColor.from_string(track); t.line.fill.background(); _nostyle(t)
    if share > 0:
        a = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, px(cx - d / 2), px(cy - d / 2), px(d), px(d))
        start = 270.0; end = (270.0 + 360.0 * min(share, 0.9999)) % 360
        a.adjustments[0] = start * 0.6; a.adjustments[1] = end * 0.6; a.adjustments[2] = stroke / d  # angles in 60000ths of a degree
        a.fill.solid(); a.fill.fore_color.rgb = RGBColor.from_string(arc); a.line.fill.background(); _nostyle(a)


def L_donut_split(prs, s, deck):
    """MasterDonutSplit: one share in a ring on white; Blue 10 half with bullets (template pp. 98–101)."""
    sl = Slide(prs, "Two-liner + panel")
    share = float(s["value"])
    share = share / 100 if share > 1 else share
    ring(sl.s, 496, 520, 560, share)
    text(sl.s, 216, 430, 560, 110, s.get("label_value", f"{round(share * 100)}%"), size=96, weight=700, color=INK, align="c", anchor="m", lh=1.0)
    if s.get("caption"):
        text(sl.s, 286, 545, 420, 90, s["caption"], size=28, color=G700, align="c", lh=1.2)
    items = s.get("bullets")
    if items:
        sl.put(1, items[:5], bullets=True, align="l", indent=48, check=("Bullets", 53, 600, 1.3))
    else:
        sl.put(1, s.get("title"), align="l")
    return sl.finish(s)


def L_timeline(prs, s, deck):
    """MasterTimeline: events alternate above and below one horizontal line."""
    sl, top, bot = content_slide(prs, s)
    ev = s["events"][:7]
    n = len(ev); mid = (top + bot) / 2
    hline(sl.s, 66, mid, 1786, color=INK, weight=2)
    span = 1786 / n
    for i, e in enumerate(ev):
        x = 66 + i * span + 10
        dot = sl.s.shapes.add_shape(MSO_SHAPE.OVAL, px(x - 10), px(mid - 10), px(20), px(20))
        dot.fill.solid(); dot.fill.fore_color.rgb = RGBColor.from_string(INK); dot.line.fill.background(); _nostyle(dot)
        w = min(span * 1.9, 520) - 30
        lines = [x for x in (e.get("date"), e.get("text")) if x]
        h_title = est_lines(e["title"], 37, w, 600) * 37 * 1.2
        h_det = sum(est_lines(l, 26, w) for l in lines) * 26 * 1.3
        block = h_title + 8 + h_det
        if i % 2 == 0:   # above
            y = mid - 60 - block
            vline(sl.s, x, y + 10, mid - y - 10, color=INK, weight=1)
        else:
            y = mid + 60
            vline(sl.s, x, mid, y - mid + block * 0.2, color=INK, weight=1)
        text(sl.s, x + 20, y, w, h_title + 4, e["title"], size=37, weight=600, color=INK, lh=1.1)
        if lines:
            text(sl.s, x + 20, y + h_title + 8, w, h_det + 10, lines, size=26, color=G500, lh=1.25)
        if (i % 2 == 0 and y < top) or (i % 2 == 1 and y + block > bot):
            warn(f"Timeline event {i + 1}: text too long for the space; shorten it.")
    note(sl.s, s)
    return sl.finish(s)


def L_from_to(prs, s, deck):
    """MasterFromTo: paired pills, old way on Blue 10, new way on Appinio Blue."""
    sl, top, bot = content_slide(prs, s)
    rows = s["rows"][:6]
    colw = (1786 - 66) / 2
    text(sl.s, 66, top + 10, colw, 50, s.get("from_label", "From…"), size=36, weight=600, color=INK)
    text(sl.s, 66 + colw + 66, top + 10, colw, 50, s.get("to_label", "To…"), size=36, weight=600, color=INK)
    y0 = top + 84; rh = (bot - y0 - (len(rows) - 1) * 20) / len(rows)
    for i, (a, b) in enumerate(rows):
        y = y0 + i * (rh + 20)
        for j, (t, fill, col) in enumerate(((a, BLUE10, INK), (b, BLUE, WHITE))):
            x = 66 + j * (colw + 66)
            rect(sl.s, x, y, colw, rh, fill, radius=12)
            text(sl.s, x + 24, y, colw - 48, rh, t, size=29, weight=600, color=col, align="c", anchor="m", lh=1.15)
            check_fit(f"From/to row {i + 1}", t, 29, colw - 48, rh, 600, 1.15)
    note(sl.s, s)
    return sl.finish(s)


def L_comparison(prs, s, deck):
    """MasterComparison: argument left; ink panel right with the usual way vs Appinio."""
    sl = Slide(prs, "Blank")
    rect(sl.s, 993, 0, 927, 1080, INK)
    text(sl.s, 66, 67, 734, 180, s["title"], size=45, weight=800, color=INK, lh=1.1, spc=-0.5)
    text(sl.s, 66, 300, 734, 600, s.get("body", ""), size=26, color=G700, lh=1.45, para_gap=20)
    check_fit("Comparison body", s.get("body", ""), 26, 734, 600, 400, 1.45)
    for j, (key, fill, lab_col, default) in enumerate((("old", G10, G500, "The usual way"), ("new", "C4D7F1", BLUE, "With Appinio"))):
        blk = s[key]
        y = 160 + j * 400; h = 360
        rect(sl.s, 1109, y, 696, h, fill, radius=16)
        text(sl.s, 1149, y + 36, 616, 32, blk.get("label", default), size=24, weight=700, color=lab_col, spc=2, caps=True)
        text(sl.s, 1149, y + 90, 616, h - 120, blk["items"][:4], size=30, weight=600, color=INK, lh=1.35)
        check_fit(f"Comparison card {j + 1}", "\n".join(blk["items"][:4]), 30, 616, h - 120, 600, 1.35)
    return sl.finish(s)


def L_stat_mosaic(prs, s, deck):
    """MasterStatMosaic: statement on Blue 10 with three white proof cards."""
    sl = Slide(prs, "Blank")
    rect(sl.s, 0, 0, 1920, 1080, BLUE10)
    text(sl.s, 66, 67, 1786, 400, s["statement"], size=99, weight=800, color=INK, lh=1.05, spc=-1)
    check_fit("Statement (max 3 lines)", s["statement"], 99, 1786, 400, 800, 1.05)
    cards = s["cards"][:3]
    for i, c in enumerate(cards):
        x = 66 + i * (551 + 66); y = 620; h = 320
        rect(sl.s, x, y, 551, h, WHITE, radius=16)
        if c.get("value"):
            text(sl.s, x + 40, y + 40, 471, 80, c["value"], size=61, weight=800, color=INK)
        elif c.get("icon"):
            picture(sl.s, asset("icons", c["icon"], "png"), x + 40, y + 40, 96, 96, fit="contain")
        text(sl.s, x + 40, y + 170, 471, 120, c["text"], size=32, weight=700, color=INK, lh=1.2)
    logo(sl.s)
    return sl.finish(s)


def L_highlights(prs, s, deck):
    """MasterHighlights: six headline figures in white on the light gradient."""
    sl = Slide(prs, "Statement gradient", dark=True)
    if s.get("gradient"):
        to_back(sl.s, picture(sl.s, gradient_path(s["gradient"]), 0, 0, 1920, 1080))
    text(sl.s, 66, 67, 1786, 60, s.get("title", "Highlights"), size=45, weight=800, color=WHITE)
    for i, f in enumerate(s["figures"][:6]):
        col, row = i % 2, i // 2
        x = 66 + col * 927; y = 230 + row * 230
        text(sl.s, x, y, 860, 110, f["value"], size=93, weight=700, color=WHITE, lh=1.0)
        text(sl.s, x, y + 115, 860, 40, f["label"], size=24, weight=600, color=WHITE)
    return sl.finish(s)


def L_gradient_cards(prs, s, deck):
    """MasterGradientCards: 3–5 tall gradient cards with white titles."""
    sl, top, bot = content_slide(prs, s)
    cards = s["cards"][:5]
    n = len(cards); w = (1786 - (n - 1) * 32) / n
    grads = s.get("gradients") or ["lilac", "cover", "contact"]
    for i, t in enumerate(cards):
        x = 66 + i * (w + 32)
        picture(sl.s, gradient_path(grads[i % 3]), x, top, w, bot - top, radius=16)
        text(sl.s, x + 30, top, w - 60, bot - top, t, size=43, weight=800, color=WHITE, align="c", anchor="m", lh=1.1)
    note(sl.s, s)
    return sl.finish(s)


def L_logo_wall(prs, s, deck):
    """MasterLogoWall: client logos in a grid (official files only)."""
    sl = Slide(prs, "Blank")
    text(sl.s, 66, 67, 1786, 90, s.get("title", "Join the 3,000+ clients working with Appinio"), size=67, weight=800, color=INK, spc=-0.5)
    logos = s["logos"][:15]
    cols = 5 if len(logos) > 12 else 4
    rows = (len(logos) + cols - 1) // cols
    cw = 1786 / cols; ch = (900 - 230) / max(rows, 1)
    for i, f in enumerate(logos):
        x = 66 + (i % cols) * cw; y = 230 + (i // cols) * ch
        picture(sl.s, photo_path(f), x + 40, y + 30, cw - 80, ch - 60, fit="contain")
    return sl.finish(s)


def L_team(prs, s, deck):
    sl = Slide(prs, "Team")
    sl.put(0, s.get("title", "Meet the team"))
    for i, p in enumerate(s["people"][:3]):
        sl.pic([2, 3, 4][i], p.get("photo"))
        if not p.get("photo"):
            warn(f"Team member {i + 1}: no photo; add their real photo.")
        sl.put([1, 6, 8][i], p["name"]); sl.put([13, 14, 15][i], p.get("role"))
        sl.put([5, 7, 9][i], [x for x in (p.get("email"), p.get("phone"), p.get("text")) if x])
    return sl.finish(s)


def L_thank_you(prs, s, deck):
    """MasterThankYou: "Thank you!" or "Q&A" with one wide photo."""
    sl = Slide(prs, "Title + wide photo")
    sl.put(0, s.get("title", "Thank you!")); sl.put(2, s.get("subtitle"))
    if s.get("photo"):
        sl.pic(3, s["photo"])
    return sl.finish(s)


REGIONS = {"north america": (560, 470), "europe": (1000, 420), "asia": (1340, 470),
           "south america": (700, 700), "oceania": (1420, 760)}
FLAGS = {"usa": "usa", "us": "usa", "united states": "usa", "uk": "uk", "gb": "uk", "united kingdom": "uk",
         "spain": "spain", "es": "spain", "france": "france", "fr": "france", "germany": "germany", "de": "germany"}
DISCLAIMER = ("The content described herein is for informational purposes only to outline Appinio’s general product "
              "direction. The development, release, and timing of any features or functionality described herein remain "
              "at the sole discretion of Appinio and may be subject to change. Nothing contained in this roadmap is a "
              "commitment to deliver any material, code, or functionality, at any particular time or within any given "
              "timeframe and should not be relied on in making purchasing decisions.")


def flag_path(country):
    key = FLAGS.get(str(country).strip().lower())
    if not key:
        warn(f"No approved flag for '{country}' (available: USA, UK, Spain, France, Germany); shown without a flag.")
        return None
    return asset("flags", key, "png")


def L_map(prs, s, deck):
    """MasterMap: Blue 10 world map with blue markers and pill labels (template pp. 105–106)."""
    sl, top, bot = content_slide(prs, s)
    picture(sl.s, asset("maps", "world", "png"), 267, 220, 1320, 700, fit="contain")
    for i, m in enumerate(s["markers"][:6]):
        if "region" in m and m["region"].lower() in REGIONS:
            x, y = REGIONS[m["region"].lower()]
        elif "x" in m and "y" in m:
            x, y = m["x"], m["y"]
        else:
            warn(f"Map marker {i + 1}: give a region ({', '.join(REGIONS)}) or x/y on the 1920×1080 canvas."); continue
        dot = sl.s.shapes.add_shape(MSO_SHAPE.OVAL, px(x), px(y), px(24), px(24))
        dot.fill.solid(); dot.fill.fore_color.rgb = RGBColor.from_string(BLUE); dot.line.fill.background(); _nostyle(dot)
        label = m.get("label") or m.get("region", "")
        pw = min(420, 40 + len(label) * 15)
        if m.get("value") is not None:   # per-market KPI: small ring instead of a pill
            share = float(m["value"]); share = share / 100 if share > 1 else share
            ring(sl.s, x + 36 + 60, y + 12, 120, share, stroke=16)
            text(sl.s, x + 36, y - 48, 120, 120, f"{round(share * 100)}%", size=28, weight=700, color=INK, align="c", anchor="m")
            text(sl.s, x + 36, y + 80, 300, 40, label, size=26, weight=700, color=INK)
        else:
            rect(sl.s, x + 36, y - 18, pw, 48, BLUE, radius=24)
            text(sl.s, x + 56, y - 18, pw - 40, 48, label, size=26, weight=700, color=WHITE, anchor="m")
            if m.get("text"):
                text(sl.s, x + 36, y + 38, 300, 70, m["text"], size=24, color=G700, lh=1.3)
    note(sl.s, s)
    return sl.finish(s)


def L_markets(prs, s, deck):
    """MasterFlags: one column per market with its round flag, name and an optional figure."""
    sl, top, bot = content_slide(prs, s)
    mk = s["markets"][:5]
    n = len(mk); w = (1786 - (n - 1) * 32) / n
    fill, (ax, ay, aw, ah) = boxes(sl.s, top, bot)
    w = (aw - (n - 1) * 32) / n
    for i, m in enumerate(mk):
        x = ax + i * (w + 32)
        rect(sl.s, x, ay, w, ah, fill, radius=16)
        f = flag_path(m["country"])
        if f:
            picture(sl.s, f, x + (w - 160) / 2, ay + 48, 160, 160, fit="contain")
        text(sl.s, x + 24, ay + 232, w - 48, 50, m.get("label", m["country"]), size=32, weight=700, color=INK, align="c")
        if m.get("value"):
            text(sl.s, x + 16, ay + 300, w - 32, 80, m["value"], size=52, weight=800, color=INK, align="c", lh=1.0)
            check_fit(f"Market {i + 1} value", m["value"], 52, w - 32, 80, 800, 1.0)
        if m.get("text"):
            text(sl.s, x + 32, ay + 430, w - 64, ah - 470, m["text"], size=24, color=G700, align="c", lh=1.4)
    note(sl.s, s)
    return sl.finish(s)


def L_case_study(prs, s, deck):
    """MasterCaseStudy: facts on white left, challenge / approach / impact / quote on Blue 10 right."""
    sl = Slide(prs, "Blank")
    rect(sl.s, 619, 0, 1301, 1080, BLUE10)
    if s.get("client_logo"):
        picture(sl.s, photo_path(s["client_logo"]), 67, 68, 300, 110, fit="contain")
    else:
        warn("Case study: add the official client logo file (client_logo).")
    text(sl.s, 67, 222, 480, 260, s["headline"], size=40, weight=700, color=INK, lh=1.2)
    check_fit("Case study headline", s["headline"], 40, 480, 260, 700, 1.2)
    y = 530
    for k, v in list(s.get("facts", {}).items())[:4]:
        text(sl.s, 67, y, 480, 32, k, size=24, weight=600, color=G300)
        text(sl.s, 67, y + 32, 480, 34, v, size=24, weight=700, color=INK); y += 90
    cells = [(685, 68), (1302, 68), (685, 532), (1302, 532)]
    for (x, cy), (head, body) in zip(cells[:2], (("Challenge", s["challenge"]), (s.get("approach_label", "The Appinio approach"), s["approach"]))):
        text(sl.s, x, cy, 551, 50, head, size=40, weight=700, color=INK)
        text(sl.s, x, cy + 70, 551, 330, body, size=30, color=G700, lh=1.45)
        check_fit(head, body, 30, 551, 330, 400, 1.45)
    x, cy = cells[2]
    text(sl.s, x, cy, 551, 50, "Impact", size=40, weight=700, color=INK)
    text(sl.s, x, cy + 66, 551, 90, s["impact_value"], size=80, weight=800, color=BLUE, lh=1.0)
    text(sl.s, x, cy + 170, 551, 230, s.get("impact_text", ""), size=30, color=G700, lh=1.45)
    if s.get("quote"):
        x, cy = cells[3]
        text(sl.s, x, cy, 551, 400, ["“" + s["quote"].strip("“”\"") + "”", f"[[{s.get('quote_by', '')}]]"],
             size=34, weight=700, color=INK, lh=1.35, hl=G500)
        check_fit("Case study quote", s["quote"], 34, 551, 330, 700, 1.35)
    logo(sl.s)
    return sl.finish(s)


def L_disclaimer(prs, s, deck):
    """MasterDisclaimer: Legal's fixed text, bottom-aligned (template p. 140). The text is not editable."""
    sl = Slide(prs, "Blank")
    text(sl.s, 67, 520, 1786, 70, "Legal disclaimer", size=48, weight=600, color=INK)
    text(sl.s, 67, 600, 1786, 320, DISCLAIMER, size=29, weight=300, color=INK, lh=1.4)
    return sl.finish(s)


# ---------------------------------------------------------------- funnel & growth (template slides 112–114)
FUNNEL_FILLS = ["2648A3", "3463E0", "5987FF", "96B3FF", "BFCDFF", "E5EBFF"]


def _poly(slide, pts, fill):
    fb = slide.shapes.build_freeform(px(pts[0][0]), px(pts[0][1]), scale=1.0)
    fb.add_line_segments([(px(x), px(y)) for x, y in pts[1:]], close=True)
    shp = fb.convert_to_shape()
    shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor.from_string(fill)
    shp.line.fill.background(); _nostyle(shp)
    return shp


def _pct(v):
    if isinstance(v, str):
        return v
    return f"{v * 100:.0f}%" if v <= 1 else f"{v:g}%"


def L_funnel(prs, s, deck):
    """Funnel summary (template slides 112/113): 3–6 stages, big share + label (+ conversion rate) per column,
    blue blocks stepping down from dark to light. Both styles (brand funnels in results decks)."""
    sl, top, bot = content_slide(prs, s)
    st = s["stages"][:6]
    if len(s["stages"]) > 6:
        warn("Funnel: at most 6 stages.")
    n = len(st); W = 1786 / n
    base = min(bot, 889)
    nums = [parse_num(x["value"]) for x in st]
    mx = max(v for v in nums if v is not None) or 1
    stop = top + (190 if any(x.get("conversion") for x in st) else 160)
    H = base - stop
    hs = [max(24, (v or 0) / mx * H) for v in nums]
    fills = FUNNEL_FILLS[:n] if n < 6 else FUNNEL_FILLS
    for i, x in enumerate(st):
        x0 = 66 + i * W
        hl = hs[i]; hr = hs[i + 1] if i + 1 < n else hs[i] * 0.55
        _poly(sl.s, [(x0, base), (x0, base - hl), (x0 + W, base - hr), (x0 + W, base)], fills[i])
        vline(sl.s, x0, top, base - top, color=LINE, weight=2)
        v = _pct(x["value"])
        text(sl.s, x0 + 32, top, W - 48, 80, f"[[{v}]]" if s.get("highlight") == i else v,
             size=67, weight=800, color=INK, lh=1.0)
        text(sl.s, x0 + 32, top + 84, W - 48, 70, x["label"], size=24, color=INK, lh=1.2)
        check_fit(f"Stage {i + 1} label", x["label"], 24, W - 48, 32 if x.get("conversion") else 64, 400, 1.2)
        if x.get("conversion"):
            text(sl.s, x0 + 32, top + 118, W - 48, 40, x["conversion"], size=24, weight=700, color=INK, lh=1.2)
            check_fit(f"Stage {i + 1} conversion", x["conversion"], 24, W - 48, 40, 700, 1.2)
    note(sl.s, s)
    return sl.finish(s)


def L_growth(prs, s, deck):
    """Growth plan (template slide 114): 3–6 rising bars cut from one gradient, headline + text above each."""
    sl, top, bot = content_slide(prs, s)
    st = s["steps"][:6]
    n = len(st); gap = 36
    w = (1786 - (n - 1) * gap) / n
    base = 927 if not is_results() else bot
    tmax = top + 150
    hmin, hmax = 118, base - tmax
    grad = gradient_path(s.get("gradient", "cover"))
    ytop = base - hmax
    for i, k in enumerate(st):
        x = 66 + i * (w + gap)
        h = hmin + (hmax - hmin) * (i / (n - 1) if n > 1 else 1)
        y = base - h
        pic = sl.s.shapes.add_picture(grad, px(x), px(y), px(w), px(h))
        pic.crop_left = (x - 66) / 1786; pic.crop_right = 1 - (x + w - 66) / 1786
        pic.crop_top = (y - ytop) / hmax; pic.crop_bottom = 0
        body = k.get("text", "")
        tb_h = 32 + est_lines(body, 24, w, 400) * 24 * 1.25 if body else 32
        text(sl.s, x, y - tb_h - 16, w, tb_h, k["title"],
             size=24, weight=700, color=INK, lh=1.25)
        if body:
            text(sl.s, x, y - tb_h - 16 + 32, w, tb_h - 32, body, size=24, color=G700, lh=1.25)
        if est_lines(k["title"], 24, w, 700) > 1:
            warn(f"Growth step {i + 1}: title wraps; keep it to 2–3 words.")
        if tb_h > 150:
            warn(f"Growth step {i + 1}: text too long for the column (max ~12 words).")
    note(sl.s, s)
    return sl.finish(s)


LAYOUTS = {k[2:]: v for k, v in globals().items() if k.startswith("L_")}
RESULTS_ONLY = {"profile"}
MASTER_ONLY = {"growth", "timeline", "from_to", "comparison", "stat_mosaic", "highlights", "gradient_cards",
               "logo_wall", "donut_split", "text_images", "thank_you", "ideas", "columns", "photos", "panel",
               "map", "case_study", "disclaimer"}


# ---------------------------------------------------------------- main
def build(outline, out):
    style = outline.get("style")
    if style not in ("results", "master"):
        raise SystemExit('Set "style" at the top of the outline: "results" (research results) or "master" (sales, company, pitch and everything else).')
    STYLE["name"] = style
    prs = Presentation(TEMPLATE)
    deck = dict(outline)
    for i, s in enumerate(outline["slides"], 1):
        t = s.get("type")
        if t not in LAYOUTS:
            raise SystemExit(f"Slide {i}: unknown type '{t}'. Known: {', '.join(sorted(LAYOUTS))}")
        before = len(WARN)
        if style == "master" and t in RESULTS_ONLY:
            warn(f"'{t}' is a results-deck layout; consider a simpler chart in a master-style deck.")
        if style == "results" and t in MASTER_ONLY:
            warn(f"'{t}' is a master-template layout; a results deck should only borrow it where nothing else fits.")
        try:
            LAYOUTS[t](prs, s, deck)
        except KeyError as e:
            raise SystemExit(f"Slide {i} ({t}): missing field {e}")
        for j in range(before, len(WARN)):
            WARN[j] = f"Slide {i} ({t}) – {WARN[j]}"
    if outline["slides"][-1].get("type") != "back_cover":
        warn("Deck does not end with a back_cover slide.")
    prs.core_properties.title = outline.get("title", "")
    prs.core_properties.author = "Appinio"
    prs.save(out)
    print(f"Saved {out} ({len(outline['slides'])} slides, {style} style)")
    for w in WARN:
        print("WARNING:", w)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("outline", nargs="?")
    ap.add_argument("-o", "--out", default="deck.pptx")
    ap.add_argument("--list-assets", action="store_true")
    a = ap.parse_args()
    if a.list_assets:
        list_assets(); sys.exit()
    if not a.outline:
        ap.error("outline.json required")
    build(json.load(open(a.outline, encoding="utf-8")), a.out)
