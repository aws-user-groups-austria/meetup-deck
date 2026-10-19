"""Slides made in the deck's own design, not pictures of the stream.

They start from the layout of the template's "Shoutout" slide (so the master
puts the user group logo top right), copy its triangle, and set text in the
deck's fonts and colours: PT Sans, titles bold #9E5EFB, text white, as the
template's own slides do. Everything stays editable in Google Slides.
"""

from __future__ import annotations

import copy
import io
import tempfile
from datetime import date
from pathlib import Path

import segno
from PIL import Image, ImageDraw
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.presentation import Presentation as PresentationT
from pptx.slide import Slide
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

PURPLE = RGBColor(0x9E, 0x5E, 0xFB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0xD8, 0xD8, 0xD8)
FONT = "Amazon Ember"

# The theme's type scale (SPEC.md, "Theme"), in points: 1 px on the 1920 frame = 0.375 pt.
DISPLAY, TITLE, HEADING, LEAD, BODY, SMALL, LABEL = 51, 36, 25.5, 21, 15.75, 12.75, 10.5


class Style:
    """Where the new slides take their look from in the template."""

    def __init__(self, prs: PresentationT, source: Slide | None):
        self.prs = prs
        self.layout = source.slide_layout if source is not None else _blank(prs)
        self.triangle = None
        if source is not None:
            for sh in source.shapes:
                if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and "rtTriangle" in sh._element.xml:
                    self.triangle = sh._element
                    break
        self.tmp = Path(tempfile.mkdtemp(prefix="meetup-deck-"))

    @property
    def w(self) -> int:
        return self.prs.slide_width

    @property
    def h(self) -> int:
        return self.prs.slide_height


def _blank(prs: PresentationT):
    layouts = list(prs.slide_layouts)
    return next((l for l in layouts if l.name.upper() == "BLANK"), layouts[-1])


def new_slide(style: Style, notes: str, triangle: bool = True) -> Slide:
    slide = style.prs.slides.add_slide(style.layout)
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)
    if triangle and style.triangle is not None:
        slide.shapes._spTree.insert(2, copy.deepcopy(style.triangle))
    slide.notes_slide.notes_text_frame.text = notes
    return slide


def text(
    slide: Slide,
    lines: list[str] | list[tuple[str, dict]],
    x: float,
    y: float,
    w: float,
    h: float,
    size: float,
    color: RGBColor = WHITE,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.CENTER,
    anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
    spacing: float = 1.0,
):
    """A text box in inches. A line is a string, or (string, overrides of size/color/bold)."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    for i, line in enumerate(lines):
        s, o = (line, {}) if isinstance(line, str) else line
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = s
        r.font.name = FONT
        r.font.size = Pt(o.get("size", size))
        r.font.bold = o.get("bold", bold)
        r.font.color.rgb = o.get("color", color)
    return box


def title(slide: Slide, s: str, y: float = 0.82, x: float = 0.31, w: float = 9.38) -> None:
    """The big purple title, where the template's title slides have it.

    52pt as in the template; a long one is set smaller and lower, so it stays
    clear of the user group logo in the top right corner.
    """
    if len(s) > 14:
        text(slide, [s], x, y + 0.28, w, 1.0, TITLE, PURPLE, bold=True, anchor=MSO_ANCHOR.BOTTOM)
    else:
        text(slide, [s], x, y, w, 1.0, DISPLAY, PURPLE, bold=True, anchor=MSO_ANCHOR.BOTTOM)


def subtitle(slide: Slide, s: str, y: float = 1.9, x: float = 0.31, w: float = 9.38, size: float = LEAD) -> None:
    text(slide, [s], x, y, w, 0.6, size, WHITE)


def round_photo(style: Style, src: Path, name: str) -> Path:
    """A circle of the photo with a thin purple ring, as on the Canva cards."""
    with Image.open(src) as im:
        im = im.convert("RGB")
        side = min(im.size)
        im = im.crop(((im.width - side) // 2, (im.height - side) // 2, (im.width + side) // 2, (im.height + side) // 2))
        im = im.resize((480, 480), Image.LANCZOS)
    out = Image.new("RGBA", (500, 500), (0, 0, 0, 0))
    mask = Image.new("L", (480, 480), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, 479, 479), fill=255)
    ring = ImageDraw.Draw(out)
    ring.ellipse((0, 0, 499, 499), fill=(0x9E, 0x5E, 0xFB, 255))
    out.paste(im, (10, 10), mask)
    path = style.tmp / f"{name}.png"
    out.save(path)
    return path


def qr_png(style: Style, url: str, name: str) -> Path:
    path = style.tmp / f"{name}.png"
    buf = io.BytesIO()
    segno.make(url, error="m").save(buf, kind="png", scale=12, border=2, dark="#151E2C", light="#FFFFFF")
    path.write_bytes(buf.getvalue())
    return path


def nice_date(iso: str) -> str:
    d = date.fromisoformat(iso)
    return f"{d.strftime('%A')}, {d.day} {d.strftime('%B %Y')}"


def short_date(iso: str) -> str:
    d = date.fromisoformat(iso)
    return f"{d.day} {d.strftime('%B')}"


# --- the slides -------------------------------------------------------------


def agenda(style: Style, rows: list[dict], heading: str, notes: str) -> Slide:
    """rows: {time, title, speakers?, talk: bool}. Like 'Introducing Our User Group Members': title top left."""
    slide = new_slide(style, notes, triangle=False)
    text(slide, ["Agenda"], 0.58, 0.45, 6.0, 0.7, TITLE, PURPLE, bold=True, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.BOTTOM)
    text(slide, [heading], 0.58, 1.12, 6.4, 0.35, SMALL, GREY, align=PP_ALIGN.LEFT)
    # Every row the same height, room for the title and a speakers line, so the times on the left are
    # evenly spaced whatever a row holds; the rows fill the space under the heading. One text box per
    # row for the time and one for the title (+ speakers), both anchored at the row's top.
    top, bottom = 1.6, 5.5
    pitch = (bottom - top) / max(1, len(rows))
    size = BODY if pitch >= 0.46 else SMALL
    for i, r in enumerate(rows):
        y = top + i * pitch
        text(slide, [r["time"]], 0.58, y, 1.75, pitch, size, PURPLE, bold=True, align=PP_ALIGN.LEFT)
        lines: list = [(r["title"], {"bold": bool(r.get("talk"))})]
        if r.get("speakers"):
            lines.append((r["speakers"], {"size": SMALL, "color": GREY}))
        text(slide, lines, 2.45, y, 7.0, pitch, size, WHITE, align=PP_ALIGN.LEFT)
    return slide


def break_slide(style: Style, back_at: str, nxt: dict | None, notes: str) -> Slide:
    slide = new_slide(style, notes, triangle=False)
    title(slide, "Break")
    subtitle(slide, f"Back at {back_at}")
    if nxt:
        lines: list = [("UP NEXT", {"size": LABEL, "color": PURPLE, "bold": True}), (nxt["title"], {"size": HEADING, "bold": True})]
        if nxt.get("speakers"):
            lines.append((nxt["speakers"], {"size": BODY, "color": GREY}))
        text(slide, lines, 0.31, 3.05, 9.38, 1.6, HEADING, WHITE, spacing=1.1)
    text(slide, ["Time for drinks and networking"], 0.31, 4.75, 9.38, 0.4, BODY, GREY)
    return slide


def closing(style: Style, speakers: list[str], host: str | None, feedback: str, until: str, notes: str) -> Slide:
    slide = new_slide(style, notes, triangle=False)
    title(slide, "Thank you")
    thanks = f"to {', '.join(speakers[:-1])} and {speakers[-1]}" if len(speakers) > 1 else (f"to {speakers[0]}" if speakers else "")
    subtitle(slide, thanks)
    if host:
        text(slide, [f"and to {host} for hosting us"], 0.31, 2.45, 9.38, 0.4, BODY, GREY)
    qr = qr_png(style, feedback, "feedback-qr")
    slide.shapes.add_picture(str(qr), Inches(6.125), Inches(3.0), Inches(1.75), Inches(1.75))
    text(
        slide,
        [("Please leave feedback", {"size": HEADING, "bold": True}), ("Scan to rate tonight on Meetup", {"size": SMALL, "color": GREY})],
        2.125, 3.25, 3.75, 1.2, HEADING, WHITE, align=PP_ALIGN.RIGHT, spacing=1.1,
    )
    text(slide, [f"Networking until {until}"], 0.31, 4.9, 9.38, 0.5, HEADING, PURPLE, bold=True, align=PP_ALIGN.CENTER)
    return slide


def link(box, url: str) -> None:
    """The whole text box links to `url`, clickable in PowerPoint and the PDF. On the shape, not the
    runs, so the text keeps its colour instead of the template's orange, underlined link style."""
    box.click_action.hyperlink.address = url


def upcoming(
    style: Style,
    day: dict,
    then: dict | None,
    photos: list[tuple[str, Path | None]],
    notes: str,
    screenshot: Path | None = None,
) -> Slide:
    """The next AWS Community Day: name, date, place and who speaks there on the left; its website
    (screenshot), the link and a QR code on the right, so it can be found again. Picture, link and
    QR are all clickable."""
    slide = new_slide(style, notes, triangle=False)
    title(slide, "Next AWS Community Day")
    url = day["url"]
    host = url.replace("https://", "").replace("http://", "").rstrip("/")
    when = nice_date(day["date"]) if not day.get("endDate") else f"{short_date(day['date'])} to {nice_date(day['endDate'])}"
    left, col = 0.31, 4.9
    # Below the title (which a long name sets lower, down to about 2.1 in).
    text(slide, [(day["name"], {"size": HEADING, "bold": True})], left, 2.25, col, 0.45, HEADING, WHITE, align=PP_ALIGN.LEFT)
    text(slide, [f"{when}  |  {day['location']}"], left, 2.72, col, 0.35, BODY, GREY, align=PP_ALIGN.LEFT)
    if photos:
        text(slide, ["SPEAKING THERE"], left, 3.22, col, 0.28, LABEL, PURPLE, bold=True, align=PP_ALIGN.LEFT)
        size, gap = 0.85, 0.45
        for i, (name, img) in enumerate(photos[:4]):
            x = left + i * (size + gap)
            if img is not None:
                slide.shapes.add_picture(str(img), Inches(x), Inches(3.52), Inches(size), Inches(size))
            text(slide, [name], x - 0.25, 4.42, size + 0.5, 0.3, LABEL, WHITE)
    if then:
        # its bottom on the same line as the QR code's and the url text's (bottom below)
        text(slide, [f"Then {then['name']}, {short_date(then['date'])}"], left, 2.3 + 3.85 * 9 / 16 + 0.62 - 0.35, col, 0.35, SMALL, GREY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.BOTTOM)
    # Right: the website, the link, the QR code.
    # The screenshot's top on the community day's name; the QR on a white sticker over its bottom
    # right corner, its bottom on the bottom of the url text below (and of the left column).
    sx, sy, sw = 5.55, 2.3, 3.85
    sh = sw * 9 / 16
    if screenshot is not None:
        pic = slide.shapes.add_picture(str(screenshot), Inches(sx), Inches(sy), Inches(sw), Inches(sh))
        pic.line.color.rgb = PURPLE
        pic.line.width = Pt(1)
        pic.click_action.hyperlink.address = url
    bottom = sy + sh + 0.62
    qr_size = 1.05
    # minus its white border (3 pt, drawn outside the picture), so what shows ends on the text's last line
    qx, qy = sx + sw - qr_size * 0.72, bottom - qr_size - 0.05
    qr = slide.shapes.add_picture(str(qr_png(style, url, f"upcoming-qr-{host}")), Inches(qx), Inches(qy), Inches(qr_size), Inches(qr_size))
    qr.line.color.rgb = WHITE
    qr.line.width = Pt(3)
    qr.click_action.hyperlink.address = url
    url_box = text(slide, [(host, {"size": BODY, "bold": True, "color": PURPLE}), ("Scan or click to open the website", {"size": LABEL, "color": GREY})], sx, bottom - 0.55, qx - sx - 0.15, 0.55, BODY, WHITE, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.BOTTOM)
    link(url_box, url)
    return slide


def use_font(prs: PresentationT, font: str = FONT) -> int:
    """Set every text in the deck, the template's own slides included, in one font.

    Changes each typeface named in the slides, layouts and masters, and the
    theme's heading and body fonts. Returns how many font references changed.
    """
    import re as _re

    count = 0
    pattern = _re.compile(rb'(<a:(?:latin|ea|cs|sym)\b[^>]*\btypeface=")([^"]*)(")')
    for part in prs.part.package.iter_parts():
        if not str(part.partname).endswith(".xml"):
            continue
        if hasattr(part, "_element"):
            for el in part._element.iter():
                tag = el.tag.rsplit("}", 1)[-1]
                if tag in ("latin", "ea", "cs", "sym") and el.get("typeface") not in (None, "", font) and not el.get("typeface", "").startswith("+"):
                    el.set("typeface", font)
                    count += 1
        elif "/theme/" in str(part.partname):
            blob, n = pattern.subn(lambda m: m.group(1) + font.encode() + m.group(3) if m.group(2) and not m.group(2).startswith(b"+") else m.group(0), part.blob)
            if n:
                part._blob = blob
                count += n
    return count
