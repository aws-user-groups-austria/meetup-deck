"""The group's slide deck, filled in for one meetup from its event.json.

The template is the deck the organizers already use (Google Slides, exported
as .pptx). Its slides, design and hidden backup slides stay as they are; only
the parts that change every meetup are touched, found by what they say, not
by their position, so the team can keep editing the template:

- the welcome slide's "Meetup #52" becomes the event's number;
- the header slide ("#X Meetup", one full-slide picture) gets the Meetup
  event's header image;
- every visible "Shoutout" slide gets the sponsors' logos, or their names
  when there is no logo yet; then the template's shoutouts (at the start and
  at the end) become a short "Thank you" to the host, and the shoutout itself,
  the host's own slot, is copied to right after the first talk's slide: we
  introduce the first talk, the host introduces themselves, the talk begins;
- the visible "See you soon" slide points at the next meetup;
- new slides from the stream's Remotion stills: the agenda after the header
  slide, then one per talk (the speaker handover) and break, and the closing,
  after "Introducing Our User Group Members".

The template's "#Todo" boxes sit beside the slide, outside the visible area.
The ones this fills in are removed; anything still to do gets a new one.
"""

from __future__ import annotations

import copy
import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.presentation import Presentation as PresentationT
from pptx.slide import Slide
from pptx.util import Inches, Pt

from . import native

TODO_MARK = "#Todo"
WELCOME_NUMBER = re.compile(r"^Meetup #\d+$")
HEADER_TITLE = re.compile(r"^#\S* Meetup$|^Meetup #\S+$")
MEMBERS_TITLE = "Introducing Our User Group Members"
SHOUTOUT_TITLE = "Shoutout"
SEE_YOU_TITLE = "See you soon"
WHAT_WE_DO_TITLE = "What we do"
WHO_WE_ARE_TITLE = "Who we are"
# The template is dark; text added to it is white, as its subtitles are.
ON_DARK = RGBColor(0xFF, 0xFF, 0xFF)


@dataclass
class Report:
    """What the build changed, and what it could not fill in."""

    changed: list[str] = field(default_factory=list)
    todo: list[str] = field(default_factory=list)
    outline: list[str] = field(default_factory=list)


# --- reading the template ---------------------------------------------------


def is_visible(slide: Slide) -> bool:
    return slide._element.get("show") != "0"


def texts(slide: Slide) -> list[str]:
    return [sh.text_frame.text.strip() for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip()]


def has_text(slide: Slide, wanted: str) -> bool:
    return any(t == wanted for t in texts(slide))


def text_shape(slide: Slide, pattern: re.Pattern[str]):
    return next((sh for sh in slide.shapes if sh.has_text_frame and pattern.match(sh.text_frame.text.strip())), None)


def pictures(slide: Slide) -> list:
    return [sh for sh in slide.shapes if sh.shape_type == 13]


def is_todo(shape) -> bool:
    return shape.has_text_frame and TODO_MARK in shape.text_frame.text


def full_slide_picture(prs: PresentationT, slide: Slide):
    """A picture covering (nearly) the whole slide, as on the header slide."""
    for pic in pictures(slide):
        if pic.width >= prs.slide_width * 0.95 and pic.height >= prs.slide_height * 0.95:
            return pic
    return None


# --- changing it ------------------------------------------------------------


def set_lines(shape, lines: list[str]) -> None:
    """Replace a text box's text, keeping the first run's formatting.

    The lines become line breaks inside one paragraph, as the template's
    subtitles use them (a vertical tab in Google Slides).
    """
    tf = shape.text_frame
    paras = tf._txBody.findall(qn("a:p"))
    first = paras[0]
    for extra in paras[1:]:
        tf._txBody.remove(extra)
    runs = first.findall(qn("a:r"))
    rpr = runs[0].find(qn("a:rPr")) if runs else None
    for child in list(first):
        if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            first.remove(child)
    end = first.find(qn("a:endParaRPr"))
    for i, line in enumerate(lines):
        if i:
            br = first.makeelement(qn("a:br"), {})
            if rpr is not None:
                br.append(copy.deepcopy(rpr))
            _insert_before(first, br, end)
        r = first.makeelement(qn("a:r"), {})
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = r.makeelement(qn("a:t"), {})
        t.text = line
        r.append(t)
        _insert_before(first, r, end)


def _insert_before(parent, child, anchor) -> None:
    if anchor is None:
        parent.append(child)
    else:
        anchor.addprevious(child)


def replace_picture(slide: Slide, old, image: Path) -> None:
    """Put another image in a picture's place: same box, same layer."""
    new = slide.shapes.add_picture(str(image), old.left, old.top, old.width, old.height)
    old._element.addprevious(new._element)
    old._element.getparent().remove(old._element)


def swap_image(slide: Slide, pic, image: Path) -> None:
    """Another image in the same picture: its shape (a circle on "Who we are"), box and layer stay; any crop is cleared."""
    _, rid = slide.part.get_or_add_image_part(str(image))
    blip_fill = pic._element.blipFill
    blip_fill.blip.set(qn("r:embed"), rid)
    src = blip_fill.find(qn("a:srcRect"))
    if src is not None:
        blip_fill.remove(src)


def layout_who_we_are(prs: PresentationT, slide: Slide, team: dict, people: Path) -> None:
    W, H = prs.slide_width, prs.slide_height
    inch = Inches(1)
    rect = next((sh for sh in slide.shapes if sh.shape_type == 1 and sh.left > W // 4 and sh.height > H * 0.9), None)
    left, width = (rect.left, rect.width) if rect is not None else (W * 2 // 5, W * 3 // 5)
    # Who is on the slide: a text box that starts with a team member's name, inside the slide.
    members = []
    for sh in list(slide.shapes):
        if not sh.has_text_frame or sh.left >= W:
            continue
        first = sh.text_frame.text.strip().split("\n")[0].strip()
        if first in team:
            mid = sh.left + sh.width // 2
            above = [p for p in pictures(slide) if p.left <= mid <= p.left + p.width and p.top + p.height <= sh.top + Inches(0.1) and sh.top - (p.top + p.height) < Inches(0.6)]
            members.append((first, sh, max(above, key=lambda p: p.top) if above else None))
    order = list(team)
    members.sort(key=lambda m: order.index(m[0]))
    if members:
        rows = [members[:3], members[3:6]] if len(members) > 3 else [members]
        photo, gap_name, name_h, mail_h, row_gap = Inches(1.0), Inches(0.1), Inches(0.32), Inches(0.24), Inches(0.5)
        # The widest name ("Philipp Bergsmann") is about 1.6 in: with that as the item width the space
        # between the names equals the space from the outer names to the purple area's edges.
        item_w = Inches(1.62)
        box_w = Inches(1.9)
        block = photo + gap_name + name_h + mail_h
        top = (H - (len(rows) * block + (len(rows) - 1) * row_gap)) // 2
        for r, row in enumerate(rows):
            n = len(row)
            space = (width - n * item_w) // (n + 1)  # the same space between the items and to the edges
            y = top + r * (block + row_gap)
            for i, (name, label, pic) in enumerate(row):
                x = left + space + i * (item_w + space)
                if pic is not None:
                    pic.left, pic.top, pic.width, pic.height = x + (item_w - photo) // 2, y, photo, photo
                    file = next((people / f"{slug(name)}{ext}" for ext in (".png", ".jpg", ".jpeg") if (people / f"{slug(name)}{ext}").is_file()), None)
                    if file is not None:
                        swap_image(slide, pic, file)
                remove(label)
                lines: list = [(name, {"size": 13.5})]
                if team[name].get("email"):
                    lines.append((team[name]["email"], {"size": 9, "color": native.GREY}))
                tx = x + (item_w - box_w) // 2
                native.text(slide, lines, tx / inch, (y + photo + gap_name) / inch, box_w / inch, (name_h + mail_h) / inch, 13.5, ON_DARK, spacing=1.05)
    # The website group in the dark panel: text, arrow, QR code, centred on the panel; the arrow and
    # the text at the middle of the QR code.
    note = next((sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip().startswith("Check our new website")), None)
    arrow = next((sh for sh in slide.shapes if sh.shape_type == 1 and "rightArrow" in sh._element.xml), None)
    qr_pic = next((p for p in pictures(slide) if p.left < left and p.top > H // 2 and abs(p.width - p.height) < Inches(0.05)), None)
    if note is not None and arrow is not None and qr_pic is not None:
        frame = next((sh for sh in slide.shapes if sh.shape_type == 1 and sh is not arrow and sh is not rect and abs(sh.left - qr_pic.left) < Inches(0.15) and abs(sh.top - qr_pic.top) < Inches(0.15)), None)
        qr_box = frame.width if frame is not None else qr_pic.width
        note_w, gap = Inches(1.69), Inches(0.12)  # about the width of "Check our new website", so the group is centred by what shows
        total = note_w + gap + arrow.width + gap + qr_box
        x0 = (left - total) // 2
        cy = qr_pic.top + qr_pic.height // 2
        note.left, note.width = x0, note_w
        note.top = cy - note.height // 2
        note.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        note.text_frame.word_wrap = False
        # Two lines, never three: an explicit break before the link (the template relied on wrapping).
        p0 = note.text_frame.paragraphs[0]
        runs = p0.runs
        if len(runs) >= 2 and p0._p.find(qn("a:br")) is None:
            runs[0].text = runs[0].text.rstrip()
            runs[0]._r.addnext(runs[0]._r.makeelement(qn("a:br"), {}))
        for p in note.text_frame.paragraphs:
            p.alignment = PP_ALIGN.RIGHT
        arrow.left = x0 + note_w + gap
        arrow.top = cy - arrow.height // 2
        qx = arrow.left + arrow.width + gap
        if frame is not None:
            dx, dy = qr_pic.left - frame.left, qr_pic.top - frame.top
            frame.left = qx
            qr_pic.left = qx + dx
        else:
            qr_pic.left = qx


def slug(name: str) -> str:
    """As people/ names its files: "Darko Mesaroš" -> "darko-mesaros"."""
    import unicodedata
    flat = unicodedata.normalize("NFD", name).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", flat).strip("-")


def fit(image: Path, box_w: int, box_h: int) -> tuple[int, int]:
    with Image.open(image) as im:
        w, h = im.size
    scale = min(box_w / w, box_h / h)
    return int(w * scale), int(h * scale)


def deck_font(slide: Slide) -> str | None:
    """The first font named on the slide (PT Sans in the group's deck)."""
    for sh in slide.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name:
                        return r.font.name
    return None


PICTURE_TYPES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}


def logo_file(event_dir: Path, ref: str | None) -> tuple[Path | None, str | None]:
    """A sponsor logo as a local picture file PowerPoint can take, or why not.

    `logo` in event.json is written for the web player: `events/<id>/assets/x`
    (from the events repo root) or `assets/x` (from the event folder). An SVG
    is used through a PNG of the same name next to it, since a .pptx picture
    cannot be an SVG on its own.
    """
    if not ref:
        return None, None
    if ref.startswith(("http://", "https://")):
        return None, f"logo {ref} is a URL; put the file in the event's assets/ for the deck"
    candidates = [event_dir / ref, event_dir.parent.parent / ref]
    found = next((c for c in candidates if c.is_file()), None)
    if found is None:
        return None, f"logo {ref} not found"
    if found.suffix.lower() == ".svg":
        png = found.with_suffix(".png")
        return (png, None) if png.is_file() else (None, f"logo {found.name} is an SVG; add {png.name} next to it for the deck")
    if found.suffix.lower() not in PICTURE_TYPES:
        return None, f"logo {found.name} is not a picture PowerPoint can show"
    return found, None


def remove(shape) -> None:
    shape._element.getparent().remove(shape._element)


def remove_todos(slide: Slide) -> None:
    for sh in [s for s in slide.shapes if is_todo(s)]:
        remove(sh)


def add_todo(prs: PresentationT, slide: Slide, text: str) -> None:
    """A note beside the slide, where the template keeps its own #Todo boxes."""
    box = slide.shapes.add_textbox(prs.slide_width + Inches(0.34), 0, Inches(1.81), Inches(1.6))
    box.text_frame.word_wrap = True
    set_plain(box, [f"{TODO_MARK}:", text], Pt(12))


def set_plain(shape, lines: list[str], size: Pt, bold: bool = False, color: RGBColor | None = None, font: str | None = None) -> None:
    tf = shape.text_frame
    tf.text = lines[0]
    for line in lines[1:]:
        tf.add_paragraph().text = line
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = size
            r.font.bold = bold
            if color is not None:
                r.font.color.rgb = color
            if font:
                r.font.name = font


def duplicate_slide(prs: PresentationT, src: Slide) -> Slide:
    """A copy of a slide at the end of the deck: its shapes, background and pictures (relationships re-pointed)."""
    dst = prs.slides.add_slide(src.slide_layout)
    for shape in list(dst.shapes):
        remove(shape)
    rid_map: dict[str, str] = {}
    for rid, rel in src.part.rels.items():
        if rel.is_external or rel.reltype.endswith(("/slideLayout", "/notesSlide")):
            continue
        rid_map[rid] = dst.part.relate_to(rel._target, rel.reltype)
    bg = src._element.cSld.find(qn("p:bg"))
    if bg is not None:
        dst._element.cSld.insert(0, copy.deepcopy(bg))
    for el in src.shapes._spTree.iterchildren():
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        clone = copy.deepcopy(el)
        for node in clone.iter():
            for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                if node.get(attr) in rid_map:
                    node.set(attr, rid_map[node.get(attr)])
        dst.shapes._spTree.append(clone)
    return dst


def set_notes(slide: Slide, text: str) -> None:
    slide.notes_slide.notes_text_frame.text = text


def move_slide(prs: PresentationT, slide: Slide, index: int) -> None:
    ids = prs.slides._sldIdLst
    for el in list(ids):
        if prs.slides.part.related_part(el.get(qn("r:id"))) is slide.part:
            ids.remove(el)
            ids.insert(index, el)
            return


def index_of(prs: PresentationT, slide: Slide) -> int:
    return next(i for i, s in enumerate(prs.slides) if s.slide_id == slide.slide_id)


def blank_layout(prs: PresentationT):
    """The template's BLANK layout, else its last one (placeholders are removed anyway)."""
    layouts = list(prs.slide_layouts)
    return next((l for l in layouts if l.name.upper() == "BLANK"), layouts[-1])


def add_image_slide(prs: PresentationT, image: Path, index: int, notes: str) -> Slide:
    slide = prs.slides.add_slide(blank_layout(prs))
    for ph in list(slide.placeholders):
        remove(ph)
    slide.shapes.add_picture(str(image), 0, 0, prs.slide_width, prs.slide_height)
    slide.notes_slide.notes_text_frame.text = notes
    move_slide(prs, slide, index)
    return slide


# --- the event --------------------------------------------------------------


def local(iso: str, tz: str) -> datetime:
    return datetime.fromisoformat(iso).astimezone(ZoneInfo(tz))


def hhmm(iso: str, tz: str) -> str:
    return local(iso, tz).strftime("%H:%M")


def meetup_number(event: dict) -> str | None:
    m = re.search(r"#(\d+)", event["title"])
    return m.group(1) if m else None


def host_full_name(event: dict) -> str | None:
    """The host as the closing thanks it: the sponsor's full name, else the short one."""
    s = next(iter(event.get("sponsors", [])), None)
    return s["name"].removesuffix(" AG").strip() if s else host_name(event)


def host_name(event: dict) -> str | None:
    """Who hosts: "hosted by RBI" in the title, else the first sponsor."""
    m = re.search(r"hosted by (.+)$", event["title"])
    if m:
        return m.group(1).strip()
    s = next(iter(event.get("sponsors", [])), None)
    return (s.get("short") or s["name"]) if s else None


def speakers_line(seg: dict) -> str:
    return ", ".join(s["name"] for s in (seg.get("talk") or {}).get("speakers", []))


def host_of(url: str) -> str:
    return re.sub(r"^https?://(www\.)?", "", url).rstrip("/")


def canva_image(canva: Path | None, name: str) -> Path | None:
    if canva is None:
        return None
    f = canva / f"{name}.png"
    return f if f.is_file() else None


@dataclass
class Inputs:
    template: Path
    event_file: Path
    stills: Path
    header: Path | None = None
    # Canva exports: header.png and one <segment id>.png per talk (talk-1.png, ...).
    canva: Path | None = None
    # One font for the whole deck (SPEC.md, "Theme"); None keeps the template's fonts.
    font: str | None = native.FONT


@dataclass
class EveningSlide:
    """One slide of the evening: a picture (talks) or a native slide (break, closing, upcoming)."""

    kind: str  # "image" | "break" | "closing" | "upcoming"
    notes: str
    image: Path | None = None
    data: dict = field(default_factory=dict)


def evening_slides(event: dict, stills: Path, canva: Path | None = None) -> list[EveningSlide]:
    """The slides after the members slide, in running order.

    A talk is a picture: its Canva card (`<canva>/<segment id>.png`, the speaker
    card the team made for the announcement) when there is one, else the
    stream's speaker handover still. Breaks, the closing and the next AWS
    Community Day are native slides in the deck's design.
    """
    tz = event["timezone"]
    out: list[EveningSlide] = []
    segments = event["segments"]
    for i, seg in enumerate(segments):
        when = f"{hhmm(seg['start'], tz)}-{hhmm(seg['end'], tz)}"
        if seg["kind"] == "talk":
            talk = seg.get("talk") or {"title": seg["title"]}
            out.append(
                EveningSlide(
                    "image",
                    f"{when}  {talk['title']}\nHand over to {speakers_line(seg) or 'the speaker'}.",
                    image=canva_image(canva, seg["id"]) or stills / f"SpeakerHandover--{seg['id']}.png",
                )
            )
        elif seg["kind"] == "break":
            nxt = next((s for s in segments[i + 1 :]), None)
            after = f" Next: {nxt['title']} at {hhmm(nxt['start'], tz)}." if nxt else ""
            data = {"back_at": hhmm(seg["end"], tz), "next": None}
            if nxt:
                t = nxt.get("talk") or {}
                data["next"] = {"title": t.get("title", nxt["title"]), "speakers": speakers_line(nxt)}
            out.append(EveningSlide("break", f"{when}  {seg['title']}. Back at {hhmm(seg['end'], tz)}.{after}", data=data))
    closing = next((s for s in segments if s["kind"] == "closing"), None)
    end = hhmm(event["end"], tz)
    names = [sp["name"] for s in segments for sp in (s.get("talk") or {}).get("speakers", [])]
    out.append(
        EveningSlide(
            "closing",
            f"{hhmm(closing['start'], tz) if closing else ''}  Thank the speakers and the sponsor, ask for feedback "
            f"({event['links']['feedback']}), networking until {end}.".strip(),
            data={"speakers": names, "feedback": event["links"]["feedback"], "until": end},
        )
    )
    coming = event.get("upcoming") or []
    if coming:
        day = coming[0]
        who = ", ".join(s["name"] for s in day.get("speakers", [])) or "nobody listed"
        out.append(
            EveningSlide(
                "upcoming",
                f"Next: {day['name']}, {day['date']}, {day['location']} ({day['url']}). Speaking from our community: {who}.",
                data={"day": day, "then": coming[1] if len(coming) > 1 else None},
            )
        )
    return out


def photo_file(event_dir: Path, ref: str | None) -> Path | None:
    """A local photo for `events/<id>/assets/x`, `assets/x` or `people/x`; None for URLs and missing files."""
    if not ref or ref.startswith(("http://", "https://")):
        return None
    for c in (event_dir / ref, event_dir.parent.parent / ref):
        if c.is_file() and c.suffix.lower() in PICTURE_TYPES:
            return c
    return None


def build(inputs: Inputs, out: Path) -> Report:
    event = json.loads(inputs.event_file.read_text())
    event_dir = inputs.event_file.parent
    tz = event["timezone"]
    prs = Presentation(str(inputs.template))
    report = Report()
    number = meetup_number(event)
    sponsors = event.get("sponsors", [])
    host = host_name(event)

    evening = evening_slides(event, inputs.stills, inputs.canva)
    for f in [e.image for e in evening if e.image is not None]:
        if not f.is_file():
            raise FileNotFoundError(f"{f} not found: render the stills for this event first (make stills EVENT={event['id']})")

    visible = [s for s in prs.slides if is_visible(s)]

    # Welcome: "Meetup #52" -> this meetup's number.
    for slide in visible:
        shape = text_shape(slide, WELCOME_NUMBER)
        if shape is None:
            continue
        if number:
            set_lines(shape, [f"Meetup #{number}"])
            remove_todos(slide)
            report.changed.append(f"slide {index_of(prs, slide) + 1}: Meetup #{number}")
        else:
            report.todo.append("welcome slide: the event title has no #number, left as it was")

    # Header slide: the Meetup header image over the whole slide.
    header_slide = None
    for slide in visible:
        title = text_shape(slide, HEADER_TITLE)
        pic = full_slide_picture(prs, slide)
        if title is None or pic is None:
            continue
        header_slide = slide
        if number:
            set_lines(title, [f"Meetup #{number}"])
        subtitle = next((sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip().startswith("hosted by")), None)
        if subtitle is not None and host:
            set_lines(subtitle, [f"hosted by {host}"])
        header = canva_image(inputs.canva, "header") or (inputs.header if inputs.header and inputs.header.is_file() else None)
        if header:
            replace_picture(slide, pic, header)
            remove_todos(slide)
            report.changed.append(f"slide {index_of(prs, slide) + 1}: the Meetup header image")
        else:
            report.todo.append("header slide: no header image (the event has none, or it was not downloaded), old image kept")
        break

    # Sponsor shoutouts: logos, or names until there is a logo.
    logos = []
    for s in sponsors:
        path, problem = logo_file(event_dir, s.get("logo"))
        if path:
            logos.append((s, path))
        elif problem:
            report.todo.append(f"{s['name']}: {problem}")
    for slide in visible:
        if not has_text(slide, SHOUTOUT_TITLE):
            continue
        pics = pictures(slide)
        box = (pics[0].left, pics[0].top, pics[0].width, pics[0].height) if pics else (
            Inches(4.66), Inches(1.4), Inches(4.9), Inches(3.6))
        for p in pics:
            remove(p)
        remove_todos(slide)
        left, top, width, height = box
        n = index_of(prs, slide) + 1
        if logos and len(logos) == len(sponsors):
            gap = Inches(0.3)
            each = (width - gap * (len(logos) - 1)) // len(logos)
            for i, (_, path) in enumerate(logos):
                w, h = fit(path, each, height)
                x = left + i * (each + gap) + (each - w) // 2
                slide.shapes.add_picture(str(path), x, top + (height - h) // 2, w, h)
            report.changed.append(f"slide {n}: sponsor logos")
        else:
            names = [s["name"] for s in sponsors] or ["(no sponsor on the Meetup page)"]
            tb = slide.shapes.add_textbox(left, top, width, height)
            tb.text_frame.word_wrap = True
            tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            set_plain(tb, names, Pt(28), bold=True, color=ON_DARK, font=deck_font(slide))
            for p in tb.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
            add_todo(prs, slide, "Add the sponsor logo: logo in source.json overrides, file in events/<id>/assets/")
            report.changed.append(f"slide {n}: sponsor names (no logo yet)")
            report.todo.append("Shoutout slide: sponsor logo missing, names shown instead")

    # Who we are: laid out from measurements, not nudged. The organizers on the slide (those inside
    # it) in rows of three and two, spread evenly over the purple area (the same space between them
    # and to its edges), each with the photo from people/<name>.png, the name on one line and the
    # email from community.json. Left: "Check our new website", the arrow and the QR code as one
    # group, centred in the dark panel, the arrow at the middle of the QR code.
    people = event_dir.parent.parent / "people"
    team = {m["name"]: m for m in (event.get("community") or {}).get("team", [])}
    who = next((sl for sl in visible if has_text(sl, WHO_WE_ARE_TITLE)), None)
    if who is not None and team:
        layout_who_we_are(prs, who, team, people)
        report.changed.append(f"slide {index_of(prs, who) + 1}: the team laid out evenly, photos from people/, emails from community.json")

    # What we do: each group's Meetup page picture links to the group, and a QR code per group under the text.
    groups = [g for g in (event.get("community") or {}).get("groups", []) if g.get("url")]
    what = next((sl for sl in visible if has_text(sl, WHAT_WE_DO_TITLE)), None)
    if what is not None and groups:
        style0 = native.Style(prs, None)
        half = prs.slide_width // 2
        # the group pages are the pictures on the right, top to bottom then left to right (Vienna, Women's, Linz)
        pages = sorted((p for p in pictures(what) if p.left >= half and p.top < prs.slide_height * 0.45), key=lambda p: (round(p.top / Inches(0.3)), p.left))
        for pic, g in zip(pages, groups):
            pic.click_action.hyperlink.address = g["url"]
        # Centred in the left panel (left of the purple box), with the same space between and beside them.
        shown = groups[:3]
        # The dark panel is twice as wide as the centre of its title ("What we do" is centred in it).
        heading = text_shape(what, re.compile(rf"^{WHAT_WE_DO_TITLE}$"))
        width = 2 * (heading.left + heading.width / 2) / Inches(1) if heading is not None else 5.0
        size, y = 0.78, 4.42
        gap = (width - len(shown) * size) / (len(shown) + 1)
        for i, g in enumerate(shown):
            x = gap + i * (size + gap)
            qr = what.shapes.add_picture(str(native.qr_png(style0, g["url"], f"group-qr-{i}")), Inches(x), Inches(y), Inches(size), Inches(size))
            qr.click_action.hyperlink.address = g["url"]
            label = g["name"].replace("AWS ", "").replace("User Group", "UG")  # "UG Vienna", "Women's UG Vienna", "UG Linz"
            box = native.text(what, [label], x - 0.3, y + size + 0.04, size + 0.6, 0.3, native.LABEL, ON_DARK)
            native.link(box, g["url"])
        report.changed.append(f"slide {index_of(prs, what) + 1}: links on the group pages, {min(len(groups), 3)} QR codes")

    # See you soon: the next meetup, which the page for this one cannot know.
    for slide in visible:
        if not has_text(slide, SEE_YOU_TITLE):
            continue
        sub = next(
            (sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip().startswith("at our")), None
        )
        if sub is not None:
            nxt = event["links"].get("nextEvents") or event["group"]["url"]
            set_lines(sub, ["at our next Meetup", host_of(nxt)])
            remove_todos(slide)
            report.changed.append(f"slide {index_of(prs, slide) + 1}: see you at the next meetup ({host_of(nxt)})")

    # New slides, in the deck's own design (talks: the Canva card or the stream still).
    style = native.Style(prs, next((s for s in visible if has_text(s, SHOUTOUT_TITLE)), None))
    agenda_notes = "\n".join(
        f"{hhmm(s['start'], tz)}-{hhmm(s['end'], tz)}  {s['title']}" + (f" ({speakers_line(s)})" if speakers_line(s) else "")
        for s in event["segments"]
    )
    rows = [
        {
            "time": f"{hhmm(s['start'], tz)}-{hhmm(s['end'], tz)}",
            "title": (s.get("talk") or {}).get("title", s["title"]),
            "speakers": speakers_line(s),
            "talk": s["kind"] == "talk",
        }
        for s in event["segments"]
    ]
    heading = f"{local(event['doorsOpen'], tz).strftime('%A')}, {local(event['doorsOpen'], tz).day} {local(event['doorsOpen'], tz).strftime('%B %Y')}  |  {event['venue']['name']}"
    after = header_slide or next((s for s in visible if text_shape(s, WELCOME_NUMBER)), None)
    at = index_of(prs, after) + 1 if after is not None else 0
    move_slide(prs, native.agenda(style, rows, heading, f"Agenda, from {event['links']['meetup']}\n{agenda_notes}"), at)
    report.changed.append(f"slide {at + 1}: new, the agenda")

    anchor = next((s for s in visible if has_text(s, MEMBERS_TITLE)), None)
    if anchor is not None:
        at = index_of(prs, anchor) + 1
    else:
        see = next((s for s in visible if has_text(s, SEE_YOU_TITLE)), None)
        at = index_of(prs, see) if see is not None else len(prs.slides)
    first_talk = None
    for e in evening:
        if e.kind == "image":
            added = add_image_slide(prs, e.image, at, e.notes)
            first_talk = first_talk or added
        elif e.kind == "break":
            move_slide(prs, native.break_slide(style, e.data["back_at"], e.data["next"], e.notes), at)
        elif e.kind == "closing":
            move_slide(prs, native.closing(style, e.data["speakers"], host_full_name(event), e.data["feedback"], e.data["until"], e.notes), at)
        elif e.kind == "upcoming":
            day = e.data["day"]
            photos = []
            for sp in day.get("speakers", []):
                f = photo_file(event_dir, sp.get("photo"))
                photos.append((sp["name"], native.round_photo(style, f, f"photo-{len(photos)}") if f else None))
            shot = photo_file(event_dir, day.get("screenshot"))
            move_slide(prs, native.upcoming(style, day, e.data["then"], photos, e.notes, shot), at)
        report.changed.append(f"slide {at + 1}: new, {e.kind}")
        at += 1

    # The host: thanked briefly at the start and at the end; its own slot right after the first talk's slide.
    shoutouts = [s for s in prs.slides if is_visible(s) and has_text(s, SHOUTOUT_TITLE)]
    full = host_full_name(event) or (sponsors[0]["name"] if sponsors else None)
    short = host or full
    if shoutouts and first_talk is not None:
        slot = duplicate_slide(prs, shoutouts[0])
        move_slide(prs, slot, index_of(prs, first_talk) + 1)
        set_notes(slot, f"Sponsor intro: {short or 'the host'} introduces themselves (about 10 minutes), then the first talk begins.")
        report.changed.append(f"slide {index_of(prs, slot) + 1}: the host's slot (shoutout), after the first talk's slide")
    for i, slide in enumerate(shoutouts):
        title = text_shape(slide, re.compile(rf"^{SHOUTOUT_TITLE}$"))
        if title is not None:
            set_lines(title, ["Thank you"])
        sub = next((sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip().lower().startswith("to our")), None)
        # On the slide only the thanks; what happens next is for the presenter, in the notes.
        thanks = "to our Location & Food Sponsor"
        if i == 0:
            if sub is not None:
                set_lines(sub, [thanks])
            set_notes(slide, f"A quick thank you to {short or 'the host'} for hosting. Before we kick off the first talk, {short or 'the host'} introduces themselves.")
        else:
            if sub is not None:
                set_lines(sub, [thanks])
            set_notes(slide, f"Thank {short or 'the host'} again for hosting tonight.")
        report.changed.append(f"slide {index_of(prs, slide) + 1}: thank you to the host")

    if inputs.font:
        native.use_font(prs, inputs.font)
        report.changed.append(f"every text in {inputs.font}")
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    report.outline = outline(prs)
    return report


def outline(prs: PresentationT) -> list[str]:
    """The visible slides in running order, by their first text or their notes."""
    lines = []
    for i, slide in enumerate(prs.slides, 1):
        if not is_visible(slide):
            continue
        shown = [t for sh in slide.shapes if sh.has_text_frame and not is_todo(sh) for t in [sh.text_frame.text.strip()] if t]
        notes = slide.notes_slide.notes_text_frame.text.split("\n")[0] if slide.has_notes_slide else ""
        label = " / ".join(shown[:2]).replace("\x0b", " ") or notes or "(picture)"
        lines.append(f"{i:>2}  {label[:90]}")
    return lines
