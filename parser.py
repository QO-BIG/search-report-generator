"""Parse a MyIPO "Prior Marks" PDF (concatenated TRADEMARK DETAILS printouts).

Everything is processed in memory: pass the PDF bytes, get back structured
records with the mark image (PNG bytes) or the word-mark text.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime

import fitz  # PyMuPDF
from PIL import Image, ImageChops

RENDER_DPI = 200

# Statuses shown as-is in the report; anything else is kept raw and flagged.
KNOWN_STATUSES = {"Registered", "Protected", "Under Substantive Examination", "Provisional Refusal"}

# Labels of the key/value grid at the top of each record. Used to find where
# a multi-line value ends (the next label in the same column).
_LABEL_COL_MAX_X = 150  # left label column starts at x~36
_RIGHT_LABEL_MIN_X = 280  # right label column starts at x~297
_LEFT_VALUE_X = (140, 290)
_RIGHT_VALUE_X = (400, 10_000)

_BOILERPLATE = (
    re.compile(r"^TRADEMARK DETAILS$"),
    re.compile(r"^Disclaimer: The information available is not intended"),
    re.compile(r"^\(MyIPO\) shall not be liable"),
)
_CLASS_RE = re.compile(r"^Classes and list of Goods and Services:\s*Class\s*(\d+)", re.I)


class ParseError(Exception):
    """Raised when the PDF does not look like a MyIPO Prior Marks printout."""


@dataclass
class Line:
    page: int
    x0: float
    y0: float
    x1: float
    y1: float
    text: str


@dataclass
class PriorMark:
    app_no: str
    status_raw: str = ""
    status: str = ""
    status_flagged: bool = False
    classes: list[int] = field(default_factory=list)
    goods: dict[int, str] = field(default_factory=dict)
    nature: str = ""
    denomination: str = ""
    description: str = ""
    disclaimer: str = ""
    applicant: str = ""
    applicant_address: str = ""
    date_printed: date | None = None
    mark_text: str = ""  # text shown for Word marks (standard characters)
    image_png: bytes | None = None
    image_hash: int | None = None
    pages: list[int] = field(default_factory=list)

    @property
    def is_word_mark(self) -> bool:
        return self.nature.strip().lower() == "word"

    @property
    def classes_str(self) -> str:
        return ", ".join(str(c) for c in self.classes)

    @property
    def display_name(self) -> str:
        """Best human-readable name for the mark (for prompts and UI)."""
        return self.mark_text or self.denomination or "(device mark)"


@dataclass
class ParseResult:
    records: list[PriorMark]
    search_date: date | None
    warnings: list[str]


# --------------------------------------------------------------------------- utils

def normalise_status(raw: str) -> tuple[str, bool]:
    """Return (status for the report, flagged-for-review)."""
    s = re.sub(r"\s+", " ", raw).strip()
    if s.lower().startswith("provisional refusal"):
        return "Provisional Refusal", False
    for known in KNOWN_STATUSES:
        if s.lower() == known.lower():
            return known, False
    return s, True


def _parse_date(s: str) -> date | None:
    s = s.strip()
    for fmt in ("%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _norm(s: str) -> str:
    return re.sub(r"[^A-Z0-9]+", " ", s.upper()).strip()


def _dhash(img: Image.Image, size: int = 16) -> int:
    g = img.convert("L").resize((size + 1, size), Image.LANCZOS)
    px = g.tobytes()
    bits = 0
    for row in range(size):
        for col in range(size):
            left = px[row * (size + 1) + col]
            right = px[row * (size + 1) + col + 1]
            bits = (bits << 1) | (left > right)
    return bits


def hash_distance(a: int | None, b: int | None) -> int:
    if a is None or b is None:
        return 10_000
    return bin(a ^ b).count("1")


def _trim(img: Image.Image, pad: int = 6) -> Image.Image | None:
    rgb = img.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert("L").point(lambda v: 255 if v > 18 else 0)
    box = diff.getbbox()
    if not box:
        return None
    l, t, r, b = box
    l, t = max(0, l - pad), max(0, t - pad)
    r, b = min(rgb.width, r + pad), min(rgb.height, b + pad)
    return rgb.crop((l, t, r, b))


def _xref_image(doc: fitz.Document, xref: int, smask: int) -> Image.Image | None:
    try:
        pix = fitz.Pixmap(doc, xref)
        if pix.colorspace and pix.colorspace.n > 3:  # CMYK etc.
            pix = fitz.Pixmap(fitz.csRGB, pix)
        if smask:
            pix = fitz.Pixmap(pix, fitz.Pixmap(doc, smask))
        img = Image.open(io.BytesIO(pix.tobytes("png")))
    except Exception:  # noqa: BLE001 - unusual encodings fall back to the rendered crop
        return None
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        img = Image.alpha_composite(bg, img)
    return img.convert("RGB")


def _native_mark_image(doc: fitz.Document, page: fitz.Page, clip: fitz.Rect) -> Image.Image | None:
    """MyIPO stretches every mark into a square box; the embedded raster keeps
    the true proportions. Use it (several images composed in their printed
    layout) when the mark area contains only rasters."""
    placed = []
    for info in page.get_images(full=True):
        xref, smask = info[0], info[1]
        for rect in page.get_image_rects(xref):
            if clip.contains(rect) or (rect & clip).get_area() > 0.8 * rect.get_area():
                placed.append((rect, xref, smask))
    if not placed or len(placed) > 4:
        return None
    covered = fitz.Rect()
    for rect, _, _ in placed:
        covered |= rect
    # any text/vector content in the box outside the rasters -> keep the rendered crop
    if page.get_text("text", clip=clip).strip() or any(
            fitz.Rect(d["rect"]).intersects(clip) and not covered.contains(fitz.Rect(d["rect"]))
            for d in page.get_drawings()):
        return None
    images = []
    for rect, xref, smask in sorted(placed, key=lambda p: (round(p[0].y0), p[0].x0)):
        img = _xref_image(doc, xref, smask)
        if img is None:
            return None
        images.append((rect, _trim(img, pad=2) or img))
    if len(images) == 1:
        return images[0][1]
    vertical = (images[-1][0].y0 - images[0][0].y0) > (images[-1][0].x0 - images[0][0].x0)
    gap = 20
    if vertical:
        width = max(im.width for _, im in images)
        scaled = [im.resize((width, max(1, round(im.height * width / im.width)))) for _, im in images]
        canvas = Image.new("RGB", (width, sum(s.height for s in scaled) + gap * (len(scaled) - 1)), "white")
        y = 0
        for s in scaled:
            canvas.paste(s, (0, y))
            y += s.height + gap
    else:
        height = max(im.height for _, im in images)
        scaled = [im.resize((max(1, round(im.width * height / im.height)), height)) for _, im in images]
        canvas = Image.new("RGB", (sum(s.width for s in scaled) + gap * (len(scaled) - 1), height), "white")
        x = 0
        for s in scaled:
            canvas.paste(s, (x, 0))
            x += s.width + gap
    return canvas


# --------------------------------------------------------------------------- extraction

def _page_lines(page: fitz.Page, page_no: int) -> tuple[list[Line], float, str | None]:
    """Content lines (footer/boilerplate removed), footer top y, Date Printed text."""
    raw: list[Line] = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for ln in block["lines"]:
            text = "".join(s["text"] for s in ln["spans"]).strip()
            if text:
                x0, y0, x1, y1 = ln["bbox"]
                raw.append(Line(page_no, x0, y0, x1, y1, text))

    footer_top = page.rect.height
    date_printed = None
    dp = [l for l in raw if l.text == "Date Printed"]
    if dp:
        footer_top = dp[0].y0 - 1
        same_row = [l for l in raw if abs(l.y0 - dp[0].y0) < 3 and l.x0 > dp[0].x1 and l.text != ":"]
        if same_row:
            date_printed = min(same_row, key=lambda l: l.x0).text.lstrip(": ").strip()

    lines = [
        l for l in raw
        if l.y0 < footer_top and not any(p.search(l.text) for p in _BOILERPLATE)
    ]
    lines.sort(key=lambda l: (round(l.y0), l.x0))
    return lines, footer_top, date_printed


def _is_label(line: Line) -> bool:
    return line.x0 < _LABEL_COL_MAX_X or _RIGHT_LABEL_MIN_X <= line.x0 < _RIGHT_LABEL_MIN_X + 20


def _field(lines: list[Line], label: str) -> str | None:
    """Value of a grid field such as 'Legal Status' (None if label absent)."""
    pat = re.compile(rf"^{re.escape(label)}\s*:?\s*(?::\s*(.*))?$")
    for i, lab in enumerate(lines):
        m = pat.match(lab.text)
        if not m or not _is_label(lab):
            continue
        right_col = lab.x0 >= _RIGHT_LABEL_MIN_X
        vx = _RIGHT_VALUE_X if right_col else _LEFT_VALUE_X
        # next label in the same column bounds a multi-line value
        nxt = next(
            (l.y0 for l in lines[i + 1:]
             if l.page == lab.page and abs(l.x0 - lab.x0) < 6 and l.y0 > lab.y0 + 2),
            lab.y0 + 60,
        )
        parts = [m.group(1)] if m.group(1) else []
        for l in lines:
            if l.page == lab.page and vx[0] <= l.x0 < vx[1] and lab.y0 - 2 <= l.y0 < nxt - 1:
                parts.append(l.text)
        value = " ".join(p.strip() for p in parts)
        return re.sub(r"\s+", " ", value).lstrip(": ").strip()
    return None


def _find(lines: list[Line], pred) -> int | None:
    return next((i for i, l in enumerate(lines) if pred(l)), None)


def _extract_record(doc: fitz.Document, lines: list[Line], footer_tops: dict[int, float],
                    warnings: list[str]) -> PriorMark:
    app_no = _field(lines, "Application number") or ""
    rec = PriorMark(app_no=app_no.split()[0] if app_no else "")
    rec.pages = sorted({l.page for l in lines})

    status = _field(lines, "Legal Status")
    if status is None:
        status = _field(lines, "Application status")
    rec.status_raw = status or ""
    rec.status, rec.status_flagged = normalise_status(rec.status_raw)
    if rec.status_flagged:
        warnings.append(f"{rec.app_no}: unrecognised status '{rec.status_raw}' – please review.")

    rec.nature = _field(lines, "Nature of Mark") or ""
    rec.disclaimer = _field(lines, "Disclaimer/Condition") or ""

    # Applicant block: left column between "Applicant:" and Denomination/Mark
    i_app = _find(lines, lambda l: l.text == "Applicant:")
    i_mark = _find(lines, lambda l: l.text == "Mark:" and l.x0 < _LABEL_COL_MAX_X)
    i_den = _find(lines, lambda l: l.text.startswith("Denomination:"))
    if i_app is not None:
        stop = min(x for x in (i_den, i_mark, len(lines)) if x is not None)
        app_lines = [l.text for l in lines[i_app + 1:stop] if l.x0 < _RIGHT_LABEL_MIN_X - 10]
        if app_lines:
            # National records: name lines precede "Default Address:"; IR records: name is line 1
            n_name = next((k for k, t in enumerate(app_lines) if t.startswith("Default Address:")), 1)
            rec.applicant = " ".join(app_lines[:max(n_name, 1)]).strip()
            rec.applicant_address = " ".join(app_lines[max(n_name, 1):]).replace("Default Address:", "").strip()

    if i_den is not None:
        den = [lines[i_den].text.split(":", 1)[1].strip()]
        stop = i_mark if i_mark is not None and i_mark > i_den else i_den + 1
        den += [l.text for l in lines[i_den + 1:stop]]
        rec.denomination = " ".join(d for d in den if d).strip()

    i_desc = _find(lines, lambda l: l.text.startswith("Mark Description:"))
    i_color = _find(lines, lambda l: l.text.startswith("Color Condition:"))
    if i_desc is not None:
        desc = [lines[i_desc].text.split(":", 1)[1].strip()]
        if i_color is not None and i_color > i_desc:
            desc += [l.text for l in lines[i_desc + 1:i_color]]
        rec.description = " ".join(d for d in desc if d).strip()

    # Classes + goods/services text
    current: int | None = None
    for l in lines:
        m = _CLASS_RE.match(l.text)
        if m:
            current = int(m.group(1))
            rec.goods.setdefault(current, "")
            if current not in rec.classes:
                rec.classes.append(current)
            continue
        if current is not None:
            rec.goods[current] = (rec.goods[current] + " " + l.text).strip()
    rec.classes.sort()

    # Mark representation
    if i_mark is not None:
        mark_line = lines[i_mark]
        end_line = lines[i_desc] if i_desc is not None else None
        in_box = [
            l.text for l in lines[i_mark + 1:(i_desc if i_desc is not None else i_mark + 1)]
            if l.page == mark_line.page
        ]
        if rec.is_word_mark:
            rec.mark_text = " ".join(in_box).strip() or rec.denomination
        else:
            page = doc[mark_line.page]
            bottom = (end_line.y0 - 1 if end_line is not None and end_line.page == mark_line.page
                      else footer_tops.get(mark_line.page, page.rect.height) - 2)
            clip = fitz.Rect(20, mark_line.y1 + 1, page.rect.width - 20, bottom)
            if clip.height > 4:
                img = _native_mark_image(doc, page, clip)
                if img is None:
                    pix = page.get_pixmap(dpi=RENDER_DPI, clip=clip, alpha=False)
                    img = _trim(Image.open(io.BytesIO(pix.tobytes("png"))))
                if img is not None:
                    buf = io.BytesIO()
                    img.save(buf, format="PNG", optimize=True)
                    rec.image_png = buf.getvalue()
                    rec.image_hash = _dhash(img)
            if rec.image_png is None:
                warnings.append(f"{rec.app_no}: could not crop the mark image.")
                rec.mark_text = rec.denomination
    else:
        warnings.append(f"{rec.app_no}: 'Mark:' label not found.")
        rec.mark_text = rec.denomination

    if not rec.classes:
        warnings.append(f"{rec.app_no}: no class found.")
    return rec


def parse_prior_marks(pdf_bytes: bytes) -> ParseResult:
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:  # noqa: BLE001 - surface any open error to the UI
        raise ParseError(f"Could not open the PDF: {exc}") from exc

    warnings: list[str] = []
    all_lines: list[Line] = []
    footer_tops: dict[int, float] = {}
    printed_by_page: dict[int, date] = {}
    for i, page in enumerate(doc):
        lines, ftop, dp = _page_lines(page, i)
        footer_tops[i] = ftop
        all_lines.extend(lines)
        if dp and (d := _parse_date(dp)):
            printed_by_page[i] = d
    printed = list(printed_by_page.values())

    if not any(l.text == "Application number" for l in all_lines) and not any(
        l.text.startswith("Application number") for l in all_lines
    ):
        if not "".join(l.text for l in all_lines).strip():
            raise ParseError("The PDF has no text layer (scanned?). Please upload the MyIPO printout PDF.")
        raise ParseError("No 'Application number' found – this does not look like a MyIPO Prior Marks PDF.")

    # Split into records at each "Application number" label
    chunks: list[list[Line]] = []
    for l in all_lines:
        if l.text.startswith("Application number") and l.x0 < _LABEL_COL_MAX_X:
            chunks.append([])
        if chunks:
            chunks[-1].append(l)

    records = [_extract_record(doc, chunk, footer_tops, warnings) for chunk in chunks]
    for rec in records:
        rec.date_printed = next((printed_by_page[p] for p in rec.pages if p in printed_by_page), None)

    search_date = max(set(printed), key=printed.count) if printed else None
    if search_date is None:
        warnings.append("'Date Printed' not found – please enter the search date manually.")
    elif len(set(printed)) > 1:
        warnings.append("Records were printed on different dates; using the most common one as the search date.")
    return ParseResult(records=records, search_date=search_date, warnings=warnings)


# --------------------------------------------------------------------------- merging

def suggest_groups(records: list[PriorMark], max_hash_distance: int = 20) -> list[int]:
    """Group number per record (1-based, PDF order). Records sharing a group
    are shown as one table row: same applicant, same mark text/denomination,
    same nature and (for device marks) a near-identical image."""
    groups: list[int] = []
    next_group = 1
    for i, rec in enumerate(records):
        match = None
        for j in range(i):
            other = records[j]
            if not rec.applicant or _norm(rec.applicant) != _norm(other.applicant):
                continue
            if _norm(rec.display_name) != _norm(other.display_name) or rec.nature != other.nature:
                continue
            if rec.is_word_mark or hash_distance(rec.image_hash, other.image_hash) <= max_hash_distance:
                match = groups[j]
                break
        if match is None:
            match = next_group
            next_group += 1
        groups.append(match)
    return groups
