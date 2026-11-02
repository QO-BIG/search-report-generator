"""Parse a Singapore (IPOS Digital Hub) "Prior Marks" PDF.

The PDF is a set of browser printouts of "Details of Trade Mark" (or "Details of
Logogram") pages, one record after another. Each record starts with
"Trade Mark No. <number>" (or "Application No. <number>" for a logogram) and has
one embedded raster for the mark (a series of marks is a single image).

Records come back as the same PriorMark objects as the MyIPO parser, so the rest
of the app treats both countries alike. Everything is processed in memory.
"""
from __future__ import annotations

import io
import re
from collections import Counter
from datetime import date

import fitz  # PyMuPDF

from parser import ParseError, ParseResult, PriorMark, _dhash, _trim, _xref_image

KNOWN_STATUSES_SG = {"Registered", "Pending (Under Examination)", "Recorded", "Expired (Renewal Possible)",
                     "Removed"}

_START_RE = re.compile(r"^(?:Trade Mark No\.|Application No\.)\s+(\S+)$")
_STATUS_RE = re.compile(r"^Mark Status \(unless otherwise indicated below\)\s*(.*)$")
_LOGOGRAM_STATUS_RE = re.compile(r"^Status\s+(?!Date\b|Update\b)(.+)$")
# browser print header, e.g. "9/4/26, 3:51 PM" (US, month first) or "04/09/2026, 15:51" (day first)
_PRINTED_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2}|\d{4}), \d{1,2}:\d{2}(\s*[AP]M)?$")


class SGPriorMark(PriorMark):
    """IPOS record: every mark is shown as its image; classes as in the SG reports ("03, 05, 44")."""

    @property
    def classes_str(self) -> str:
        return ", ".join(f"{c:02d}" for c in self.classes) or "-"

    @property
    def display_name(self) -> str:
        return self.mark_text or "(image)"


def normalise_status_sg(raw: str) -> tuple[str, bool]:
    s = re.sub(r"\s+", " ", raw).strip()
    for known in KNOWN_STATUSES_SG:
        if s.lower() == known.lower():
            return known, False
    return s, True


def _printed_date(text: str) -> date | None:
    m = _PRINTED_RE.match(text.strip())
    if not m:
        return None
    a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    y += 2000 if y < 100 else 0
    month, day = (a, b) if m.group(4) else (b, a)  # AM/PM clock = US browser locale = month first
    try:
        return date(y, month, day)
    except ValueError:
        return None


def _spec_classes(page: fitz.Page) -> list[int]:
    """Class numbers in the "Specification of Goods or Services" table on this page.

    The table header row reads "S/N | Class No. | Goods/Services | Class Status | ...";
    the class numbers are the two-digit words in the "Class No." column below it.
    """
    words = page.get_text("words")  # (x0, y0, x1, y1, text, ...)
    classes = []
    for w in words:
        if w[4] != "Class":
            continue
        row = [o[4] for o in words if abs(o[1] - w[1]) < 3]
        if not {"S/N", "No.", "Goods/Services"} <= set(row):
            continue
        for o in words:
            if (abs(o[0] - w[0]) < 4 and o[1] > w[1] + 5 and re.fullmatch(r"\d{2}", o[4])
                    and 1 <= int(o[4]) <= 45):
                classes.append(int(o[4]))
    return classes


def _mark_image(doc: fitz.Document, page: fitz.Page):
    """The largest raster on the record's first page below the IPOS logo."""
    best = None
    for info in page.get_images(full=True):
        xref, smask = info[0], info[1]
        for rect in page.get_image_rects(xref):
            if rect.y0 < 150:  # IPOS logo in the page header
                continue
            if best is None or rect.get_area() > best[0].get_area():
                best = (rect, xref, smask)
    if best is None:
        return None
    img = _xref_image(doc, best[1], best[2])
    return (_trim(img, pad=2) or img) if img is not None else None


def parse_prior_marks_sg(pdf_bytes: bytes) -> ParseResult:
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:  # noqa: BLE001 - surface any open error to the UI
        raise ParseError(f"Could not open the PDF: {exc}") from exc

    pages = [[l.strip() for l in page.get_text("text").splitlines() if l.strip()] for page in doc]
    if not any(pages):
        raise ParseError("The PDF has no text layer (scanned?). Please upload the IPOS printout PDF.")

    starts = [(i, m.group(1)) for i, lines in enumerate(pages) for m in map(_START_RE.match, lines) if m]
    if not starts:
        raise ParseError("No 'Trade Mark No.' found – this does not look like an IPOS Prior Marks PDF.")

    warnings: list[str] = []
    records: list[PriorMark] = []
    printed: list[date] = []
    for k, (first, app_no) in enumerate(starts):
        last = starts[k + 1][0] if k + 1 < len(starts) else len(doc)
        lines = [l for p in range(first, last) for l in pages[p]]
        rec = SGPriorMark(app_no=app_no, pages=list(range(first, last)))

        status = next((m.group(1) for m in map(_STATUS_RE.match, lines) if m), None)
        if status is None:  # logogram records have a plain "Status" field
            status = next((m.group(1) for m in map(_LOGOGRAM_STATUS_RE.match, lines) if m), "")
        rec.status_raw = status
        rec.status, rec.status_flagged = normalise_status_sg(status)
        if rec.status_flagged:
            warnings.append(f"{app_no}: unrecognised status '{status}' – please review.")

        rec.classes = sorted({c for p in range(first, last) for c in _spec_classes(doc[p])})

        dates = [d for d in map(_printed_date, lines) if d]
        rec.date_printed = dates[0] if dates else None
        printed += dates[:1]

        img = _mark_image(doc, doc[first])
        if img is not None:
            buf = io.BytesIO()
            img.save(buf, format="PNG", optimize=True)
            rec.image_png = buf.getvalue()
            rec.image_hash = _dhash(img)
        else:
            warnings.append(f"{app_no}: could not find the mark image.")
        records.append(rec)

    search_date = Counter(printed).most_common(1)[0][0] if printed else None
    if search_date is None:
        warnings.append("Print date not found – please enter the search date manually.")
    return ParseResult(records=records, search_date=search_date, warnings=warnings)
