"""Search Report Generator – Streamlit app.

Staff type the client and mark details and upload the MyIPO "Prior Marks" PDF;
the app fills the firm's Search Report template and returns an editable .docx.
The prior-marks table is filled from the PDF; the assessment parts are left
(marked in red) for staff to write in Word.
All processing happens in memory; nothing is written to disk.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import io
import re
from datetime import date

import pandas as pd
import streamlit as st
from PIL import Image

from docx_builder import ClassSection, ReportData, build_report, class_no, report_filename
from parser import ParseError, parse_prior_marks
from parser_sg import merge_prior_marks_pdfs, parse_prior_marks_sg, parse_prior_marks_sg_files
from report_model import initial_selection, salutation_from_attention, table_rows

st.set_page_config(page_title="Search Report Generator", page_icon="📄", layout="wide")
S = st.session_state


# ----------------------------------------------------------------- password gate
def _secret(name: str, default: str | None = None) -> str | None:
    try:
        return st.secrets.get(name, default)
    except Exception:  # noqa: BLE001 - no secrets.toml at all (StreamlitSecretNotFoundError)
        return default


def password_gate() -> None:
    expected = _secret("APP_PASSWORD")
    if not expected:
        st.error("APP_PASSWORD is not configured in the app secrets.")
        st.stop()
    if S.get("authenticated"):
        return
    st.title("Search Report Generator")
    with st.form("login"):
        pw = st.text_input("Password", type="password")
        if st.form_submit_button("Enter"):
            if hmac.compare_digest(pw.encode(), str(expected).encode()):
                S.authenticated = True
                st.rerun()
            st.error("Incorrect password.")
    st.stop()


password_gate()


# ----------------------------------------------------------------- helpers
def thumb_data_url(png: bytes | None, max_px: int = 160) -> str | None:
    if not png:
        return None
    im = Image.open(io.BytesIO(png))
    im.thumbnail((max_px, max_px))
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def to_png(data: bytes) -> bytes:
    im = Image.open(io.BytesIO(data))
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGBA")
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


def parse_classes(text: str) -> list[int]:
    return sorted({int(n) for n in re.findall(r"\d+", text) if 1 <= int(n) <= 45})


def selection_df(selection: list[dict], records) -> pd.DataFrame:
    df = pd.DataFrame(selection)
    df.insert(1, "preview", [thumb_data_url(None if r.is_word_mark else r.image_png) for r in records])
    order = ["order"] if "order" in df else []  # SG only
    return df[order + ["include", "preview", "group", "app_no", "mark", "classes", "status", "flag"]]


def moved_order(orders: list) -> list[int]:
    """Row positions after the Order edits: each changed row moves to the number typed, the rest keep
    their order around it (blank = last)."""
    n = len(orders)
    target = {i: n if pd.isna(o) else min(max(int(o), 1), n) for i, o in enumerate(orders) if o != i + 1}
    rows = [i for i in range(n) if i not in target]
    for i in sorted(target, key=lambda i: (target[i], i)):
        rows.insert(target[i] - 1, i)
    return rows


def set_order(c: int, positions: list[int]) -> None:
    """Put class c's marks (records and table rows) in the given order and renumber the Order column."""
    parsed = S[f"parsed_{c}"]
    parsed.records = [parsed.records[i] for i in positions]
    selection = [S[f"selection_{c}"][i] for i in positions]
    for n, s in enumerate(selection, start=1):
        s["order"] = n
    S[f"selection_{c}"] = selection


def load_prior_marks(c: int, pdf, country: str) -> None:
    """Parse the Prior Marks PDF uploaded for class c (once per file); state is kept per class."""
    if pdf is None:
        for k in ("pdf_digest", "parsed", "selection"):
            S.pop(f"{k}_{c}", None)
        return
    digest = country + hashlib.sha256(pdf.getvalue()).hexdigest()
    if S.get(f"pdf_digest_{c}") == digest:
        return
    try:
        parsed = (parse_prior_marks_sg if country == "SG" else parse_prior_marks)(pdf.getvalue())
    except ParseError as exc:
        st.error(f"Could not read the Prior Marks PDF for Class {class_no(c)}: {exc}")
        return
    S[f"pdf_digest_{c}"] = digest
    S[f"parsed_{c}"] = parsed
    S[f"selection_{c}"] = initial_selection(parsed.records)
    S[f"editor_version_{c}"] = S.get(f"editor_version_{c}", 0) + 1


def load_prior_marks_sg(c: int, pdfs: list) -> None:
    """SG: several IPOS PDFs per class, joined in upload order (staff download the marks one by one).

    Adding or removing a file keeps the Include/Status edits and the order of marks already in the
    table; marks from new files go to the end.
    """
    if not pdfs:
        for k in ("pdf_digest", "parsed", "selection", "merged_pdf", "merged_pdf_key"):
            S.pop(f"{k}_{c}", None)
        return
    digest = "SG" + "".join(hashlib.sha256(f.getvalue()).hexdigest() for f in pdfs)
    if S.get(f"pdf_digest_{c}") == digest:
        return
    try:
        parsed = parse_prior_marks_sg_files([(f.name, f.getvalue()) for f in pdfs])
    except ParseError as exc:
        st.error(f"Could not read the Prior Marks PDFs for Class {class_no(c)}: {exc}")
        return
    old = {s["app_no"]: s for s in S.get(f"selection_{c}", [])}
    old_pos = {no: i for i, no in enumerate(old)}
    selection = initial_selection(parsed.records)
    for s in selection:
        if s["app_no"] in old:
            s["include"], s["status"] = old[s["app_no"]]["include"], old[s["app_no"]]["status"]
    S[f"pdf_digest_{c}"] = digest
    S[f"parsed_{c}"] = parsed
    S[f"selection_{c}"] = selection
    S[f"editor_version_{c}"] = S.get(f"editor_version_{c}", 0) + 1
    set_order(c, sorted(range(len(selection)), key=lambda i: (old_pos.get(selection[i]["app_no"], len(old)), i)))


def merged_pdf_sg(c: int, pdfs: list) -> bytes:
    """The combined Prior Marks PDF for class c, in the table's order (rebuilt only when that changes)."""
    key = (S[f"pdf_digest_{c}"], tuple(r.app_no for r in S[f"parsed_{c}"].records))
    if S.get(f"merged_pdf_key_{c}") != key:
        S[f"merged_pdf_{c}"] = merge_prior_marks_pdfs([f.getvalue() for f in pdfs], S[f"parsed_{c}"].records)
        S[f"merged_pdf_key_{c}"] = key
    return S[f"merged_pdf_{c}"]


def signature_png() -> bytes | None:
    """Signature from the app secrets (SIGNATURE_PNG_BASE64); None falls back to the local file."""
    b64 = _secret("SIGNATURE_PNG_BASE64")
    return base64.b64decode(b64) if b64 else None


# ----------------------------------------------------------------- page
st.title("Search Report Generator")
st.caption("Fill in the details, upload the Prior Marks PDF for each class, download the Word report. "
           "Nothing is stored; refresh the page to clear.")
country_label = st.radio("Country", ["Malaysia", "Singapore"], horizontal=True)
country = "SG" if country_label == "Singapore" else "MY"
registry = "IPOS" if country == "SG" else "MyIPO"

# ---------------- Step 1
st.header("1 · Report details")
c1, c2 = st.columns(2)
with c1:
    ref_no = st.text_input("Our Ref. No.", placeholder=f"TM{country}LQOS26xxxxx")
    letter_date = st.date_input("Report date", value=date.today(), format="DD/MM/YYYY")
    attention = st.text_input("Attention", placeholder="Ros Imanina")
    derived = salutation_from_attention(attention) if attention else ""
    if S.get("_last_derived_salutation") != derived and S.get("salutation", "") in ("", S.get("_last_derived_salutation")):
        S.salutation = derived
    S._last_derived_salutation = derived
    st.text_input("Salutation line", key="salutation")
    company = st.text_input("Client company", placeholder="ABC Sdn. Bhd.")
    address = st.text_area("Address (one line per row)", height=110)
with c2:
    subj_kind = st.radio("Client's mark", ["Logo (image)", "Word mark (text only)"], horizontal=True)
    subject_text = st.text_input("Mark name (used for word marks and the file name)", placeholder="BRANDNAME")
    subject_png = None
    if subj_kind.startswith("Logo"):
        up = st.file_uploader("Logo image", type=["png", "jpg", "jpeg"])
        if up:
            subject_png = to_png(up.getvalue())
            st.image(subject_png, width=120)
    classes_text = st.text_input("Class(es) searched", placeholder="5   or   5, 10")
    classes = parse_classes(classes_text)
    intended_goods: dict[int, str] = {}
    for c in classes:
        intended_goods[c] = st.text_area(f"Intended goods/services – Class {class_no(c)}", key=f"goods_{c}",
                                         height=90)
        if country == "SG":
            pdfs = st.file_uploader(f"Prior Marks PDFs ({registry}) – Class {class_no(c)} (one or more files)",
                                    type=["pdf"], accept_multiple_files=True, key=f"pdf_{country}_{c}")
            load_prior_marks_sg(c, pdfs)
            if pdfs and f"parsed_{c}" in S:
                st.download_button(f"Download combined Prior Marks PDF – Class {class_no(c)}", merged_pdf_sg(c, pdfs),
                                   file_name=f"[SG] Prior Marks Class {class_no(c)}.pdf", mime="application/pdf",
                                   key=f"merged_dl_{c}")
        else:
            load_prior_marks(c, st.file_uploader(f"Prior Marks PDF ({registry}) – Class {class_no(c)}",
                                                 type=["pdf"], key=f"pdf_{country}_{c}"), country)

# ---------------- Step 2
st.header("2 · Prior marks")
if not classes:
    st.info("Enter the class(es) searched, then upload the Prior Marks PDF for each class.")
    st.stop()
loaded = [c for c in classes if f"parsed_{c}" in S]
if not loaded:
    st.info("Upload the Prior Marks PDF for each class (below its intended goods).")
    st.stop()

# all PDFs normally share one print date; if not, the earliest is used
printed = [S[f"parsed_{c}"].search_date for c in loaded if S[f"parsed_{c}"].search_date]
search_date = st.date_input("Search date (taken from 'Date Printed' in the PDFs)",
                            value=min(printed) if printed else date.today(), format="DD/MM/YYYY")
if len(set(printed)) > 1:
    st.warning("The PDFs have different print dates; the earliest is used. Change it above if needed.")
st.caption("Untick **Include** to leave a mark out. Marks with the same **Group** number share one merged "
           "table row (a series of the same mark). You can correct the Status if needed."
           + (" Change a mark's **Order** number to move it (the Word table and combined PDF follow this order)."
              if country == "SG" else ""))

for c in classes:
    st.subheader(f"Class {class_no(c)} prior marks")
    if f"parsed_{c}" not in S:
        st.info(f"Upload the Prior Marks PDF for Class {class_no(c)}.")
        continue
    records = S[f"parsed_{c}"].records
    for w in S[f"parsed_{c}"].warnings:
        st.warning(w)
    st.caption(f"{len(records)} prior marks found.")
    # the editor's base frame only changes on a new PDF; user edits live in the widget state
    version = S.get(f"editor_version_{c}", 0)
    if S.get(f"sel_base_version_{c}") != version:
        S[f"sel_base_{c}"] = selection_df(S[f"selection_{c}"], records)
        S[f"sel_base_version_{c}"] = version
    edited = st.data_editor(
        S[f"sel_base_{c}"],
        key=f"sel_editor_{c}_{version}",
        hide_index=True,
        width="stretch",
        disabled=["preview", "app_no", "mark", "classes", "flag"],
        column_config={
            "order": st.column_config.NumberColumn("Order", min_value=1, step=1, width="small"),
            "include": st.column_config.CheckboxColumn("Include", width="small"),
            "preview": st.column_config.ImageColumn("Mark", width="small"),
            "group": st.column_config.NumberColumn("Group", min_value=1, step=1, width="small"),
            "app_no": "Application No.",
            "mark": "Name",
            "classes": "Class",
            "status": "Status",
            "flag": "Check",
        },
    )
    S[f"selection_{c}"] = [
        {k: (int(v) if k == "group" and pd.notna(v) else v) for k, v in row.items() if k != "preview"}
        for row in edited.to_dict("records")
    ]
    if country == "SG":
        positions = moved_order([s.get("order") for s in S[f"selection_{c}"]])
        if positions != list(range(len(positions))):
            set_order(c, positions)
            S[f"editor_version_{c}"] += 1
            st.rerun()

# ---------------- Step 3
st.header("3 · Word report")
no_pdf = [class_no(c) for c in classes if f"parsed_{c}" not in S]
missing = [m for m, ok in [("mark name", subject_text.strip()),
                           ("intended goods", all(v.strip() for v in intended_goods.values())),
                           ("Ref. No.", ref_no.strip()), ("Attention", attention.strip()),
                           ("Prior Marks PDF for Class " + ", ".join(no_pdf), not no_pdf)] if not ok]
if missing:
    st.caption("Still to fill in: " + ", ".join(missing))
if st.button("Build Word report", type="primary", disabled=bool(no_pdf) or not subject_text.strip()):
    sections = [ClassSection(cls=c, rows=table_rows(S[f"parsed_{c}"].records, S[f"selection_{c}"]))
                for c in classes]
    empty = [class_no(s.cls) for s in sections if not s.rows]
    if empty:
        st.warning("No prior marks in the table for Class " + ", ".join(empty) + ".")
    data = ReportData(
        ref_no=ref_no.strip(), letter_date=letter_date, attention=attention.strip(), company=company.strip(),
        address_lines=[l for l in address.splitlines() if l.strip()], salutation=S.salutation.strip(),
        subject_text=subject_text.strip(), subject_image=subject_png, classes=classes,
        intended_goods=intended_goods, search_date=search_date, sections=sections,
        signature_png=signature_png(), country=country,
    )
    try:
        S.docx_bytes = build_report(data)
        S.docx_name = report_filename(subject_text, country)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not build the report: {type(exc).__name__}: {exc}")

if S.get("docx_bytes"):
    st.download_button("Download .docx", S.docx_bytes, file_name=S.docx_name,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.caption("The red lines in the report mark where to write the analysis and opinion.")
