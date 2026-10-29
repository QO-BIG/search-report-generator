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
    return df[["include", "preview", "group", "app_no", "mark", "classes", "status", "flag"]]


def load_prior_marks(c: int, pdf) -> None:
    """Parse the Prior Marks PDF uploaded for class c (once per file); state is kept per class."""
    if pdf is None:
        for k in ("pdf_digest", "parsed", "selection"):
            S.pop(f"{k}_{c}", None)
        return
    digest = hashlib.sha256(pdf.getvalue()).hexdigest()
    if S.get(f"pdf_digest_{c}") == digest:
        return
    try:
        parsed = parse_prior_marks(pdf.getvalue())
    except ParseError as exc:
        st.error(f"Could not read the Prior Marks PDF for Class {class_no(c)}: {exc}")
        return
    S[f"pdf_digest_{c}"] = digest
    S[f"parsed_{c}"] = parsed
    S[f"selection_{c}"] = initial_selection(parsed.records)
    S[f"editor_version_{c}"] = S.get(f"editor_version_{c}", 0) + 1


def signature_png() -> bytes | None:
    """Signature from the app secrets (SIGNATURE_PNG_BASE64); None falls back to the local file."""
    b64 = _secret("SIGNATURE_PNG_BASE64")
    return base64.b64decode(b64) if b64 else None


# ----------------------------------------------------------------- page
st.title("Search Report Generator")
st.caption("Fill in the details, upload the Prior Marks PDF for each class, download the Word report. "
           "Nothing is stored; refresh the page to clear.")

# ---------------- Step 1
st.header("1 · Report details")
c1, c2 = st.columns(2)
with c1:
    ref_no = st.text_input("Our Ref. No.", placeholder="TMMYLQOS26xxxxx")
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
        load_prior_marks(c, st.file_uploader(f"Prior Marks PDF – Class {class_no(c)}", type=["pdf"],
                                             key=f"pdf_{c}"))

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
           "table row (a series of the same mark). You can correct the Status if needed.")

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
        signature_png=signature_png(),
    )
    try:
        S.docx_bytes = build_report(data)
        S.docx_name = report_filename(subject_text)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not build the report: {type(exc).__name__}: {exc}")

if S.get("docx_bytes"):
    st.download_button("Download .docx", S.docx_bytes, file_name=S.docx_name,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.caption("The red lines in the report mark where to write the analysis and opinion.")
