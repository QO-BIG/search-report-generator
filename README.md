# Search Report Generator

A small web app for Quality Oracle staff. Type the client and mark details, upload the MyIPO
"Prior Marks" PDF for each class searched, and download the Search Report as an editable Word file
(`[MY] Search Report - <mark>.docx`).

## What the app fills in
| In the report | Comes from |
|---|---|
| Ref. No., date, attention, company, address, salutation | typed in the app |
| The mark (logo image or word) and class(es) in the RE line and table headings | typed / uploaded in the app |
| Search date ("as of …") | the "Date Printed" in the PDF (can be changed) |
| Intended goods/services per class | typed in the app |
| Prior-marks table: mark image, application no., class, status | read from the Prior Marks PDF |
| Signature | app secrets (`SIGNATURE_PNG_BASE64`) |

Left for staff to write in Word: the summary table (only its Class column is filled), the Part I analysis, each Opinion and the
Similarity column. Red lines in the report show where to write; delete them afterwards.

**Malaysia or Singapore:** choose the country at the top. Malaysia reads MyIPO printouts;
Singapore reads IPOS Digital Hub printouts ("Details of Trade Mark" pages, printed from the
browser). The SG report uses the same layout with the wording of QO's SG reports
(Trademark Register of Singapore, "their related classes", SG Opinion/Part I closings, etc.).

Each searched class has its own Prior Marks PDF, uploaded below that class's intended goods.
The report has one prior-marks table and Opinion per class (as in past reports), filled from
that class's PDF. If the PDFs have different print dates, the earliest is used.

## Files
- `app.py` – the Streamlit app (password gate, inputs, prior-marks review, download).
- `parser.py` – reads the MyIPO (Malaysia) Prior Marks PDF (fields and mark images).
- `parser_sg.py` – reads the IPOS (Singapore) Prior Marks PDF.
- `report_model.py` – turns the reviewed marks into table rows.
- `docx_builder.py` – writes the Word report into `assets/template.docx` (QO letterhead).

Nothing is stored: uploads are processed in memory and nothing is written to disk or logged.

## Secrets (Streamlit → App settings → Secrets)
```
APP_PASSWORD = "..."
SIGNATURE_PNG_BASE64 = "..."   # the line from "signature for Streamlit secrets.txt"
```

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```
Put the same secrets in `.streamlit/secrets.toml` (see the `.example` file).

Tests (`pip install -r requirements-dev.txt`, then `python -m pytest -q tests`) use the sample
PDFs in `samples/`, which are kept on the developer's PC only.
