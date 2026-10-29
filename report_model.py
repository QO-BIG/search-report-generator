"""Glue between parsed records, the editable UI table and the report tables.

The UI keeps one "selection" row per parsed record:
    {"include": bool, "group": int, "app_no": str, "mark": str, "classes": str, "status": str,
     "flag": str}
There is one selection per searched class (each class has its own Prior Marks PDF).
Included records sharing a group number become one (merged) table row.
"""
from __future__ import annotations

from docx_builder import TableRow
from parser import PriorMark, suggest_groups


def initial_selection(records: list[PriorMark]) -> list[dict]:
    groups = suggest_groups(records)
    return [
        {"include": True, "group": g, "app_no": r.app_no, "mark": r.display_name,
         "classes": r.classes_str, "status": r.status,
         "flag": "status?" if r.status_flagged else ""}
        for r, g in zip(records, groups)
    ]


def grouped(records: list[PriorMark], selection: list[dict]) -> list[list[tuple[PriorMark, dict]]]:
    """Included records grouped by group number, in order of first appearance."""
    order: list[int] = []
    buckets: dict[int, list] = {}
    for rec, sel in zip(records, selection):
        if not sel.get("include"):
            continue
        g = int(sel.get("group") or 0)
        if g not in buckets:
            buckets[g] = []
            order.append(g)
        buckets[g].append((rec, sel))
    return [buckets[g] for g in order]


def table_rows(records: list[PriorMark], selection: list[dict]) -> list[TableRow]:
    """Rows of the prior-marks table."""
    rows = []
    for members in grouped(records, selection):
        first_rec, first_sel = members[0]
        rows.append(TableRow(
            numbers=[s["app_no"] for _, s in members],
            classes=[s["classes"] for _, s in members],
            status=first_sel.get("status") or first_rec.status,
            image_png=None if first_rec.is_word_mark else first_rec.image_png,
            mark_text=first_rec.mark_text or first_rec.denomination,
        ))
    return rows


def salutation_from_attention(attention: str) -> str:
    """'Mr. John Tan' -> 'Dear Mr. John,'"""
    parts = attention.split()
    if not parts:
        return "Dear Sir/Madam,"
    titles = {"mr", "mr.", "ms", "ms.", "mrs", "mrs.", "dr", "dr.", "miss", "mdm", "mdm.", "madam", "prof", "prof.",
              "tuan", "puan", "encik", "cik", "dato'", "dato", "datuk", "datin"}
    lead = []
    while parts and parts[0].lower() in titles and len(lead) < 3:
        lead.append(parts.pop(0))
    name = parts[0] if parts else ""
    return f"Dear {' '.join(lead + [name]).strip()},"
