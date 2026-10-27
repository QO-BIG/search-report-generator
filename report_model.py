"""Glue between parsed records, the editable UI table and the report tables.

The UI keeps one "selection" row per parsed record:
    {"include": bool, "group": int, "app_no": str, "mark": str, "classes": str, "status": str,
     "in_<class>": bool}   (one key per searched class, only used when several classes are searched)
Included records sharing a group number become one (merged) table row.
"""
from __future__ import annotations

import re

from docx_builder import TableRow
from parser import PriorMark, suggest_groups


def in_key(cls: int) -> str:
    return f"in_{cls}"


def initial_selection(records: list[PriorMark]) -> list[dict]:
    groups = suggest_groups(records)
    return [
        {"include": True, "group": g, "app_no": r.app_no, "mark": r.display_name,
         "classes": r.classes_str, "status": r.status,
         "flag": "status?" if r.status_flagged else ""}
        for r, g in zip(records, groups)
    ]


def _mark_classes(text: str) -> set[int]:
    return {int(n) for n in re.findall(r"\d+", text or "")}


def ensure_class_keys(selection: list[dict], classes: list[int]) -> list[dict]:
    """Add an 'in Class X table' tick for every searched class.

    Default: a mark filed in one of the searched classes goes into that class's table;
    a mark in none of them (a related class) goes into every table, for staff to untick.
    """
    out = []
    for s in selection:
        own = _mark_classes(s.get("classes", "")) & set(classes)
        out.append({**s, **{in_key(c): s.get(in_key(c), (c in own) if own else True) for c in classes}})
    return out


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


def table_rows(records: list[PriorMark], selection: list[dict], cls: int | None = None) -> list[TableRow]:
    """Rows of the prior-marks table (for one searched class when cls is given)."""
    rows = []
    for members in grouped(records, selection):
        if cls is not None and not any(s.get(in_key(cls), True) for _, s in members):
            continue
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
