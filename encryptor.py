"""
encryptor.py
------------
Core logic for anonymising and restoring Excel data.

Two APIs are provided:
  DataFrame API  – anonymise() / restore()
      Works on pandas DataFrames.  Handy for testing and for generating
      in-app previews.  Does NOT preserve Excel formatting.

  Workbook API   – anonymise_workbook() / restore_workbook()
      Works on openpyxl Workbook objects.  Replaces only cell *values*
      and leaves every style, border, font, merged cell, conditional
      format, chart, etc. completely intact.  Use this for all actual
      file I/O so that formatting is preserved end-to-end.

Strategy
--------
Categorical columns: each unique value gets a stable token like
  "<ColumnName>__001", "<ColumnName>__002", …
Numeric columns: left untouched (no information leaked).

The mapping dict that the anonymise functions return is all that is needed
to reverse the transformation with the matching restore function.
"""

from __future__ import annotations

import io
import json
import re
from typing import Any

import openpyxl
import pandas as pd


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _safe_prefix(col_name: str) -> str:
    """
    Turn an arbitrary column name into a token-safe prefix.
    Keeps alphanumeric characters and underscores; collapses the rest to '_'.
    """
    return re.sub(r"[^A-Za-z0-9_]", "_", str(col_name))


def _clone_workbook(wb: openpyxl.Workbook) -> openpyxl.Workbook:
    """Return a fresh independent copy of a workbook via a bytes round-trip."""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return openpyxl.load_workbook(buf)


def wb_to_bytes(wb: openpyxl.Workbook) -> bytes:
    """Serialise a workbook to bytes (for Streamlit download buttons)."""
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def wb_to_dataframe(wb: openpyxl.Workbook) -> pd.DataFrame:
    """Convert the active sheet of a workbook to a DataFrame (for previews)."""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return pd.read_excel(buf, engine="openpyxl")


# ---------------------------------------------------------------------------
# Workbook API  (preserves all Excel formatting)
# ---------------------------------------------------------------------------

def anonymise_workbook(
    wb: openpyxl.Workbook,
    columns: list[str],
) -> tuple[openpyxl.Workbook, dict[str, dict[str, Any]]]:
    """
    Anonymise the selected columns in the active sheet of *wb*.

    Returns a (cloned) workbook with only the cell values changed and a
    mapping dict that `restore_workbook()` can use to reverse the operation.
    All formatting, styles, merged cells, charts, etc. are preserved.

    Parameters
    ----------
    wb:
        Source workbook (will not be mutated – a clone is returned).
    columns:
        Column names (from the first row / header row) to anonymise.

    Returns
    -------
    anonymised_wb:
        Clone of *wb* with selected columns replaced by tokens.
    mapping:
        { col_name: { token: original_value, … }, … }
    """
    wb = _clone_workbook(wb)
    ws = wb.active

    # Build a map of 1-based column index → column name for selected columns.
    idx_to_col: dict[int, str] = {}
    for cell in ws[1]:  # row 1 = header
        if cell.value in columns:
            idx_to_col[cell.column] = str(cell.value)

    missing = set(columns) - set(idx_to_col.values())
    if missing:
        raise ValueError(f"Column(s) not found in sheet: {missing}")

    mapping: dict[str, dict[str, Any]] = {col: {} for col in columns}
    counters: dict[str, int] = {col: 1 for col in columns}
    reverse_lookup: dict[str, dict[Any, str]] = {col: {} for col in columns}

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            col_name = idx_to_col.get(cell.column)
            if col_name is None or cell.value is None:
                continue

            val = cell.value
            if val not in reverse_lookup[col_name]:
                prefix = _safe_prefix(col_name)
                token = f"{prefix}__{counters[col_name]:03d}"
                reverse_lookup[col_name][val] = token
                mapping[col_name][token] = val
                counters[col_name] += 1

            cell.value = reverse_lookup[col_name][val]

    return wb, mapping


def restore_workbook(
    wb: openpyxl.Workbook,
    mapping: dict[str, dict[str, Any]],
) -> openpyxl.Workbook:
    """
    Reverse the anonymisation produced by `anonymise_workbook()`.

    Only cell values are touched; all formatting is left intact.
    Tokens not present in *mapping* (e.g. values added by the AI) are kept
    as-is, so nothing is lost.

    Parameters
    ----------
    wb:
        Workbook to restore (will not be mutated – a clone is returned).
    mapping:
        The dict returned by (or loaded from JSON from) `anonymise_workbook()`.

    Returns
    -------
    restored_wb:
        Clone of *wb* with original values reinstated wherever a token is
        found.
    """
    wb = _clone_workbook(wb)
    ws = wb.active

    cols_of_interest = set(mapping.keys())

    # Build column-index → column-name map from the header row.
    idx_to_col: dict[int, str] = {}
    for cell in ws[1]:
        if cell.value in cols_of_interest:
            idx_to_col[cell.column] = str(cell.value)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            col_name = idx_to_col.get(cell.column)
            if col_name is None or cell.value is None:
                continue

            token_map = mapping[col_name]
            # Cast to str for lookup so that Excel's auto-type-conversion
            # (e.g. storing the token as a string) doesn't cause misses.
            original = token_map.get(str(cell.value))
            if original is not None:
                cell.value = original

    return wb


# ---------------------------------------------------------------------------
# DataFrame API  (no formatting – useful for tests and in-app previews)
# ---------------------------------------------------------------------------

def anonymise(
    df: pd.DataFrame,
    columns: list[str],
) -> tuple[pd.DataFrame, dict[str, dict[str, Any]]]:
    """DataFrame-level anonymisation (formatting not preserved)."""
    df = df.copy()
    mapping: dict[str, dict[str, Any]] = {}

    for col in columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in the DataFrame.")

        prefix = _safe_prefix(col)
        col_mapping: dict[str, Any] = {}
        reverse_lookup: dict[Any, str] = {}

        counter = 1
        for val in df[col]:
            if pd.isna(val):
                continue
            if val not in reverse_lookup:
                token = f"{prefix}__{counter:03d}"
                reverse_lookup[val] = token
                col_mapping[token] = val
                counter += 1

        df[col] = df[col].apply(
            lambda v: reverse_lookup[v] if not pd.isna(v) else v
        )
        mapping[col] = col_mapping

    return df, mapping


def restore(
    df: pd.DataFrame,
    mapping: dict[str, dict[str, Any]],
) -> pd.DataFrame:
    """DataFrame-level restoration (formatting not preserved)."""
    df = df.copy()

    for col, token_map in mapping.items():
        if col not in df.columns:
            continue
        df[col] = df[col].apply(
            lambda v: token_map.get(str(v), v) if not pd.isna(v) else v
        )

    return df


# ---------------------------------------------------------------------------
# JSON serialisation helpers
# ---------------------------------------------------------------------------

def mapping_to_json(mapping: dict[str, dict[str, Any]]) -> str:
    """Serialise a mapping dict to a pretty-printed JSON string."""
    return json.dumps(mapping, ensure_ascii=False, indent=2, default=str)


def mapping_from_json(json_str: str) -> dict[str, dict[str, Any]]:
    """Deserialise a mapping dict from a JSON string."""
    return json.loads(json_str)


def mapping_to_file(mapping: dict[str, dict[str, Any]], path: str) -> None:
    """Write a mapping dict to a JSON file."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(mapping_to_json(mapping))


def mapping_from_file(path: str) -> dict[str, dict[str, Any]]:
    """Load a mapping dict from a JSON file."""
    with open(path, encoding="utf-8") as f:
        return json.load(f)
