"""
app.py  –  Excel Anonymiser & Restorer
=======================================
Run with:
    streamlit run app.py
"""

from __future__ import annotations

import io
import zipfile
from typing import Any

import openpyxl
import streamlit as st

from encryptor import (
    anonymise_workbook,
    mapping_from_json,
    mapping_to_json,
    restore_workbook,
    wb_to_bytes,
    wb_to_dataframe,
)

# ---------------------------------------------------------------------------
# page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Excel Anonymiser",
    page_icon="🔒",
    layout="wide",
)

st.title("🔒 Excel Anonymiser & Restorer")
st.caption(
    "Anonymise sensitive columns before sending an Excel file to an AI chatbot, "
    "then restore the original values in the processed result."
)

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def load_workbook_from_upload(uploaded_file: Any) -> openpyxl.Workbook:
    return openpyxl.load_workbook(io.BytesIO(uploaded_file.read()))


def make_zip(excel_bytes: bytes, mapping_json: str) -> bytes:
    """Bundle the anonymised Excel and the mapping JSON into a single ZIP."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("anonymised.xlsx", excel_bytes)
        zf.writestr("mapping_key.json", mapping_json.encode("utf-8"))
    return buf.getvalue()


def guess_categorical(wb: openpyxl.Workbook) -> list[str]:
    """
    Return column names from the header row that look categorical.
    A column is considered categorical if at least one of its non-empty
    data cells contains a string value.
    """
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    categorical = []
    for col_idx, header in enumerate(headers, start=1):
        if header is None:
            continue
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            val = row[0].value
            if isinstance(val, str):
                categorical.append(str(header))
                break
    return categorical


# ---------------------------------------------------------------------------
# session-state initialisation
# ---------------------------------------------------------------------------

if "anon_zip" not in st.session_state:
    st.session_state.anon_zip = None          # bytes of the ZIP to download
if "anon_preview" not in st.session_state:
    st.session_state.anon_preview = None      # DataFrame for the preview

if "restored_excel" not in st.session_state:
    st.session_state.restored_excel = None    # bytes of the restored Excel
if "restored_preview" not in st.session_state:
    st.session_state.restored_preview = None  # DataFrame for the preview

# ---------------------------------------------------------------------------
# tabs
# ---------------------------------------------------------------------------

tab_anon, tab_restore = st.tabs(["1 · Anonymise", "2 · Restore"])

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 – Anonymise
# ════════════════════════════════════════════════════════════════════════════
with tab_anon:
    st.header("Step 1 · Upload your Excel file")
    uploaded = st.file_uploader(
        "Choose an Excel file (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="upload_anon",
    )

    if uploaded is not None:
        wb_orig = load_workbook_from_upload(uploaded)
        df_orig = wb_to_dataframe(wb_orig)
        all_columns = [str(h) for h in [c.value for c in wb_orig.active[1]] if h is not None]

        st.subheader("Preview")
        st.dataframe(df_orig, use_container_width=True)

        st.subheader("Step 2 · Select columns to anonymise")
        suggested = guess_categorical(wb_orig)

        selected_cols = st.multiselect(
            "Columns to anonymise",
            options=all_columns,
            default=suggested,
            help=(
                "Text/categorical columns are pre-selected. "
                "Numeric columns are left as-is unless you add them here."
            ),
        )

        if not selected_cols:
            st.info("Select at least one column to anonymise.")
        else:
            if st.button("🔐 Anonymise", type="primary"):
                try:
                    wb_anon, mapping = anonymise_workbook(wb_orig, selected_cols)
                    mapping_json = mapping_to_json(mapping)
                    excel_bytes = wb_to_bytes(wb_anon)

                    # Persist results so the download button survives reruns.
                    st.session_state.anon_zip = make_zip(excel_bytes, mapping_json)
                    st.session_state.anon_preview = wb_to_dataframe(wb_anon)

                except Exception as exc:
                    st.error(f"Error during anonymisation: {exc}")
                    st.session_state.anon_zip = None
                    st.session_state.anon_preview = None

    # Rendered outside the button block so it stays visible after clicking.
    if st.session_state.anon_zip is not None:
        st.success("Anonymisation complete!")

        st.subheader("Anonymised preview")
        st.dataframe(st.session_state.anon_preview, use_container_width=True)

        st.download_button(
            label="⬇️ Download anonymised.xlsx + mapping_key.json (ZIP)",
            data=st.session_state.anon_zip,
            file_name="anonymised_package.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )
        st.info(
            "The ZIP contains **anonymised.xlsx** (safe to share with the AI) "
            "and **mapping_key.json** (your secret key — keep it safe). "
            "You will need the JSON file in Tab 2 to restore original values."
        )

# ════════════════════════════════════════════════════════════════════════════
# TAB 2 – Restore
# ════════════════════════════════════════════════════════════════════════════
with tab_restore:
    st.header("Step 1 · Upload the AI-processed Excel file")
    processed_file = st.file_uploader(
        "Choose the processed Excel file (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="upload_restore",
    )

    st.header("Step 2 · Upload the mapping key")
    mapping_file = st.file_uploader(
        "Choose the mapping key file (.json)",
        type=["json"],
        key="upload_mapping",
    )

    if processed_file is not None and mapping_file is not None:
        try:
            wb_proc = load_workbook_from_upload(processed_file)
            mapping: dict[str, dict] = mapping_from_json(
                mapping_file.read().decode("utf-8")
            )

            st.subheader("Processed file preview")
            st.dataframe(wb_to_dataframe(wb_proc), use_container_width=True)

            if st.button("🔓 Restore original values", type="primary"):
                wb_restored = restore_workbook(wb_proc, mapping)
                st.session_state.restored_excel = wb_to_bytes(wb_restored)
                st.session_state.restored_preview = wb_to_dataframe(wb_restored)

        except Exception as exc:
            st.error(f"Error during restoration: {exc}")
            st.session_state.restored_excel = None
            st.session_state.restored_preview = None

    elif processed_file is None and mapping_file is None:
        st.info("Upload both files above to proceed.")
    elif processed_file is None:
        st.warning("Please upload the processed Excel file.")
    else:
        st.warning("Please upload the mapping key JSON file.")

    # Rendered outside the button block so it stays visible after clicking.
    if st.session_state.restored_excel is not None:
        st.success("Restoration complete!")

        st.subheader("Restored preview")
        st.dataframe(st.session_state.restored_preview, use_container_width=True)

        st.download_button(
            label="⬇️ Download restored.xlsx",
            data=st.session_state.restored_excel,
            file_name="restored.xlsx",
            mime=EXCEL_MIME,
            type="primary",
            use_container_width=True,
        )
