from __future__ import annotations
import streamlit as st

from amount_rounder import process_excel, save_outputs

st.set_page_config(page_title="Excel Converter", page_icon="📊", layout="wide")

st.title("📊Balance sheet Amount Converter")
st.caption("Select lakh or thousand, upload your Excel file, and download the converted workbook.")

# --- Only required choice ---
mode = st.radio("Convert amounts into:", ["lakh", "thousand"], index=0, horizontal=True)

# --- Only required action: upload file ---
uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)

if uploaded:
    st.success(f"Loaded file: {uploaded.name}")

    # Process with minimal/default settings; logic lives in amount_rounder.process_excel
    uploaded_bytes = uploaded.read()
    out_bytes, summary = process_excel(
        input_bytes=uploaded_bytes,
        mode=mode,                # "lakh" or "thousand"
        header_row=1,             # keep default behavior for any day rounding detection
        lakh_edge_threshold=50000 # unused for "thousand" and simple "lakh" division, safe default
    )

    # Download button
    st.download_button(
        label="⬇️ Download processed workbook",
        data=out_bytes,
        file_name=f"processed__{uploaded.name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Save a copy + log to backup/
    out_path, log_path = save_outputs(
        output_bytes=out_bytes,
        summary=summary,
        original_filename=uploaded.name,
        out_dir="backup",
        tag=mode,
    )

    st.info("Processing complete. A copy and log were saved to the server.")
else:
    st.warning("Upload an Excel file to begin.")

