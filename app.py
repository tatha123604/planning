from __future__ import annotations

from pathlib import Path

import streamlit as st

from processor import build_output_workbook


st.set_page_config(page_title="CLI Matrix Updater", layout="wide")

st.title("CLI Matrix Updater")
st.write(
    "Upload the latest CLI Matrix file and the previous workbook. "
    "The app rebuilds Sheet 1 and Sheet 2, keeps the remaining sheets from the template, "
    "and gives you a downloadable Excel file."
)

source_file = st.file_uploader("Latest CLI Matrix (input file)", type=["xlsx"])
template_file = st.file_uploader("Template / previous workbook (output base)", type=["xlsx"])

default_name = "CLI_Matrix_updated.xlsx"
if source_file is not None and source_file.name:
    source_name = Path(source_file.name).stem
    default_name = f"{source_name}_updated.xlsx"

if st.button("Generate output file", type="primary"):
    if source_file is None or template_file is None:
        st.error("Upload both files first.")
    else:
        try:
            output = build_output_workbook(source_file, template_file, source_file.name)
        except Exception as exc:
            st.error(f"Processing failed: {exc}")
        else:
            st.success("Output file is ready.")
            st.download_button(
                "Download updated workbook",
                data=output.getvalue(),
                file_name=default_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
