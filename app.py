import streamlit as st

from processor import build_output_workbook

st.set_page_config(page_title="CLI Matrix Updater", page_icon="📊", layout="wide")
st.title("CLI Matrix Overdue Updater")
st.write(
    "Upload the latest CLI Matrix (source) and the template workbook. "
    "Click **Generate output file** to refresh Sheet 1 and Sheet 2, then download the updated Excel."
)

col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("Latest CLI Matrix (source)", type=["xlsx"])
with col2:
    template_file = st.file_uploader("Template / previous workbook", type=["xlsx"])

generate = st.button("Generate output file", type="primary")

if generate:
    if not source_file or not template_file:
        st.error("Please upload both files first.")
    else:
        try:
            output, summary_df, master_df = build_output_workbook(
                source_file, template_file
            )
        except Exception as exc:  # pragma: no cover - surfaced in UI
            st.error(f"Processing failed: {exc}")
        else:
            st.success("Output ready. Download the refreshed workbook below.")
            st.download_button(
                label="Download updated Excel",
                data=output.getvalue(),
                file_name="CLI_Matrix_updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            with st.expander("Preview: Summary position of FP OVERDUE (Sheet 1)"):
                st.dataframe(summary_df)

            with st.expander("Preview: All overdue metrics (Sheet 2)"):
                st.dataframe(master_df)
