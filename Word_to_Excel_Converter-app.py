import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Word to Excel Converter", layout="wide")

st.title("📄➡️📊 Word to Excel Converter")

# === File uploader ===
uploaded_file = st.file_uploader("Upload a Word (.docx) file", type=["docx"])

if uploaded_file is not None:
    # Load Word document
    doc = Document(uploaded_file)

    # Extract tables
    all_tables = []
    for i, table in enumerate(doc.tables, start=1):
        data = []
        for row in table.rows:
            data.append([cell.text.strip() for cell in row.cells])
        df = pd.DataFrame(data)
        all_tables.append((f"Table_{i}", df))

    # User choice: merge or separate
    choice = st.radio(
        "How do you want to export the tables?",
        ("Merge all tables into ONE sheet", "Each table in a SEPARATE sheet")
    )

    if choice == "Merge all tables into ONE sheet":
        merged_df = pd.concat([df for _, df in all_tables], ignore_index=True)
        st.subheader("🔎 Preview of merged data")
        st.dataframe(merged_df)

        # Save to Excel (in memory)
        output = BytesIO()
        merged_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="📥 Download Excel (Merged)",
            data=output,
            file_name="Word_to_Excel_Merged.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:  # Separate sheets
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine="openpyxl")

        for sheet_name, df in all_tables:
            df.to_excel(writer, index=False, sheet_name=sheet_name)

        writer.close()
        output.seek(0)

        st.subheader(f"🔎 Preview of {all_tables[0][0]}")
        st.dataframe(all_tables[0][1])  # show preview of first table

        st.download_button(
            label="📥 Download Excel (Separate Sheets)",
            data=output,
            file_name="Word_to_Excel_SeparateSheets.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
