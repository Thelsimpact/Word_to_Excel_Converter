# Word to Excel Extractor

This project extracts **tables from Microsoft Word (`.docx`) files** and saves them into a clean Excel (`.xlsx`) file.  
It is useful for working with government reports, financial tables, or any document that contains structured data.

---

## 🚀 Features
- Extracts all tables from a Word document.
- Merges tables into a single Excel sheet.
- Preserves row and column structure.
- Works on any `.docx` file containing tables.

---

# ⚠️ But note the limitations:

- It only extracts tables. If your Word file has paragraphs, titles, or sections outside of tables, they will be ignored.

- If a document is purely text (no tables), the result will be an empty Excel.

