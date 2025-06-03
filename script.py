# Automating vendor comparison

import pandas as pd
import xlsxwriter
import io
import streamlit as st
from xlsxwriter.utility import xl_col_to_name

st.title("Vendor Supplier Comparison")

st.markdown("<h4><strong>1. How many suppliers would you like to compare?</strong></h4>",
            unsafe_allow_html=True)

options = [""] + list(range(1, 21))

num_suppliers = st.selectbox(
    label="",
    options=options,
    format_func=lambda x: "Select number" if x == "" else str(x)
)

if num_suppliers != "":
    num_suppliers = int(num_suppliers)

output = io.BytesIO()

# Create a workbook and worksheet using XlsxWriter
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Supplier Quotation")
    writer.sheets["Supplier Quotation"] = worksheet

    # Formats
    bold_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    bold_left = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1})

    # Merge for "QUOTATION NAME"
    worksheet.write('A1', 'QUOTATION NAME:', bold_left)

    # Header row 2
    headers_base = ['ITEM CODE', 'DESCRIPTION', 'QTY']
    supplier_headers = [f"Supplier {i + 1}" for i in range(num_suppliers)]

    # Merge supplier header
    worksheet.merge_range(0, 3, 0, 3 + num_suppliers - 1, 'Suppliers', bold_center)

    # Row 3: Column headers
    for col, header in enumerate(headers_base + supplier_headers):
        worksheet.write(1, col, header, bold_center)

    for i in range(num_suppliers):
        worksheet.write(2, 3 + i, "UP", bold_center)        

    # Set column widths (optional)
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 25)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:Z", 18)

    writer.close()
    excel_data = output.getvalue()

st.markdown("<h4><strong>2. Download the template below to add the quotations from different suppliers.</strong></h4>", 
            unsafe_allow_html=True)

st.download_button(
    label="📥 Download Supplier Comparison Template",
    data=excel_data,
    file_name="supplier_comparison_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("""
---

### Instructions:

- Fill in the **quotation name**, **correct item codes**, **quantities required**, and **unit prices** for each vendor.
- Do **not** change the any column headers.
- Replace cells with Supplier 1, Supplier 2, etc., with the **actual names** of the suppliers. 
- Save the file **in Excel format (.xlsx)** before uploading.
- Once completed, return here and upload the file using the uploader below.

---
""", unsafe_allow_html=True)

# User uploads file

st.markdown("<h4><strong>3. Upload your completed Excel file below.</strong></h4>", 
            unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    label="Upload here", 
    type=["xlsx"]
)

#TODO add check that correct file (no changed headers, etc.) is uploaded or throw an error below

if uploaded_file is not None:
    # Try to read the file
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ File uploaded successfully!")
        st.write("Preview of uploaded data:")
        st.dataframe(df)
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")

# Initialize session state
if 'names' not in st.session_state:
    st.session_state.names = []

if 'input_text' not in st.session_state:
    st.session_state.input_text = ""

st.markdown("<h4><strong>4. Enter supplier names (must match with the names in the uploaded file)</strong></h4>", 
            unsafe_allow_html=True)

multi_input = st.text_area(
    label="",  # no label here
    value=st.session_state.input_text,
    key="name_input_area"
)

# Add names
if st.button("Add Names"):
    if multi_input:
        new_names = [name.strip() for part in multi_input.splitlines() for name in part.split(",") if name.strip()]
        
        for name in new_names:
            if name not in st.session_state.names:
                st.session_state.names.append(name)
        
        # Clear input
        st.session_state.input_text = ""
        st.rerun()
    else:
        st.warning("Please enter at least one name.")

#TODO add check that number of suppliers given matches number of suppliers selected in the first step

# Display added names with remove buttons
st.markdown(f"<h4><strong>Added suppliers:</strong>",unsafe_allow_html=True)
for i, name in enumerate(st.session_state.names):
    col1, col2 = st.columns([5, 1])
    with col1:
        st.success(name)
    with col2:
        if st.button("❌", key=f"remove_{i}"):
            st.session_state.names.pop(i)
            st.rerun()

# Start processing the doc
supplier_names_input = st.session_state.names
input_suppliers_lower = set(s.lower() for s in supplier_names_input)

# print(uploaded_file)

file = pd.read_excel(uploaded_file, sheet_name='Supplier Quotation', header=[1,2])
print(file)

# new_columns = []
# for col in template.columns:
#     if isinstance(col, tuple) and col[0].lower() in input_suppliers_lower:
#         new_columns.append('_'.join(col).strip())  # e.g., HERMES_UP
#     else:
#         new_columns.append(col[0] if isinstance(col, tuple) else col)  # e.g., ITEM CODE, QTY
# template.columns = new_columns

# print(template.columns)

def modify_uploaded_file(uploaded_file, supplier_names):
    """
    Args:
    uploaded_file: DataFrame containing the uploaded Excel file.
    supplier_names: List of supplier names to be processed (#TODO: make this input flexy as needed)

    Modifies the uploaded excel file
    - finds unit price columns
    - adds total price columns for each supplier
    - add summary row
    - Highlights lowest unit prices per row and lowest total in summary


    """

    # 1. Make unit price columns and total price columns

    new_columns = []
    for col in uploaded_file.columns:

        if isinstance(col, tuple) and col[0].lower() in input_suppliers_lower:
            new_columns.append('_'.join(col).strip())  # e.g., HERMES_UP
        else:
            new_columns.append(col[0] if isinstance(col, tuple) else col)  # e.g., ITEM CODE, QTY
    uploaded_file.columns = new_columns

    for supplier in supplier_names:
        supplier_col = f"{supplier}_UP"
        col_index = uploaded_file.columns.get_loc(supplier_col)
        if supplier_col in uploaded_file.columns:
            total_values = uploaded_file[supplier_col] * file['QTY']
            uploaded_file.insert(loc=col_index + 1, column=f"{supplier}_TOTAL", value=total_values)
        else:
            print(f"Warning: Column {supplier_col} not found in the template.")

    blank_row = pd.DataFrame([{col: "" for col in uploaded_file.columns}])

    # 2: Add summary row for each supplier
    summary_row = {'ITEM CODE': 'TOTAL_QUOTE'}  # you can also label 'DESCRIPTION' or a new column
    for col in uploaded_file.columns:
        if col.endswith("_TOTAL"):
            summary_row[col] = uploaded_file[col].sum()

    summary_row_df = pd.DataFrame([summary_row])

    # 3: Concatenate everything
    final_df = pd.concat([uploaded_file, blank_row, summary_row_df], ignore_index=True)

    # 4: Apply highlighting 
    # for each row, highlight lowest UP per supplier
    # lowest total per supplier
    # rearrange to show supplier with lowest total cols after qty
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Quotation')
        workbook = writer.book
        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})  # light green fill, dark green text
        worksheet = writer.sheets['Quotation']

        # Highlight lowest UP per row
    # green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})

# Highlight lowest UP per row
        up_cols = [f"{supplier}_UP" for supplier in supplier_names if f"{supplier}_UP" in final_df.columns]
        for row in range(1, len(uploaded_file) + 1):  # Data rows only
            col_letters = [xl_col_to_name(final_df.columns.get_loc(col)) for col in up_cols]
            row_num = row + 1  # 1-based Excel row
            if col_letters:
                range_expr = ",".join(f"{letter}{row_num}" for letter in col_letters)
                formula = f"MIN({range_expr})"
                for letter in col_letters:
                    cell = f"{letter}{row_num}"
                    worksheet.conditional_format(cell, {
                        'type': 'formula',
                        'criteria': f"{cell}={formula}",
                        'format': green_format
                    })

        # Highlight lowest total in summary row
        total_cols = [f"{supplier}_TOTAL" for supplier in supplier_names if f"{supplier}_TOTAL" in final_df.columns]
        summary_row_index = len(final_df) + 1  # 1-based
        summary_letters = [xl_col_to_name(final_df.columns.get_loc(col)) for col in total_cols]
        if summary_letters:
            range_expr = ",".join(f"{letter}{summary_row_index}" for letter in summary_letters)
            formula = f"MIN({range_expr})"
            for letter in summary_letters:
                cell = f"{letter}{summary_row_index}"
                worksheet.conditional_format(cell, {
                    'type': 'formula',
                    'criteria': f"{cell}={formula}",
                    'format': green_format
                })

        output_buffer.seek(0)
    return final_df, output_buffer    

modified_df, excel_buffer = modify_uploaded_file(supplier_names=supplier_names_input, uploaded_file=file)

st.dataframe(modified_df)

st.download_button(
    label="Download Quotation with Highlights",
    data=excel_buffer,
    file_name="highlighted_quotation.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)



