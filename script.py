# Automating vendor comparison

import pandas as pd
import xlsxwriter
import io
import streamlit as st
from xlsxwriter.utility import xl_col_to_name
from funcs import generate_supplier_template, modify_uploaded_file

st.title("Vendor Supplier Comparison")

# st.markdown("<h4><strong>1. How many suppliers would you like to compare?</strong></h4>",
            # unsafe_allow_html=True)

options = list(range(1, 21))

num_suppliers = st.selectbox(
    label = "1. How many suppliers would you like to compare?",
    placeholder = "Select number of suppliers",
    # label_visibility= None,
    options=options,
    # format_func=lambda x: "Select number" if x == "" else str(x)
)

# if num_suppliers != "":
#     num_suppliers = int(num_suppliers)

output = io.BytesIO()

# Create a workbook and worksheet using XlsxWriter

buffer = generate_supplier_template(num_suppliers=num_suppliers, num_rows=100)

st.markdown("<h5><strong>1. Download the template below to add the quotations from different suppliers.</strong></h4>", 
            unsafe_allow_html=True)

st.markdown("""

###### Instructions:

- Fill in the **quotation name**, **correct item codes**, **quantities required**, and **unit prices** for each vendor.
- Do **not** change the any column headers.
- Replace cells with Supplier 1, Supplier 2, etc., with the **actual names** of the suppliers. 
- Select the availability of each item from the dropdowns in the **AVAILABLE** columns.
- Save the file **in Excel format (.xlsx)** before uploading.
- Once completed, return here and upload the file using the uploader below.

""", unsafe_allow_html=True)

st.download_button(
    label="📥 Download Supplier Comparison Template",
    data=buffer,
    file_name="supplier_comparison_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# User uploads file

st.markdown("<h5><strong>2. Upload your completed Excel file below.</strong></h4>", 
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

st.markdown("<h5><strong>3. Enter supplier names (all upper case and must match with the names in the uploaded file)</strong></h4>", 
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
st.markdown(f"<h5><strong>Added suppliers (ensure these are correct):</strong>",unsafe_allow_html=True)
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
# input_suppliers_lower = set(s.lower() for s in supplier_names_input)

# print(uploaded_file)

# file = pd.read_excel(uploaded_file, sheet_name='Supplier Quotation', header=[1,2])
# print(file)

# new_columns = []
# for col in template.columns:
#     if isinstance(col, tuple) and col[0].lower() in input_suppliers_lower:
#         new_columns.append('_'.join(col).strip())  # e.g., HERMES_UP
#     else:
#         new_columns.append(col[0] if isinstance(col, tuple) else col)  # e.g., ITEM CODE, QTY
# template.columns = new_columns

# print(template.columns)

#TODO next to enable functionality to work with merged supplier header because of added availability columns and highlighted functionality for yellow for unavailable products    

if uploaded_file is not None and st.session_state.names:
    try:
        file = pd.read_excel(uploaded_file, sheet_name='Supplier Quotation', header=[1, 2])
        modified_df, excel_buffer = modify_uploaded_file(supplier_names=supplier_names_input, uploaded_file=file)

        st.download_button(
            label="📥 Download Quotation with Highlights",
            data=excel_buffer,
            file_name="highlighted_quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error processing the file: {e}")

# modified_df, excel_buffer = modify_uploaded_file(supplier_names=supplier_names_input, uploaded_file=file)

# st.download_button(
#     label="Download Quotation with Highlights",
#     data=excel_buffer,
#     file_name="highlighted_quotation.xlsx",
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# )



