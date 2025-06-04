import pandas as pd
import xlsxwriter
import io
import streamlit as st
from xlsxwriter.utility import xl_col_to_name

def generate_supplier_template(num_suppliers: int = 1, num_rows: int = 100):
    output = io.BytesIO()

    # Build empty DataFrame with required structure
    headers_base = ['ITEM CODE', 'DESCRIPTION', 'QTY']
    supplier_headers = []
    for i in range(num_suppliers):
        supplier_headers.extend([f"Supplier {i + 1}_UP", f"Supplier {i + 1}_AVAILABLE"])

    all_columns = headers_base + supplier_headers
    final_df = pd.DataFrame(columns=all_columns)
    final_df = final_df.reindex(range(num_rows))  # Add empty rows

    # Start Excel writer
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Let pandas create the worksheet
        final_df.to_excel(writer, sheet_name="Supplier Quotation", startrow=3, index=False, header=False)

        workbook  = writer.book
        worksheet = writer.sheets["Supplier Quotation"]

        # Define formats
        bold_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        bold_left   = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1})

        # Row 1: QUOTATION NAME
        worksheet.write('A1', 'QUOTATION NAME:', bold_left)

        # Row 2: base headers
        for col, header in enumerate(headers_base):
            worksheet.write(1, col, header, bold_center)

        # Row 2: merged supplier headers
        for i in range(num_suppliers):
            col_start = 3 + i * 2
            col_end = col_start + 1
            worksheet.merge_range(1, col_start, 1, col_end, f"Supplier {i + 1}", bold_center)

        # Row 3: UP / AVAILABLE
        for i in range(num_suppliers):
            worksheet.write(2, 3 + i * 2, "UP", bold_center)
            worksheet.write(2, 4 + i * 2, "AVAILABLE", bold_center)

        # Row 1: merged "Suppliers"
        worksheet.merge_range(0, 3, 0, 3 + (2 * num_suppliers) - 1, 'Suppliers', bold_center)

        # Column widths
        worksheet.set_column("A:A", 15)
        worksheet.set_column("B:B", 25)
        worksheet.set_column("C:C", 10)
        worksheet.set_column("D:Z", 18)

        # Data validation: dropdown for all AVAILABLE columns
        validation_options = ['YES', 'NO', 'NOT SURE']
        for i in range(num_suppliers):
            available_col_index = 4 + (i * 2)
            col_letter = xl_col_to_name(available_col_index)
            # print(i, col_letter)
            cell_range = f"{col_letter}4:{col_letter}{3 + num_rows}"  # 1-based row numbers in Excel

            worksheet.data_validation(cell_range, {
                'validate': 'list',
                'source': validation_options,
                'input_message': 'Choose: YES, NO, or NOT SURE',
                'error_title': 'Invalid Input',
                'error_message': 'Only YES, NO, or NOT SURE are allowed',
                # 'show_error_message': True
            })

    output.seek(0)
    return output

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
    input_suppliers_lower = set(s.lower() for s in supplier_names)
    new_columns = []
    for col in uploaded_file.columns:

        if isinstance(col, tuple) and col[0].lower() in input_suppliers_lower:
            new_columns.append('_'.join(col).strip())  # e.g., HERMES_UP
        else:
            new_columns.append(col[0] if isinstance(col, tuple) else col)  # e.g., ITEM CODE, QTY

    uploaded_file.columns = new_columns
    # print(uploaded_file.columns)

    for supplier in supplier_names:
        up_col = f"{supplier}_UP"
        # avail_col = f"{supplier}_AVAILABLE"

        up_index = uploaded_file.columns.get_loc(up_col)
        # avail_index = uploaded_file.columns.get_loc(avail_col)

        if up_col in uploaded_file.columns:
            total_values = uploaded_file[up_col] * uploaded_file['QTY']
            uploaded_file.insert(loc=up_index + 1, column=f"{supplier}_TOTAL", value=total_values)
        else:
            print(f"Warning: Column {up_col} not found in the template.")

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
        red_format   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})   # Light red
        orange_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})  # Light orange

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

    # Highlighting the availability columns
        avail_cols = [f"{supplier}_AVAILABLE" for supplier in supplier_names if f"{supplier}_AVAILABLE" in final_df.columns]

        for row in range(1, len(uploaded_file) + 1):  # Adjusted for Excel's 1-based indexing
            row_num = row + 1  # Excel row (data starts at row 4 = index 3)
            for col in avail_cols:
                col_letter = xl_col_to_name(final_df.columns.get_loc(col))
                cell = f"{col_letter}{row_num}"
                # print(cell)

                worksheet.conditional_format(cell, {
                    'type': 'cell',
                    'criteria': '==',
                    'value': '"NO"',
                    'format': red_format
                })

                worksheet.conditional_format(cell, {
                    'type': 'cell',
                    'criteria': '==',
                    'value': '"NOT SURE"',
                    'format': orange_format
                })

    return final_df, output_buffer 