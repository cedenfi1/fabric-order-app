import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil

st.set_page_config(page_title="Fabric Order Processor", layout="wide")
st.title("🧵 Fabric Order Processor")

uploaded_file = st.file_uploader("Upload your CSV file", type="csv")

if uploaded_file:
    # Load and clean input data
    data = pd.read_csv(uploaded_file, low_memory=False)
    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    data = data[[col for col in key_columns if col in data.columns]]

    data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')
    min_order = int(data['Order #'].min())
    max_order = int(data['Order #'].max())
    order_range_text = f"{min_order} to {max_order}"

    data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
    data = data[data['Brand'].isin(['FABRIC', 'KIT'])]  # BUNDLE removed for now

    kits = data[data['Brand'] == 'KIT']
    fabrics = data[data['Brand'] == 'FABRIC']

    fabrics_grouped = fabrics.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
    fabrics_grouped['Color'] = ""

    kits_grouped = kits.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name', 'Color'], as_index=False)['Quantity'].sum()

    main_data = pd.concat([fabrics_grouped, kits_grouped], ignore_index=True)
    main_data = main_data.sort_values(by=['Sku', 'Quantity'])

    main_tally = main_data.groupby(['Sku', 'Brand', 'Product Name', 'Color', 'Quantity']).size().reset_index(name='Count')

    # Pivot the tally table
    pivot = main_tally.pivot_table(index=['Sku', 'Brand', 'Product Name', 'Color'],
                                   columns='Quantity',
                                   values='Count',
                                   fill_value=0).reset_index()

    # Rename columns (e.g., "1 QTY (0.5 yd)")
    new_cols = []
    for col in pivot.columns:
        if isinstance(col, (int, float)):
            yards = 0.5 * col
            label = f"{int(col)} QTY ({yards} yd)"
            new_cols.append(label)
        else:
            new_cols.append(col)
    pivot.columns = new_cols

    # Calculate total quantity
    qty_cols = [col for col in pivot.columns if "QTY" in col]
    total_qty = []
    for _, row in pivot.iterrows():
        total = 0
        for col in qty_cols:
            qty = int(col.split()[0])
            count = row[col]
            try:
                total += qty * int(count)
            except:
                pass
        total_qty.append(total)
    pivot['Total Quantity'] = total_qty

    # Reorder columns (drop Brand, keep it in memory)
    main_final = pivot[['Sku', 'Product Name', 'Color', 'Total Quantity'] + qty_cols]

    # --- Excel Template Output ---
    template_path = "Cut Sheet Template (1).xlsx"
    output_path = "cut_sheet_output.xlsx"
    shutil.copy(template_path, output_path)

    wb = load_workbook(output_path)
    ws = wb.active

    # Headers
    ws["A4"] = "SKU"
    ws["B4"] = "FABRIC NAME"
    ws["C4"] = "IN STORE KITS"

    start_col = 5  # E column
    for i, qty_col in enumerate(qty_cols):
        col_letter = get_column_letter(start_col + i)
        if "0.5" in qty_col:
            ws[f"{col_letter}4"] = "1/2 YD CUTS"
        else:
            qty_num = qty_col.split()[0]
            ws[f"{col_letter}4"] = f"{qty_num}YD CUTS"

    # Order Range
    ws.merge_cells("L2:M2")
    ws["L2"] = f"Order Range: {order_range_text}"

    # Write data to sheet
    for i, row in main_final.iterrows():
        base_row = 5 + i
        ws.cell(row=base_row, column=1, value=row["Sku"])
        ws.cell(row=base_row, column=2, value=row["Product Name"])
        ws.cell(row=base_row, column=3, value=row["Color"])
        for j, qty_col in enumerate(qty_cols):
            val = row[qty_col]
            if val != 0:
                ws.cell(row=base_row, column=start_col + j, value=val)

    # Save as stream
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ File processed successfully!")
    st.download_button(
        label="📥 Download Filled Cut Sheet",
        data=output,
        file_name="cut_sheet_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
