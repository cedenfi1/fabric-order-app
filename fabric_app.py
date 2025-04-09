import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import shutil
from datetime import datetime, timedelta

st.set_page_config(page_title="Fabric Order Processor", layout="wide")
st.title("🧵 Fabric Order Processor")

uploaded_file = st.file_uploader("Upload your CSV file", type="csv")

if uploaded_file:
    data = pd.read_csv(uploaded_file, low_memory=False)
    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    data = data[[col for col in key_columns if col in data.columns]]

    data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')
    min_order = int(data['Order #'].min())
    max_order = int(data['Order #'].max())
    order_range_text = f"{min_order} to {max_order}"

    data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
    data = data[data['Brand'].isin(['FABRIC', 'KIT'])]

    kits = data[data['Brand'] == 'KIT']
    fabrics = data[data['Brand'] == 'FABRIC']

    fabrics_grouped = fabrics.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
    fabrics_grouped['Color'] = ""

    kits_grouped = kits.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name', 'Color'], as_index=False)['Quantity'].sum()

    main_data = pd.concat([fabrics_grouped, kits_grouped], ignore_index=True)
    main_data = main_data.sort_values(by=['Sku', 'Quantity'])

    main_tally = main_data.groupby(['Sku', 'Brand', 'Product Name', 'Color', 'Quantity']).size().reset_index(name='Count')

    pivot = main_tally.pivot_table(index=['Sku', 'Brand', 'Product Name', 'Color'],
                                   columns='Quantity',
                                   values='Count',
                                   fill_value=0).reset_index()

    new_cols = []
    for col in pivot.columns:
        if isinstance(col, (int, float)):
            yards = 0.5 * col
            label = f"{int(col)} QTY ({yards} yd)"
            new_cols.append(label)
        else:
            new_cols.append(col)
    pivot.columns = new_cols

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

    main_final = pivot[['Sku', 'Product Name', 'Color', 'Total Quantity'] + qty_cols]

    # Load template, copy to output
    template_path = "Cut Sheet Template (1).xlsx"
    output_path = "cut_sheet_output.xlsx"
    shutil.copy(template_path, output_path)
    wb = load_workbook(output_path)
    ws = wb.active

    # Insert "Date of Sale" (H2:I2)
    sale_date = (datetime.now() - timedelta(days=2)).strftime("%m/%d/%Y")
    ws.merge_cells("H2:I2")
    cell = ws["H2"]
    cell.value = f"Date of Sale\n{sale_date}"
    cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

    # Insert Order Range (L2:M2)
    ws.merge_cells("L2:M2")
    order_cell = ws["L2"]
    order_cell.value = f"Order Range: {order_range_text}"
    order_cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

    # Write data starting from row 3
    for i, row in main_final.iterrows():
        base_row = 3 + i
        ws.cell(row=base_row, column=1, value=row["Sku"])
        ws.cell(row=base_row, column=2, value=row["Product Name"])
        ws.cell(row=base_row, column=3, value=row["Color"])
        ws.cell(row=base_row, column=4, value=row["Total Quantity"])
        for j, qty_col in enumerate(qty_cols):
            val = row[qty_col]
            if val != 0:
                ws.cell(row=base_row, column=5 + j, value=val)

    # Final output
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
