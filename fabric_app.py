import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill

# Page setup
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
    data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]

    kits = data[data['Brand'] == 'KIT']
    fabrics = data[data['Brand'] == 'FABRIC']
    bundles = data[data['Brand'] == 'BUNDLE']

    fabrics_grouped = fabrics.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
    fabrics_grouped['Color'] = ""

    kits_grouped = kits.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name', 'Color'], as_index=False)['Quantity'].sum()
    bundles_grouped = bundles.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
    bundles_grouped['Color'] = ""

    main_data = pd.concat([fabrics_grouped, kits_grouped], ignore_index=True)
    main_data = main_data.sort_values(by=['Sku', 'Quantity'])

    main_tally = main_data.groupby(['Sku', 'Brand', 'Product Name', 'Color', 'Quantity']).size().reset_index(name='Count')
    bundle_tally = bundles_grouped.groupby(['Sku', 'Brand', 'Product Name', 'Color', 'Quantity']).size().reset_index(name='Count')

    def pivot_and_format(df, is_bundle=False):
        pivot = df.pivot_table(index=['Sku', 'Brand', 'Product Name', 'Color'],
                               columns='Quantity',
                               values='Count',
                               fill_value=0).reset_index()

        new_cols = []
        for col in pivot.columns:
            if isinstance(col, (int, float)):
                if is_bundle:
                    yards = 0.25 * col if col not in [1, 2, 4] else {1: 0.25, 2: 0.5, 4: 1.0}[col]
                else:
                    yards = 0.5 * col
                new_cols.append(f"{int(col)} QTY ({yards} yd)")
            else:
                new_cols.append(col)
        pivot.columns = new_cols

        # Replace 0s with blank strings
        qty_cols = [c for c in pivot.columns if "QTY" in c]
        for col in qty_cols:
            pivot[col] = pivot[col].replace(0, "")
        return pivot

    def add_total_quantity(df):
        qty_cols = [col for col in df.columns if "QTY" in col]
        total_qty = []
        for _, row in df.iterrows():
            total = 0
            for col in qty_cols:
                qty = int(col.split()[0])
                count = row[col]
                count = int(count) if str(count).isdigit() else 0
                total += qty * count
            total_qty.append(total)
        df['Total Quantity'] = total_qty
        return df

    main_pivot = pivot_and_format(main_tally, is_bundle=False)
    bundle_pivot = pivot_and_format(bundle_tally, is_bundle=True)

    main_pivot = add_total_quantity(main_pivot)
    bundle_pivot = add_total_quantity(bundle_pivot)

    def reorder(df):
        qty_cols = [col for col in df.columns if "QTY" in col]
        return df[['Brand', 'Sku', 'Product Name', 'Color', 'Total Quantity'] + qty_cols]

    main_final = reorder(main_pivot)
    bundle_final = reorder(bundle_pivot)

    main_final["Order Range"] = ""
    bundle_final["Order Range"] = ""

    wb = Workbook()
    ws = wb.active
    ws.title = "Fabric + Kits + Bundles"

    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    alt_row_fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")

    def write_dataframe(ws, df, start_row=1):
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
            ws.append(row)
            is_header = (r_idx == start_row)
            for c_idx, cell in enumerate(ws[r_idx], 1):
                if is_header:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.fill = header_fill
                else:
                    if (r_idx - start_row) % 2 == 1:
                        cell.fill = alt_row_fill

    write_dataframe(ws, main_final, start_row=1)
    ws.cell(row=2, column=ws.max_column).value = order_range_text

    for _ in range(5):
        ws.append([])

    bundle_start = ws.max_row + 1
    write_dataframe(ws, bundle_final, start_row=bundle_start)
    ws.cell(row=bundle_start + 1, column=ws.max_column).value = order_range_text

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ File processed successfully!")
    st.download_button(
        label="📥 Download Formatted Excel File",
        data=output,
        file_name="formatted_order_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
