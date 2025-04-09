import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="Fabric Order Processor", layout="wide")
st.title("🧵 Fabric Order Processor")

uploaded_file = st.file_uploader("Upload your CSV file", type="csv")

if uploaded_file:
    data = pd.read_csv(uploaded_file, low_memory=False)

    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    data = data[[col for col in key_columns if col in data.columns]]

    # Clean Order #
    data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')
    min_order = int(data['Order #'].min())
    max_order = int(data['Order #'].max())
    order_range_col = f"Order Range: {min_order} to {max_order}"

    # Standardize Brand
    data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
    data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]

    # Split by brand
    kits = data[data['Brand'] == 'KIT']
    fabrics = data[data['Brand'] == 'FABRIC']
    bundles = data[data['Brand'] == 'BUNDLE']

    # Group fabric & kits
    fabrics_grouped = fabrics.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
    fabrics_grouped['Color'] = ""

    kits_grouped = kits.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name', 'Color'], as_index=False)['Quantity'].sum()
    bundles_grouped = bundles.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
    bundles_grouped['Color'] = ""

    # Combine main data
    main_data = pd.concat([fabrics_grouped, kits_grouped], ignore_index=True)
    main_data = main_data.sort_values(by=['Sku', 'Quantity'])

    # Count cuts
    main_tally = main_data.groupby(['Sku', 'Brand', 'Product Name', 'Color', 'Quantity']).size().reset_index(name='Count')
    bundle_tally = bundles_grouped.groupby(['Sku', 'Brand', 'Product Name', 'Color', 'Quantity']).size().reset_index(name='Count')

    # Pivot
    def pivot_and_format(df, is_bundle=False):
        pivot = df.pivot_table(index=['Sku', 'Brand', 'Product Name', 'Color'],
                               columns='Quantity',
                               values='Count',
                               fill_value=0).reset_index()

        # Rename columns with yd conversion
        new_cols = []
        for col in pivot.columns:
            if isinstance(col, (int, float)):
                if is_bundle:
                    if col == 1:
                        yards = 0.25
                    elif col == 2:
                        yards = 0.5
                    elif col == 4:
                        yards = 1.0
                    else:
                        yards = 0.25 * col
                else:
                    yards = 0.5 * col
                new_cols.append(f"{int(col)} QTY ({yards} yd)")
            else:
                new_cols.append(col)
        pivot.columns = new_cols
        return pivot

    main_pivot = pivot_and_format(main_tally, is_bundle=False)
    bundle_pivot = pivot_and_format(bundle_tally, is_bundle=True)

    # Add Total Yardage
    def add_total_yardage(df, is_bundle=False):
        qty_cols = [col for col in df.columns if "QTY" in col]
        def compute(row):
            total = 0
            for col in qty_cols:
                qty = int(col.split()[0])
                count = row[col]
                if is_bundle:
                    if qty == 1:
                        yards = 0.25
                    elif qty == 2:
                        yards = 0.5
                    elif qty == 4:
                        yards = 1.0
                    else:
                        yards = 0.25 * qty
                else:
                    yards = 0.5 * qty
                total += count * yards
            return total
        df['Total Yardage'] = df.apply(compute, axis=1)
        return df

    main_pivot = add_total_yardage(main_pivot, is_bundle=False)
    bundle_pivot = add_total_yardage(bundle_pivot, is_bundle=True)

    # Reorder columns
    def reorder(df):
        qty_cols = [col for col in df.columns if "QTY" in col]
        return df[['Brand', 'Sku', 'Product Name', 'Color', 'Total Yardage'] + qty_cols]

    main_final = reorder(main_pivot)
    bundle_final = reorder(bundle_pivot)

    main_final[order_range_col] = ""
    bundle_final[order_range_col] = ""

    # Build Excel file — single sheet, merged
    wb = Workbook()
    ws = wb.active
    ws.title = "Fabric + Kits + Bundles"

    # Write Fabric + Kits
    for r_idx, row in enumerate(dataframe_to_rows(main_final, index=False, header=True), 1):
        ws.append(row)
        for cell in ws[r_idx]:
            if r_idx == 1:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Add 5 empty spacer rows
    for _ in range(5):
        ws.append([])

    # Write Bundles below
    start_row = ws.max_row + 1
    for r_idx, row in enumerate(dataframe_to_rows(bundle_final, index=False, header=True), start_row):
        ws.append(row)
        for cell in ws[r_idx]:
            if r_idx == start_row:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Save and export
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
