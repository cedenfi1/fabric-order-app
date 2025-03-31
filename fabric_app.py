import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

# Configure Streamlit
st.set_page_config(page_title="Fabric Order Processor", layout="wide")
st.title("🧵 Fabric Order Processor")

uploaded_file = st.file_uploader("Upload a CSV File", type="csv")

if uploaded_file:
    data = pd.read_csv(uploaded_file, low_memory=False, encoding='utf-8')

    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    data = data[[col for col in key_columns if col in data.columns]]

    if 'Order #' in data.columns:
        data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')
        original_min_order = int(data['Order #'].min())
        original_max_order = int(data['Order #'].max())
        order_range_col_name = f"Order Range: {original_min_order} to {original_max_order}"
    else:
        order_range_col_name = "Order Range"

    if 'Brand' in data.columns:
        data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
        data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]

    kits = data[data['Brand'] == 'KIT']
    non_kits = data[data['Brand'].isin(['FABRIC', 'BUNDLE'])]

    if not non_kits.empty:
        non_kits_grouped = non_kits.groupby(
            ['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False
        )['Quantity'].sum()
    else:
        non_kits_grouped = pd.DataFrame(columns=['Customer Name', 'Sku', 'Brand', 'Product Name', 'Quantity'])

    if not kits.empty:
        kits_grouped = kits.groupby(
            ['Customer Name', 'Sku', 'Brand', 'Product Name', 'Color'], as_index=False
        )['Quantity'].sum().drop(columns=['Color'])
    else:
        kits_grouped = pd.DataFrame(columns=['Customer Name', 'Sku', 'Brand', 'Product Name', 'Quantity'])

    combined_data = pd.concat([non_kits_grouped, kits_grouped], ignore_index=True)
    combined_data = combined_data.sort_values(by=['Sku', 'Quantity'], ascending=[True, False])

    cut_tally = combined_data.groupby(['Sku', 'Brand', 'Product Name', 'Quantity']).size().reset_index(name='Count')

    pivot_table = cut_tally.pivot_table(
        index=['Sku', 'Brand', 'Product Name'],
        columns='Quantity',
        values='Count',
        fill_value=0
    ).reset_index()

    pivot_table.columns.name = None
    renamed_columns = []
    for col in pivot_table.columns:
        if isinstance(col, (int, float)):
            yards = 0.5 * int(col)
            renamed_columns.append(f"{int(col)} QTY ({yards} yd)")
        else:
            renamed_columns.append(col)
    pivot_table.columns = renamed_columns

    quantity_cols = [col for col in pivot_table.columns if "QTY" in col]
    pivot_table = pivot_table[['Brand', 'Sku', 'Product Name'] + quantity_cols]

    def get_total_yards(row):
        total = 0
        for col in quantity_cols:
            qty_num = int(col.split()[0])
            count = row[col]
            total += count * (0.5 * qty_num)
        return total

    pivot_table['Total Yardage'] = pivot_table.apply(get_total_yards, axis=1)
    final_columns = ['Brand', 'Sku', 'Product Name', 'Total Yardage'] + quantity_cols
    pivot_table = pivot_table[final_columns]

    pivot_table[order_range_col_name] = ""

    # 🧾 Create Excel file with formatting
    wb = Workbook()
    ws = wb.active
    ws.title = "Order Summary"

    for r_idx, row in enumerate(dataframe_to_rows(pivot_table, index=False, header=True), 1):
        ws.append(row)
        for c_idx, cell in enumerate(ws[r_idx], 1):
            # Format header row
            if r_idx == 1:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif c_idx == 3:  # Product Name
                cell.alignment = Alignment(wrap_text=True)

    # Adjust column widths
    for column_cells in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
        adjusted_width = max_length + 2
        ws.column_dimensions[column_cells[0].column_letter].width = min(adjusted_width, 50)

    # Save to BytesIO
    excel_data = BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)

    # Download section
    st.success("✅ File processed and formatted!")
    st.markdown("---")
    st.subheader("📥 Download Your Formatted Excel File")
    st.download_button(
        label="Download .xlsx File",
        data=excel_data,
        file_name="formatted_order_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
