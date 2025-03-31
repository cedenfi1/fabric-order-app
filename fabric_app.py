import pandas as pd
import streamlit as st
from io import BytesIO

# Configure the Streamlit app layout and title
st.set_page_config(page_title="Fabric Order Processor", layout="wide")
st.title("🧵 Fabric Order Processor")

# File uploader
uploaded_file = st.file_uploader("Upload a CSV File", type="csv")

if uploaded_file:
    data = pd.read_csv(uploaded_file, low_memory=False, encoding='utf-8')

    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    data = data[[col for col in key_columns if col in data.columns]]

    # Parse order numbers
    if 'Order #' in data.columns:
        data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')
        original_min_order = int(data['Order #'].min())
        original_max_order = int(data['Order #'].max())
        order_range_col_name = f"Order Range: {original_min_order} to {original_max_order}"
    else:
        order_range_col_name = "Order Range"

    # Filter valid brands
    if 'Brand' in data.columns:
        data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
        data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]

    # Split kits from others
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

    # Merge
    combined_data = pd.concat([non_kits_grouped, kits_grouped], ignore_index=True)
    combined_data = combined_data.sort_values(by=['Sku', 'Quantity'], ascending=[True, False])

    # Cut Tally
    cut_tally = combined_data.groupby(['Sku', 'Brand', 'Product Name', 'Quantity']).size().reset_index(name='Count')

    # Pivot to wide format
    pivot_table = cut_tally.pivot_table(
        index=['Sku', 'Brand', 'Product Name'],
        columns='Quantity',
        values='Count',
        fill_value=0
    ).reset_index()

    # Rename QTY columns (e.g., "2 QTY (1 yd)")
    pivot_table.columns.name = None
    renamed_columns = []
    for col in pivot_table.columns:
        if isinstance(col, (int, float)):
            yards = 0.5 * int(col)
            renamed_columns.append(f"{int(col)} QTY ({yards} yd)")
        else:
            renamed_columns.append(col)
    pivot_table.columns = renamed_columns

    # Reorder columns: Brand, Sku, Product Name, Total Yardage, QTYs
    quantity_cols = [col for col in pivot_table.columns if "QTY" in col]
    pivot_table = pivot_table[['Brand', 'Sku', 'Product Name'] + quantity_cols]

    # Calculate total yardage
    def get_total_yards(row):
        total = 0
        for col in quantity_cols:
            qty_num = int(col.split()[0])  # from "2 QTY (1 yd)"
            count = row[col]
            total += count * (0.5 * qty_num)
        return total

    pivot_table['Total Yardage'] = pivot_table.apply(get_total_yards, axis=1)

    # Final column order with yardage inserted after Product Name
    final_columns = ['Brand', 'Sku', 'Product Name', 'Total Yardage'] + quantity_cols
    pivot_table = pivot_table[final_columns]

    # Add the Order Range column header at the end
    pivot_table[order_range_col_name] = ""

    # Convert to CSV
    output = BytesIO()
    pivot_table.to_csv(output, index=False)

    # --- Download Section ---
    st.success("✅ File processed successfully!")
    st.markdown("---")
    st.subheader("📥 Download Your Processed File")
    st.download_button(
        label="Download Processed CSV File",
        data=output.getvalue(),
        file_name="processed_order_summary.csv",
        mime="text/csv"
    )
