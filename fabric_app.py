import pandas as pd
import streamlit as st
from io import BytesIO

# Configure the Streamlit app layout and title
st.set_page_config(page_title="Fabric Order Processor", layout="wide")
st.title("🧵 Fabric Order Processor")

# Allow user to upload a CSV file
uploaded_file = st.file_uploader("Upload a CSV File", type="csv")

if uploaded_file:
    data = pd.read_csv(uploaded_file, low_memory=False, encoding='utf-8')

    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    existing_columns = [col for col in key_columns if col in data.columns]
    data = data[existing_columns]

    # Ensure 'Order #' is numeric
    if 'Order #' in data.columns:
        data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')
        original_min_order = int(data['Order #'].min())
        original_max_order = int(data['Order #'].max())
    else:
        original_min_order = None
        original_max_order = None

    # Standardize Brand column
    if 'Brand' in data.columns:
        data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
        data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]

    # Separate Kits from others
    kits = data[data['Brand'] == 'KIT']
    non_kits = data[data['Brand'].isin(['FABRIC', 'BUNDLE'])]

    # Group non-kits normally
    if not non_kits.empty:
        non_kits_grouped = non_kits.groupby(
            ['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False
        )['Quantity'].sum()
    else:
        non_kits_grouped = pd.DataFrame(columns=['Customer Name', 'Sku', 'Brand', 'Product Name', 'Quantity'])

    # Group kits using Color to avoid merging different kit types
    if not kits.empty:
        kits_grouped = kits.groupby(
            ['Customer Name', 'Sku', 'Brand', 'Product Name', 'Color'], as_index=False
        )['Quantity'].sum()
        # Drop Color after grouping, since we don't need it in the final output
        kits_grouped = kits_grouped.drop(columns=['Color'])
    else:
        kits_grouped = pd.DataFrame(columns=['Customer Name', 'Sku', 'Brand', 'Product Name', 'Quantity'])

    # Combine grouped results
    combined_data = pd.concat([non_kits_grouped, kits_grouped], ignore_index=True)

    # Sort for consistency
    combined_data = combined_data.sort_values(by=['Sku', 'Quantity'], ascending=[True, False])

    # Create quantity tally
    cut_tally = combined_data.groupby(['Sku', 'Brand', 'Product Name', 'Quantity']).size().reset_index(name='Count')

    # Pivot to wide format
    pivot_table = cut_tally.pivot_table(
        index=['Sku', 'Brand', 'Product Name'],
        columns='Quantity',
        values='Count',
        fill_value=0
    ).reset_index()

    # Rename columns (e.g., "1 QTY", "2 QTY")
    pivot_table.columns.name = None
    pivot_table.columns = [f"{int(col)} QTY" if isinstance(col, (int, float)) else col for col in pivot_table.columns]

    # Reorder columns: Brand, Sku, Product Name, then QTYs
    quantity_cols = [col for col in pivot_table.columns if "QTY" in col]
    pivot_table = pivot_table[['Brand', 'Sku', 'Product Name'] + quantity_cols]

    # Insert original order range row
    if original_min_order is not None and original_max_order is not None:
        header_row = pd.DataFrame(
            [["Original Order Range:", f"{original_min_order} to {original_max_order}"] + [""] * (len(pivot_table.columns) - 2)],
            columns=pivot_table.columns
        )
        final_output = pd.concat([header_row, pivot_table], ignore_index=True)
    else:
        final_output = pivot_table

    # Convert to downloadable CSV
    output = BytesIO()
    final_output.to_csv(output, index=False)

    # --- Download Section ---
    st.success("✅ File processed successfully! Ready to download.")
    st.markdown("---")
    st.subheader("📥 Download Your Processed File")
    st.download_button(
        label="Download Processed CSV File",
        data=output.getvalue(),
        file_name="processed_order_summary.csv",
        mime="text/csv"
    )
