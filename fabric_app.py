import pandas as pd
import streamlit as st
from io import BytesIO

# Configure the Streamlit app layout and title
st.set_page_config(page_title="Fabric Order Processor", layout="wide")
st.title("🧵 Fabric Order Processor")

# Allow user to upload a CSV file
uploaded_file = st.file_uploader("Upload a CSV File", type="csv")

# Process the uploaded file if it exists
if uploaded_file:
    # Read the uploaded CSV file
    data = pd.read_csv(uploaded_file, low_memory=False, encoding='utf-8')

    # Select only relevant columns for processing
    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    existing_columns = [col for col in key_columns if col in data.columns]
    data = data[existing_columns]

    # Ensure 'Order #' column is numeric for filtering
    if 'Order #' in data.columns:
        data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')

    # Display original order number range if possible
    if 'Order #' in data.columns:
        original_min_order = data['Order #'].min()
        original_max_order = data['Order #'].max()
        st.write("Original Order Number Range:")
        st.write(f"Min Order #: {original_min_order}")
        st.write(f"Max Order #: {original_max_order}")
    else:
        st.warning("Order # column not found. Cannot calculate original order range.")

    # Clean and filter 'Brand' column to include only valid product types
    if 'Brand' in data.columns:
        data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
        data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]
        st.write(f"Rows after filtering by Brand (FABRIC, BUNDLE, KIT): {len(data)}")
    else:
        st.warning("Brand column not found. Skipping brand filter.")

    # Display filtered order number range if applicable
    if 'Order #' in data.columns and not data.empty:
        filtered_min_order = data['Order #'].min()
        filtered_max_order = data['Order #'].max()
        st.write("Filtered Order Number Range:")
        st.write(f"Min Order #: {filtered_min_order}")
        st.write(f"Max Order #: {filtered_max_order}")

    # Combine duplicate orders by customer and SKU, summing quantities
    if all(col in data.columns for col in ['Customer Name', 'Sku', 'Brand', 'Product Name']):
        grouped = data.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
        combined_data = grouped.sort_values(by=['Sku', 'Quantity'], ascending=[True, False])
        st.write("Combined duplicate orders with summed quantities:")
        st.dataframe(combined_data.head())
    else:
        st.warning("Cannot combine data: Required columns are missing.")

    # Create a tally of how many cuts of each quantity are needed per SKU
    cut_tally = combined_data.groupby(['Sku', 'Brand', 'Product Name', 'Quantity']).size().reset_index(name='Count')

    # Pivot the table so each row is a SKU and columns show counts of each quantity
    pivot_table = cut_tally.pivot_table(index=['Sku', 'Brand', 'Product Name'],
                                        columns='Quantity',
                                        values='Count',
                                        fill_value=0).reset_index()

    # Format column headers (e.g., "1 QTY", "2 QTY", etc.)
    pivot_table.columns.name = None
    pivot_table.columns = [f"{int(col)} QTY" if isinstance(col, (int, float)) else col for col in pivot_table.columns]

    # Convert the final pivot table to CSV format in memory
    output = BytesIO()
    pivot_table.to_csv(output, index=False)

    # Show success message and provide download link
    st.success("✅ File processed successfully!")

    with st.container():
        st.markdown("---")
        st.subheader("📥 Your processed file is ready:")
        st.download_button(
            label="Download Processed CSV File",
            data=output.getvalue(),
            file_name="processed_order_summary.csv",
            mime="text/csv"
        )
