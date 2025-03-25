import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Fabric Order Processor", layout="wide")
st.title("🧵 Fabric Order Processor")

uploaded_file = st.file_uploader("Upload a CSV File", type="csv")

if uploaded_file:
    data = pd.read_csv(uploaded_file, low_memory=False, encoding='utf-8')

    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    existing_columns = [col for col in key_columns if col in data.columns]
    data = data[existing_columns]

    data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
    data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]

    data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')

    grouped = data.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
    grouped = grouped.sort_values(by=['Sku', 'Quantity'], ascending=[True, False])

    cut_tally = grouped.groupby(['Sku', 'Brand', 'Product Name', 'Quantity']).size().reset_index(name='Count')

    pivot_table = cut_tally.pivot_table(index=['Sku', 'Brand', 'Product Name'],
                                        columns='Quantity',
                                        values='Count',
                                        fill_value=0).reset_index()

    pivot_table.columns.name = None
    pivot_table.columns = [f"{int(col)} QTY" if isinstance(col, (int, float)) else col for col in pivot_table.columns]

    output = BytesIO()
    pivot_table.to_csv(output, index=False)

    st.success("File processed successfully!")
    st.download_button(
        label="📥 Download Processed CSV",
        data=output.getvalue(),
        file_name="processed_order_summary.csv",
        mime="text/csv"
    )
