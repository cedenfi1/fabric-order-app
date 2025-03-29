import pandas as pd
import streamlit as st
from io import BytesIO
import streamlit.components.v1 as components

# Page configuration
st.set_page_config(page_title="Fabric Order Processor", layout="wide")

# --- Persistent Animated Cat and Ball ---
components.html(
    """
    <style>
    .cat {
        position: fixed;
        bottom: 60px;
        left: -120px;
        width: 80px;
        z-index: 1000;
        animation: moveCat 10s linear infinite;
        pointer-events: none;
    }

    .ball {
        position: fixed;
        bottom: 20px;
        left: -60px;
        width: 40px;
        z-index: 999;
        animation: moveBall 10s linear infinite;
        pointer-events: none;
    }

    @keyframes moveCat {
        0% { left: -120px; }
        100% { left: 110%; }
    }

    @keyframes moveBall {
        0% { left: -60px; transform: rotate(0deg); }
        100% { left: 105%; transform: rotate(1080deg); }
    }
    </style>

    <img class="cat" src="https://media.giphy.com/media/JIX9t2j0ZTN9S/giphy.gif" />
    <img class="ball" src="https://upload.wikimedia.org/wikipedia/commons/thumb/0/09/Basketball.png/50px-Basketball.png" />
    """,
    height=0,
)

# --- App Box Styling ---
st.markdown("""
<style>
.app-box {
    background: #ffffffcc;
    width: 90%;
    max-width: 1000px;
    margin: 2rem auto;
    padding: 3rem;
    border-radius: 1rem;
    box-shadow: 0 15px 30px rgba(0, 0, 0, 0.15);
    position: relative;
    z-index: 1;
}
</style>
<div class="app-box">
""", unsafe_allow_html=True)

# --- Title Section ---
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 2rem 0 1rem;
    }
    .main-header h1 {
        font-size: 2.5rem;
        margin-bottom: 0.2rem;
        color: #222;
    }
    .main-header p {
        color: #666;
        font-size: 1.1rem;
        margin-top: 0;
    }
    hr {
        border: none;
        border-top: 1px solid #eee;
        margin: 2rem 0;
    }
</style>
<div class="main-header">
    <h1>🧵 Fabric Order Processor</h1>
    <p>Turn messy CSVs into clean fabric cut summaries</p>
</div>
<hr>
""", unsafe_allow_html=True)

# --- File Upload Section ---
st.subheader("Step 1: Upload Your CSV File")
uploaded_file = st.file_uploader("Select a fabric orders CSV file", type="csv")

if uploaded_file:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("Step 2: Processed Summary")

    # Load and preprocess CSV
    data = pd.read_csv(uploaded_file, low_memory=False, encoding='utf-8')
    key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
    data = data[[col for col in key_columns if col in data.columns]]

    # Ensure 'Order #' is numeric
    if 'Order #' in data.columns:
        data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')

    # Show original order number range
    if 'Order #' in data.columns:
        st.markdown("**Original Order Number Range:**")
        st.write(f"Min Order #: `{int(data['Order #'].min())}`")
        st.write(f"Max Order #: `{int(data['Order #'].max())}`")
    else:
        st.warning("'Order #' column not found.")

    # Filter by Brand
    if 'Brand' in data.columns:
        data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
        data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]
        st.markdown("**Filtered Rows by Brand:**")
        st.write(f"Remaining Rows: `{len(data)}`")
    else:
        st.warning("'Brand' column not found.")

    # Show filtered order number range
    if 'Order #' in data.columns and not data.empty:
        st.markdown("**Filtered Order Number Range:**")
        st.write(f"Min Order #: `{int(data['Order #'].min())}`")
        st.write(f"Max Order #: `{int(data['Order #'].max())}`")

    # Combine duplicates
    if all(col in data.columns for col in ['Customer Name', 'Sku', 'Brand', 'Product Name']):
        grouped = data.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
        combined_data = grouped.sort_values(by=['Sku', 'Quantity'], ascending=[True, False])
    else:
        st.warning("Missing required columns to combine data.")

    # Tally cut counts by SKU and Quantity
    cut_tally = combined_data.groupby(['Sku', 'Brand', 'Product Name', 'Quantity']).size().reset_index(name='Count')
    pivot_table = cut_tally.pivot_table(index=['Sku', 'Brand', 'Product Name'],
                                        columns='Quantity',
                                        values='Count',
                                        fill_value=0).reset_index()

    pivot_table.columns.name = None
    pivot_table.columns = [f"{int(col)} QTY" if isinstance(col, (int, float)) else col for col in pivot_table.columns]

    # Convert to downloadable CSV
    output = BytesIO()
    pivot_table.to_csv(output, index=False)

    # --- Download Section ---
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("Step 3: Download Your Summary")
    st.success("✅ File processed successfully! Ready for download.")

    st.download_button(
        label="📥 Download Processed CSV",
        data=output.getvalue(),
        file_name="processed_order_summary.csv",
        mime="text/csv",
    )

# --- Close the pop-up box no matter what ---
st.markdown("</div>", unsafe_allow_html=True)
