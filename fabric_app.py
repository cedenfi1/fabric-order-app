# Reload the newly uploaded file for processing
data = pd.read_csv("/mnt/data/report-type11-2025-03-31.csv", low_memory=False)

# Key columns
key_columns = ['Order #', 'Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']
data = data[[col for col in key_columns if col in data.columns]]

# Clean up columns
data['Order #'] = pd.to_numeric(data['Order #'], errors='coerce')
original_min_order = int(data['Order #'].min())
original_max_order = int(data['Order #'].max())
order_range_col_name = f"Order Range: {original_min_order} to {original_max_order}"

# Standardize Brand
data['Brand'] = data['Brand'].astype(str).str.upper().str.strip()
data = data[data['Brand'].isin(['FABRIC', 'BUNDLE', 'KIT'])]

# Separate by brand
kits = data[data['Brand'] == 'KIT']
fabrics = data[data['Brand'] == 'FABRIC']
bundles = data[data['Brand'] == 'BUNDLE']

# Grouping
kits_grouped = kits.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name', 'Color'], as_index=False)['Quantity'].sum()
kits_grouped = kits_grouped[['Customer Name', 'Sku', 'Brand', 'Product Name', 'Color', 'Quantity']]

fabrics_grouped = fabrics.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
fabrics_grouped['Color'] = ""

bundles_grouped = bundles.groupby(['Customer Name', 'Sku', 'Brand', 'Product Name'], as_index=False)['Quantity'].sum()
bundles_grouped['Color'] = ""

# Combine fabric + kits
main_combined = pd.concat([fabrics_grouped, kits_grouped], ignore_index=True)
main_combined = main_combined.sort_values(by=['Sku', 'Quantity'], ascending=[True, False])

# Bundle handled separately
bundles_combined = bundles_grouped.sort_values(by=['Sku', 'Quantity'], ascending=[True, False])

# Build cut tallies for both
main_tally = main_combined.groupby(['Sku', 'Brand', 'Product Name', 'Color', 'Quantity']).size().reset_index(name='Count')
bundle_tally = bundles_combined.groupby(['Sku', 'Brand', 'Product Name', 'Color', 'Quantity']).size().reset_index(name='Count')

# Pivot tables
main_pivot = main_tally.pivot_table(index=['Sku', 'Brand', 'Product Name', 'Color'], columns='Quantity', values='Count', fill_value=0).reset_index()
bundle_pivot = bundle_tally.pivot_table(index=['Sku', 'Brand', 'Product Name', 'Color'], columns='Quantity', values='Count', fill_value=0).reset_index()

# Rename columns
def rename_columns(pivot_df, bundle_mode=False):
    new_columns = []
    for col in pivot_df.columns:
        if isinstance(col, (int, float)):
            if bundle_mode:
                if int(col) == 1:
                    yards = 0.25
                elif int(col) == 2:
                    yards = 0.5
                elif int(col) == 4:
                    yards = 1.0
                else:
                    yards = 0.25 * int(col)
            else:
                yards = 0.5 * int(col)
            new_columns.append(f"{int(col)} QTY ({yards} yd)")
        else:
            new_columns.append(col)
    pivot_df.columns = new_columns
    return pivot_df

main_pivot = rename_columns(main_pivot, bundle_mode=False)
bundle_pivot = rename_columns(bundle_pivot, bundle_mode=True)

# Add total yardage
def add_total_yardage(df, bundle_mode=False):
    qty_cols = [col for col in df.columns if "QTY" in col]
    def get_total(row):
        total = 0
        for col in qty_cols:
            qty = int(col.split()[0])
            count = row[col]
            if bundle_mode:
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
    df['Total Yardage'] = df.apply(get_total, axis=1)
    return df

main_pivot = add_total_yardage(main_pivot, bundle_mode=False)
bundle_pivot = add_total_yardage(bundle_pivot, bundle_mode=True)

# Final column order
def reorder_columns(df):
    qty_cols = [col for col in df.columns if "QTY" in col]
    ordered = ['Brand', 'Sku', 'Product Name', 'Color', 'Total Yardage'] + qty_cols
    return df[ordered]

main_final = reorder_columns(main_pivot)
bundle_final = reorder_columns(bundle_pivot)

# Add order range column header
main_final[order_range_col_name] = ""
bundle_final[order_range_col_name] = ""

# Output Excel
wb = Workbook()
ws = wb.active
ws.title = "Fabric + Kits"

for r_idx, row in enumerate(dataframe_to_rows(main_final, index=False, header=True), 1):
    ws.append(row)
    for cell in ws[r_idx]:
        if r_idx == 1:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# New sheet for Bundles
ws2 = wb.create_sheet("Bundles")

for r_idx, row in enumerate(dataframe_to_rows(bundle_final, index=False, header=True), 1):
    ws2.append(row)
    for cell in ws2[r_idx]:
        if r_idx == 1:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Save
excel_output = BytesIO()
wb.save(excel_output)
excel_output.seek(0)
excel_output.name = "formatted_order_summary.xlsx"

import ace_tools as tools; tools.download_file(excel_output, "Fabric + Kits + Bundles Summary")
