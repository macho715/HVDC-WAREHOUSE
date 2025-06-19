import pandas as pd
from datetime import datetime
import os
import sys
import subprocess

# --- Configuration (English) ---
os.makedirs("outputs", exist_ok=True)
DEADSTOCK_DAYS = 90 # Dead Stock criteria (days)

# File paths remain the same as they point to physical files
file_map = {
    'HITACHI': 'data/HVDC WAREHOUSE_HITACHI(HE).xlsx',
    'HITACHI_LOCAL': 'data/HVDC WAREHOUSE_HITACHI(HE_LOCAL).xlsx',
    'HITACHI_LOT': 'data/HVDC WAREHOUSE_HITACHI(HE-0214,0252).xlsx',
    'SIEMENS': 'data/HVDC WAREHOUSE_SIMENSE(SIM).xlsx',
}
sheet_name_map = {
    'HITACHI': 'CASE LIST',
    'HITACHI_LOCAL': 'CASE LIST',
    'HITACHI_LOT': 'CASE LIST',
    'SIEMENS': 'CASE LIST',
}
warehouse_cols_map = {
    'HITACHI': ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'Hauler Indoor', 'DSV MZP', 'MOSB'],
    'HITACHI_LOCAL': ['DSV Outdoor', 'DSV Al Markaz', 'DSV MZP', 'MOSB'],
    'HITACHI_LOT': ['DSV Indoor', 'DHL WH', 'DSV Al Markaz', 'AAA Storage'],
    'SIEMENS': ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'MOSB', 'AAA Storage'],
}
site_cols = ['DAS', 'MIR', 'SHU', 'AGI']
target_month = "2025-06"

# --- Classification Data (English) ---
# Defines warehouse types for later classification
indoor_warehouses = {'DSV Indoor', 'Hauler Indoor', 'DSV Al Markaz', 'AAA Storage', 'DHL WH'}
dangerous_warehouses = {'AAA Storage'}

location_type_data = []
for supplier, warehouse_cols in warehouse_cols_map.items():
    for loc in warehouse_cols:
        # Translate '위험' to 'Dangerous'
        classification = 'Dangerous' if loc in dangerous_warehouses else ('Indoor' if loc in indoor_warehouses else 'Outdoor')
        # Translate column names '공급사', '창고', '구분'
        location_type_data.append({'Supplier': supplier, 'Warehouse': loc, 'Classification': classification})
location_type_df = pd.DataFrame(location_type_data)

# --- Data Processing Function (Outputs English DataFrames) ---
def process_supplier_file(excel_path, supplier_name, warehouse_cols, sheet_name):
    """
    Processes a single supplier's file and returns monthly data and final case statuses.
    All DataFrame columns and relevant values are in English.
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except Exception as e:
        print(f" ⚠️ File Error: {excel_path} / {e}")
        return None, None # Return a tuple to match expected return values

    # Ensure 'Quantity' column exists and is numeric
    if 'Quantity' not in df.columns:
        df['Quantity'] = 1
    else:
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(1)

    # Ensure all location columns exist and are datetime objects
    for col in warehouse_cols + site_cols:
        if col not in df.columns: df[col] = pd.NaT
    df[warehouse_cols + site_cols] = df[warehouse_cols + site_cols].apply(pd.to_datetime, errors='coerce')
    
    all_months = sorted(list(set(df[warehouse_cols + site_cols].unstack().dropna().dt.to_period('M'))))
    month_strs = [str(m) for m in all_months]
    
    event_map, case_final_status = [], []

    # Iterate through each row (case) to create a chronological event list
    for idx, row in df.iterrows():
        case, quantity = row.get('Case No.', f"Row_{idx}"), row['Quantity']
        events = []
        for w in warehouse_cols:
            if pd.notna(row[w]): events.append({'date': row[w], 'loc': w, 'type': 'warehouse', 'qty': quantity})
        for s in site_cols:
            if pd.notna(row[s]): events.append({'date': row[s], 'loc': s, 'type': 'site', 'qty': quantity})
        if not events: continue
        events.sort(key=lambda x: x['date'])
        
        # Track the final status for Dead Stock analysis
        last_event = events[-1]
        case_final_status.append({
            'Supplier': supplier_name, 'Case No.': case, 'Final_Location_Type': last_event['type'],
            'Current_Location': last_event['loc'], 'Last_Arrival_Date': last_event['date'], 'Quantity': last_event['qty']
        })
        
        # Process events into In/Out/Site_In types for monthly aggregation
        prev_loc, prev_type = None, None
        for event in events:
            mon = str(event['date'].to_period('M'))
            if event['type'] == 'warehouse':
                event_map.append({'type': 'In', 'loc': event['loc'], 'month': mon, 'quantity': event['qty']})
                prev_loc, prev_type = event['loc'], 'warehouse'
            elif event['type'] == 'site':
                if prev_type == 'warehouse' and prev_loc:
                    event_map.append({'type': 'Out', 'loc': prev_loc, 'month': mon, 'quantity': event['qty']})
                event_map.append({'type': 'Site_In', 'loc': event['loc'], 'month': mon, 'quantity': event['qty']})
                prev_loc, prev_type = None, 'site' # Item is now at a site
    
    # Aggregate events into a monthly report
    consolidated_data, warehouse_stock, site_cumulative_in = [], {w: 0 for w in warehouse_cols}, {s: 0 for s in site_cols}
    for m in month_strs:
        if m > target_month: continue
        row_data = {'Month': m, 'Supplier': supplier_name}
        for w in warehouse_cols:
            in_qty = sum(e['quantity'] for e in event_map if e['type'] == 'In' and e['loc'] == w and e['month'] == m)
            out_qty = sum(e['quantity'] for e in event_map if e['type'] == 'Out' and e['loc'] == w and e['month'] == m)
            warehouse_stock[w] += in_qty - out_qty
            row_data[f'{w}_In'], row_data[f'{w}_Out'], row_data[f'{w}_Stock'] = in_qty, out_qty, warehouse_stock[w]
        for s in site_cols:
            in_qty = sum(e['quantity'] for e in event_map if e['type'] == 'Site_In' and e['loc'] == s and e['month'] == m)
            site_cumulative_in[s] += in_qty
            row_data[f'{s}_In'], row_data[f'{s}_Cumulative_In'] = in_qty, site_cumulative_in[s]
        consolidated_data.append(row_data)
        
    return pd.DataFrame(consolidated_data), pd.DataFrame(case_final_status)

def format_excel_sheet(df, writer, sheet_name, is_pivot=False):
    """Utility function to write and format a sheet in the Excel file."""
    # For pivot table, we need to handle multi-index headers
    if is_pivot:
        df.to_excel(writer, sheet_name=sheet_name, index=True)
    else:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    workbook, worksheet = writer.book, writer.sheets[sheet_name]
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
    num_format = workbook.add_format({'num_format': '#,##0'})
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    total_num_format = workbook.add_format({'bold': True, 'num_format': '#,##0', 'fg_color': '#F2F2F2'})
    total_txt_format = workbook.add_format({'bold': True, 'fg_color': '#F2F2F2'})
    
    # Write headers
    if is_pivot:
        # Write index names
        for i, name in enumerate(df.index.names):
            worksheet.write(0, i, name, header_format)
        # Write column headers
        for c, val in enumerate(df.columns.values):
            worksheet.write(0, len(df.index.names) + c, val, header_format)
    else:
        for c, val in enumerate(df.columns.values):
            worksheet.write(0, c, val, header_format)

    # Write data
    # (Simplified formatting for brevity, can be expanded as before)
    for i, col in enumerate(df.columns):
        # Adjust column width
        # This part is tricky for multi-index and left for simplicity.
        # Can be set to a fixed width or calculated more carefully.
        worksheet.set_column(i + (len(df.index.names) if is_pivot else 0), i + (len(df.index.names) if is_pivot else 0), 15)


# --- Main Execution (English) ---
def main():
    all_monthly_data, all_case_statuses = [], []
    print("🚀 Starting data processing...")
    for supplier, path in file_map.items():
        print(f"   - Processing: {supplier}")
        monthly_df, status_df = process_supplier_file(path, supplier, warehouse_cols_map[supplier], sheet_name_map[supplier])
        if monthly_df is not None: all_monthly_data.append(monthly_df)
        if status_df is not None: all_case_statuses.append(status_df)

    if not all_monthly_data:
        print("⚠️ No data to process. Please check file paths and content.")
        return

    # 1. Prepare the detailed monthly status DataFrame
    consolidated_df = pd.concat(all_monthly_data, ignore_index=True)
    summary_rows = []
    for supplier in file_map.keys():
        s_data = consolidated_df[consolidated_df['Supplier'] == supplier]
        if not s_data.empty:
            s_row = {'Month': 'TOTAL', 'Supplier': supplier}
            for col in s_data.columns[2:]: s_row[col] = s_data[col].iloc[-1] if 'Stock' in col or 'Cumulative' in col else s_data[col].sum()
            summary_rows.append(s_row)
    if summary_rows:
        consolidated_df = pd.concat([consolidated_df, pd.DataFrame(summary_rows)], ignore_index=True)
    consolidated_df.sort_values(by=['Supplier', 'Month'], inplace=True, ignore_index=True)

    print("   - Generating summary data...")
    total_rows_df = pd.DataFrame(summary_rows)
    
    # 2. Prepare data for "Overall_Supplier_Summary" sheet
    # (Logic remains the same)
    summary_list = []
    if not total_rows_df.empty:
        for _, row in total_rows_df.iterrows():
            supplier, wh_cols = row['Supplier'], warehouse_cols_map[row['Supplier']]
            summary_list.append({
                'Supplier': supplier,
                'Total Warehouse In': sum(row.get(f'{w}_In', 0) for w in wh_cols),
                'Total Warehouse Out': sum(row.get(f'{w}_Out', 0) for w in wh_cols),
                'Final Warehouse Stock': sum(row.get(f'{w}_Stock', 0) for w in wh_cols),
                'Final Site Cumulative In': sum(row.get(f'{s}_Cumulative_In', 0) for s in site_cols)
            })
    overall_summary_df = pd.DataFrame(summary_list)
    if not overall_summary_df.empty:
        grand_total = overall_summary_df.drop(columns='Supplier').sum()
        grand_total['Supplier'] = 'GRAND TOTAL'
        overall_summary_df = pd.concat([overall_summary_df, pd.DataFrame([grand_total])], ignore_index=True)

    # 3. Prepare data for "Warehouse_Stock_Summary" sheet
    # (Logic remains the same)
    warehouse_summary_list = []
    if not total_rows_df.empty:
        for _, row in total_rows_df.iterrows():
            for warehouse in warehouse_cols_map[row['Supplier']]:
                warehouse_summary_list.append({
                    'Supplier': row['Supplier'], 'Warehouse': warehouse,
                    'Total In': row.get(f'{warehouse}_In', 0),
                    'Total Out': row.get(f'{warehouse}_Out', 0),
                    'Current Stock': row.get(f'{warehouse}_Stock', 0)
                })
    warehouse_summary_df = pd.DataFrame(warehouse_summary_list)
    if not warehouse_summary_df.empty:
        warehouse_summary_df = pd.merge(warehouse_summary_df, location_type_df, on=['Supplier', 'Warehouse'], how='left')
        warehouse_summary_df = warehouse_summary_df[['Supplier', 'Warehouse', 'Classification', 'Total In', 'Total Out', 'Current Stock']]
    
    # 4. NEW: Prepare data for "Pivoted_Monthly_Summary" sheet
    print("   - Generating pivoted summary data...")
    monthly_data_only = consolidated_df[consolidated_df['Month'] != 'TOTAL'].copy()
    
    # Melt the DataFrame from wide to long format
    id_vars = ['Month', 'Supplier']
    value_vars = [col for col in monthly_data_only.columns if '_' in col and 'Cumulative' not in col]
    long_df = monthly_data_only.melt(id_vars=id_vars, value_vars=value_vars, var_name='Location_Metric', value_name='Value')

    # Split 'Location_Metric' into 'Warehouse' and 'Metric'
    long_df[['Warehouse', 'Metric']] = long_df['Location_Metric'].str.rsplit('_', n=1, expand=True)
    long_df.drop(columns='Location_Metric', inplace=True)

    # Merge with classification data to get warehouse type
    pivoted_summary_df = pd.merge(long_df, location_type_df, on=['Supplier', 'Warehouse'], how='left')
    
    # Aggregate by Month, Supplier, Classification, and Metric
    pivoted_summary_df = pivoted_summary_df.groupby(['Month', 'Supplier', 'Classification', 'Metric'])['Value'].sum().reset_index()

    # Create the pivot table
    pivoted_summary_df = pivoted_summary_df.pivot_table(
        index=['Month', 'Classification', 'Metric'],
        columns='Supplier',
        values='Value',
        fill_value=0,
        aggfunc='sum'
    ).sort_index()


    # 5. Prepare data for "DeadStock_Analysis" sheet
    # (Logic remains the same)
    dead_stock_df = pd.DataFrame() 
    if all_case_statuses:
        case_status_df = pd.concat(all_case_statuses, ignore_index=True)
        in_warehouse_df = case_status_df[case_status_df['Final_Location_Type'] == 'warehouse'].copy()
        if not in_warehouse_df.empty:
            in_warehouse_df['Days_Passed'] = (datetime.now() - in_warehouse_df['Last_Arrival_Date']).dt.days
            dead_stock_df = in_warehouse_df[in_warehouse_df['Days_Passed'] >= DEADSTOCK_DAYS].copy()
            dead_stock_df = dead_stock_df[['Supplier', 'Case No.', 'Current_Location', 'Last_Arrival_Date', 'Days_Passed', 'Quantity']]
            dead_stock_df.rename(columns={'Current_Location': 'Warehouse'}, inplace=True)
            dead_stock_df = dead_stock_df.sort_values(by=['Supplier', 'Days_Passed'], ascending=[True, False])

    # --- Save results to a multi-sheet Excel file ---
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"Consolidated_Inventory_Report_{timestamp}.xlsx"
    output_path = os.path.join("outputs", output_filename)

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        format_excel_sheet(consolidated_df, writer, 'Consolidated_Status')
        print("   ✅ 'Consolidated_Status' sheet created.")
        if not overall_summary_df.empty:
            format_excel_sheet(overall_summary_df, writer, 'Overall_Supplier_Summary')
            print("   ✅ 'Overall_Supplier_Summary' sheet created.")
        if not warehouse_summary_df.empty:
            format_excel_sheet(warehouse_summary_df, writer, 'Warehouse_Stock_Summary')
            print("   ✅ 'Warehouse_Stock_Summary' sheet created.")
        if not pivoted_summary_df.empty:
            format_excel_sheet(pivoted_summary_df, writer, 'Pivoted_Monthly_Summary', is_pivot=True)
            print("   ✅ 'Pivoted_Monthly_Summary' sheet created.")
        if not dead_stock_df.empty:
            format_excel_sheet(dead_stock_df, writer, f'DeadStock_Analysis ({DEADSTOCK_DAYS}+ days)')
            print(f"   ✅ 'DeadStock_Analysis' sheet created.")

    print(f"\n📦 '{output_filename}' has been created successfully!")
    try:
        if os.name == 'nt': subprocess.run(['start', output_path], shell=True, check=True)
        elif sys.platform == 'darwin': subprocess.run(['open', output_path], check=True)
        else: subprocess.run(['xdg-open', output_path], check=True)
    except Exception as e:
        print(f"⚠️ Failed to open the Excel file automatically: {e}\nPlease open it manually from the '{output_path}' path.")

if __name__ == '__main__':
    main() 