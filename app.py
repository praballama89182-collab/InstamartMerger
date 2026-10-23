import os
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(
    page_title="Swiggy Granular Ad & Sales Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
    :root {
        --bg-main: #EDF2F7;
        --card-bg: #FFFFFF;
        --border-color: #CBD5E0;
        --text-main: #2D3748;
        --accent-dark: #1A202C;
        --accent-primary: #3182CE;
    }
    .stApp {
        background: var(--bg-main);
        color: var(--text-main);
    }
    [data-testid="stHeader"] {
        background: rgba(237, 242, 247, 0.95);
    }
    h1, h2, h3 {
        color: var(--accent-dark) !important;
        font-weight: 700 !important;
    }
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #FFFFFF 0%, #F7FAFC 100%);
        border: 1px solid var(--border-color);
        border-left: 5px solid var(--accent-primary);
        border-radius: 12px;
        padding: 14px 18px;
        box-shadow: 0 3px 10px rgba(0, 0, 0, 0.04);
    }
    [data-testid="stMetricLabel"] {
        color: #4A5568 !important;
        font-weight: 600 !important;
    }
    [data-testid="stMetricValue"] {
        color: var(--accent-dark) !important;
        font-weight: 750 !important;
    }
    /* Boxed Curved Selection Boxes / Inputs */
    div[data-baseweb="select"] > div, div.stSelectbox > div > div {
        background-color: #FFFFFF !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.03);
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background: #E2E8F0;
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        border-radius: 8px;
        color: #4A5568;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #2D3748 0%, #1A202C 100%) !important;
        color: #FFFFFF !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15);
    }
    [data-testid="stFileUploader"] {
        background: var(--card-bg);
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 10px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
    }
    [data-testid="stFileUploaderDropzone"] {
        background: #F7FAFC;
        border: 1px dashed #A0AEC0;
        border-radius: 10px;
    }
    [data-testid="stFileUploader"] button,
    .stDownloadButton button,
    .stButton button {
        background: linear-gradient(135deg, #2D3748, #1A202C) !important;
        color: #FFFFFF !important;
        border: 0 !important;
        border-radius: 8px !important;
        font-weight: 650 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.10);
    }
    [data-testid="stFileUploader"] button:hover,
    .stDownloadButton button:hover,
    .stButton button:hover {
        background: #1A202C !important;
        color: #FFFFFF !important;
    }
    [data-testid="stDataFrame"] {
        border: 1px solid var(--border-color);
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03);
        background: #FFFFFF;
    }
    [data-testid="stDataFrame"] th,
    [data-testid="stDataFrame"] td {
        text-align: center !important;
        vertical-align: middle !important;
    }
    [data-testid="stDataFrame"] th {
        background: #CBD5E0 !important;
        color: #1A202C !important;
        font-weight: 700 !important;
    }
    hr {
        border-color: #CBD5E0 !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Swiggy Granular Ad Analytics & Overall Sales Merger Dashboard")
st.write("Upload Swiggy granular advertising spreadsheets/CSV files and Overall Sales reports. Analyze TACOS, TROAS, Ad Sales, Overall Sales, ACOS, ROAS, Daily Tracker, and multi-metric trends.")

def style_and_export_pivot(pivot_df, sheet_name="Comparison"):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    
    if pivot_df is None or pivot_df.empty:
        ws.cell(row=1, column=1, value="No data available.")
        wb.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()
    
    top_header_fill = PatternFill(start_color="2D3748", end_color="2D3748", fill_type="solid")
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    sec_header_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="1A202C")
    index_fill = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")
    total_fill = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    data_font = Font(name="Calibri", size=10, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E0'), right=Side(style='thin', color='CBD5E0'),
        top=Side(style='thin', color='CBD5E0'), bottom=Side(style='thin', color='CBD5E0')
    )

    is_daily_tracker = ("Daily_Tracker" in sheet_name or "Daily" in sheet_name)

    if isinstance(pivot_df.columns, pd.MultiIndex):
        metrics = pivot_df.columns.get_level_values(1)
        entity_title = pivot_df.index.name if pivot_df.index.name else ""
        ws.cell(row=1, column=1, value="").fill = top_header_fill
        ws.cell(row=1, column=1).border = thin_border
        
        current_col = 2
        for month in pivot_df.columns.get_level_values(0).unique():
            sub_cols = [c for c in pivot_df.columns if c[0] == month]
            num_sub = len(sub_cols)
            if num_sub > 1:
                ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + num_sub - 1)
            cell = ws.cell(row=1, column=current_col, value=str(month))
            cell.fill = top_header_fill
            cell.font = top_header_font
            cell.alignment = align_center
            for c in range(current_col, current_col + num_sub):
                ws.cell(row=1, column=c).border = thin_border
                ws.cell(row=1, column=c).fill = top_header_fill
            current_col += num_sub

        cell_a2 = ws.cell(row=2, column=1, value=str(entity_title))
        cell_a2.fill = index_fill
        cell_a2.font = index_font
        cell_a2.alignment = align_center
        cell_a2.border = thin_border

        for col_idx, metric in enumerate(metrics, start=2):
            cell = ws.cell(row=2, column=col_idx, value=str(metric))
            cell.fill = sec_header_fill
            cell.font = sec_header_font
            cell.alignment = align_center
            cell.border = thin_border

        start_data_row = 3
        for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=start_data_row):
            is_grand_total = (str(idx_val).strip().lower() in ['grand total', 'aggregate / average'])
            idx_cell = ws.cell(row=r_idx, column=1, value=str(idx_val))
            idx_cell.fill = total_fill if is_grand_total else index_fill
            idx_cell.font = total_font if is_grand_total else index_font
            idx_cell.alignment = align_left if is_daily_tracker else align_center
            idx_cell.border = thin_border

            for c_idx, (m_col, val) in enumerate(zip(metrics, row_data), start=2):
                val_cell = ws.cell(row=r_idx, column=c_idx)
                val_cell.font = total_font if is_grand_total else data_font
                if is_grand_total:
                    val_cell.fill = total_fill
                val_cell.alignment = align_center
                val_cell.border = thin_border
                
                m_upper = str(m_col).upper()
                if isinstance(val, (int, float, np.number)):
                    if m_upper in ['ACOS', 'TACOS', 'T-ACOS']:
                        val_cell.value = val / 100.0 if val > 1 else val
                        val_cell.number_format = '0.00%'
                    elif m_upper in ['SPEND', 'SPENDS', 'SALES', 'AD SALES', 'OV SALES']:
                        val_cell.value = int(round(val))
                        val_cell.number_format = '#,##0'
                    else:
                        val_cell.value = round(val, 2)
                        val_cell.number_format = '#,##0' if float(val).is_integer() else '#,##0.00'
                else:
                    val_cell.value = val
    else:
        for col_idx, col_name in enumerate(pivot_df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=str(col_name))
            cell.fill = top_header_fill
            cell.font = top_header_font
            cell.alignment = align_center
            cell.border = thin_border

        for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=2):
            is_grand_total = (str(idx_val).strip().lower() in ['grand total', 'aggregate / average'] or 'VS' in str(row_data.iloc[0]).strip().upper())
            
            for c_idx, (col_name, val) in enumerate(zip(pivot_df.columns, row_data), start=1):
                val_cell = ws.cell(row=r_idx, column=c_idx)
                val_cell.font = total_font if is_grand_total else data_font
                if is_grand_total:
                    val_cell.fill = total_fill
                val_cell.alignment = align_left if (c_idx == 1 and is_daily_tracker) else align_center
                val_cell.border = thin_border
                
                col_upper = str(col_name).upper()
                if isinstance(val, (int, float, np.number)):
                    if col_upper in ['ACOS', 'TACOS', 'T-ACOS']:
                        val_cell.value = val / 100.0 if val > 1 else val
                        val_cell.number_format = '0.00%'
                    elif col_upper in ['SPEND', 'SPENDS', 'SALES', 'AD SALES', 'OV SALES']:
                        val_cell.value = int(round(val))
                        val_cell.number_format = '#,##0'
                    else:
                        val_cell.value = round(val, 2)
                        val_cell.number_format = '#,##0' if float(val).is_integer() else '#,##0.00'
                else:
                    val_cell.value = val

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def convert_all_pivots_to_excel(pivot_dict):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, pivot_df in pivot_dict.items():
        if pivot_df is not None and not pivot_df.empty:
            sheet_bytes = style_and_export_pivot(pivot_df, sheet_name=sheet_name)
            temp_wb = openpyxl.load_workbook(io.BytesIO(sheet_bytes))
            temp_ws = temp_wb.active
            temp_ws.title = sheet_name[:31]
            wb._sheets.append(temp_ws)
    if not wb._sheets:
        ws = wb.create_sheet(title="No_Data")
        ws.cell(row=1, column=1, value="No data available.")
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

st.markdown("### 📥 Upload Swiggy Ad Campaign Files (Up to 5)")
col1, col2, col3, col4, col5 = st.columns(5)
uploaded_files = []
with col1:
    f1 = st.file_uploader("Upload Ad File 1", type=["csv", "xlsx", "xls"], key="file1")
    if f1: uploaded_files.append(f1)
with col2:
    f2 = st.file_uploader("Upload Ad File 2", type=["csv", "xlsx", "xls"], key="file2")
    if f2: uploaded_files.append(f2)
with col3:
    f3 = st.file_uploader("Upload Ad File 3", type=["csv", "xlsx", "xls"], key="file3")
    if f3: uploaded_files.append(f3)
with col4:
    f4 = st.file_uploader("Upload Ad File 4", type=["csv", "xlsx", "xls"], key="file4")
    if f4: uploaded_files.append(f4)
with col5:
    f5 = st.file_uploader("Upload Ad File 5", type=["csv", "xlsx", "xls"], key="file5")
    if f5: uploaded_files.append(f5)

st.markdown("### 📥 Upload Overall Sales Files (Up to 5)")
scol1, scol2, scol3, scol4, scol5 = st.columns(5)
uploaded_sales_files = []
with scol1:
    s1 = st.file_uploader("Upload Sales File 1", type=["csv", "xlsx", "xls"], key="sales1")
    if s1: uploaded_sales_files.append(s1)
with scol2:
    s2 = st.file_uploader("Upload Sales File 2", type=["csv", "xlsx", "xls"], key="sales2")
    if s2: uploaded_sales_files.append(s2)
with scol3:
    s3 = st.file_uploader("Upload Sales File 3", type=["csv", "xlsx", "xls"], key="sales3")
    if s3: uploaded_sales_files.append(s3)
with scol4:
    s4 = st.file_uploader("Upload Sales File 4", type=["csv", "xlsx", "xls"], key="sales4")
    if s4: uploaded_sales_files.append(s4)
with scol5:
    s5 = st.file_uploader("Upload Sales File 5", type=["csv", "xlsx", "xls"], key="sales5")
    if s5: uploaded_sales_files.append(s5)

# Robust Sales File Ingestion (handling header offsets & metadata sheets)
sales_dfs_list = []
if uploaded_sales_files:
    for s_file in uploaded_sales_files:
        try:
            if s_file.name.lower().endswith('.csv'):
                sdf = pd.read_csv(s_file)
                sales_dfs_list.append(sdf)
            else:
                s_xls = pd.ExcelFile(s_file)
                for s_sheet in s_xls.sheet_names:
                    temp_df = pd.read_excel(s_xls, sheet_name=s_sheet, nrows=5)
                    has_date_col = any(c in temp_df.columns for c in ['ORDERED_DATE', 'Order Date', 'date', 'Date'])
                    if not has_date_col:
                        sdf = pd.read_excel(s_xls, sheet_name=s_sheet, header=1)
                    else:
                        sdf = pd.read_excel(s_xls, sheet_name=s_sheet, header=0)
                    
                    if any(c in sdf.columns for c in ['ORDERED_DATE', 'Order Date', 'date', 'Date', 'GMV', 'UNITS_SOLD']):
                        sales_dfs_list.append(sdf)
        except Exception as e:
            st.error(f"Error reading sales file {s_file.name}: {str(e)}")

master_sales_df = pd.DataFrame()
if sales_dfs_list:
    master_sales_df = pd.concat(sales_dfs_list, ignore_index=True)
    date_cand = next((c for c in ['ORDERED_DATE', 'Order Date', 'date', 'Date'] if c in master_sales_df.columns), None)
    gmv_cand = next((c for c in ['GMV', 'Total Gross Bill Amount', 'Gross Amount', 'Total Amount'] if c in master_sales_df.columns), None)
    qty_cand = next((c for c in ['UNITS_SOLD', 'Quantity', 'Qty', 'Units'] if c in master_sales_df.columns), None)
    
    if date_cand:
        master_sales_df['_sales_dt'] = pd.to_datetime(master_sales_df[date_cand], errors='coerce')
        master_sales_df['Month'] = master_sales_df['_sales_dt'].dt.strftime('%B').str.upper()
        master_sales_df['Date_str'] = master_sales_df['_sales_dt'].dt.strftime('%Y-%m-%d')
        master_sales_df['Day'] = master_sales_df['_sales_dt'].dt.day
    if gmv_cand:
        master_sales_df['_gross_amt'] = pd.to_numeric(master_sales_df[gmv_cand], errors='coerce').fillna(0.0)
    else:
        master_sales_df['_gross_amt'] = 0.0
    if qty_cand:
        master_sales_df['_qty'] = pd.to_numeric(master_sales_df[qty_cand], errors='coerce').fillna(0.0)
    else:
        master_sales_df['_qty'] = 1.0

if uploaded_files:
    consolidated_dfs = []
    raw_files_dict = {}
    consolidated_raw_tabs = {}
    
    for uploaded_file in uploaded_files:
        fallback_month_name = os.path.splitext(uploaded_file.name)[0].upper()
        raw_files_dict[uploaded_file.name] = {}
        
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, skiprows=6)
                sheet_name = "Sheet1"
                raw_files_dict[uploaded_file.name][sheet_name] = df.copy()
                
                date_col = 'METRICS_DATE' if 'METRICS_DATE' in df.columns else None
                if date_col:
                    dt_series = pd.to_datetime(df[date_col], errors='coerce')
                    month_series = dt_series.dt.strftime('%B').str.upper()
                    df['Month'] = month_series.fillna(fallback_month_name)
                    df['_date_dt'] = dt_series
                    df['Date'] = dt_series.dt.strftime('%Y-%m-%d')
                else:
                    df['Month'] = fallback_month_name
                    df['_date_dt'] = pd.NaT
                    df['Date'] = np.nan
                
                df_consolidated = df.copy()
                df_consolidated.insert(0, 'Month', df_consolidated.pop('Month') if 'Month' in df_consolidated.columns else fallback_month_name)
                df_consolidated.insert(1, 'Tab Name', sheet_name)
                consolidated_dfs.append(df_consolidated)
                
                if sheet_name not in consolidated_raw_tabs:
                    consolidated_raw_tabs[sheet_name] = []
                consolidated_raw_tabs[sheet_name].append(df.copy())
            else:
                xls = pd.ExcelFile(uploaded_file)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    raw_files_dict[uploaded_file.name][sheet_name] = df.copy()
                    
                    date_col = 'METRICS_DATE' if 'METRICS_DATE' in df.columns else ('Date' if 'Date' in df.columns else None)
                    if date_col:
                        dt_series = pd.to_datetime(df[date_col], errors='coerce')
                        month_series = dt_series.dt.strftime('%B').str.upper()
                        df['Month'] = month_series.fillna(fallback_month_name)
                        df['_date_dt'] = dt_series
                        df['Date'] = dt_series.dt.strftime('%Y-%m-%d')
                    else:
                        df['Month'] = fallback_month_name
                        df['_date_dt'] = pd.NaT
                        df['Date'] = np.nan
                        
                    if sheet_name not in consolidated_raw_tabs:
                        consolidated_raw_tabs[sheet_name] = []
                    consolidated_raw_tabs[sheet_name].append(df.copy())
                    
                    df_consolidated = df.copy()
                    df_consolidated.insert(0, 'Month', df_consolidated.pop('Month') if 'Month' in df_consolidated.columns else fallback_month_name)
                    df_consolidated.insert(1, 'Tab Name', sheet_name)
                    consolidated_dfs.append(df_consolidated)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)
        
        def get_numeric_col(df, possible_cols):
            for col in possible_cols:
                if col in df.columns:
                    return pd.to_numeric(df[col], errors='coerce').fillna(0)
            return pd.Series(0, index=df.index)

        final_df['_impressions'] = get_numeric_col(final_df, ['TOTAL_IMPRESSIONS', 'Impressions'])
        final_df['_atc'] = get_numeric_col(final_df, ['TOTAL_A2C', 'A2C'])
        final_df['_orders'] = get_numeric_col(final_df, ['TOTAL_CONVERSIONS', 'Orders'])
        final_df['_sales'] = get_numeric_col(final_df, ['TOTAL_GMV', 'GMV', 'Sales'])
        final_df['_budget_consumed'] = get_numeric_col(final_df, ['TOTAL_BUDGET_BURNT', 'Budget Burnt', 'Spend'])
        
        final_df['ROAS'] = final_df.apply(lambda r: round(r['_sales'] / r['_budget_consumed'], 2) if r['_budget_consumed'] > 0 else 0.0, axis=1)
        final_df['ACOS'] = final_df.apply(lambda r: round((r['_budget_consumed'] / r['_sales']) * 100, 2) if r['_sales'] > 0 else 0.0, axis=1)
        final_df['CPM'] = final_df.apply(lambda r: round((r['_budget_consumed'] / r['_impressions']) * 1000, 2) if r['_impressions'] > 0 else 0.0, axis=1)

        has_sales_data = not master_sales_df.empty and '_sales_dt' in master_sales_df.columns

        MONTH_ORDER = {
            "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4,
            "MAY": 5, "JUNE": 6, "JULY": 7, "AUGUST": 8,
            "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12
        }

        def get_calendar_months(df_input):
            months = [str(x).strip().upper() for x in df_input['Month'].dropna().unique()]
            return sorted(months, key=lambda x: MONTH_ORDER.get(x, 99))

        def create_monthly_summary_table(df_input):
            working_df = df_input.dropna(subset=['Month']).copy()
            if working_df.empty:
                return pd.DataFrame()
            
            all_months = get_calendar_months(working_df)
            has_sales = not master_sales_df.empty and '_sales_dt' in master_sales_df.columns

            rows_list = []
            for m in all_months:
                m_df = working_df[working_df['Month'] == m]
                imp = m_df['_impressions'].sum()
                atc = m_df['_atc'].sum()
                ords = m_df['_orders'].sum()
                spends = round(m_df['_budget_consumed'].sum(), 2)
                sales = round(m_df['_sales'].sum(), 2)
                cpm = round((spends / imp) * 1000, 2) if imp > 0 else 0.0
                roas = round(sales / spends, 2) if spends > 0 else 0.0
                acos = round((spends / sales) * 100, 2) if sales > 0 else 0.0
                
                row_dict = {
                    'Month': m,
                    'Impressions': imp,
                    'CPM': cpm,
                    'ATC': atc,
                    'Spends': spends,
                    'Ad Sales': sales,
                }
                
                if has_sales:
                    m_sales_df = master_sales_df[master_sales_df['Month'] == m]
                    ov_sales = round(m_sales_df['_gross_amt'].sum(), 2)
                    ov_orders = len(m_sales_df)
                    tacos = round((spends / ov_sales) * 100, 2) if ov_sales > 0 else 0.0
                    troas = round(ov_sales / spends, 2) if spends > 0 else 0.0
                    
                    row_dict['Overall Sales'] = ov_sales
                    row_dict['Ad ROAS'] = roas
                    row_dict['TROAS'] = troas
                    row_dict['Ad Orders'] = ords
                    row_dict['Overall Orders'] = ov_orders
                    row_dict['TACOS'] = f"{tacos:.2f}%"
                    row_dict['ACOS'] = f"{acos:.2f}%"
                else:
                    row_dict['Ad ROAS'] = roas
                    row_dict['Ad Orders'] = ords
                    row_dict['ACOS'] = f"{acos:.2f}%"

                rows_list.append(row_dict)

            res_df = pd.DataFrame(rows_list)
            if len(res_df) >= 2:
                prev_row = res_df.loc[len(res_df) - 2]
                curr_row = res_df.loc[len(res_df) - 1]
                pct_row = {'Month': f"{curr_row['Month']} vs {prev_row['Month']} %"}
                for metric in [c for c in res_df.columns if c != 'Month']:
                    try:
                        p_v = float(str(prev_row[metric]).replace('%', ''))
                        c_v = float(str(curr_row[metric]).replace('%', ''))
                        if p_v > 0:
                            chg = round(((c_v - p_v) / p_v) * 100, 2)
                            pct_row[metric] = f"{chg:+.2f}%"
                        else:
                            pct_row[metric] = "0.00%"
                    except:
                        pct_row[metric] = "-"
                res_df = pd.concat([res_df, pd.DataFrame([pct_row])], ignore_index=True)
            return res_df

        def create_daily_tracker_table(selected_month):
            if master_sales_df.empty or '_sales_dt' not in master_sales_df.columns:
                return pd.DataFrame()

            m_df = master_sales_df[master_sales_df['Month'] == selected_month].copy()
            if m_df.empty:
                return pd.DataFrame()

            ad_df = final_df[final_df['Month'] == selected_month].copy() if 'Month' in final_df.columns else pd.DataFrame()

            rows = []
            for day_i in range(1, 32):
                day_sales = m_df[m_df['Day'] == day_i]
                ov_sales = round(day_sales['_gross_amt'].sum(), 2)
                ov_orders = len(day_sales)
                ov_qty = int(day_sales['_qty'].sum()) if '_qty' in day_sales.columns else ov_orders
                aov = round(ov_sales / ov_orders, 2) if ov_orders > 0 else 0.0
                
                day_spends = 0.0
                day_ad_sales = 0.0
                day_ad_orders = 0

                if not ad_df.empty and 'Date' in ad_df.columns:
                    matched_ad = ad_df[ad_df['_date_dt'].dt.day == day_i]
                    day_spends = round(matched_ad['_budget_consumed'].sum(), 2)
                    day_ad_sales = round(matched_ad['_sales'].sum(), 2)
                    day_ad_orders = int(matched_ad['_orders'].sum())

                roas = round(day_ad_sales / day_spends, 2) if day_spends > 0 else 0.0
                acos = round((day_spends / day_ad_sales) * 100, 2) if day_ad_sales > 0 else 0.0
                tacos = round((day_spends / ov_sales) * 100, 2) if ov_sales > 0 else 0.0

                rows.append({
                    'Date': f"Day {day_i}",
                    'Spend': day_spends,
                    'Orders': int(day_ad_orders),
                    'Sales': day_ad_sales,
                    'ACOS': acos / 100.0,
                    'ROAS': roas,
                    'Qty Sold': ov_qty,
                    'OV Sales': ov_sales,
                    'T-ACOS': tacos / 100.0,
                    'AOV': aov
                })

            res_df = pd.DataFrame(rows)
            
            tot_spend = res_df['Spend'].sum()
            tot_orders = res_df['Orders'].sum()
            tot_sales = res_df['Sales'].sum()
            tot_qty = res_df['Qty Sold'].sum()
            tot_ov_sales = res_df['OV Sales'].sum()
            
            avg_acos = round((tot_spend / tot_sales), 4) if tot_sales > 0 else 0.0
            avg_roas = round(tot_sales / tot_spend, 2) if tot_spend > 0 else 0.0
            avg_tacos = round((tot_spend / tot_ov_sales), 4) if tot_ov_sales > 0 else 0.0
            avg_aov = round(tot_ov_sales / tot_orders, 2) if tot_orders > 0 else 0.0

            summary_dict = {
                'Date': 'Aggregate / Average',
                'Spend': tot_spend,
                'Orders': int(tot_orders),
                'Sales': tot_sales,
                'ACOS': avg_acos,
                'ROAS': avg_roas,
                'Qty Sold': int(tot_qty),
                'OV Sales': tot_ov_sales,
                'T-ACOS': avg_tacos,
                'AOV': avg_aov
            }
            res_df = pd.concat([res_df, pd.DataFrame([summary_dict])], ignore_index=True)
            return res_df

        def create_mom_comparison_table(df_input, entity_col):
            if entity_col not in df_input.columns:
                return pd.DataFrame()
            
            working_df = df_input.dropna(subset=[entity_col, 'Month']).copy()
            if working_df.empty:
                return pd.DataFrame()
            
            grouped = working_df.groupby([entity_col, 'Month']).agg(
                Impressions=('_impressions', 'sum'),
                ATC=('_atc', 'sum'),
                Orders=('_orders', 'sum'),
                Spends=('_budget_consumed', 'sum'),
                Sales=('_sales', 'sum')
            ).reset_index()
            
            grouped['CPM'] = grouped.apply(lambda r: round((r['Spends'] / r['Impressions']) * 1000, 2) if r['Impressions'] > 0 else 0.0, axis=1)
            grouped['ROAS'] = grouped.apply(lambda r: round(r['Sales'] / r['Spends'], 2) if r['Spends'] > 0 else 0.0, axis=1)
            grouped['ACOS'] = grouped.apply(lambda r: round((r['Spends'] / r['Sales']) * 100, 2) if r['Sales'] > 0 else 0.0, axis=1)
            
            pivot_df = grouped.pivot(
                index=entity_col,
                columns='Month',
                values=['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS']
            )
            metrics_order = ['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS']
            all_months = get_calendar_months(df_input)
            
            pivot_df = pivot_df.reorder_levels([1, 0], axis=1)
            sorted_cols = pd.MultiIndex.from_product([all_months, metrics_order], names=['Month', 'Metric'])
            pivot_df = pivot_df.reindex(columns=sorted_cols).fillna(0)
            
            grand_total_series = {}
            for month in all_months:
                month_df = working_df[working_df['Month'] == month]
                total_imp = month_df['_impressions'].sum()
                total_atc = month_df['_atc'].sum()
                total_orders = month_df['_orders'].sum()
                total_spends = round(month_df['_budget_consumed'].sum(), 2)
                total_sales = round(month_df['_sales'].sum(), 2)
                total_cpm = round((total_spends / total_imp) * 1000, 2) if total_imp > 0 else 0.0
                total_roas = round(total_sales / total_spends, 2) if total_spends > 0 else 0.0
                total_acos = round((total_spends / total_sales) * 100, 2) if total_sales > 0 else 0.0
                
                grand_total_series[(month, 'Impressions')] = total_imp
                grand_total_series[(month, 'CPM')] = total_cpm
                grand_total_series[(month, 'ATC')] = total_atc
                grand_total_series[(month, 'Orders')] = total_orders
                grand_total_series[(month, 'Spends')] = total_spends
                grand_total_series[(month, 'Sales')] = total_sales
                grand_total_series[(month, 'ROAS')] = total_roas
                grand_total_series[(month, 'ACOS')] = total_acos
                
            pivot_df.loc['Grand Total'] = grand_total_series
            pivot_df.index.name = entity_col
            return pivot_df

        def _format_dashboard_value(value, column_name=""):
            if pd.isna(value): return value
            col = str(column_name).strip().upper()
            is_percentage = ('%' in col) or ('ACOS' in col) or ('TACOS' in col) or ('T-ACOS' in col)
            if isinstance(value, (int, np.integer)) and not isinstance(value, bool):
                return f"{int(value):,}" if not is_percentage else f"{int(value)}%"
            if isinstance(value, (float, np.floating)) and np.isfinite(value):
                if is_percentage: return f"{value * 100:.2f}%" if value <= 1.0 else f"{value:.2f}%"
                if col in ['SPEND', 'SPENDS', 'SALES', 'AD SALES', 'TOTAL SALES', 'OVERALL SALES', 'OV SALES', 'AOV']:
                    return f"{int(round(value)):,}" if col != 'AOV' else f"{value:,.2f}"
                if float(value).is_integer(): return f"{int(value):,}"
                return f"{value:,.2f}"
            return value

        def format_dashboard_dataframe(df):
            out = df.copy()
            if isinstance(out.columns, pd.MultiIndex):
                for col in out.columns:
                    out[col] = out[col].map(lambda v: _format_dashboard_value(v, f"{col[0]} {col[1]}"))
            else:
                for col in out.columns:
                    out[col] = out[col].map(lambda v: _format_dashboard_value(v, col))
            return out

        def render_unified_single_table(df_to_show, key_prefix="mom"):
            if df_to_show is None or df_to_show.empty: return
            display_df = format_dashboard_dataframe(df_to_show)
            st.dataframe(
                display_df.style.set_properties(**{
                    'text-align': 'center', 'vertical-align': 'middle'
                }),
                use_container_width=True,
                hide_index=False if isinstance(df_to_show.index, pd.MultiIndex) or df_to_show.index.name else True,
                key=f"{key_prefix}_single_grid"
            )

        # Global Dashboard Metrics
        st.markdown("### 🔍 Global Dashboard Filters")
        available_months_asc = get_calendar_months(final_df)
        selected_month = st.selectbox("Select Month Across Dashboard (Excluding Comparison Tables)", ["All Months"] + available_months_asc)

        filtered_df = final_df.copy()
        if selected_month != "All Months":
            filtered_df = filtered_df[filtered_df['Month'] == selected_month]

        total_ad_sales = round(filtered_df['_sales'].sum(), 2)
        total_ad_spends = round(filtered_df['_budget_consumed'].sum(), 2)
        overall_ad_roas = round(total_ad_sales / total_ad_spends, 2) if total_ad_spends > 0 else 0.0

        st.markdown("### 📈 Overall Swiggy Campaign Performance Dashboard")
        if has_sales_data:
            f_sales_df = master_sales_df if selected_month == "All Months" else master_sales_df[master_sales_df['Month'] == selected_month]
            total_overall_sales = round(f_sales_df['_gross_amt'].sum(), 2)
            total_overall_orders = len(f_sales_df)
            overall_tacos = round((total_ad_spends / total_overall_sales) * 100, 2) if total_overall_sales > 0 else 0.0
            overall_troas = round(total_overall_sales / total_ad_spends, 2) if total_ad_spends > 0 else 0.0

            # 6 fields in 2 rows: Ad Sales, Overall Sales, Ad ROAS / TACOS, TROAS, Ov Orders
            dcol1, dcol2, dcol3 = st.columns(3)
            with dcol1: st.metric("Ad Sales", f"₹{int(round(total_ad_sales)):,}")
            with dcol2: st.metric("Overall Sales", f"₹{int(round(total_overall_sales)):,}")
            with dcol3: st.metric("Ad ROAS", f"{overall_ad_roas:.2f}x")

            dcol4, dcol5, dcol6 = st.columns(3)
            with dcol4: st.metric("TACOS", f"{overall_tacos:.2f}%")
            with dcol5: st.metric("TROAS", f"{overall_troas:.2f}")
            with dcol6: st.metric("Ov Orders", f"{total_overall_orders:,}")
        else:
            dcol1, dcol2, dcol3 = st.columns(3)
            with dcol1: st.metric("Ad Sales", f"₹{int(round(total_ad_sales)):,}")
            with dcol2: st.metric("Spends", f"₹{int(round(total_ad_spends)):,}")
            with dcol3: st.metric("Ad ROAS", f"{overall_ad_roas:.2f}x")

        st.divider()
        st.markdown("### 📑 Navigation & Swiggy Analytics Tabs")

        main_tab1, main_tab2, main_tab3, main_tab4, main_tab5 = st.tabs([
            "📄 Raw Files Preview",
            "📌 Consolidated Master",
            "📊 Comparison Tables (MoM)",
            "📅 Daily Tracker",
            "🎯 Campaign / City Performance"
        ])

        with main_tab1:
            st.caption("Inspect individual sheets/files tab-by-tab.")
            selected_file_name = st.selectbox("Select Uploaded File to Preview:", list(raw_files_dict.keys()), key="raw_file_sel")
            if selected_file_name:
                sheets = raw_files_dict[selected_file_name]
                selected_sheet = st.selectbox("Select Sheet Tab:", list(sheets.keys()), key="raw_sheet_sel")
                if selected_sheet:
                    st.dataframe(format_dashboard_dataframe(sheets[selected_sheet].head(100)), use_container_width=True, hide_index=True)

        with main_tab2:
            st.caption("Preview the combined dataset across all uploaded files.")
            st.dataframe(format_dashboard_dataframe(final_df.head(100)), use_container_width=True, hide_index=True)

        with main_tab3:
            st.subheader("📊 Month-on-Month Comparison Tables")
            mom_dict = {
                "Monthly_Summary": create_monthly_summary_table(final_df),
                "Campaign_MoM": create_mom_comparison_table(final_df, 'CAMPAIGN_NAME') if 'CAMPAIGN_NAME' in final_df.columns else None,
                "City_MoM": create_mom_comparison_table(final_df, 'CITY') if 'CITY' in final_df.columns else None,
                "Product_MoM": create_mom_comparison_table(final_df, 'PRODUCT_NAME') if 'PRODUCT_NAME' in final_df.columns else None,
                "AdProperty_MoM": create_mom_comparison_table(final_df, 'AD_PROPERTY') if 'AD_PROPERTY' in final_df.columns else None
            }
            all_pivots_bytes = convert_all_pivots_to_excel(mom_dict)
            st.download_button(
                label="📥 Download All MoM Comparison Tables (.xlsx - Master Workbook)",
                data=all_pivots_bytes,
                file_name="Swiggy_All_MoM_Comparison_Tables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_all_mom",
                type="primary"
            )
            st.divider()

            comp_sub_tab0, comp_sub_tab1, comp_sub_tab2 = st.tabs([
                "📅 Monthly Summary",
                "🎯 Campaign Comparison",
                "🏙️ City Comparison"
            ])
            with comp_sub_tab0:
                ms_df = create_monthly_summary_table(final_df)
                render_unified_single_table(ms_df, key_prefix="ms_tab")
                st.download_button("📥 Download Monthly Summary (.xlsx)", style_and_export_pivot(ms_df, "Monthly_Summary"), "Monthly_Summary.xlsx", key="dl_ms")
            with comp_sub_tab1:
                cp_df = create_mom_comparison_table(final_df, 'CAMPAIGN_NAME') if 'CAMPAIGN_NAME' in final_df.columns else pd.DataFrame()
                render_unified_single_table(cp_df, key_prefix="cp_tab")
                st.download_button("📥 Download Campaign MoM (.xlsx)", style_and_export_pivot(cp_df, "Campaign_MoM"), "Campaign_MoM.xlsx", key="dl_cp")
            with comp_sub_tab2:
                city_df = create_mom_comparison_table(final_df, 'CITY') if 'CITY' in final_df.columns else pd.DataFrame()
                render_unified_single_table(city_df, key_prefix="city_tab")
                st.download_button("📥 Download City MoM (.xlsx)", style_and_export_pivot(city_df, "City_MoM"), "City_MoM.xlsx", key="dl_city")

        with main_tab4:
            st.subheader("📅 Daily Tracker")
            if has_sales_data:
                tracker_month = st.selectbox("Select Month for Daily Tracker:", available_months_asc, key="tracker_month_sel")
                if tracker_month:
                    daily_df = create_daily_tracker_table(tracker_month)
                    if not daily_df.empty:
                        render_unified_single_table(daily_df, key_prefix="daily_tracker_tab")
                        excel_daily = style_and_export_pivot(daily_df, sheet_name="Daily_Tracker")
                        st.download_button(
                            label="📥 Download Daily Tracker Table (.xlsx)",
                            data=excel_daily,
                            file_name=f"Daily_Tracker_{tracker_month}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_dl_daily_tracker"
                        )
                    else:
                        st.info("No sales data available for the selected month.")
            else:
                st.info("Please upload Overall Sales reports to view the Daily Tracker with TACOS, TROAS, and AOV.")

        with main_tab5:
            st.subheader("🎯 Campaign & City Granular Performance")
            if 'CAMPAIGN_NAME' in final_df.columns:
                c_opts = ["All Campaigns"] + sorted([str(x) for x in final_df['CAMPAIGN_NAME'].dropna().unique()])
                sel_camp = st.selectbox("Select Campaign:", c_opts)
                sub_df = final_df if sel_camp == "All Campaigns" else final_df[final_df['CAMPAIGN_NAME'].astype(str) == sel_camp]
                
                agg_df = sub_df.groupby('CAMPAIGN_NAME').agg(
                    Impressions=('_impressions', 'sum'),
                    ATC=('_atc', 'sum'),
                    Orders=('_orders', 'sum'),
                    Spends=('_budget_consumed', 'sum'),
                    Sales=('_sales', 'sum')
                ).reset_index()
                agg_df['ROAS'] = agg_df.apply(lambda r: round(r['Sales'] / r['Spends'], 2) if r['Spends'] > 0 else 0.0, axis=1)
                agg_df['ACOS'] = agg_df.apply(lambda r: round((r['Spends'] / r['Sales']) * 100, 2) if r['Sales'] > 0 else 0.0, axis=1)
                render_unified_single_table(agg_df, key_prefix="camp_perf")
        
        st.divider()
        st.subheader("💾 Download Complete Master Workbook")
        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            for raw_tab_name, df_list in consolidated_raw_tabs.items():
                pd.concat(df_list, ignore_index=True).to_excel(writer, sheet_name=raw_tab_name[:31], index=False)
            final_df.to_excel(writer, sheet_name='Consolidated_Master', index=False)
        buffer_multi.seek(0)
        st.download_button(
            label="📥 Download Complete Excel Workbook (.xlsx)",
            data=buffer_multi,
            file_name="Swiggy_Consolidated_Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.info("Please upload valid ad campaign reports to start analysis.")
else:
    st.info("Awaiting file uploads. Please upload your Swiggy Ad Campaign reports and Overall Sales reports.")
