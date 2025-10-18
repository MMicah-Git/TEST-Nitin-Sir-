# app_takeoffnsw.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="TakeoffNSW Formatter", layout="wide")

# -----------------------
# Helper functions (same logic as your pipeline)
# -----------------------
def read_input_file(uploaded_file):
    # Accept CSV or Excel
    try:
        if uploaded_file.name.lower().endswith((".xls", ".xlsx")):
            return pd.read_excel(uploaded_file, dtype=str).fillna("")
        else:
            # try common encodings via pandas
            return pd.read_csv(uploaded_file, dtype=str, encoding="utf-8").fillna("")
    except Exception:
        try:
            return pd.read_csv(uploaded_file, dtype=str, encoding="latin1").fillna("")
        except Exception:
            return pd.read_csv(uploaded_file, dtype=str, sep=",").fillna("")

def normalize_text(x):
    if x is None:
        return ""
    s = str(x)
    return " ".join(s.split()).strip().lower()

def parse_qty(x):
    try:
        s = str(x).strip().replace(",", "").replace("$", "")
        if s == "" or s.lower() in ("nan", "none"):
            return 0.0
        return float(s)
    except:
        return 0.0

def neck_numeric(s):
    if s is None:
        return float("inf")
    s = str(s)
    nums = re.findall(r"\d+\.?\d*", s)
    return float(nums[0]) if nums else float("inf")

# default mapping (template -> raw column)
DEFAULT_MAPPING = {
    'PRODUCT': 'Subject',
    'SCHEDULE BRAND': 'MANUFACTURER',
    'SCHEDULE MODEL': 'MODEL',
    'BRAND': 'MANUFACTURER',
    'MODEL': 'MODEL',
    'QTY': 'Count',
    'TAG': 'Label',
    'NECK SIZE': 'NECK SIZE',
    'MODULE SIZE': 'FACE SIZE',
    'DUCT SIZE': 'DUCT SIZE',
    'TYPE': 'TYPE',
    'MOUNTING': 'MOUNTING',
    'ACCESSORIES1': 'ACCESSORIES',
    'ACCESSORIES2': 'ACCESSORIES',
    'REMARK': 'REMARK'
}

BOLD_VALUE_COLS_DEFAULT = ['PRODUCT', 'SCHEDULE BRAND', 'BRAND', 'TAG', 'MODULE SIZE', 'TYPE', 'ACCESSORIES1']

# core pipeline to produce df_out (unformatted strings)
def takeoff_pipeline(df_raw, mapping=DEFAULT_MAPPING, brand_rules=True,
                     model_blank=True, empty_dot=True, bold_value_cols=BOLD_VALUE_COLS_DEFAULT):
    # normalize column names
    df_raw = df_raw.copy()
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    # apply mapping to template headers
    template_headers = list(mapping.keys())
    rows = []
    for _, r in df_raw.iterrows():
        out = {}
        for th in template_headers:
            rc = mapping.get(th, "")
            val = r.get(rc, "") if rc in r.index else ""
            if isinstance(val, str):
                val = " ".join(val.split())
            out[th] = val
        rows.append(out)
    df = pd.DataFrame(rows, columns=template_headers)

    # brand rules (clear then apply)
    if 'BRAND' not in df.columns:
        df['BRAND'] = ''
    else:
        df['BRAND'] = ''

    if brand_rules and 'PRODUCT' in df.columns:
        df.loc[df['PRODUCT'].str.contains(r'AD[-\s]?GRD', case=False, na=False), 'BRAND'] = 'PRICE'
        df.loc[df['PRODUCT'].str.contains(r'\bFAN\b', case=False, na=False), 'BRAND'] = 'LOREN COOK'
        df.loc[df['PRODUCT'].str.contains(r'SPLIT SYSTEM HEAT PUMP', case=False, na=False), 'BRAND'] = 'SAMSUNG'

    # clear MODEL if requested
    if model_blank and 'MODEL' in df.columns:
        df['MODEL'] = ''

    # sorting helpers
    df['_neck_num'] = df['NECK SIZE'].apply(neck_numeric) if 'NECK SIZE' in df.columns else float('inf')
    df['_product_norm'] = df['PRODUCT'].astype(str).str.strip().str.upper() if 'PRODUCT' in df.columns else ''
    df['_tag_norm'] = df['TAG'].astype(str).str.strip().str.upper() if 'TAG' in df.columns else ''
    df_sorted = df.sort_values(by=['_product_norm','_tag_norm','_neck_num','PRODUCT'], ascending=[True, True, True, True], kind='mergesort')

    # grouping keys
    GROUP_FIELDS = [
        'PRODUCT','SCHEDULE BRAND','SCHEDULE MODEL','BRAND','MODEL',
        'NECK SIZE','MODULE SIZE','DUCT SIZE','TYPE','MOUNTING',
        'ACCESSORIES1','ACCESSORIES2','REMARK'
    ]
    df_sorted['_key'] = df_sorted.apply(lambda r: tuple(normalize_text(r.get(c,'')) for c in GROUP_FIELDS), axis=1)
    df_sorted['_qty'] = df_sorted['QTY'].apply(parse_qty) if 'QTY' in df_sorted.columns else 0.0

    grouped_rows = []
    product_order = sorted(df_sorted['_product_norm'].unique(), key=lambda x: x or '')
    for prod in product_order:
        prod_group = df_sorted[df_sorted['_product_norm'] == prod]
        if prod_group.empty:
            continue
        tags_in_prod = sorted(prod_group['_tag_norm'].unique(), key=lambda x: x or '')
        for tag in tags_in_prod:
            tag_group = prod_group[prod_group['_tag_norm'] == tag]
            if tag_group.empty:
                continue
            grouped = tag_group.groupby('_key', sort=False)
            for _, sub in grouped:
                rep = sub.iloc[0].to_dict()
                rep['QTY'] = f"{sub['_qty'].sum():.2f}"
                if 'MODEL' in rep:
                    rep['MODEL'] = '' if model_blank else rep.get('MODEL','')
                grouped_rows.append(rep)
            subtotal = tag_group['_qty'].sum()
            prod_text = prod_group.iloc[0]['PRODUCT']
            tag_label = tag if tag else ''
            subtotal_row = {c:'' for c in df_sorted.columns}
            subtotal_row['PRODUCT'] = f"{prod_text} - {tag_label} TOTAL = ({subtotal:.2f})"
            subtotal_row['TAG'] = f"{tag_label} TOTAL"
            subtotal_row['QTY'] = f"{subtotal:.2f}"
            if model_blank:
                subtotal_row['MODEL'] = ''
            grouped_rows.append(subtotal_row)

    grand_total = df_sorted['_qty'].sum()
    grand_row = {c:'' for c in df_sorted.columns}
    grand_row['PRODUCT'] = "Grand Total"
    grand_row['TAG'] = "Grand Total"
    grand_row['QTY'] = f"{grand_total:.2f}"
    if model_blank:
        grand_row['MODEL'] = ''
    grouped_rows.append(grand_row)

    df_out = pd.DataFrame(grouped_rows)
    # drop helper columns if present
    drop_cols = [c for c in ['_neck_num','_product_norm','_tag_norm','_key','_qty'] if c in df_out.columns]
    df_out = df_out.drop(columns=drop_cols, errors='ignore')

    # Fill empties with dot or keep model blank
    if empty_dot:
        # Replace pure-empty strings with '.' but keep MODEL blank if requested
        for col in df_out.columns:
            if col == 'MODEL' and model_blank:
                df_out[col] = df_out[col].replace({None:'', pd.NA: ''}).astype(str).replace(r'^\s*$', '', regex=True)
            else:
                df_out[col] = df_out[col].replace({None:'', pd.NA: ''}).astype(str).replace(r'^\s*$', '.', regex=True)
    else:
        # keep empties as ''
        for col in df_out.columns:
            if col == 'MODEL' and model_blank:
                df_out[col] = df_out[col].replace({None:'', pd.NA: ''}).astype(str).replace(r'^\s*$', '', regex=True)
            else:
                df_out[col] = df_out[col].replace({None:'', pd.NA: ''}).astype(str)

    # ensure final column order as in mapping
    final_cols = [c for c in mapping.keys() if c in df_out.columns]
    for c in df_out.columns:
        if c not in final_cols:
            final_cols.append(c)
    df_out = df_out.loc[:, final_cols]

    return df_out, final_cols

def export_styled_excel_bytes(df_out, final_cols, bold_value_cols, header_fill_hex="FFD966"):
    # create workbook and style
    wb = Workbook()
    ws = wb.active
    ws.title = "Takeoff NSW"
    thin = Side(border_style="thin", color="000000")
    header_fill = PatternFill(start_color=header_fill_hex, end_color=header_fill_hex, fill_type="solid")
    subtotal_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    grand_fill = PatternFill(start_color="CFE2F3", end_color="CFE2F3", fill_type="solid")
    qty_red = "FF0000"

    # header
    for i, col in enumerate(final_cols, start=1):
        cell = ws.cell(row=1, column=i, value=col)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # rows
    for r_idx, (_, row) in enumerate(df_out.iterrows(), start=2):
        is_subtotal = str(row.get('TAG','')).strip().endswith('TOTAL') and 'Grand' not in str(row.get('TAG',''))
        is_grand = str(row.get('TAG','')).strip().lower() == 'grand total'
        for c_idx, col in enumerate(final_cols, start=1):
            val = row.get(col, '.')
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            if col == 'QTY':
                # attempt numeric conversion
                v = str(val).replace(',', '').replace('(', '').replace(')', '').replace('$','').strip()
                try:
                    if v not in ('', '.'):
                        cell.value = float(v)
                        cell.number_format = '0.00'
                    cell.font = Font(color=qty_red)
                except:
                    cell.font = Font(color=qty_red)
            if is_subtotal:
                if col == 'QTY':
                    cell.font = Font(bold=True, color=qty_red)
                else:
                    cell.font = Font(bold=True)
                cell.fill = subtotal_fill
            if is_grand:
                cell.font = Font(bold=True)
                cell.fill = grand_fill
            if col in bold_value_cols and not is_grand:
                # make value bold
                current_font = cell.font or Font()
                cell.font = Font(bold=True, name=current_font.name, size=current_font.size, color=current_font.color)

    # adjust widths
    for i, col in enumerate(final_cols, start=1):
        col_letter = get_column_letter(i)
        max_len = max(len(str(col)), max((len(str(ws.cell(row=r, column=i).value or "")) for r in range(2, ws.max_row+1)), default=0)) + 4
        ws.column_dimensions[col_letter].width = min(max_len, 60)

    # return bytes
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# -----------------------
# Streamlit UI
# -----------------------
st.title("TakeoffNSW Formatter — Streamlit")
st.write("Upload a raw CSV or Excel and convert to the TakeoffNSW formatted, grouped & styled file.")

uploaded = st.file_uploader("Upload raw CSV / Excel file", type=["csv","xls","xlsx"])
st.sidebar.header("Options")
brand_rules = st.sidebar.checkbox("Apply BRAND auto-fill rules (AD-GRD / FAN / SPLIT SYSTEM HEAT PUMP)", value=True)
model_blank = st.sidebar.checkbox("Keep MODEL column blank", value=True)
empty_dot = st.sidebar.checkbox("Replace empty cells with '.'", value=True)
bold_cols_input = st.sidebar.text_area("Bold value columns (comma separated)", value="PRODUCT, SCHEDULE BRAND, BRAND, TAG, MODULE SIZE, TYPE, ACCESSORIES1")
# parse bold columns
bold_value_cols = [c.strip() for c in bold_cols_input.split(",") if c.strip()]

if uploaded is not None:
    st.info(f"Processing file: {uploaded.name}")
    df_raw = read_input_file(uploaded)
    try:
        df_out, final_cols = takeoff_pipeline(df_raw, mapping=DEFAULT_MAPPING,
                                             brand_rules=brand_rules,
                                             model_blank=model_blank,
                                             empty_dot=empty_dot,
                                             bold_value_cols=bold_value_cols)
    except Exception as e:
        st.error(f"Error processing file: {e}")
        raise

    st.subheader("Preview (first 200 rows)")
    st.dataframe(df_out.head(200), use_container_width=True)

    # prepare downloads
    csv_bytes = df_out.to_csv(index=False).encode("utf-8")
    excel_bio = export_styled_excel_bytes(df_out, final_cols, bold_value_cols)

    st.download_button("Download CSV", data=csv_bytes, file_name="TakeoffNSW_Converted.csv", mime="text/csv")
    st.download_button("Download Excel", data=excel_bio, file_name="TakeoffNSW_Converted.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.success("Done — download files above. You can adjust options in the sidebar and re-upload another file.")
else:
    st.info("No file uploaded yet. Upload a raw CSV/XLSX to begin.")
