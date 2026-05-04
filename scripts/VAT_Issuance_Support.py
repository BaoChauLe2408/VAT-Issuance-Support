# -*- coding: utf-8 -*-

from google.colab import drive
drive.mount('/content/drive')

"""# **Import Library**"""

# Import necessary libraries
import pandas as pd
import numpy as np
import openpyxl
import re
from datetime import datetime

"""# **File Path Configuration**"""

import pandas as pd

# File path - Using export link to read Excel from Google Sheets
# Security Note: Replaced specific URL with a placeholder
file_path = "https://docs.google.com/spreadsheets/d/[PROTECTED_ID]/export?format=xlsx"

# Read multiple sheets (index 0 and 2)
all_sheets = pd.read_excel(file_path, sheet_name=[0, 2], header=0)

# System product names (used for merging and retrieving VAT info)
Tenhanghoa_he_thong = all_sheets[0]
Tenhanghoa_he_thong

Tenhanghoa_xuat_VAT = all_sheets[2]
Tenhanghoa_xuat_VAT

# --- Bug Fix: Convert URL to CSV export format ---
# File ID: [PROTECTED_ID]
# GID: [PROTECTED_GID]
file_path_csv = 'https://docs.google.com/spreadsheets/d/[PROTECTED_ID]/export?format=csv&gid=[PROTECTED_GID]'

# 1. Use pd.read_csv()
# 2. header=1: Use the second row (index 1) as column headers
Hoadontheothoigian = pd.read_csv(file_path_csv, header=1)

# Hoadontheothoigian.drop(index=0): Remove the first row of data
Hoadontheothoigian = Hoadontheothoigian.drop(index=0)
Hoadontheothoigian

import pandas as pd

# --- Bug Fix: Convert URL to CSV export format ---
file_path_csv_fixed = 'https://docs.google.com/spreadsheets/d/[PROTECTED_ID]/export?format=csv&gid=[PROTECTED_GID]'

# 1. Use pd.read_csv()
# 2. header=1: Use the second row (index 1) as column headers
Doichieuhoadon = pd.read_csv(file_path_csv_fixed, header=1)
Doichieuhoadon

"""# **Processing Conversion Items for VAT Invoicing**

## **Merging to get Invoicing Names and VAT**
"""

# 1. Standardize joining column on the Left DataFrame
Tenhanghoa_he_thong['Tên_chuan'] = (
    Tenhanghoa_he_thong['Tên']
    .astype(str)    # Ensure string type
    .str.lower()    # Convert to lowercase
    .str.strip()    # Remove leading/trailing whitespace
)

# 2. Standardize joining column on the Right DataFrame
Tenhanghoa_xuat_VAT['Tên_HT_chuan'] = (
    Tenhanghoa_xuat_VAT['Tên hệ thống']
    .astype(str)    # Ensure string type
    .str.lower()    # Convert to lowercase
    .str.strip()    # Remove leading/trailing whitespace
)

# 3. Perform Merge using standardized columns
output_menu_rooftop_VAT = pd.merge(
    Tenhanghoa_he_thong[['Mã món','Cửa hàng','Tên','Giá','Đơn vị','Nhóm','Tên nhóm','VAT (%)', 'Tên_chuan']],
    Tenhanghoa_xuat_VAT[['Tên hệ thống','Tên xuất hóa đơn', 'Tên_HT_chuan']],
    left_on = 'Tên_chuan',         # Using standardized column
    right_on = 'Tên_HT_chuan',     # Using standardized column
    how = 'inner'
)
output_menu_rooftop_VAT = output_menu_rooftop_VAT.drop(columns=['Tên','Tên_chuan','Tên_HT_chuan'])
output_menu_rooftop_VAT

# Define desired new column order
new_order = ['Tên xuất hóa đơn','Tên hệ thống','Mã món','Cửa hàng','Giá','Đơn vị','Nhóm','Tên nhóm','VAT (%)' ]

# Reorder columns
output_menu_rooftop_VAT = output_menu_rooftop_VAT.reindex(columns=new_order)
# Remove duplicates based on system name
output_menu_rooftop_VAT = output_menu_rooftop_VAT.drop_duplicates(subset=['Tên hệ thống'])

"""## **Push output_menu_rooftop_VAT data to Google Sheets**"""

# Note: This section is commented out for security/authentication purposes during export
# [Code for Google Auth and gspread skipped as per original]

"""# **Transfer Invoices Over Time**"""

# Keep only necessary columns
Hoadontheothoigian = Hoadontheothoigian[[ 'Cửa hàng', 'Mã hoá đơn','Mã hoá đơn gốc','Số hoá đơn', 'Ngày vào', 'Bàn', 'Tên hàng', 'Số lượng', 'Đơn vị', 'Đơn giá', 'Thành tiền','Số khách','Loại thành viên','Tên khách','SĐT','Tổng hóa đơn','Giảm giá','Phương thức thanh toán']]

# Rename columns for standardization
Hoadontheothoigian = Hoadontheothoigian.rename(columns={
    'Ngày vào': 'Ngày hóa đơn',
    'Tên hàng': 'Tên hàng hóa/dịch vụ (*)',
    'Đơn vị': 'ĐVT',
    'Tên khách':'Tên'
})

Hoadontheothoigian

"""## **Handle date format errors (auto-detected as mm/dd/yyyy) => convert to dd/mm/yyyy**"""

def normalize_date(val):
    if pd.isna(val):
        return val

    val = str(val).strip()

    # Case 1: Already in DD/MM/YYYY format -> Keep as is
    if "/" in val:
        try:
            d, m, y = val.split("/")
            return f"{d.zfill(2)}/{m.zfill(2)}/{y}"
        except:
            return val

    # Case 2: Format 'YYYY-DD-MM HH:MM:SS'
    if " " in val and "-" in val:
        try:
            date_part = val.split(" ")[0]   # YYYY-DD-MM
            y, d, m = date_part.split("-")
            return f"{d.zfill(2)}/{m.zfill(2)}/{y}"
        except:
            return val

    # Case 3: Format 'YYYY-DD-MM' (Backup case)
    if "-" in val:
        try:
            y, d, m = val.split("-")
            return f"{d.zfill(2)}/{m.zfill(2)}/{y}"
        except:
            return val

    return val

Hoadontheothoigian['Ngày hóa đơn'] = (
    Hoadontheothoigian['Ngày hóa đơn']
    .astype(str)
    .apply(normalize_date)
)

# Enforce 'Ngày hóa đơn' to dd/mm/yyyy format
Hoadontheothoigian['Ngày hóa đơn'] = pd.to_datetime(
    Hoadontheothoigian['Ngày hóa đơn'], format="%d/%m/%Y"
).dt.strftime("%d/%m/%Y")

Hoadontheothoigian

"""## **Merge Invoice Code and Invoice Number**"""

# 1) Create new column based on conditions
# ---------------------------------------------
Hoadontheothoigian['Số Hóa Đơn Mới'] = np.where(
    Hoadontheothoigian['Mã hoá đơn gốc'].notna() & 
    (Hoadontheothoigian['Mã hoá đơn gốc'].astype(str).str.strip() != ''),

    # If original invoice code exists ➜ use it
    Hoadontheothoigian['Mã hoá đơn gốc'].astype(str).str.strip() + '-' + 
    Hoadontheothoigian['Số hoá đơn'].astype(str).str.strip(),

    # Otherwise ➜ use current invoice code
    Hoadontheothoigian['Mã hoá đơn'].astype(str).str.strip() + '-' + 
    Hoadontheothoigian['Số hoá đơn'].astype(str).str.strip()
)

# ---------------------------------------------
# 2) Data Cleaning: Remove 'nan' strings and artifacts
# ---------------------------------------------
Hoadontheothoigian['Số Hóa Đơn Mới'] = (
    Hoadontheothoigian['Số Hóa Đơn Mới']
      .str.replace(r'(^nan-)|(-nan$)|(^nan$)|(nan-nan)', '', regex=True)
      .replace(['nan', '', 'NaN', 'None'], np.nan)
)

# ---------------------------------------------
# 3) Drop old columns and rename the new one
# ---------------------------------------------
Hoadontheothoigian = (
    Hoadontheothoigian
      .drop(columns=['Mã hoá đơn', 'Mã hoá đơn gốc', 'Số hoá đơn'])
      .rename(columns={'Số Hóa Đơn Mới': 'Số hoá đơn'})
      .reset_index(drop=True)
)

# ---------------------------------------------
# 4) Ensure no lingering 'nan' strings
# ---------------------------------------------
Hoadontheothoigian['Số hoá đơn'] = (
    Hoadontheothoigian['Số hoá đơn']
      .replace(['nan', 'NaN', 'None', ''], np.nan)
)

Hoadontheothoigian

"""## **Fill 'None' values for specific cases to ensure fillna logic is correct**"""

# Condition 1: Name is NaN but Invoice Number is NOT NaN
mask_nan_name_and_has_invoice = Hoadontheothoigian['Tên'].isna() & Hoadontheothoigian['Số hoá đơn'].notna()
Hoadontheothoigian.loc[mask_nan_name_and_has_invoice, ['Loại thành viên', 'Tên', 'SĐT']] = ['Retail Customer','None','None']

# Condition 2: If Name is 'IPOS - 020' → assign 'None' (External integration source)
mask_ipos = Hoadontheothoigian['Tên'] == 'iPOS-O2O'
Hoadontheothoigian.loc[mask_ipos, ['Loại thành viên', 'SĐT']] = ['Retail Customer','None']

# Condition 3: If all 4 key columns are NaN → set to "None"
# Logic: Prevents forward-fill (ffill) from incorrectly propagating customer data 
# from the previous row into unrelated blank rows.
mask_all_nan = (
    Hoadontheothoigian['Tên'].isna() 
    & Hoadontheothoigian['SĐT'].isna() 
    & Hoadontheothoigian['Loại thành viên'].isna() 
    & Hoadontheothoigian['Số hoá đơn'].isna()
)
Hoadontheothoigian.loc[mask_all_nan, ['Tên', 'SĐT', 'Loại thành viên']] = 'None'

# Condition 4: If Invoice exists with Customer Name but Member Type is NaN => New Customer
mask_has_invoice_and_customer_name_and_NaN_member_type = Hoadontheothoigian['Số hoá đơn'].notna() & Hoadontheothoigian['Tên'].notna() & Hoadontheothoigian['Loại thành viên'].isna()
Hoadontheothoigian.loc[mask_has_invoice_and_customer_name_and_NaN_member_type, 'Loại thành viên'] = 'New Customer'

Hoadontheothoigian

"""## **Standardize data using fillna and other transformations**"""

# Fill logic: Rows with item names but missing price/total should be 0 (e.g., toppings, combo items)
for index, row in Hoadontheothoigian.iterrows():
    if not pd.isna(row['Tên hàng hóa/dịch vụ (*)']):
        if pd.isna(row['Đơn giá']):
            Hoadontheothoigian.at[index, 'Đơn giá'] = 0
        if pd.isna(row['Thành tiền']):
            Hoadontheothoigian.at[index, 'Thành tiền'] = 0

# Handle rows where specific invoice metadata needs to be marked before propagation
for index, row in Hoadontheothoigian.iterrows():
    if not pd.isna(row['Số hoá đơn']):
        columns_to_fill = ['Tên hàng hóa/dịch vụ (*)', 'Số lượng', 'ĐVT','Đơn giá', 'Thành tiền']
        Hoadontheothoigian.loc[index, columns_to_fill] = 'None'

# Forward fill missing values
Hoadontheothoigian.fillna(method='ffill', inplace=True)

# Update: Fill 'Set' unit for COMBOs to avoid incorrect fillna inheritance
Hoadontheothoigian.loc[
    Hoadontheothoigian["Tên hàng hóa/dịch vụ (*)"].str.startswith(("COMBO -", "Set"), na=False), 
    "ĐVT"
] = "Set"

# Filter out unnecessary rows and drop columns used for processing
Hoadontheothoigian = Hoadontheothoigian.drop(columns=['Đơn giá','Loại thành viên','Tên','SĐT','Tổng hóa đơn','Giảm giá'])
Hoadontheothoigian = Hoadontheothoigian[Hoadontheothoigian['Tên hàng hóa/dịch vụ (*)'] != 'None']

# Clean 'Thành tiền' column: Extract numeric part and convert to numeric type
Hoadontheothoigian['Thành tiền'] = Hoadontheothoigian['Thành tiền'].astype(str).str.split().str[0]
Hoadontheothoigian['Thành tiền'] = pd.to_numeric(Hoadontheothoigian['Thành tiền'], errors='coerce')

# Remove rows with zero total after processing
Hoadontheothoigian = Hoadontheothoigian[Hoadontheothoigian['Thành tiền'] != 0]

# Calculate 'Đơn giá' (Unit Price)
Hoadontheothoigian['Thành tiền'] = pd.to_numeric(Hoadontheothoigian['Thành tiền'], errors='coerce')
Hoadontheothoigian['Số lượng'] = pd.to_numeric(Hoadontheothoigian['Số lượng'], errors='coerce')

# Handle division by zero/NaN for unit price calculation
Hoadontheothoigian['Đơn giá'] = np.where(
    Hoadontheothoigian['Số lượng'] != 0, 
    Hoadontheothoigian['Thành tiền'] / Hoadontheothoigian['Số lượng'], 
    np.nan
)

"""Handling specific surcharge logic for 2026 Holiday (Surcharge includes 8% VAT)"""

# Overwrite unit price for Holiday Surcharges (Pre-VAT calculation)
Hoadontheothoigian.loc[
    Hoadontheothoigian['Tên hàng hóa/dịch vụ (*)'] == 'Phụ thu Tết', 
    'Đơn giá'
] = 27.7777

Hoadontheothoigian.loc[
    Hoadontheothoigian['Tên hàng hóa/dịch vụ (*)'] == 'Phụ thu Tết (30, Mồng 1)', 
    'Đơn giá'
] = 92.5925

# Recalculate 'Thành tiền' = Unit Price * Quantity, rounded to 0 decimals
Hoadontheothoigian['Thành tiền'] = np.where(
    Hoadontheothoigian['Số lượng'].notna() & Hoadontheothoigian['Đơn giá'].notna(),
    (Hoadontheothoigian['Đơn giá'] * Hoadontheothoigian['Số lượng']).round(0),
    np.nan
)

# Summary check for NaN prices
so_luong_nan = Hoadontheothoigian['Đơn giá'].isna().sum()
print(f"Number of NaN rows in 'Đơn giá': {so_luong_nan}")

"""## **Merge to convert system items to invoicing names**"""

# Filter for specific branches only (Security Note: Names generalized)
danh_sach_cua_hang = ['Branch_A', 'Branch_B'] 
Hoadontheothoigian = Hoadontheothoigian[Hoadontheothoigian['Cửa hàng'].isin(danh_sach_cua_hang)]

# 1. Standardize joining column on Left DataFrame
Hoadontheothoigian['Ten_HH_chuan'] = (
    Hoadontheothoigian['Tên hàng hóa/dịch vụ (*)']
    .astype(str).str.lower().str.strip()
)

# 2. Standardize joining column on Right DataFrame
output_menu_rooftop_VAT['Ten_HT_chuan'] = (
    output_menu_rooftop_VAT['Tên hệ thống']
    .astype(str).str.lower().str.strip()
)

# 3. Perform Merge to map system names to official invoicing names
Hoadontheothoigian = pd.merge(
    Hoadontheothoigian,
    output_menu_rooftop_VAT[['Tên xuất hóa đơn', 'Tên hệ thống', 'Ten_HT_chuan','VAT (%)']], 
    left_on='Ten_HH_chuan', 
    right_on='Ten_HT_chuan', 
    how='left'
)

# 4. Clean up temporary standardization columns
Hoadontheothoigian = Hoadontheothoigian.drop(columns=['Ten_HH_chuan', 'Ten_HT_chuan'])

# Exception handling for specific items with incorrect units
Hoadontheothoigian.loc[Hoadontheothoigian['Tên hệ thống'] == 'Specific_Product_Name', 'ĐVT'] = 'BOTTLE'

# --- Standardize Payment Methods using Regex ---
# Match keywords for Bank Transfers or Card payments
regex_chuyen_khoan = r'.*(TRANSFER|ATM|VISA).*'
Hoadontheothoigian['Phương thức thanh toán'] = (
    Hoadontheothoigian['Phương thức thanh toán']
    .astype(str)
    .str.replace(regex_chuyen_khoan, 'Bank Transfer', regex=True)
)

# Match COD keywords to Cash
Hoadontheothoigian['Phương thức thanh toán'] = (
    Hoadontheothoigian['Phương thức thanh toán']
    .astype(str)
    .str.replace(r'.*COD.*', 'Cash', regex=True)
)

"""# **Final Audit and Preparation for Export**"""

# Check for rows missing Invoicing Names
so_luong_nan = Hoadontheothoigian['Tên xuất hóa đơn'].isna().sum()
print(f"Rows missing 'Tên xuất hóa đơn': {so_luong_nan}")

"""## **Filter and final column selection**"""

cols_to_keep = ['Số hoá đơn','Ngày hóa đơn', 'Tên xuất hóa đơn', 'ĐVT','Số lượng','Đơn giá','Thành tiền', 'VAT (%)','Phương thức thanh toán']
Hoadontheothoigian = Hoadontheothoigian[cols_to_keep]

# Rename to final desired headers
Hoadontheothoigian = Hoadontheothoigian.rename(columns={
    'Tên xuất hóa đơn': 'Tên hàng hóa/dịch vụ (*)',
    'VAT (%)':'Thuế suất GTGT (%)'
})

# Calculate total value for Bank Transfers
tong_thanh_tien_chuyen_khoan = Hoadontheothoigian.loc[
    Hoadontheothoigian["Phương thức thanh toán"] == "Bank Transfer", 
    "Thành tiền"
].sum()

# Calculate VAT Amount column
Hoadontheothoigian['Tiền thuế GTGT'] = (
    Hoadontheothoigian['Thành tiền'] * Hoadontheothoigian['Thuế suất GTGT (%)'] / 100
)
Hoadontheothoigian

"""# **Transfer: Invoice Reconciliation**"""

Doichieuhoadon

# Keep necessary columns and rename for standardization
cols_to_keep = ['Mã hoá đơn','Số HĐ', 'Tên công ty','MST','Địa chỉ','Email']
Doichieuhoadon = Doichieuhoadon[cols_to_keep].rename(columns={
    'Tên công ty': 'Tên đơn vị mua hàng',
    'MST':'Mã số thuế',
    'Số HĐ':'Số hoá đơn'
})

# Convert Tax ID to string
Doichieuhoadon['Mã số thuế'] = Doichieuhoadon['Mã số thuế'].astype(str)

# Map specific buyer descriptions to standard format
Doichieuhoadon["Tên đơn vị mua hàng"] = Doichieuhoadon["Tên đơn vị mua hàng"].replace(
    "Người mua không lấy hóa đơn",
    "Retail Customer (No Invoice Requested)"
)
Doichieuhoadon

# 1) Generate new "Số hoá đơn" using logic: #<last 5 chars of Order ID>-<Original Invoice ID>
# Only perform concatenation where both columns are NOT NaN
mask = Doichieuhoadon['Mã hoá đơn'].notna() & Doichieuhoadon['Số hoá đơn'].notna()

Doichieuhoadon.loc[mask, 'Số hoá đơn'] = (
    '#' +
    Doichieuhoadon.loc[mask, 'Mã hoá đơn'].str[-5:] + # Last 5 characters
    '-' +
    Doichieuhoadon.loc[mask, 'Số hoá đơn'].astype(str).str.strip()
)

# 2) Assign NaN if data is insufficient
Doichieuhoadon.loc[~mask, 'Số hoá đơn'] = np.nan

# 3) Cleanup: Replace any 'nan-nan' strings with actual NaN
Doichieuhoadon['Số hoá đơn'] = Doichieuhoadon['Số hoá đơn'].replace('nan-nan', np.nan)

# Drop redundant column
Doichieuhoadon = Doichieuhoadon.drop(columns=['Mã hoá đơn'])
Doichieuhoadon

"""# **Merge: Invoices over time and Reconciliation data**"""

Thong_tin_xuat_VAT = pd.merge(
    Hoadontheothoigian,
    Doichieuhoadon,
    on='Số hoá đơn',
    how='left'
)
Thong_tin_xuat_VAT

# Logic: Keep 'Bank Transfer' records; for 'Cash' records, only keep those with VAT invoicing info

# 1. Define filtering conditions:
# Condition 1: (Payment Method is 'Cash' AND Tax ID is NaN) -> These are excluded
dieu_kien_loai_1 = (Thong_tin_xuat_VAT['Phương thức thanh toán'] == 'Cash') & \
                   (Thong_tin_xuat_VAT['Mã số thuế'].isna())

# 2. Combine exclusion conditions
cac_dong_can_loai = dieu_kien_loai_1

# 3. Filter DataFrame: Keep rows that do NOT satisfy the exclusion condition
Thong_tin_xuat_VAT = Thong_tin_xuat_VAT[~cac_dong_can_loai]

Thong_tin_xuat_VAT

# Review unique payment methods remaining
unique_values = Thong_tin_xuat_VAT['Phương thức thanh toán'].unique()
print(unique_values)

# --- Standardize Address Column ---
# 1. Replace placeholder dots ('.') with NaN for filling
Thong_tin_xuat_VAT['Địa chỉ'] = Thong_tin_xuat_VAT['Địa chỉ'].replace('.', np.nan)

# 2. Fill all NaN addresses with a standard notice
Thong_tin_xuat_VAT['Địa chỉ'] = Thong_tin_xuat_VAT['Địa chỉ'].fillna('Buyer did not provide address')

# Standardize payment method string for combined types
Thong_tin_xuat_VAT['Phương thức thanh toán'] = 'Cash/Bank Transfer'

Thong_tin_xuat_VAT

# Add supplementary columns for compliance reporting
Thong_tin_xuat_VAT['Người mua hàng'] = 'Retail Customer (No Invoice Requested)'
Thong_tin_xuat_VAT['Số điện thoại'] = np.nan
Thong_tin_xuat_VAT['Căn cước công dân'] = np.nan # ID Card/Citizen ID

Thong_tin_xuat_VAT

# Define final column sequence for export
new_order = [
    'Số hoá đơn','Ngày hóa đơn','Tên đơn vị mua hàng','Mã số thuế','Địa chỉ',
    'Người mua hàng','Email','Số điện thoại','Căn cước công dân',
    'Phương thức thanh toán','Tên hàng hóa/dịch vụ (*)', 'ĐVT', 'Số lượng', 
    'Đơn giá', 'Thành tiền', 'Thuế suất GTGT (%)','Tiền thuế GTGT'
]

# Reindex and sort data by date and invoice number
Thong_tin_xuat_VAT = Thong_tin_xuat_VAT.reindex(columns=new_order)
Thong_tin_xuat_VAT = Thong_tin_xuat_VAT.sort_values(by=['Ngày hóa đơn','Số hoá đơn'], ascending=True)

# Data cleaning for Tax ID column
Thong_tin_xuat_VAT['Mã số thuế'] = Thong_tin_xuat_VAT['Mã số thuế'].replace('.', np.nan)

# --- Generate Sequential Invoice Index ---
# 1. Use .factorize() to assign a unique integer ID to each unique invoice number
ma_so_hoa_don, _ = Thong_tin_xuat_VAT['Số hoá đơn'].factorize()

# 2. Create 'Invoice Sequence Number (*)' starting from 1
Thong_tin_xuat_VAT['Số thứ tự hóa đơn (*)'] = ma_so_hoa_don + 1

# Move Sequence Number to the first column position
cols = Thong_tin_xuat_VAT.columns.tolist()
cols.remove('Số thứ tự hóa đơn (*)')
cols.insert(0, 'Số thứ tự hóa đơn (*)')
Thong_tin_xuat_VAT = Thong_tin_xuat_VAT[cols]

# Reset index for clean export
Thong_tin_xuat_VAT = Thong_tin_xuat_VAT.reset_index(drop=True)

Thong_tin_xuat_VAT

"""# **Auto-push data to Google Sheets**"""

# Note: Security placeholders used for spreadsheet URLs
from google.colab import auth
auth.authenticate_user()

import gspread
from google.auth import default
from gspread_dataframe import set_with_dataframe

creds, _ = default()
gc = gspread.authorize(creds)

# --- PUSH MENU VAT DATA ---
spreadsheet_menu = gc.open_by_url("https://docs.google.com/spreadsheets/d/[PROTECTED_ID_1]")
worksheet_menu = spreadsheet_menu.get_worksheet(0)
worksheet_menu.clear()

set_with_dataframe(
    worksheet_menu,
    output_menu_rooftop_VAT,
    include_index=False,
    include_column_header=True,
    row=1
)
print("✅ Menu data successfully overwritten.")

# --- PUSH FINAL INVOICE DATA ---
spreadsheet_final = gc.open_by_url("https://docs.google.com/spreadsheets/d/[PROTECTED_ID_2]")
worksheet_final = spreadsheet_final.get_worksheet(0)
worksheet_final.clear()

set_with_dataframe(
    worksheet_final,
    Thong_tin_xuat_VAT,
    include_index=False,
    include_column_header=True,
    row=1
)
print("✅ Final invoice data successfully overwritten.")

"""# **Export to Excel file**"""

Thong_tin_xuat_VAT.to_excel('Invoicing_Export_Final.xlsx', index=False)
