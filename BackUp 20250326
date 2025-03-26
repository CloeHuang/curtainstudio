import os
import pandas as pd
from datetime import datetime

# === Function to convert .xls to .xlsx ===
def convert_xls_to_xlsx(xls_path, xlsx_path):
    try:
        import xlrd
    except ImportError:
        raise ImportError("Please install xlrd with: pip install xlrd")
    df = pd.read_excel(xls_path, engine='xlrd')
    df.to_excel(xlsx_path, index=False)

# === Function to process incoming stock data ===
def process_incoming_stock(file_path):
    print("\n📥 Reading incoming stock data...")
    df = pd.read_csv(file_path, encoding='ISO-8859-1', low_memory=False)
    df['Duedate'] = pd.to_datetime(df['Duedate'], format='%d/%m/%Y', errors='coerce')
    df['Stockcode'] = df['Stockcode'].astype(str)
    df['YearMonth'] = df['Duedate'].dt.to_period('M').astype(str)
    grouped = df.groupby(['Stockcode', 'YearMonth'])['Ord_quant'].sum().reset_index()
    print("✅ Incoming stock processed.")
    return grouped

# === Pivot incoming stock by month ===
def incoming_by_month(df):
    print("\n📊 Pivoting incoming stock by month...")
    pivot = df.pivot_table(index='Stockcode', columns='YearMonth', values='Ord_quant', aggfunc='sum', fill_value=0).reset_index()
    print("✅ Pivot complete.")
    return pivot

# === Load individual datasets ===
def process_sku_list(file_path):
    print("\n📥 Loading SKU list...")
    return pd.read_csv(file_path, low_memory=False)

def process_stock_value(file_path):
    print("\n📥 Loading stock value...")
    return pd.read_csv(file_path, low_memory=False)

def concat_sales_files(folder):
    print("\n📁 Combining sales files...")
    files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.csv')]
    combined = pd.concat([pd.read_csv(f, low_memory=False) for f in files], ignore_index=True)
    print(f"✅ {len(files)} sales files combined.")
    return combined

def mtd_sales(df):
    print("\n📆 Calculating month-to-date sales...")
    df['TransactionDate'] = pd.to_datetime(df['TransactionDate'], dayfirst=True, errors='coerce')
    df['Month'] = df['TransactionDate'].dt.to_period('M')
    result = df.groupby('Month').sum(numeric_only=True).reset_index()
    print("✅ MTD sales calculated.")
    return result

def rolling_averages(df):
    print("\n📈 Calculating 3-month rolling averages...")
    df['TransactionDate'] = pd.to_datetime(df['TransactionDate'], dayfirst=True, errors='coerce')
    df['Month'] = df['TransactionDate'].dt.to_period('M')
    grouped = df.groupby('Month').sum(numeric_only=True).reset_index()
    if 'InvoiceValueTaxExclusive' not in grouped.columns:
        raise ValueError("❌ 'InvoiceValueTaxExclusive' not found in columns: " + ", ".join(grouped.columns))
    grouped['Rolling3M'] = grouped['InvoiceValueTaxExclusive'].rolling(window=3).mean()
    print("✅ Rolling averages done.")
    return grouped

def committed_file(file_path):
    print("\n📥 Reading committed sales file...")
    return pd.read_excel(file_path, engine='openpyxl')

def sales_last_13_months(df):
    print("\n📆 Getting last 13 months of sales...")
    df['TransactionDate'] = pd.to_datetime(df['TransactionDate'], dayfirst=True, errors='coerce')
    df['Month'] = df['TransactionDate'].dt.to_period('M')
    result = df.groupby('Month').sum(numeric_only=True).tail(13).reset_index()
    print("✅ Last 13 months extracted.")
    return result

# === Main Execution ===

incoming_stock_file = "Incoming.csv"
sku_list_file = "SKU_List_1.csv"
stock_value_file = "Stock_Value.csv"
sales_folder = "Sales"
committed_sales_file = "Com.XLS"
converted_committed_file = "Com.xlsx"
output_file = "Final_Merged_Data.xlsx"

# Convert XLS if needed
if committed_sales_file.endswith(".XLS") and not os.path.exists(converted_committed_file):
    convert_xls_to_xlsx(committed_sales_file, converted_committed_file)
    committed_sales_file = converted_committed_file
else:
    committed_sales_file = converted_committed_file if os.path.exists(converted_committed_file) else committed_sales_file

# Load all data
incoming = process_incoming_stock(incoming_stock_file)
incoming_pivot = incoming_by_month(incoming)
sku_list = process_sku_list(sku_list_file)
stock_value = process_stock_value(stock_value_file)
sales_combined = concat_sales_files(sales_folder)
sales_mtd = mtd_sales(sales_combined)
sales_rolling = rolling_averages(sales_combined)
sales_committed = committed_file(committed_sales_file)
sales_13_months = sales_last_13_months(sales_combined)

# Save all sheets in one file using xlsxwriter
print("\n💾 Writing all sheets to Excel using xlsxwriter...")
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    sku_list.to_excel(writer, sheet_name='Sku_list', index=False)
    incoming_pivot.to_excel(writer, sheet_name='IncomingByMonth', index=False)
    sales_rolling.to_excel(writer, sheet_name='Ave_Sales', index=False)
    sales_mtd.to_excel(writer, sheet_name='Mtd_Sales', index=False)
    sales_committed.to_excel(writer, sheet_name='Sales_Committed', index=False)
    stock_value.to_excel(writer, sheet_name='SOH', index=False)
    sales_13_months.to_excel(writer, sheet_name='Sales_by_Month', index=False)

print("\n✅ Done – Final_Merged_Data.xlsx has been successfully generated with xlsxwriter!")