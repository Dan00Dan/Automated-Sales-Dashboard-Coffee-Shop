import pandas as pd
import xlsxwriter
from datetime import datetime

print("Reading data...")
df = pd.read_excel('Coffee Shop Sales.xlsx')

# Clean and calculate
print("Processing data...")
df['Revenue'] = df['transaction_qty'] * df['unit_price']
# Convert transaction_date to datetime if it's not
df['transaction_date'] = pd.to_datetime(df['transaction_date'])
df['Month'] = df['transaction_date'].dt.strftime('%Y-%m')

# 1. KPIs
total_revenue = df['Revenue'].sum()
total_orders = df['transaction_id'].nunique() # Assuming transaction_id is unique per order line, actually each line is a transaction. Let's just use count.
total_orders = len(df)
avg_order_val = total_revenue / total_orders if total_orders > 0 else 0

# 2. Revenue by Date (Line Chart data)
rev_by_date = df.groupby('transaction_date')['Revenue'].sum().reset_index()
rev_by_date['transaction_date'] = rev_by_date['transaction_date'].dt.strftime('%Y-%m-%d')

# 3. Revenue by Category (Bar Chart data)
rev_by_category = df.groupby('product_category')['Revenue'].sum().reset_index().sort_values(by='Revenue', ascending=False)

# 4. Revenue by Store (Pie Chart data)
rev_by_store = df.groupby('store_location')['Revenue'].sum().reset_index()

# Create Excel writer
output_file = 'Coffee_Shop_Dashboard_VN.xlsx'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = writer.book

print("Creating Dashboard...")
# Add Dashboard Sheet
dash_sheet = workbook.add_worksheet('Bảng Điều Khiển')

# Formats
title_format = workbook.add_format({'bold': True, 'font_size': 20, 'font_color': '#ffffff', 'bg_color': '#4A2311', 'align': 'center', 'valign': 'vcenter'})
kpi_title_format = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#4A2311', 'align': 'center'})
kpi_val_format = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#D2691E', 'align': 'center', 'num_format': '#,##0.00'})

# Set column widths
dash_sheet.set_column('A:A', 2)
dash_sheet.set_column('B:D', 20)
dash_sheet.set_column('E:H', 15)

# Title
dash_sheet.merge_range('B2:H3', '☕ BẢNG ĐIỀU KHIỂN DOANH THU QUÁN CÀ PHÊ', title_format)

# KPIs
dash_sheet.write('B5', 'Tổng Doanh Thu', kpi_title_format)
dash_sheet.write('B6', total_revenue, kpi_val_format)

dash_sheet.write('D5', 'Tổng Đơn Hàng', kpi_title_format)
dash_sheet.write('D6', total_orders, kpi_val_format)

dash_sheet.write('F5', 'Trung Bình / Đơn', kpi_title_format)
dash_sheet.write('F6', avg_order_val, kpi_val_format)

# Write summary data to a hidden sheet to power the charts
summary_sheet = workbook.add_worksheet('SummaryData')
summary_sheet.hide()

# Write data for Revenue by Date
summary_sheet.write_row('A1', ['Date', 'Revenue'])
summary_sheet.write_column('A2', rev_by_date['transaction_date'])
summary_sheet.write_column('B2', rev_by_date['Revenue'])

# Write data for Revenue by Category
summary_sheet.write_row('D1', ['Category', 'Revenue'])
summary_sheet.write_column('D2', rev_by_category['product_category'])
summary_sheet.write_column('E2', rev_by_category['Revenue'])

# Write data for Revenue by Store
summary_sheet.write_row('G1', ['Store', 'Revenue'])
summary_sheet.write_column('G2', rev_by_store['store_location'])
summary_sheet.write_column('H2', rev_by_store['Revenue'])

# Add Charts
# 1. Line Chart (Revenue Trend)
line_chart = workbook.add_chart({'type': 'line'})
line_chart.add_series({
    'name': 'Revenue',
    'categories': ['SummaryData', 1, 0, len(rev_by_date), 0],
    'values':     ['SummaryData', 1, 1, len(rev_by_date), 1],
    'line':       {'color': '#D2691E'}
})
line_chart.set_title({'name': 'Xu Hướng Doanh Thu Theo Thời Gian'})
line_chart.set_x_axis({'name': 'Ngày'})
line_chart.set_y_axis({'name': 'Doanh Thu ($)'})
line_chart.set_legend({'none': True})
dash_sheet.insert_chart('B9', line_chart, {'x_scale': 1.8, 'y_scale': 1.2})

# 2. Bar Chart (Revenue by Category)
bar_chart = workbook.add_chart({'type': 'column'})
bar_chart.add_series({
    'name': 'Revenue',
    'categories': ['SummaryData', 1, 3, len(rev_by_category), 3],
    'values':     ['SummaryData', 1, 4, len(rev_by_category), 4],
    'fill':       {'color': '#8B4513'}
})
bar_chart.set_title({'name': 'Doanh Thu Theo Danh Mục'})
bar_chart.set_legend({'none': True})
dash_sheet.insert_chart('B28', bar_chart, {'x_scale': 1.2, 'y_scale': 1.1})

# 3. Pie Chart (Revenue by Store)
pie_chart = workbook.add_chart({'type': 'pie'})
pie_chart.add_series({
    'name': 'Revenue',
    'categories': ['SummaryData', 1, 6, len(rev_by_store), 6],
    'values':     ['SummaryData', 1, 7, len(rev_by_store), 7],
})
pie_chart.set_title({'name': 'Tỷ Trọng Doanh Thu Theo Cửa Hàng'})
dash_sheet.insert_chart('F28', pie_chart, {'x_scale': 1.0, 'y_scale': 1.1})


# Write Raw Data for Filtering
print("Writing Raw Data...")
df.to_excel(writer, sheet_name='Dữ Liệu Gốc', index=False)

# Add Filter to Raw Data
worksheet = writer.sheets['Dữ Liệu Gốc']
worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

# Format Headers in Raw Data
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)

writer.close()
print(f"Done! Dashboard saved to: {output_file}")
