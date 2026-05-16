# Automated Sales Dashboard (Coffee Shop)

[🇻🇳 Xem bản Tiếng Việt ở dưới](#vn-bản-tiếng-việt)

## 📌 Introduction
This project automates the creation of a comprehensive Sales Dashboard for a Coffee Shop using Python. It transforms raw transactional data into an interactive Excel Dashboard with key performance indicators (KPIs) and visual charts, completely eliminating the need for manual Excel manipulation.

## 👤 Author
- **Dan** ([@Dan00Dan](https://github.com/Dan00Dan))

## 📊 Data Description
The dataset used in this project is sourced from Kaggle: [Coffee Shop Sales](https://www.kaggle.com/datasets/ahmedabbas757/coffee-sales).
It contains real-world transactional data from a coffee shop, including:
- **transaction_date & transaction_time**: When the order was placed.
- **transaction_qty & unit_price**: Quantity sold and the price per item.
- **store_location**: The branch where the sale occurred.
- **product_category & product_type**: The classification of the sold item (e.g., Coffee, Tea, Bakery).

*Note: The raw data contains over 150,000 rows, which is processed and aggregated efficiently using Pandas.*

## ⚙️ How It Works (Methodology)
1. **Data Extraction & Cleaning**: 
   - The Python script (`create_dashboard.py`) loads the raw `Coffee Shop Sales.xlsx` file using `pandas`.
   - It calculates the `Revenue` for each transaction (Quantity × Unit Price) and extracts date information for trend analysis.
2. **Data Aggregation**:
   - The script groups the data to calculate KPIs: Total Revenue, Total Orders, and Average Order Value.
   - It also prepares summarized DataFrames for Revenue by Date, Revenue by Category, and Revenue by Store.
3. **Dashboard Automation**:
   - Using the `xlsxwriter` library, the script programmatically creates a new Excel file (`Coffee_Shop_Dashboard.xlsx`).
   - It inserts formatted KPI text and generates native Excel Charts (Line Chart, Column Chart, Pie Chart).
   - Finally, it outputs the raw data with Excel AutoFilters enabled, allowing users to interactively slice and dice the data.

## 🚀 How to Run
1. Install the required Python libraries:
   ```bash
   pip install pandas xlsxwriter openpyxl
   ```
2. Place the `Coffee Shop Sales.xlsx` dataset in the same directory.
3. Run the script:
   ```bash
   python create_dashboard.py
   ```
4. Open the newly generated `Coffee_Shop_Dashboard.xlsx` to view your dashboard!

---

<a name="vn-bản-tiếng-việt"></a>
# 🇻🇳 Bản Tiếng Việt

## 📌 Giới thiệu
Dự án này tự động hóa quá trình tạo Dashboard phân tích doanh thu cho một quán cà phê bằng Python. Hệ thống sẽ biến dữ liệu giao dịch thô thành một Dashboard Excel trực quan với các chỉ số kinh doanh quan trọng (KPIs) và biểu đồ sinh động, loại bỏ hoàn toàn việc phải thao tác thủ công trên Excel.

## 👤 Tác giả
- **Dan** ([@Dan00Dan](https://github.com/Dan00Dan))

## 📊 Mô tả dữ liệu
Tập dữ liệu sử dụng trong dự án được lấy từ Kaggle: [Coffee Shop Sales](https://www.kaggle.com/datasets/ahmedabbas757/coffee-sales).
Đây là dữ liệu giao dịch thực tế của một quán cà phê, bao gồm:
- **transaction_date & transaction_time**: Ngày và giờ đặt hàng.
- **transaction_qty & unit_price**: Số lượng bán và đơn giá.
- **store_location**: Chi nhánh diễn ra giao dịch.
- **product_category & product_type**: Phân loại sản phẩm (VD: Cà phê, Trà, Bánh ngọt).

*Lưu ý: Dữ liệu thô gồm hơn 150.000 dòng, được xử lý và tổng hợp cực kỳ nhanh chóng thông qua thư viện Pandas.*

## ⚙️ Cách hệ thống hoạt động
1. **Trích xuất và Làm sạch dữ liệu**: 
   - Script Python (`create_dashboard.py`) sẽ đọc file dữ liệu thô `Coffee Shop Sales.xlsx` thông qua `pandas`.
   - Tính toán `Doanh thu (Revenue)` cho từng giao dịch (Số lượng × Đơn giá) và trích xuất ngày tháng để phân tích xu hướng.
2. **Tổng hợp dữ liệu**:
   - Hệ thống nhóm dữ liệu để tính các chỉ số KPIs: Tổng doanh thu, Tổng số đơn hàng, và Giá trị trung bình mỗi đơn.
   - Chuẩn bị sẵn dữ liệu phân nhóm theo Ngày, theo Danh mục sản phẩm và theo Chi nhánh.
3. **Tự động hóa Dashboard**:
   - Sử dụng thư viện `xlsxwriter`, script sẽ tự lập trình xuất ra một file Excel mới (`Coffee_Shop_Dashboard.xlsx`).
   - Tự động chèn các chỉ số KPIs với định dạng màu sắc đẹp mắt và vẽ các biểu đồ (Biểu đồ đường, Biểu đồ cột, Biểu đồ tròn).
   - Cuối cùng, in ra bảng dữ liệu gốc kèm theo tính năng Bộ lọc (AutoFilters), giúp người dùng dễ dàng lọc và tìm kiếm dữ liệu.

## 🚀 Hướng dẫn sử dụng
1. Cài đặt các thư viện Python cần thiết:
   ```bash
   pip install pandas xlsxwriter openpyxl
   ```
2. Đặt file dữ liệu `Coffee Shop Sales.xlsx` vào cùng thư mục với script.
3. Chạy lệnh:
   ```bash
   python create_dashboard.py
   ```
4. Mở file `Coffee_Shop_Dashboard.xlsx` vừa được sinh ra để xem thành quả!
