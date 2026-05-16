# Automated Sales Dashboard (Coffee Shop)

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
