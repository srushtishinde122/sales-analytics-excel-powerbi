# Sales Analytics Dashboard — Excel + Power BI

## Project objective
Analyze sales, profit, products, categories, and regional performance using a sample retail dataset. The data is synthetic and intended for learning/portfolio demonstration.

## Files
- `sales_data.csv` — source dataset (600 sample orders)
- `Sales_Analytics_Dashboard.xlsx` — Excel workbook with source data, region/category summaries, and monthly trend charts
- `build_project.py` — runnable Python source to rebuild the Excel workbook
- `PowerBI_DAX_and_Visuals.txt` — DAX measures and dashboard build plan

## Run the Python source
1. Install Python 3.10+.
2. Open a terminal in this folder.
3. Run: `pip install pandas openpyxl numpy`
4. Run: `python build_project.py`
5. The Excel workbook will be created/replaced in this folder.

## Build the Power BI report
1. Open Power BI Desktop.
2. Select **Home > Get data > Text/CSV** and import `sales_data.csv`.
3. Click **Transform Data** if you need to check data types; set `Order_Date` to Date and numeric fields to Decimal/Whole Number.
4. Select **Close & Apply**.
5. Create the DAX measures in `PowerBI_DAX_and_Visuals.txt`.
6. Build the visuals listed there and add slicers for date, region, and category.
7. Save the report as `Sales_Analytics_Dashboard.pbix`.

Note: Power BI Desktop's `.pbix` report is created inside Power BI Desktop; this package provides the data, runnable Excel-building code, and DAX instructions rather than a prebuilt `.pbix` file.
