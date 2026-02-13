# Coffee-Sales-Analysis-Excel
## Project Overview
This project involves building an end-to-end interactive Excel dashboard to analyze coffee sales data. The goal was to transform raw sales records into actionable insights, visualizing Total Sales over time, Sales by Country, and top customers.

## Files
*   **Coffee Sales Dashboard.xlsx**: The final Excel file containing the raw data, processing steps, and the interactive dashboard.

## Key Features & Skills Applied

### 1. Data Gathering & Transformation
*   **Data Integration:** Consolidated data from three distinct worksheets ("Orders", "Customers", and "Products") into a single master dataset.
*   **Advanced Formulas:**
    *   **XLOOKUP:** Used to retrieve customer information (Customer Name, Email, Country) based on Customer IDs.
    *   **INDEX MATCH:** Implemented a dynamic double-match formula to populate product details (Coffee Type, Roast Type, Size, Unit Price) ensuring scalability if column positions change.
    *   **IF Functions:** Created nested IF formulas to expand abbreviations (e.g., converting "Rob" to "Robusta") and handle missing data errors cleanly.

### 2. Data Cleaning & Modeling
*   **Formatting:** Applied custom number formatting to display unit sizes (e.g., "0.5 kg") and currency, and standardized date formats to "DD-MMM-YYYY" for readability.
*   **Duplicate Removal:** Performed data validation checks to ensure no duplicate transaction records existed.
*   **Excel Tables:** Converted the raw data range into a structured Excel Table ("Orders Table") to allow for automatic data range updates when refreshing Pivot Tables.

### 3. Dashboard Design & Visualization
*   **Pivot Tables & Charts:** Built Pivot Tables to aggregate data and created corresponding visualizations:
    *   **Line Chart:** Total Sales over Time (grouped by Month/Year).
    *   **Bar Charts:** Sales by Country and Top 5 Customers.
*   **Interactivity:**
    *   **Timeline:** Added a timeline filter to analyze sales trends across specific time periods.
    *   **Slicers:** Integrated slicers for "Roast Type," "Size," and "Loyalty Card" status.
    *   **Report Connections:** Linked all slicers and timelines to every chart, ensuring the entire dashboard updates dynamically based on user selection.

## Snapshots
https://github.com/aryaapareek/Coffee-Sales-Analysis-Excel/blob/main/CoffeeSalesAnalysisDashboard.png
