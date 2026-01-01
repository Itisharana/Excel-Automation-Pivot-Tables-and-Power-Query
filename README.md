# Excel-Automation-Pivot-Tables-and-Power-Query
User Guide – Excel Sales Automation Dashboard
________________________________________
1️ How to Open the File
•	Locate the Excel file Excel_Sales_Automation.xlsx on your system.
•	Double-click the file to open it in Microsoft Excel (2019 or later).
•	If a security warning appears, click Enable Content to allow data connections.
•	Wait a few seconds for all queries and connections to load completely.
________________________________________
2️ How to Refresh the Data
Use this process whenever the source file sales_data.csv is updated.
•	Open Excel_Sales_Automation.xlsx.
•	Navigate to the Data tab in the Excel ribbon.
•	Click Refresh All.
•	Wait until the refresh process finishes successfully.
  The following components update automatically:
•	Power Query transformations
•	Data Model (Power Pivot)
•	Pivot Tables
•	Charts and visualizations (if available)
  Do not manually edit any pivot table values.
________________________________________
3️ How to Read the Pivot Tables
•	Rows represent categories such as Product, Region, or Date.
•	Values represent calculated metrics such as Total Sales or Quantity Sold.
•	Use filter dropdowns to analyze specific regions, products, or time periods.
•	Pivot Tables are dynamic and refresh automatically after data updates.
     Example Insights You Can Derive:
•	Sales performance by product
•	Regional sales comparison
•	Monthly or yearly sales trends
________________________________________
4️ Important Instructions (Do Not Modify)
 To ensure automation, accuracy, and stability of the dashboard:
 Do not edit Power Query steps
 Do not delete or rename any queries
 Do not modify the Data Model
 Do not manually overwrite Pivot Table results
 Only update the source CSV file and use Refresh All.
________________________________________
5️ Technical Notes
•	Data is imported, cleaned, and transformed using Power Query.
•	All transformation logic is automated and stored as Applied Steps.
•	Pivot Tables are connected to the Excel Data Model (Power Pivot).
•	The file is designed for scalable, repeatable, and automated reporting.
________________________________________
Support & Troubleshooting
If the data does not refresh correctly:
•	Verify that sales_data.csv exists in the original folder location.
•	Ensure that column names in the CSV have not been changed.
•	Confirm that the CSV file is closed before refreshing.
•	Contact the file creator or dashboard owner for assistance.

