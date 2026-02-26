# Enterprise Expense Standardization, Aggregation & Visualization Tool

This VBA automation tool is designed to streamline enterprise expense tracking and quarterly reporting by consolidating data across multiple worksheets into a single, standardized quarterly report.

_What traditionally requires manual copying, formatting, header management, and recalculation across multiple sheets is reduced to a one-click execution._

## By automating repetitive, error-prone tasks, this solution effectively decreases tedious effort and introduces consistency, scalability, and audit-readiness into financial reporting workflows

**_Run VBA Macro LaunchApp_**

1. Generates a number (user's choice) of raw data sheets, prepopulated with randomly generated data (Populate Sample Data).
2. Standardizes raw data data for a specific quarter (selected by user in a dropdown menu) - the dropdown meny cannot be used until data is populated.
3. Aggregates data (in the same event as standardizaton) into a "Quarter X Report" sheet, where X - 1, 2, 3, 4 based on user selection.
- If the user wants to generate more data, say for another quarter, the user must simply click the Populate Sample data button again. This generates more Raw data sheets (however many user requests - following the wrokflow from step 1, the previously egnerated sheets are not overwritten).
- If in following generations of new data, the quarter selected already has correspondng data generated, then the user will be given the option to overwrite data or keep existing - in that case workflow is exited.
4. There is a button to refresh workbook, so all sheets with data are deleted and there is only 1 empty sheet left (Sheet1 - to mimic default workbook set up).
5. Close button exits workflow.
6. Finally there is a Quick Analysis button.
- This launches teh Quick Analysis user form.

**_Quick Analysis Workflow_**

_User selects from combo boxes:_

- Combo box: Division - select east, west, north, south.
- Combo box: Category - select any of the possible categories available as expenses in data set.
- Check which quarters to include in the analysis: Q1, Q2, Q3, Q4 (if a quarter that has no corresponding data sheets in the workbook - fails gracefully with a message).
- Finally check what KPIs should be retrieved from Quick Analysis: Sum, Average, and Standard Deviation.

_For each selected Quarter sheet, for each row:_

- VLOOKUP retrieves the Total based on a combined key composed from Division and Category, searching only in selected quarters.
- When matches are found results are collected and processed based on what KPIs the user is looking for.
- Results are aggregated and returned to the user.

**_Pivot tables for Effective Data Analysis_**

**_PowerQuery - The M Language of PowerBI - For Data Transformation and Visualization_**

