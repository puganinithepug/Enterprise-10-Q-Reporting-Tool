# Enterprise Expense Standardization, Aggregation & Visualization Tool

The goal of the project is automation for an enterprise's expenses. Allowing a one-click VBA Macro execution of data standardization and data aggregation for reporting. Additionally, PowerBI is leveraged to enable automated data visualization for robust insights. 

**Accessible Tool Navigation**
- The script allows the user to easily navigate through different sheets of the workbook via user form that contains a drop down of all sheets in the workbook.

**Using the Tool**
- Additionally the user form has a button that allows the creation of new sheets in he workbook and their renaming.
Likewise, the script allows the user to import files into excel, by prompting the user to (multi) select files to import.
- The user form also has a button allowing the user to run the yearly report on data imported into the workbook given the same core staructure.

**Automated Data Standardization**
- The script automates the addition of appropriate headers (given similarly structured data) and formatting of the financial data as currency.

**Data Aggregation**
- Running the yearly report will compile all data from the sheets in the workbook and paste it into the yearly report sheet.
The aggregated data sheet produced is also formatted.

**Fault Tolerance**
- The script contains conditions which will prevent errors like double generation of a header on top of a set of data already decorated and formatted. Hence, adding new data and running the yearly report will only format those fles which have no format.

