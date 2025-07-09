# Excel Project(s)

The first project listed in this repository, named VBA_AutoSum_Formatting_YearlyReport demonstrates the use of macros and VBA in excel.
The goal of the project is automation.
The script allows the user to easily navigate through different sheets of the wrokbook with user form that contains a drop down of all sheets in the workbook.
Additionally the user form has a button that allows the creation of new sheets in he workbook and their renaming.
Likewise, the script allows the user to import files into excel, by prompting the user to (multi) select files to import.
The user form also has a button allowing the user to run the yearly report on data imported into the workbook given the same core staructure.
The script automates the addition of appropriate headers (given similarly structured data) and formatting of the financial data as currency.
Running the yearly report will compile all data from the sheets in the workbook and paste it into the yearly report sheet.
The yearly report is also decorated with headers and formatted appropriately.
The script contains conditions which will prevent errors like double generation of a header on top of a set of data already decorated and formatted.
Hence, adding new data and running the yearly report will only format those fles which have no format.
