This is a simple macro to extract data from MySQL database and load it on excel spreadsheet.
It contains three macros;

1. Extract list of tables from the database
2. Load data from selected table on excel spreadsheet
3. Clear loaded data and list of tables from excel spreadsheet

Copy MySQL ClassicModels database in below path. Please note, path might be different on your computer.

C:\ProgramData\MySQL\MySQL Server 8.0\Data

Enter server details on Worksheet 1. Please read steps on Worksheet 1 for more details.

ClassicModels.xlsx is connected to MySQL database using MS Query which joins data from three tables namely;
customers, orders, and employees. Data refreshes every minute.
