# SQL_INSERT

This script generates mySQL insert statements from an Excel file. The scheme name is derived from the name of the Excel file. In the
example provided the name of the Excel file is SQLinsert-Basketbal-.xlsx, therefor the scheme name is Basketbal. If no quotes must
be present in the values in the SQL script, then you must append '-NQ' to the column names. The table names are equal to the Excel 
names of the tabs. You must change the properties of all cells to 'text' right after you add a new tab. The Excel file must be 
selected using a file dialog. The script produces a text file in the same directory with the SQL INSERT statements.

Install openpyxl using pip if needed.

Hans Verkerk September 2019
