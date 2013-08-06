Import all Excel
================

Objective:
The objective of this project is to ensure import of data from all the MS Excel files (.xlsx) stored in a predetermined folder. Through the integration of ACL Analytics 10, Visual Basic and MS Excel the ACL project will be populated with separate table for each Excel Worksheet available in each of the workbook excluding the empty worksheets. 

Project Files:
There are total three files in this project:
•  ImportAllExcelInFolder.ACL - ACL Project prepared in AN10
•	Import_Excel.vbs – A vbscript which is to be called by ACL script
•	Contest_One.xlsm – A micro enabled MS Excel file
 
Pre-requisite:
To properly achieve the objective of importing data, we have to setup the environment as follows:
a)	The VBA Project requires setting up of references to the following:
-	Visual Basic For Applications
-	Microsoft Excel 12.0 Object Library
-	OLE Automation
-	Microsoft Office 12.0 Object Library
-	Microsoft Scripting Runtime
b)	A window folder namely ‘MY Folder’ created on the C drive
c)	You have to place Import_Excel.vbs, Contest_One.xlsm and all MS Excel files needs to be imported in this folder (C:\MY Folder);
d)	ACL project included in this project is to be kept in a single separate window folder
e)	Default ACL Options have been used, Tools -> Options -> Factory… 

Limitations:
This uses old syntax of IMPORT EXCEL command which includes CHARMAX parameter instead of FIELDS syntax as in available in version 9.3 onward. Resultantly we cannot ignore to import any field or exercise control over the type and format of different fields. 

By default the data files (.FIL) created as a result of IMPORT command will be stored in the folder in which ACL project has been saved. 

To get complete and correct data imported from Excel you need to have a clean spreadsheet. The vbscript is not intended to clean up the individual worksheets. 

Main logic:
The ACL scripts start to develop a blank table layout for holding the names of the workbooks and worksheets. Through EXECUTE command a vbscript is called which runs macros stored in a macro enabled workbook. The macro open each workbook stored in the predetermined folder (C:\MY Folder) and remove any empty worksheets in it. It then collects names of the workbook and all the non empty worksheets in a variable and write it to a text file.

ACL script then opens this text file through the table layout developed earlier. GROUP – LOOP structure then write a .BAT file containing IMPORT EXCEL command for each worksheet. Resulting table names are structured to include the workbook name and a serial number appended to it. At last the script commands stored in the .BAT file are executed through DO SCRIPT command and worksheets are imported to ACL as separate tables.
Setting VBA Project References:
There may be times, especially when copying a workbook from one machine to another, that the reference for that VBA Project are not properly carried  through to the new machine. This normally happens when the project is using a DLL that is not part of the core set of DLLs that make up the VBA environment for Office. A Reference in VBA Project is a “pointer” to a type library or DLL file that defines various objects and entities and the properties and methods of those entities.

In the VBA editor go to Tools menu and chose References to display the Reference dialog. Scroll down in the list and put a check next to the required item in the list. Click OK.

Clean spread sheets:
1.	Make a copy of the original spreadsheet and only work with the copy if changes are necessary to prepare it for import to ACL. 
2.	Assure the column titles cover only the first row. Column titles are truncated at 31 characters. 
3.	Remove blank lines or note lines, all rows must be uniform. 
4.	Remove subtotals or totals. 
5.	Confirm properties of each column; format if necessary. 

When importing from Excel, also consider that: 
•	Formulas will be resolved and imported (values are imported, not expressions). 
•	Hidden columns and rows are imported. 
•	Blank rows and columns that exist between populated data are imported. 
•	If special characters exist in column headings, they are replaced with underscores in the corresponding ACL field name. 
