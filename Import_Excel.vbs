Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'E:\AN10Contest\Contest_One.xlsm'!OpenAllExcelInFolder"
objExcel.ActiveWorkbook.Save 
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
WScript.Quit
