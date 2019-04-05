Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\BM\Fun Team.xlsm")

objExcel.Application.Visible = True
'objExcel.Workbooks.Add
'objExcel.Cells(1, 1).Value = "Test value"
'objExcel.ActiveWorkbook.Save "C:\BM\Fun Team.xlsm"

WScript.Sleep 10000

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
'WScript.Echo "Birthday Wishes Finished."
WScript.Quit
