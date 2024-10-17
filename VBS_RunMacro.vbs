Set argObjExcelApp = GetObject(,"Excel.Application")

' argObjExcelApp.DisplayAlerts = False
' argObjExcelApp.ScreenUpdating = False
' argObjExcelApp.Visible = False

boolWorkbookFound = False

' Loop through all workbooks, including unsaved ones
For Each wbWorkbook In argObjExcelApp.Workbooks
  
  ' Check if workbook title (without extension) contains "Market_data_Eikon_Dep_Libor" (case-insensitive)
  If UCase(wbWorkbook.Name) = UCase("Market_data_Eikon_Dep_Libor.xlsm") Then
    Set wbWorkbook = wbWorkbook
    boolWorkbookFound = True
    msgbox "Found excel: "&UCase(wbWorkbook.name)
	' Run the macro named "sheet1.checkSynchronization"
    'wbWorkbook.Sheets("sheet1").Run "checkSynchronization"
	msgbox "running macro"
	argObjExcelApp.Run "Sheet2.Run_Click"  'this part works!!!!
	'argObjExcelApp.Run "Macro1"
	msgbox "success"
	Exit For

  

  End If
Next

If Not boolWorkbookFound Then
  ' Handle case where no matching workbook found
  ' (e.g., display message, exit sub)
  MsgBox "Workbook 'Market_data_Eikon_Dep_Libor' not found.", vbCritical
End If

' Close Excel Application (Optional)
' argObjExcelApp.Quit
