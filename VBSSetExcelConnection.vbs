 Set argObjExcelApp = GetObject(,"Excel.Application")

  argObjExcelApp.DisplayAlerts = False
  argObjExcelApp.ScreenUpdating = False
  argObjExcelApp.Visible = False

  boolWorkbookFound = False

  ' Loop through all workbooks, including unsaved ones
  For Each wbWorkbook In argObjExcelApp.Workbooks
    msgbox wbWorkbook.name
	' Check if workbook title (without extension) contains "Book1" (case-insensitive)
    If UCase(wbWorkbook.Name) = UCase("Book1") Then
	'If InStr(1, UCase(Left(wbWorkbook.Name, wbWorkbook.Name)), UCase("Book1")) > 0 Then
      Set wbWorkbook = wbWorkbook
      boolWorkbookFound = True
      
	  Exit For
    End If
  Next

  If Not boolWorkbookFound Then
    ' Handle case where no matching workbook found
    ' (e.g., display message, exit sub)
    MsgBox "Workbook 'Book1' not found.", vbCritical
  End If

  ' ... rest of your code using 'wbWorkbook'