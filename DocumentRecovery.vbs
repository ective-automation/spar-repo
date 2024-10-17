' you may have to switch the order of "Check and close the document recovery Pane" with "Handle any recovery popups"

Option Explicit


Dim objExcel, objWorkbook, objDocumentRecoveryPane, excelPath, dialog

' Check if the script has received an argument for the Excel file path
If WScript.Arguments.Count = 0 Then
    WScript.Echo "No Excel file path was provided."
    WScript.Quit
End If

' Get the Excel file path from the argument
excelPath = WScript.Arguments(0)

' Create Excel Application object
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    Set objExcel = CreateObject("Excel.Application")
End If
On Error GoTo 0

' Set DisplayAlerts and EnableEvents to False to suppress pop-ups and events
objExcel.DisplayAlerts = False
objExcel.EnableEvents = False

' Open the provided Excel file
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(excelPath)
If Err.Number <> 0 Then
    WScript.Echo "Failed to open the Excel file: " & excelPath
    WScript.Quit
End If
On Error GoTo 0

' Make Excel visible (optional)
objExcel.Visible = False

' Check and close the Document Recovery Pane if it exists
If Not objExcel Is Nothing Then
    For Each objDocumentRecoveryPane In objExcel.CommandBars
        If InStr(1, objDocumentRecoveryPane.Name, "Document Recovery") > 0 Then
            objDocumentRecoveryPane.Enabled = False ' Disable the pane
            objDocumentRecoveryPane.Visible = False ' Hide the pane
        End If
    Next
End If

' Handle any recovery-related pop-ups
On Error Resume Next
For Each dialog In objExcel.Dialogs
    If InStr(1, dialog.Name, "Document Recovery") > 0 Then
        dialog.Hide ' Automatically close any document recovery-related dialogs
    End If
Next
On Error GoTo 0

' Save the workbook after closing Document Recovery Pane
objWorkbook.Save

' Close the workbook (optional) but leave Excel running
objWorkbook.Close False

' Restore DisplayAlerts and EnableEvents to their default (True)
objExcel.DisplayAlerts = True
objExcel.EnableEvents = True

' Clean up
Set objDocumentRecoveryPane = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
