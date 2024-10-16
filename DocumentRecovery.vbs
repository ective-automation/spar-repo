Option Explicit

Dim objExcel, objWorkbook, objDocumentRecoveryPane, excelPath

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

' Open the provided Excel file
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(excelPath)
If Err.Number <> 0 Then
    WScript.Echo "Failed to open the Excel file: " & excelPath
    WScript.Quit
End If
On Error GoTo 0

' Make Excel visible (optional)
objExcel.Visible = True

' Check if the Document Recovery Pane is open and close it
If Not objExcel Is Nothing Then
    For Each objDocumentRecoveryPane In objExcel.CommandBars
        If InStr(1, objDocumentRecoveryPane.Name, "Document Recovery") > 0 Then
            objDocumentRecoveryPane.Enabled = False ' Disable the pane
            objDocumentRecoveryPane.Visible = False ' Hide the pane
        End If
    Next
End If

' Clean up
Set objDocumentRecoveryPane = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
