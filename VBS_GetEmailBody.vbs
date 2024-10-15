Dim objOutlook
Dim objNamespace
Dim objMailItem
Dim objShell
Dim objFileSystem
Dim strBody

' Get the path to the .msg file from UiPath argument
strMsgFilePath = WScript.Arguments(0)


' Create Outlook and Shell objects
Set objOutlook = CreateObject("Outlook.Application")
Set objShell = CreateObject("WScript.Shell")

' Get the Namespace and open the .msg file
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objMailItem = objNamespace.OpenSharedItem(strMsgFilePath)

' Get the email body
strBody = objMailItem.Body

' Close the .msg file
objMailItem.Close olDiscard

' Release objects
Set objMailItem = Nothing
Set objNamespace = Nothing
Set objOutlook = Nothing
'Pass the argument out
WScript.Echo strBody