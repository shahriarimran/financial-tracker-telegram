Option Explicit
Dim fso, projectRoot, workbookPath, excelApp, workbook
Set fso = CreateObject("Scripting.FileSystemObject")
projectRoot = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
workbookPath = fso.BuildPath(projectRoot, "workbook\Financial_Tracker_Template.xlsm")
If Not fso.FileExists(workbookPath) Then
    WScript.Echo "Workbook not found: " & workbookPath
    WScript.Quit 2
End If
On Error Resume Next
Set excelApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Could not start Excel."
    WScript.Quit 3
End If
excelApp.Visible = False
excelApp.DisplayAlerts = False
Set workbook = excelApp.Workbooks.Open(workbookPath)
excelApp.Run "'" & workbook.Name & "'!DailyAutoRun"
If Err.Number <> 0 Then
    WScript.Echo "DailyAutoRun failed."
    workbook.Close False
    excelApp.Quit
    WScript.Quit 4
End If
On Error GoTo 0
workbook.Save
workbook.Close False
excelApp.Quit
Set workbook = Nothing
Set excelApp = Nothing
Set fso = Nothing
