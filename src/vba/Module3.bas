Attribute VB_Name = "Module3"
Sub AppendDailyPrices()
    Dim logSheet As Worksheet
    Dim frontendSheet As Worksheet
    Dim lastRow As Long
    Dim currentDate As String

    Set logSheet = ThisWorkbook.Sheets("Daily Log")
    Set frontendSheet = ThisWorkbook.Sheets("Financial Tracker")

    ' Get today's date
    currentDate = Format(Date, "yyyy-mm-dd")

    ' Check if today's date already logged
    If Application.WorksheetFunction.CountIf(logSheet.Range("A:A"), currentDate) > 0 Then
        Debug.Print "Today's prices already logged.", vbExclamation
        Exit Sub
    End If

    ' Find next empty row in log
    lastRow = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row + 1

    ' Write data
    logSheet.Cells(lastRow, 1).Value = currentDate
    logSheet.Cells(lastRow, 2).Value = frontendSheet.Range("H2").Value ' Gold price
    logSheet.Cells(lastRow, 3).Value = frontendSheet.Range("I2").Value ' Silver price
    logSheet.Cells(lastRow, 4).Value = frontendSheet.Range("J2").Value ' USD price

    Debug.Print "Prices logged successfully.", vbInformation
End Sub


