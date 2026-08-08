Attribute VB_Name = "Module5"
Option Explicit

Public Sub DailyAutoRun()

    Dim wsFT As Worksheet
    Dim wsLog As Worksheet

    Dim refreshOK As Boolean
    Dim usingFallback As Boolean

    Dim goldPrice As Double
    Dim silverPrice As Double
    Dim usdPrice As Double

    Dim fallbackDate As Date
    Dim priceSource As String

    On Error GoTo ErrorHandler

    Set wsFT = ThisWorkbook.Worksheets("Financial Tracker")
    Set wsLog = ThisWorkbook.Worksheets("Daily Log")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Refreshing financial data..."

    '========================================================
    ' 1. TRY REFRESH ALL
    '========================================================
    refreshOK = TryRefreshAll()

    '========================================================
    ' 2. IF REFRESH WORKED, USE LIVE PRICES
    '========================================================
    If refreshOK And LivePricesValid(wsFT) Then

        goldPrice = CDbl(wsFT.Range("H2").Value)
        silverPrice = CDbl(wsFT.Range("I2").Value)
        usdPrice = CDbl(wsFT.Range("J2").Value)

        priceSource = "LIVE"
        usingFallback = False

        Application.StatusBar = _
            "Live prices refreshed successfully."

        'Save refreshed prices to historical Daily Log
        Call AppendDailyPrices

    Else

        '====================================================
        ' 3. REFRESH FAILED -> USE LATEST DAILY LOG VALUES
        '====================================================
        Application.StatusBar = _
            "Refresh failed. Using Daily Log fallback..."

        If Not GetLatestDailyLogPrices( _
                    wsLog, _
                    fallbackDate, _
                    goldPrice, _
                    silverPrice, _
                    usdPrice) Then

            Err.Raise vbObjectError + 1000, _
                      "DailyAutoRun", _
                      "Refresh failed and no valid Daily Log fallback exists."

        End If

        priceSource = _
            "DAILY LOG (" & _
            Format(fallbackDate, "yyyy-mm-dd") & ")"

        usingFallback = True

    End If

    '========================================================
    ' 4. RECALCULATE WORKBOOK
    '========================================================
    Application.CalculateFull

    '========================================================
    ' 5. DEBUG / AUDIT OUTPUT
    '========================================================
    Debug.Print "======================================"
    Debug.Print "DailyAutoRun: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Debug.Print "Source: " & priceSource
    Debug.Print "Gold: " & goldPrice
    Debug.Print "Silver: " & silverPrice
    Debug.Print "USD: " & usdPrice
    Debug.Print "======================================"

    '========================================================
    ' 6. TELEGRAM
    '
    ' IMPORTANT:
    ' Only send normal BUY/SELL signals when live refresh
    ' actually succeeded.
    '
    ' Daily Log fallback is historical data, not confirmed
    ' fresh market data.
    '========================================================
    If Not usingFallback Then

        Call SendTelegramAlert

    Else

        Debug.Print _
            "Telegram BUY/SELL skipped because fallback data is being used."

    End If

    '========================================================
    ' 7. SAVE
    '========================================================
    ThisWorkbook.Save

CleanExit:

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    Exit Sub


ErrorHandler:

    Debug.Print _
        Format(Now, "yyyy-mm-dd hh:mm:ss") & _
        " | DailyAutoRun ERROR " & _
        Err.Number & _
        " | " & Err.Description

    Resume CleanExit

End Sub


'============================================================
' TRY REFRESH ALL
' Returns TRUE if Excel completed RefreshAll without an error.
'============================================================
Private Function TryRefreshAll() As Boolean

    On Error GoTo RefreshFailed

    ThisWorkbook.RefreshAll

    'Wait for Power Query / asynchronous connections
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo RefreshFailed

    Application.CalculateFull

    TryRefreshAll = True
    Exit Function

RefreshFailed:

    Err.Clear
    TryRefreshAll = False

End Function


'============================================================
' CHECK LIVE BUY PRICES
'
' Financial Tracker:
' H2 = Gold Buy
' I2 = Silver Buy
' J2 = USD Buy
'============================================================
Private Function LivePricesValid(ByVal ws As Worksheet) As Boolean

    LivePricesValid = False

    If IsError(ws.Range("H2").Value) Then Exit Function
    If IsError(ws.Range("I2").Value) Then Exit Function
    If IsError(ws.Range("J2").Value) Then Exit Function

    If Not IsNumeric(ws.Range("H2").Value) Then Exit Function
    If Not IsNumeric(ws.Range("I2").Value) Then Exit Function
    If Not IsNumeric(ws.Range("J2").Value) Then Exit Function

    If CDbl(ws.Range("H2").Value) <= 0 Then Exit Function
    If CDbl(ws.Range("I2").Value) <= 0 Then Exit Function
    If CDbl(ws.Range("J2").Value) <= 0 Then Exit Function

    LivePricesValid = True

End Function


'============================================================
' READ MOST RECENT DAILY LOG ENTRY
'
' A = Date
' B = Gold Buy
' C = Silver Buy
' D = USD Buy
'============================================================
Private Function GetLatestDailyLogPrices( _
    ByVal ws As Worksheet, _
    ByRef logDate As Date, _
    ByRef goldPrice As Double, _
    ByRef silverPrice As Double, _
    ByRef usdPrice As Double) As Boolean

    Dim lastRow As Long

    GetLatestDailyLogPrices = False

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Row 1 contains headers
    If lastRow < 2 Then Exit Function

    If Not IsDate(ws.Cells(lastRow, "A").Value) Then Exit Function

    If Not IsNumeric(ws.Cells(lastRow, "B").Value) Then Exit Function
    If Not IsNumeric(ws.Cells(lastRow, "C").Value) Then Exit Function
    If Not IsNumeric(ws.Cells(lastRow, "D").Value) Then Exit Function

    If CDbl(ws.Cells(lastRow, "B").Value) <= 0 Then Exit Function
    If CDbl(ws.Cells(lastRow, "C").Value) <= 0 Then Exit Function
    If CDbl(ws.Cells(lastRow, "D").Value) <= 0 Then Exit Function

    logDate = CDate(ws.Cells(lastRow, "A").Value)

    goldPrice = CDbl(ws.Cells(lastRow, "B").Value)
    silverPrice = CDbl(ws.Cells(lastRow, "C").Value)
    usdPrice = CDbl(ws.Cells(lastRow, "D").Value)

    GetLatestDailyLogPrices = True

End Function

