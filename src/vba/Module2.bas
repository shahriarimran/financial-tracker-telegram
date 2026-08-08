Attribute VB_Name = "Module2"
Option Explicit

'==========================================================
' TELEGRAM CONFIGURATION
'
' Expected public project layout:
'
' <project root>\
'   workbook\
'       Financial_Tracker_Template.xlsm
'   config\
'       config.ini
'       config.example.ini
'   scripts\
'
' config.ini should contain:
'
' TELEGRAM_BOT_TOKEN=your_token
' TELEGRAM_CHAT_ID=your_chat_id
'
'==========================================================


'==========================================================
' RETURNS PROJECT ROOT
'==========================================================
Private Function GetProjectRoot() As String

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ThisWorkbook.Path:
    ' <project root>\workbook
    '
    ' Therefore parent folder:
    ' <project root>

    GetProjectRoot = fso.GetParentFolderName(ThisWorkbook.Path)

    Set fso = Nothing

End Function


'==========================================================
' READS A VALUE FROM config\config.ini
'==========================================================
Private Function GetConfigValue(ByVal keyName As String) As String

    Dim configPath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim eqPos As Long
    Dim currentKey As String
    Dim currentValue As String
    Dim fileIsOpen As Boolean

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo ErrorHandler

    configPath = GetProjectRoot() & "\config\config.ini"

    '------------------------------------------------------
    ' Make sure configuration file exists
    '------------------------------------------------------
    If Dir(configPath) = "" Then

        Err.Raise vbObjectError + 3000, _
                  "GetConfigValue", _
                  "Missing configuration file: " & configPath

    End If

    '------------------------------------------------------
    ' Open configuration file
    '------------------------------------------------------
    fileNum = FreeFile

    Open configPath For Input As #fileNum

    fileIsOpen = True

    '------------------------------------------------------
    ' Read file one line at a time
    '------------------------------------------------------
    Do Until EOF(fileNum)

        Line Input #fileNum, lineText

        lineText = Trim$(lineText)

        ' Ignore blank lines
        If Len(lineText) > 0 Then

            ' Ignore comments beginning with # or ;
            If Left$(lineText, 1) <> "#" And _
               Left$(lineText, 1) <> ";" Then

                ' Split only at first "="
                eqPos = InStr(1, lineText, "=")

                If eqPos > 0 Then

                    currentKey = Trim$(Left$(lineText, eqPos - 1))
                    currentValue = Trim$(Mid$(lineText, eqPos + 1))

                    ' Case-insensitive key comparison
                    If StrComp( _
                            currentKey, _
                            keyName, _
                            vbTextCompare) = 0 Then

                        Close #fileNum
                        fileIsOpen = False

                        If Len(currentValue) = 0 Then

                            Err.Raise vbObjectError + 3001, _
                                      "GetConfigValue", _
                                      "Configuration value is blank: " & _
                                      keyName

                        End If

                        GetConfigValue = currentValue

                        Exit Function

                    End If

                End If

            End If

        End If

    Loop

    Close #fileNum
    fileIsOpen = False

    '------------------------------------------------------
    ' Requested key was not found
    '------------------------------------------------------
    Err.Raise vbObjectError + 3002, _
              "GetConfigValue", _
              "Configuration key not found: " & keyName

    Exit Function


ErrorHandler:

    ' Preserve original error before cleanup
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next

    If fileIsOpen Then
        Close #fileNum
    End If

    On Error GoTo 0

    Err.Raise errNumber, _
              errSource, _
              errDescription

End Function


'==========================================================
' DECIDES WHAT TO SEND
'==========================================================
Public Sub SendTelegramAlert()

    Dim ws As Worksheet
    Dim messageText As String
    Dim hasAlert As Boolean

    Dim goldSignal As String
    Dim silverSignal As String
    Dim usdSignal As String

    Set ws = ThisWorkbook.Worksheets("Financial Tracker")

    '------------------------------------------------------
    ' Read current signals
    '------------------------------------------------------
    goldSignal = UCase$(Trim$(CStr(ws.Range("H9").Value)))
    silverSignal = UCase$(Trim$(CStr(ws.Range("I9").Value)))
    usdSignal = UCase$(Trim$(CStr(ws.Range("J9").Value)))

    '------------------------------------------------------
    ' Message header
    '------------------------------------------------------
    messageText = _
        "Financial Tracker" & vbCrLf & _
        Format$(Date, "yyyy-mm-dd") & _
        vbCrLf & vbCrLf

    hasAlert = False


    '======================================================
    ' GOLD
    '======================================================
    If goldSignal = "BUY" Or goldSignal = "SELL" Then

        messageText = messageText & _
                      "GOLD - " & goldSignal & vbCrLf

        ' BUY = price user pays
        ' SELL = price user receives
        If goldSignal = "BUY" Then

            messageText = messageText & _
                          "Price: " & _
                          ws.Range("H2").text & vbCrLf

        Else

            messageText = messageText & _
                          "Price: " & _
                          ws.Range("H3").text & vbCrLf

        End If

        messageText = messageText & _
                      "DailyChange: " & _
                      ws.Range("H5").text & vbCrLf & _
                      "Trend: " & _
                      ws.Range("H6").text & vbCrLf & _
                      "Volatility: " & _
                      ws.Range("H7").text & vbCrLf & _
                      "Drawdown: " & _
                      ws.Range("H8").text & _
                      vbCrLf & vbCrLf

        hasAlert = True

    End If


    '======================================================
    ' SILVER
    '======================================================
    If silverSignal = "BUY" Or silverSignal = "SELL" Then

        messageText = messageText & _
                      "SILVER - " & silverSignal & vbCrLf

        If silverSignal = "BUY" Then

            messageText = messageText & _
                          "Price: " & _
                          ws.Range("I2").text & vbCrLf

        Else

            messageText = messageText & _
                          "Price: " & _
                          ws.Range("I3").text & vbCrLf

        End If

        messageText = messageText & _
                      "DailyChange: " & _
                      ws.Range("I5").text & vbCrLf & _
                      "Trend: " & _
                      ws.Range("I6").text & vbCrLf & _
                      "Volatility: " & _
                      ws.Range("I7").text & vbCrLf & _
                      "Drawdown: " & _
                      ws.Range("I8").text & _
                      vbCrLf & vbCrLf

        hasAlert = True

    End If


    '======================================================
    ' USD
    '======================================================
    If usdSignal = "BUY" Or usdSignal = "SELL" Then

        messageText = messageText & _
                      "USD - " & usdSignal & vbCrLf

        If usdSignal = "BUY" Then

            messageText = messageText & _
                          "Price: " & _
                          ws.Range("J2").text & vbCrLf

        Else

            messageText = messageText & _
                          "Price: " & _
                          ws.Range("J3").text & vbCrLf

        End If

        messageText = messageText & _
                      "DailyChange: " & _
                      ws.Range("J5").text & vbCrLf & _
                      "Trend: " & _
                      ws.Range("J6").text & vbCrLf & _
                      "Volatility: " & _
                      ws.Range("J7").text & vbCrLf & _
                      "Drawdown: " & _
                      ws.Range("J8").text & _
                      vbCrLf & vbCrLf

        hasAlert = True

    End If


    '======================================================
    ' HOLD = NO MESSAGE
    '======================================================
    If Not hasAlert Then

        Debug.Print _
            Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
            " | No actionable Telegram signal."

        Exit Sub

    End If


    '======================================================
    ' SEND CONSTRUCTED MESSAGE
    '======================================================
    SendTelegramMessage messageText

End Sub


'==========================================================
' ACTUALLY SENDS THE MESSAGE
'==========================================================
Private Sub SendTelegramMessage(ByVal messageText As String)

    Dim http As Object
    Dim url As String

    Dim botToken As String
    Dim chatId As String

    '------------------------------------------------------
    ' Load private credentials at runtime
    '------------------------------------------------------
    botToken = GetConfigValue("TELEGRAM_BOT_TOKEN")
    chatId = GetConfigValue("TELEGRAM_CHAT_ID")

    '------------------------------------------------------
    ' Construct Telegram Bot API request
    '------------------------------------------------------
    url = _
        "https://api.telegram.org/bot" & botToken & _
        "/sendMessage?chat_id=" & chatId & _
        "&text=" & URLEncode(messageText)

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False
    http.Send

    '------------------------------------------------------
    ' Validate Telegram response
    '------------------------------------------------------
    If http.Status <> 200 Then

        Err.Raise vbObjectError + 2000, _
                  "SendTelegramMessage", _
                  "Telegram HTTP error: " & _
                  CStr(http.Status)

    End If

    Debug.Print _
        Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
        " | Telegram alert sent successfully."

    Set http = Nothing

End Sub


'==========================================================
' URL ENCODING HELPER
'
' Suitable for the current English/ASCII notification text.
'==========================================================
Private Function URLEncode(ByVal text As String) As String

    Dim i As Long
    Dim c As String
    Dim result As String
    Dim asciiCode As Long

    For i = 1 To Len(text)

        c = Mid$(text, i, 1)
        asciiCode = AscW(c)

        Select Case asciiCode

            ' 0-9, A-Z, a-z
            Case 48 To 57, _
                 65 To 90, _
                 97 To 122

                result = result & c

            ' Space
            Case 32

                result = result & "%20"

            ' URL-safe characters:
            ' - . _ ~
            Case 45, 46, 95, 126

                result = result & c

            Case Else

                result = result & _
                         "%" & _
                         Right$("0" & Hex$(asciiCode), 2)

        End Select

    Next i

    URLEncode = result

End Function

