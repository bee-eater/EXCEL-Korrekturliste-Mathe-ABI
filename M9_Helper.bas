Attribute VB_Name = "M9_Helper"
Option Explicit

'-----------------------------------------------------
' UNBLOCK CHECK
'-----------------------------------------------------

' Returns True if the workbook file has the Windows Zone.Identifier NTFS stream,
' which means it was downloaded and not yet unblocked by the user.
Public Function IsFileBlocked(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Zone.Identifier is an NTFS alternate data stream appended by Windows
    ' after downloading a file from the internet.
    IsFileBlocked = fso.FileExists(filePath & ":Zone.Identifier")
End Function

' Checks whether a given downloaded file is still blocked by Windows.
' Shows a dialog with step-by-step instructions and opens the folder in Explorer.
' Loops until the user has unblocked the file (Retry) or aborts (Cancel).
' Returns True if the file is ready to open, False if the user aborted.
Public Function CheckUnblockedInteractive(filePath As String) As Boolean
    Dim folderPath As String
    folderPath = Left(filePath, InStrRev(filePath, "\"))
    Do While IsFileBlocked(filePath)
        Shell "explorer.exe """ & folderPath & """", vbNormalFocus
        Dim choice As VbMsgBoxResult
        choice = MsgBox( _
            "Die heruntergeladene Datei ist durch Windows blockiert." & vbNewLine & vbNewLine & _
            "Bitte wie folgt vorgehen:" & vbNewLine & _
            "  1. Klicke im ge" & Chr(246) & "ffneten Explorer-Fenster mit der rechten Maustaste auf die Datei." & vbNewLine & _
            "  2. W" & Chr(228) & "hle ""Eigenschaften""." & vbNewLine & _
            "  3. Aktiviere unten das H" & Chr(228) & "kchen ""Zulassen"" (Unblock) und klicke OK." & vbNewLine & vbNewLine & _
            "Danach hier auf Wiederholen klicken " & Chr(150) & " oder Abbrechen, um den Update zu stoppen.", _
            vbExclamation + vbRetryCancel, "Datei blockiert " & Chr(150) & " Entsperren erforderlich")
        If choice = vbCancel Then
            CheckUnblockedInteractive = False
            Exit Function
        End If
        ' vbRetry -> re-check the stream
    Loop
    CheckUnblockedInteractive = True
End Function

' Checks whether the workbook itself is blocked. If it is, shows a message,
' opens the containing folder in Explorer so the user can unblock it,
' and closes the workbook without saving.
Public Sub CheckUnblocked()
    Dim wbPath As String
    wbPath = ThisWorkbook.FullName
    If IsFileBlocked(wbPath) Then
        MsgBox "Diese Datei ist durch Windows blockiert, da sie aus dem Internet heruntergeladen wurde." & vbNewLine & vbNewLine & _
               "Bitte wie folgt vorgehen:" & vbNewLine & _
               "  1. Schlie" & Chr(223) & "e dieses Fenster." & vbNewLine & _
               "  2. Klicke im Explorer mit der rechten Maustaste auf die Datei." & vbNewLine & _
               "  3. W" & Chr(228) & "hle " & Chr(34) & "Eigenschaften" & Chr(34) & "." & vbNewLine & _
               "  4. Aktiviere unten das H" & Chr(228) & "kchen " & Chr(34) & "Zulassen" & Chr(34) & " (Unblock)." & vbNewLine & _
               "  5. Klicke OK und " & Chr(246) & "ffne die Datei erneut.", _
               vbCritical + vbOKOnly, "Datei blockiert " & Chr(150) & " Aktion erforderlich"
        ' Open the containing folder in Explorer so the user can right-click easily
        Shell "explorer.exe /select," & Chr(34) & wbPath & Chr(34), vbNormalFocus
        ' Close without saving
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

' Shows cmdUpdateFile on the Config sheet when no segment sheets exist yet
' (i.e. the workbook is fresh / empty), hides it otherwise.
Public Sub RefreshUpdateFileButton()
    Dim ws As Worksheet
    Dim systemSheets As String
    systemSheets = "|" & WbNameConfig & "|" & WbNameGradeKey & "|" & WbNameGradeSheet & _
                   "|" & WbNamePrintSheet & "|" & WbNameSelExConfig & "|" & WbNameTestDaten & "|"
    Dim hasSegment As Boolean
    hasSegment = False
    For Each ws In ThisWorkbook.Worksheets
        If InStr(systemSheets, "|" & ws.Name & "|") = 0 Then
            hasSegment = True
            Exit For
        End If
    Next ws
    Worksheets(WbNameConfig).Unprotect Password:=WbPw
    Worksheets(WbNameConfig).OLEObjects("cmdUpdateFile").Visible = Not hasSegment
    Worksheets(WbNameConfig).Protect Password:=WbPw
    Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
End Sub

Function WSExists(n As String) As Boolean
  Dim ws As Worksheet
  WSExists = False
  For Each ws In Worksheets
    If n = ws.Name Then
      WSExists = True
      Exit Function
    End If
  Next ws
End Function

' Returns the number of sub-exercises for the named segment sheet, or 0 if not found.
Public Function GetNumOfSubEx(sheetName As String) As Integer
    Dim i As Integer
    For i = 0 To CfgMaxSheets
        If Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, i * 2).Value = sheetName Then
            GetNumOfSubEx = CInt(Worksheets(WbNameConfig).Range(CfgExerCount).offset(0, i * 2).Value)
            Exit Function
        End If
    Next i
    GetNumOfSubEx = 0
End Function

Public Function ceil(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
    ceil = -Int(-X / Factor) * Factor
End Function

Public Function floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    floor = Int(X / Factor) * Factor
End Function

' Applies fill, borders, merge and alignment directly to rng — no .Select required.
Public Sub setBorder(rng As Range, merge As Boolean, left As Boolean, right As Boolean, top As Boolean, bottom As Boolean, style As Integer, fillColor As Long, Optional edge As Boolean, Optional horAlign As Integer, Optional verAlign As Integer)
    With rng
        If fillColor <> 0 Then .Interior.color = fillColor
        If merge Then .MergeCells = True
        If horAlign <> 0 Then .HorizontalAlignment = horAlign
        If verAlign <> 0 Then .VerticalAlignment = verAlign
        If edge Then
            ' Outer edges only
            If left Then .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = style: .Borders(xlEdgeLeft).ColorIndex = 1
            If right Then .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = style: .Borders(xlEdgeRight).ColorIndex = 1
            If top Then .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = style: .Borders(xlEdgeTop).ColorIndex = 1
            If bottom Then .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = style: .Borders(xlEdgeBottom).ColorIndex = 1
        Else
            ' All borders: outer edges + inside grid lines
            If left Then .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = style: .Borders(xlEdgeLeft).ColorIndex = 1
            If right Then .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = style: .Borders(xlEdgeRight).ColorIndex = 1
            If top Then .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = style: .Borders(xlEdgeTop).ColorIndex = 1
            If bottom Then .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = style: .Borders(xlEdgeBottom).ColorIndex = 1
            If .Rows.Count > 1 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous: .Borders(xlInsideHorizontal).Weight = style: .Borders(xlInsideHorizontal).ColorIndex = 1
            If .Columns.Count > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous: .Borders(xlInsideVertical).Weight = style: .Borders(xlInsideVertical).ColorIndex = 1
        End If
    End With
End Sub

' Applies decimal validation (0..refCell) directly to rng — no .Select required.
Public Sub setUpperLimit(rng As Range, refCell As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
            :=xlBetween, Formula1:="0", Formula2:="=" & refCell
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = False
        .ShowError = True
    End With
End Sub

Public Sub TestDatenUebernehmen()

    Application.ScreenUpdating = False
    
    Dim ws As String
    ws = ActiveSheet.Name
    
    Worksheets("TestData").Unprotect Password:=WbPw

    Sheets(WbNameTestDaten).Visible = True

    Sheets(WbNameTestDaten).Select
    Range("A1:F23").Select
    Selection.Copy
    Sheets("Analysis A").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    Range("A25:J47").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Analysis B").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    Range("A49:A71").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Stochastik A").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    ActiveWindow.SmallScroll Down:=27
    Range("A73:E95").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Stochastik B").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    ActiveWindow.SmallScroll Down:=18
    Range("A97:B119").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Geometrie A").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    ActiveWindow.SmallScroll Down:=33
    Range("A121:E143").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Geometrie B").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    
    
    Sheets(WbNameTestDaten).Select
    Range("A147:E169").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ConfigW").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets(WbNameTestDaten).Select
    Range("A173:E195").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Wahlaufgaben").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
        
    Sheets(WbNameTestDaten).Visible = False
    
    Worksheets(ws).Activate

    Application.ScreenUpdating = True

End Sub

Function IsVersionGreater(v1 As String, v2 As String) As Boolean
    Dim parts1() As String, parts2() As String
    Dim i As Integer, maxLen As Integer
    Dim num1 As Integer, num2 As Integer

    ' Remove leading "v"
    v1 = Replace(v1, "v", "")
    v2 = Replace(v2, "v", "")

    parts1 = Split(v1, ".")
    parts2 = Split(v2, ".")

    maxLen = Application.WorksheetFunction.Max(UBound(parts1), UBound(parts2))

    For i = 0 To maxLen
        If i <= UBound(parts1) Then
            num1 = Val(parts1(i))
        Else
            num1 = 0
        End If

        If i <= UBound(parts2) Then
            num2 = Val(parts2(i))
        Else
            num2 = 0
        End If

        If num1 > num2 Then
            IsVersionGreater = True
            Exit Function
        ElseIf num1 < num2 Then
            IsVersionGreater = False
            Exit Function
        End If
    Next i

    IsVersionGreater = False ' equal versions
End Function

Function CheckForUpdate(currentVersion As String)
    Dim http As Object
    Dim json As Object
    Dim response As String
    Dim latestVersion As String

    On Error GoTo UpdateCheckError
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://api.github.com/repos/bee-eater/EXCEL-Korrekturliste-Mathe-ABI/releases/latest", False
    http.setRequestHeader "User-Agent", "Excel VBA"
    http.Send
        
    Worksheets(WbNameConfig).Unprotect Password:=WbPw
    Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells

    ' Check for valid response
    If http.Status = 200 Then
        response = http.responseText
        ' Parse JSON
        Set json = JsonConverter.ParseJson(response)
        latestVersion = json("tag_name") ' e.g. "v2.0.1"

        ' Compare versions
        If IsVersionGreater(latestVersion, Version) Then
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = "Update available! " + Version + " " + ChrW(8594) + " " + latestVersion
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(0, 138, 255) ' Blue
            ThisWorkbook.Sheets(WbNameConfig).OLEObjects("cmdUpdate").Visible = True
        ElseIf IsVersionGreater(Version, latestVersion) Then
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = "Pre-Release! " + Version
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(175, 80, 0) ' Orange
            ThisWorkbook.Sheets(WbNameConfig).OLEObjects("cmdUpdate").Visible = False
        Else
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = ChrW(10003) + " " + Version
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(0, 176, 80) ' Green
            ThisWorkbook.Sheets(WbNameConfig).OLEObjects("cmdUpdate").Visible = False
        End If
    Else
        ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = http.Status
        ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(255, 0, 0) ' Red for error
        ThisWorkbook.Sheets(WbNameConfig).OLEObjects("cmdUpdate").Visible = False
    End If
    
    Worksheets(WbNameConfig).Protect Password:=WbPw
    Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
    
    Exit Function
    
UpdateCheckError:
    Worksheets(WbNameConfig).Unprotect Password:=WbPw
    Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
    ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = "Error checking for updates..."
    ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(255, 0, 0) ' Red for error
    ThisWorkbook.Sheets(WbNameConfig).OLEObjects("cmdUpdate").Visible = False
    Worksheets(WbNameConfig).Protect Password:=WbPw
    Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
    
End Function

