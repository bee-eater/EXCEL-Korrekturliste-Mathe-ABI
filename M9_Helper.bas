Attribute VB_Name = "M9_Helper"
Option Explicit

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

' Sourcecode exportieren für Versionsverwaltung
Public Function ExportSourceFiles()
    Dim destPath As String
    destPath = Application.ActiveWorkbook.Path & "\"
    Dim component As VBComponent
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.Export destPath & component.Name & ToFileExtension(component.Type)
        End If
    Next

End Function

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    
    ' Dateiendung entsprechend Typ zurückgeben
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select

End Function

Public Function ceil(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor
End Function

Public Function floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    floor = Int(X / Factor) * Factor
End Function

Public Function setBorder(merge As Boolean, left As Boolean, right As Boolean, top As Boolean, bottom As Boolean, style As Integer, color As Long, Optional edge As Boolean, Optional horAlign As Integer, Optional verAlign As Integer)
    ' Farbe übernehmen
    If color <> 0 Then
        With Selection
            .Interior.color = color
        End With
    End If
    ' mergen?
    If merge Then
        With Selection
            .MergeCells = True
        End With
    End If
    ' Alignment übergeben?
    If horAlign <> 0 Then
        With Selection
            .HorizontalAlignment = horAlign
        End With
    End If
    ' Alignment übergeben?
    If verAlign <> 0 Then
        With Selection
            .VerticalAlignment = verAlign
        End With
    End If
    If edge Then
        ' Left border?
        If left Then
            With Selection
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = style
                .Borders(xlEdgeLeft).ColorIndex = 1
            End With
        End If
        ' Rigth border?
        If right Then
            With Selection
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = style
                .Borders(xlEdgeRight).ColorIndex = 1
            End With
        End If
        ' Top border?
        If top Then
            With Selection
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = style
                .Borders(xlEdgeTop).ColorIndex = 1
            End With
        End If
        ' Bottom border?
        If bottom Then
            With Selection
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = style
                .Borders(xlEdgeBottom).ColorIndex = 1
            End With
        End If
    Else
        ' Left border?
        If left Then
            With Selection
                .Borders(xlLeft).LineStyle = xlContinuous
                .Borders(xlLeft).Weight = style
                .Borders(xlLeft).ColorIndex = 1
            End With
        End If
        ' Rigth border?
        If right Then
            With Selection
                .Borders(xlRight).LineStyle = xlContinuous
                .Borders(xlRight).Weight = style
                .Borders(xlRight).ColorIndex = 1
            End With
        End If
        ' Top border?
        If top Then
            With Selection
                .Borders(xlTop).LineStyle = xlContinuous
                .Borders(xlTop).Weight = style
                .Borders(xlTop).ColorIndex = 1
            End With
        End If
        ' Bottom border?
        If bottom Then
            With Selection
                .Borders(xlBottom).LineStyle = xlContinuous
                .Borders(xlBottom).Weight = style
                .Borders(xlBottom).ColorIndex = 1
            End With
        End If
    End If

End Function

Public Function setUpperLimit(refCell As String)

    With Selection.Validation
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
    
End Function

Public Sub TestDatenUebernehmen()

    Application.ScreenUpdating = False
    
    Dim ws As String
    ws = ActiveSheet.Name
    
    Worksheets("TestData").Unprotect Password:=WbPw

    Sheets(WbNameTestDaten).Visible = True

    Sheets(WbNameTestDaten).Select
    Range("A1:I23").Select
    Selection.Copy
    Sheets("Analysis A").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    Range("A25:L47").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Analysis B").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    Range("A49:C71").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Stochastik A").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    ActiveWindow.SmallScroll Down:=27
    Range("A73:F95").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Stochastik B").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    ActiveWindow.SmallScroll Down:=18
    Range("A97:D119").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Geometrie A").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Sheets(WbNameTestDaten).Select
    ActiveWindow.SmallScroll Down:=33
    Range("A121:F143").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Geometrie B").Select
    Range("D7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    
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
        ElseIf IsVersionGreater(Version, latestVersion) Then
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = "Futuristic! Nice! " + Version + " " + ChrW(8592) + " " + latestVersion
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(0, 176, 80) ' Green
        Else
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = ChrW(10003) + " " + Version
            ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(0, 176, 80) ' Green
        End If
    Else
        ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = http.Status
        ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(255, 0, 0) ' Red for error
    End If
    
    Worksheets(WbNameConfig).Protect Password:=WbPw
    Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
    
    Exit Function
    
UpdateCheckError:
    Worksheets(WbNameConfig).Unprotect Password:=WbPw
    Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
    ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Value = "Error checking for updates..."
    ThisWorkbook.Sheets(WbNameConfig).Range(CfgUpdateInfo).Font.color = RGB(255, 0, 0) ' Red for error
    Worksheets(WbNameConfig).Protect Password:=WbPw
    Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
    
End Function



