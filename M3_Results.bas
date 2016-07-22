Attribute VB_Name = "M3_Results"
Option Explicit

Public Function CreateResults()

    If WSExists(WbNameGradeSheet) Then
        If makeSure Then
    
            Call Init
            
            Application.DisplayAlerts = False
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
                
            ' Sheet erzeugen
            If WSExists(WbNamePrintSheet) Then
                Worksheets(WbNamePrintSheet).Delete
            End If
            Worksheets.Add(Before:=Worksheets(WbNameConfig)).Name = WbNamePrintSheet
            Worksheets(WbNamePrintSheet).Tab.color = gClrTabPrint
            
             '------------------------------------
            ' Get number of sections
            '------------------------------------
            ' Count actual sheets
            Dim i, sheetCnt As Integer
            sheetCnt = 0
            For i = 0 To CfgMaxSheets
                If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value) Then
                    sheetCnt = (sheetCnt + 1)
                End If
            Next i
            
            '------------------------------------
            ' Druckseite erzeugen
            '------------------------------------
            Call PaintPrintPage(sheetCnt)
                   
            '------------------------------------
            ' Seitenumbrüche und Druckbereich
            '------------------------------------
            With Worksheets(WbNamePrintSheet).PageSetup
                .Orientation = xlLandscape
                .FitToPagesWide = 1
                .PrintArea = "A1:Q" & CStr(5 + (sheetCnt * 4) + (gNumOfPupils) * ((4 * (sheetCnt + 1)) + 2))
                .LeftMargin = Application.CentimetersToPoints(1)
                .RightMargin = Application.CentimetersToPoints(1)
                .TopMargin = Application.CentimetersToPoints(1)
                .BottomMargin = Application.CentimetersToPoints(1)
                .CenterHorizontally = True
            End With
            For i = 1 To gNumOfPupils
                Worksheets(WbNamePrintSheet).HPageBreaks.Add Before:=Worksheets(WbNamePrintSheet).Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), 1)
            Next i
            
            '------------------------------------
            ' Druckdialog öffnen
            '------------------------------------
            Dim doPrint As Integer
            doPrint = MsgBox("Möchten Sie drucken?", vbQuestion + vbOKCancel, "Hau raus!")
            If doPrint = vbOK Then
                Application.Dialogs(xlDialogPrint).Show
            End If
            
            Application.DisplayAlerts = True
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            
        End If
    Else
        ' Notenblatt existiert nicht!!
        MsgBox ("Es existiert kein Notenblatt! Erst Tabellen erzeugen!")
    End If
    
End Function

Private Function makeSure() As Boolean
    
    makeSure = False
    
    Dim found As Boolean
    ' Prüfen ob Sheet schon existiert
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = WbNamePrintSheet Then
            ' Abfragen ob wirklich neue Tabellen erstellt werden sollen...
            Dim Request As Integer
            Request = MsgBox("Sicher dass sie die Druckseite neu erstellen wollen?" & vbCrLf & "Es sollten keine Daten verloren gehen, da alle Daten automatisch eingesammelt werden! Manuell auf der Druckseite vorgenommene Modifikationen werden jedoch überschrieben!", vbExclamation + vbOKCancel, "Sicher?")
            If Request = vbCancel Then
                Exit Function
            End If
            Exit For
        End If
    Next ws
    
    ' Not exited -> sure
    makeSure = True

End Function

Private Function PaintPrintPage(sheetCnt As Integer)
    
    Dim i, u, p As Integer
    
    '------------------------------------
    ' Set row heights
    '------------------------------------
    Worksheets(WbNamePrintSheet).Rows("1:1000").RowHeight = 15
    
    '------------------------------------
    ' Set column widths
    '------------------------------------
    ' Set column widths for sub exercises and for sum column
    Worksheets(WbNamePrintSheet).Columns(1).ColumnWidth = 16.71
    For i = 0 To 15
        Worksheets(WbNamePrintSheet).Columns(2 + i).ColumnWidth = 5.57
    Next i
        
    '------------------------------------
    ' Header
    '------------------------------------
    For i = 0 To gNumOfPupils - 1
        Worksheets(WbNamePrintSheet).Range(Cells(1 + i * ((4 * ((sheetCnt + 1))) + 2), 1), Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), 17)).Select
        Call setBorder(False, True, True, True, True, xlThin, 0, True)
        With Selection.Font
            .Size = 12
            .Bold = True
        End With
        Worksheets(WbNamePrintSheet).Range(Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), CfgPrintNameCol), Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), 12)).Select
        With Selection
            .HorizontalAlignment = xlCenterAcrossSelection
        End With
        ' Abi
        'Worksheets(WbNamePrintSheet).Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), 1).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiTitle).Column).Address, "$")(1) & CStr(Range(CfgAbiTitle).Row) & "&"" ""&" & "YEAR('" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiDate).Column).Address, "$")(1) & CStr(Range(CfgAbiDate).Row) & ")"
        Worksheets(WbNamePrintSheet).Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), 1).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiTitle).Column).Address, "$")(1) & CStr(Range(CfgAbiTitle).row) & "&"" ""&" & "TEXT('" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiDate).Column).Address, "$")(1) & CStr(Range(CfgAbiDate).row) & ",""TT.MM.JJJJ"")"
        ' Name
        Worksheets(WbNamePrintSheet).Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), CfgPrintNameCol).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstPupi).Column + 1).Address, "$")(1) & CStr(Range(CfgFirstPupi).row + i) & "&"", ""&'" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstPupi).Column + 2).Address, "$")(1) & CStr(Range(CfgFirstPupi).row + i)
        ' Kurs
        Worksheets(WbNamePrintSheet).Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), 17).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiTeacher).Column).Address, "$")(1) & CStr(Range(CfgAbiTeacher).row) & "&"", Kurs ""&'" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiClass).Column).Address, "$")(1) & CStr(Range(CfgAbiClass).row)
        Worksheets(WbNamePrintSheet).Cells(1 + i * ((4 * (sheetCnt + 1)) + 2), 17).Select
        Call setBorder(False, False, True, True, True, xlThin, 0, True, xlRight)
        ' Notenbereich formatieren
        Worksheets(WbNamePrintSheet).Range(Cells(2 + i * ((4 * (sheetCnt + 1)) + 2), 2), Cells(1 + (4 * (sheetCnt + 1)) + i * ((4 * (sheetCnt + 1)) + 2), 17)).HorizontalAlignment = xlCenter
    Next i
    
    '------------------------------------
    ' Noten eintragen
    '------------------------------------
    Dim PunkteMaxSet, PunkteMaxAct As Integer
    For i = 0 To gNumOfPupils - 1
        PunkteMaxSet = 0
        PunkteMaxAct = 0
        For u = 0 To sheetCnt
            
            If u < sheetCnt Then
                ' Teilbereiche in Liste eintragen
                Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 1).Font.Bold = True
                Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 1).Value = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, u * 2).Value
                Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 1).Value = "max BE"
                Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 1).Value = "erreichte BE"
                For p = 0 To Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value - 1
                    ' Einzelpunkte (untere Zeile mit SVERWEIS)
                    Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + p).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstSect).Column + u * 2).Address, "$")(1) & CStr(2 + Range(CfgFirstSect).row + p)
                    Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + p).Font.Bold = True
                    ' max BE
                    Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + p).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, 1 + Range(CfgFirstSect).Column + u * 2).Address, "$")(1) & CStr(2 + Range(CfgFirstSect).row + p)
                    ' achieved BE
                    Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + p).Formula = "=VLOOKUP(" & Split(Cells(1, CfgPrintNameCol).Address, "$")(1) & CStr(1 + i * ((4 * (sheetCnt + 1)) + 2)) & ",'" & Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, u * 2).Value & "'!$" & Split(Cells(1, CfgColStart + 1).Address, "$")(1) & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value).Address, "$")(1) & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils) & "," & p + 2 & ",0)" ' 1 Spalte vorher und nachher
                Next p
                ' Sum sign (unicode 0931)
                Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value).Value = ChrW(931)
                Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value).Font.Bold = True
                ' max BE sum
                Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, 1 + Range(CfgExerCount).Column + u * 2).Address, "$")(1) & CStr(Range(CfgExerCount).row)
                Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value).Font.Bold = True
                ' achieved sum
                Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value).Formula = "=SUM(B" & CStr(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2)) & ":" & Split(Cells(1, Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value + 1).Address, "$")(1) & CStr(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2)) & ")"
                Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value).Font.Bold = True
            Else
                ' Gesamtergebnis
                Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 1).Font.Bold = True
                Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 1).Value = "Gesamt"
                Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 1).Value = "max BE"
                Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 1).Value = "erreichte BE"
                For p = 0 To sheetCnt - 1
                    Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + p).Value = left(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, p * 2).Text, 3) & right(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, p * 2).Text, 1)
                    Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + p).Font.Bold = True
                    ' Punkte gesamt pro Teilbereich
                    Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + p).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, 1 + Range(CfgExerCount).Column + p * 2).Address, "$")(1) & CStr(Range(CfgExerCount).row)
                    ' Punkte erreicht pro Teilbereich
                    Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + p).Formula = "=VLOOKUP(" & Split(Cells(1, CfgPrintNameCol).Address, "$")(1) & CStr(1 + i * ((4 * (sheetCnt + 1)) + 2)) & ",'" & Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, p * 2).Value & "'!$" & Split(Cells(1, CfgColStart + 1).Address, "$")(1) & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, p * 2).Value).Address, "$")(1) & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils) & "," & Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, p * 2).Value + 2 & ",0)" ' 1 Spalte vorher und nachher
                Next p
                Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + sheetCnt).Value = ChrW(931)
                Worksheets(WbNamePrintSheet).Cells(3 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + sheetCnt).Font.Bold = True
                ' Gesamt erzielte Punkte
                Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + sheetCnt).Formula = "=SUM(B" & CStr(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2)) & ":" & Split(Cells(1, sheetCnt + 1).Address, "$")(1) & CStr(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2)) & ")"
                Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + sheetCnt).Font.Bold = True
                Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + sheetCnt).Value = "=SUM(B" & CStr(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2)) & ":" & Split(Cells(1, sheetCnt + 1).Address, "$")(1) & CStr(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2)) & ")"
                Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 2 + sheetCnt).Font.Bold = True
                ' Notenpunkte
                Worksheets(WbNamePrintSheet).Range(Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 16), Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 17)).Select
                Call setBorder(False, True, True, True, True, xlThin, RGB(255, 255, 255), True, xlHAlignCenterAcrossSelection, xlBottom)
                Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 16).Value = "NP"
                Worksheets(WbNamePrintSheet).Cells(4 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 16).Font.Bold = True
                Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 16).Formula = "=VLOOKUP(" & Split(Cells(1, 2 + sheetCnt).Address, "$")(1) & CStr(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2)) & "," & WbNameGradeKey & CfgVLookUpPoints & ")"
                Worksheets(WbNamePrintSheet).Cells(5 + u * 4 + i * ((4 * (sheetCnt + 1)) + 2), 16).Font.Bold = True
            End If
        Next u
        
    Next i
    
    ' Create Chart
    Dim chartRow As Integer
    chartRow = gNumOfPupils * (4 * (sheetCnt + 1) + 2) + 2
    Call AddGradeDistribution(WbNamePrintSheet, chartRow, 1)
    
End Function
