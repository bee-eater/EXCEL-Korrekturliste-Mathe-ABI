Attribute VB_Name = "M2_Tables"
Option Explicit

Public Function CreateTables()
    
    'If Not EnsureVBAccess() Then
    '    Exit Function
    'End If
    
    Call Init
    
    If makeSure Then
    
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
         
        Call PaintSegmentPages
        Call FillSegmentPages
        Call PaintGradePage
        Call FillGradePage
                
        ' Wenn es Wahlaufgaben gibt, Konfigseite erzeugen und anzeigen
        If CheckForSelEx() Then
            ' Funktionen für Wahlaufgaben -> Config erstellen
            ' Auf Konfig-Seite gibt es dann einen Button für das Update des Tabs
            Call PaintSelXCfgPage
            Call FillSelXCfgPage
            Worksheets(WbNameSelExConfig).Activate
            Worksheets(WbNameSelExConfig).Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx).Select
        Else
            ' Löschen wenn es ein ConfigW Blatt gibt
            If WSExists(WbNameSelExConfig) Then
                Worksheets(WbNameSelExConfig).Delete
            End If
            ' Get back to config page
            Worksheets(WbNameConfig).Activate
            Worksheets(WbNameConfig).Range(CfgFirstPupi).offset(0, 1).Select
        End If
        
        Call LockSheets
        
        ' Druckblatt löschen
        If WSExists(WbNamePrintSheet) Then
            Worksheets(WbNamePrintSheet).Delete
        End If
        
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    
    End If
    
End Function

Private Function makeSure() As Boolean

    ' Mindestens ein sheet existiert bereits -> fragen
    Dim ws As Worksheet
    Dim i As Integer
    Dim found As Boolean
    For Each ws In ThisWorkbook.Worksheets
        For i = 0 To CfgMaxSheets
            If StrComp(ws.Name, Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, i * 2).Text) = 0 Then
                found = True
            End If
        Next i
    Next ws
    
    If found Or WSExists(WbNameGradeSheet) Then
        makeSure = False
        ' Abfragen ob wirklich neue Tabellen erstellt werden sollen...
        Dim Request As Integer
        Request = MsgBox("Sicher dass sie neue Tabellen erzeugen wollen?" & vbCrLf & "Mindestens eine Tabelle wurde gefunden, die überschrieben wird!", vbExclamation + vbOKCancel, "Sicher?")
        If Request = vbCancel Then
            Exit Function
        End If
        makeSure = True
    Else
        makeSure = True
    End If
    
End Function

Private Function PaintSegmentPages()

    Dim i As Integer

    '----------------------------------------
    ' Create Worksheets
    '----------------------------------------
    Dim actSheet As Integer
    Dim actSheetName As String
    Dim ws As Worksheet

    For actSheet = 0 To CfgMaxSheets

        ' Set name for further processing
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, actSheet * 2).Value
        If actSheetName = "" Then
            Exit Function
        End If

        '------------------------------------
        ' Delete old sheet if exists
        '------------------------------------
        If WSExists(actSheetName) Then
            Worksheets(actSheetName).Delete
        End If
        ' Create new sheet and cache reference
        Worksheets.Add(Before:=Worksheets(WbNameConfig)).Name = actSheetName
        Set ws = Worksheets(actSheetName)
        ws.Tab.color = gClrTabSections

        '------------------------------------
        ' Set background color
        '------------------------------------
        ws.Cells.Interior.color = RGB(240, 240, 240)
        ws.Cells.Locked = True

        '------------------------------------
        ' Set row heights
        '------------------------------------
        ws.Rows("1:100").RowHeight = 18

        '------------------------------------
        ' Set column widths
        '------------------------------------
        ws.Columns(1).ColumnWidth = 2.71          ' Spalte A bleibt leer
        ws.Columns(CfgColStart).ColumnWidth = 2.71   ' Spalte B: Schüler-Index
        ws.Columns(CfgColStart + 1).ColumnWidth = 25 ' Spalte C: Schüler-Name

        Dim numOfSubEx As Integer
        numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).offset(0, actSheet * 2).Value
        If numOfSubEx > 0 Then
            ' Batch: all sub-exercise columns + sum column in one range assignment
            ws.Range(ws.Columns(CfgColStart + CfgColOffsetFirstEx), _
                     ws.Columns(CfgColStart + CfgColOffsetFirstEx + numOfSubEx)).ColumnWidth = 8
        End If

        '------------------------------------
        ' Create frames and formatting
        '------------------------------------
        Dim span As Integer
        span = numOfSubEx + 2 ' Anzahl der Teilaufgaben + 3 Spalten (Index,Name,Summe)

        ' Abi Zelle
        With ws.Range(ws.Cells(CfgRowStart, CfgColStart), ws.Cells(CfgRowStart, CfgColStart + CInt(floor(CDbl(span) / 2))))
            Call setBorder(.Cells, False, True, False, True, False, xlMedium, gClrHeader, True, xlLeft, xlCenter)
            .Font.Bold = True
        End With
        ' Kurs Zelle
        With ws.Range(ws.Cells(CfgRowStart, CfgColStart + CInt(floor(CDbl(span) / 2)) + 1), ws.Cells(CfgRowStart, CfgColStart + span))
            Call setBorder(.Cells, False, False, True, True, False, xlMedium, gClrHeader, True, xlRight, xlCenter)
            .Font.Bold = True
        End With
        ' Bereich Zelle
        With ws.Range(ws.Cells(CfgRowStart + 1, CfgColStart), ws.Cells(CfgRowStart + 2, CfgColStart + span))
            Call setBorder(.Cells, False, True, True, False, True, xlMedium, gClrHeader, True, xlHAlignCenterAcrossSelection, xlCenter)
            .Font.Bold = True
            .Font.Size = 12
        End With

        ' Überschrift Namen
        Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1)), False, True, True, True, True, xlMedium, gClrTheme1, True)
        ' Überschrift Aufgaben / Punkte
        Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + span)), False, True, True, True, True, xlMedium, gClrTheme1, False, xlCenter, xlBottom)

        ' Schüler-Zeilen mit alternierenden Hintergrundfarben
        Dim iPupilSect As Integer
        Dim rowClrSect As Long
        Dim physRowSect As Long
        For iPupilSect = 0 To gNumOfPupils - 1
            physRowSect = CfgRowStart + CfgRowOffsetFirstPupil + iPupilSect
            If iPupilSect Mod 2 = 0 Then
                rowClrSect = gClrTheme2
            Else
                rowClrSect = gClrTheme2a
            End If
            ' Namen
            Call setBorder(ws.Range(ws.Cells(physRowSect, CfgColStart), ws.Cells(physRowSect, CfgColStart + CfgColOffsetFirstEx - 1)), False, True, True, True, True, xlThin, rowClrSect, False)
            ' Punkte-Bereich (bleibt weiß zur Eingabe)
            With ws.Range(ws.Cells(physRowSect, CfgColStart + CfgColOffsetFirstEx), ws.Cells(physRowSect, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1))
                Call setBorder(.Cells, False, True, True, True, True, xlThin, RGB(255, 255, 255), False, xlCenter, xlCenter)
                .Locked = False
            End With
            ' Summe-Bereich — bottom border intentionally omitted: the ZK row's hairline
            ' xlEdgeTop governs that shared border; outer block bottom is set by the
            ' block border pass below, avoiding a spurious thick line before ZK rows.
            Call setBorder(ws.Cells(physRowSect, CfgColStart + span), False, True, True, True, False, xlMedium, rowClrSect, False, xlCenter, xlCenter)
        Next iPupilSect
        For i = 0 To numOfSubEx - 1
            Call setUpperLimit( _
                ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + i), _
                         ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + i)), _
                CStr(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Address))
        Next i

        ' Prozentualer Punkteschnitt — placed after the full physical block
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + span))
            Call setBorder(.Cells, False, True, True, True, True, xlMedium, gClrTheme1, True, xlCenter, xlCenter)
            .NumberFormat = "0%"
        End With
        ' Border anpassen — full physical block
        Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)), False, True, True, True, True, xlMedium, 0, True)
        ' Re-apply outer border + inside verticals across the full block incl. sum column
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), _
                      ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + span))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeLeft).ColorIndex = 1
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeRight).ColorIndex = 1
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeTop).ColorIndex = 1
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeBottom).ColorIndex = 1
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlInsideVertical).ColorIndex = 1
        End With
        ' Re-force sum column left border to xlMedium (overridden by xlInsideVertical)
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + span), _
                      ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + span))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeLeft).ColorIndex = 1
        End With
        ' Define PupilBlock named range (replaces the AddZKDKRows call)
        Call DefinePupilBlockName(ws, numOfSubEx, gNumOfPupils)

    Next actSheet

End Function

Private Function FillSegmentPages()

    '----------------------------------------
    ' Create Worksheets
    '----------------------------------------
    Dim actSheet As Integer
    Dim actSheetName As String
    Dim numOfSubEx As Integer, span As Integer
    Dim i As Integer
    Dim ws As Worksheet
    Dim wsCfg As Worksheet
    Set wsCfg = Worksheets(WbNameConfig)

    ' Pre-compute all Config-derived column letters and row numbers once,
    ' qualified against wsCfg — no unqualified Range() calls inside the loop.
    Dim colSectEx As String     ' config column for exercise names (row captions)
    Dim colSectPts As String    ' config column for max points
    Dim colSumFirst As String   ' first sub-exercise column
    Dim colSumLast As String    ' last sub-exercise column
    Dim colPupiFirst As String  ' config first-name column
    Dim colPupiLast As String   ' config last-name column
    Dim colAbiTitle As String   ' config ABI title column
    Dim colAbiDate As String    ' config ABI date column
    Dim colAbiClass As String   ' config ABI class column
    Dim colExerCount As String  ' config exercise-count base column
    Dim cfgSectBaseRow As Long  ' first data row in config for exercise names
    Dim cfgPupiFirstRow As Long ' first pupil row in config
    Dim cfgAbiTitleRow As Long
    Dim cfgAbiDateRow As Long
    Dim cfgAbiClassRow As Long
    Dim cfgExerCntRow As Long
    Dim cfgFirstSectCol As Long

    cfgFirstSectCol = wsCfg.Range(CfgFirstSect).Column
    cfgSectBaseRow = 2 + wsCfg.Range(CfgFirstSect).row
    cfgPupiFirstRow = wsCfg.Range(CfgFirstPupi).row
    cfgAbiTitleRow = wsCfg.Range(CfgAbiTitle).row
    cfgAbiDateRow = wsCfg.Range(CfgAbiDate).row
    cfgAbiClassRow = wsCfg.Range(CfgAbiClass).row
    cfgExerCntRow = wsCfg.Range(CfgExerCount).row
    colPupiFirst = ColLetter(wsCfg.Range(CfgFirstPupi).Column + 1)
    colPupiLast = ColLetter(wsCfg.Range(CfgFirstPupi).Column + 2)
    colAbiTitle = ColLetter(wsCfg.Range(CfgAbiTitle).Column)
    colAbiDate = ColLetter(wsCfg.Range(CfgAbiDate).Column)
    colAbiClass = ColLetter(wsCfg.Range(CfgAbiClass).Column)
    colExerCount = ColLetter(wsCfg.Range(CfgExerCount).Column)
    colSumFirst = ColLetter(CfgColStart + CfgColOffsetFirstEx)

    Dim arrHdr1() As Variant
    Dim arrHdr2() As Variant

    For actSheet = 0 To CfgMaxSheets

        ' Set name for further processing
        actSheetName = wsCfg.Range(CfgFirstSect).offset(0, actSheet * 2).Value
        If actSheetName = "" Then
            Exit Function
        End If
        Set ws = Worksheets(actSheetName)
        numOfSubEx = wsCfg.Range(CfgExerCount).offset(0, actSheet * 2).Value
        span = numOfSubEx + 2 ' Anzahl der Teilaufgaben + 3 Spalten (Index,Name,Summe)

        colSectEx = ColLetter(cfgFirstSectCol + actSheet * 2)
        colSectPts = ColLetter(cfgFirstSectCol + actSheet * 2 + 1)
        colSumLast = ColLetter(CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)

        '------------------------------------
        ' Header-Text
        '------------------------------------
        ws.Cells(CfgRowStart, CfgColStart).Formula = "='" & WbNameConfig & "'!" & colAbiTitle & CStr(cfgAbiTitleRow) & "&"" ""&" & "YEAR('" & WbNameConfig & "'!" & colAbiDate & CStr(cfgAbiDateRow) & ")"
        ws.Cells(CfgRowStart, CfgColStart + span).Formula = "=""Kurs ""&'" & WbNameConfig & "'!" & colAbiClass & CStr(cfgAbiClassRow)
        ws.Cells(CfgRowStart + 1, CfgColStart).Value = actSheetName

        '------------------------------------
        ' Überschrift Name und Punkte
        '------------------------------------
        ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1).Value = "Name"

        ReDim arrHdr1(1 To 1, 1 To numOfSubEx)
        ReDim arrHdr2(1 To 1, 1 To numOfSubEx)
        For i = 0 To numOfSubEx - 1
            arrHdr1(1, i + 1) = "='" & WbNameConfig & "'!" & colSectEx & CStr(cfgSectBaseRow + i)
            arrHdr2(1, i + 1) = "='" & WbNameConfig & "'!" & colSectPts & CStr(cfgSectBaseRow + i)
        Next i
        ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), _
                 ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)).Formula = arrHdr1
        ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx), _
                 ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)).Formula = arrHdr2

        ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + numOfSubEx).Value = "Summe"
        ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx).Formula = "='" & WbNameConfig & "'!" & ColLetter(wsCfg.Range(CfgExerCount).offset(0, (actSheet * 2) + 1).Column) & CStr(cfgExerCntRow)

        '------------------------------------
        ' Namen
        '------------------------------------
        For i = 0 To gNumOfPupils - 1
            Dim physRow As Long
            physRow = CfgRowStart + CfgRowOffsetFirstPupil + i
            ws.Cells(physRow, CfgColStart).Value = wsCfg.Range(CfgFirstPupi).offset(i, 0).Value
            ws.Cells(physRow, CfgColStart + 1).Formula = "='" & WbNameConfig & "'!" & colPupiFirst & CStr(cfgPupiFirstRow + i) & "&"", ""&'" & WbNameConfig & "'!" & colPupiLast & CStr(cfgPupiFirstRow + i)
        Next i

        '------------------------------------
        ' Summen
        '------------------------------------
        For i = 0 To gNumOfPupils - 1
            physRow = CfgRowStart + CfgRowOffsetFirstPupil + i
            ws.Cells(physRow, CfgColStart + CfgColOffsetFirstEx + numOfSubEx).Formula = "=SUM(" & colSumFirst & CStr(physRow) & ":" & colSumLast & CStr(physRow) & ")"
        Next i

        FillSegmentPages_WriteSubPercentagFormulas ws, numOfSubEx

    Next actSheet

End Function

Private Function PaintGradePage()

    Dim i As Integer

    '------------------------------------
    ' Delete old sheet if exists
    '------------------------------------
    If WSExists(WbNameGradeSheet) Then
        Worksheets(WbNameGradeSheet).Delete
    End If
    ' Create new sheet and cache reference
    Worksheets.Add(Before:=Worksheets(Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, 0).Value)).Name = WbNameGradeSheet
    Dim ws As Worksheet
    Set ws = Worksheets(WbNameGradeSheet)
    ws.Tab.color = gClrTabGrades

    ' Count actual sheets
    Dim sheetCnt As Integer
    sheetCnt = CountSegmentSheets()

    '------------------------------------
    ' Set background color
    '------------------------------------
    ws.Cells.Interior.color = RGB(240, 240, 240)
    ws.Cells.Locked = True

    '------------------------------------
    ' Set row heights
    '------------------------------------
    ws.Rows("1:100").RowHeight = 18

    '------------------------------------
    ' Set column widths
    '------------------------------------
    ws.Columns(1).ColumnWidth = 2.71  ' Spalte A bleibt leer
    ws.Columns(2).ColumnWidth = 2.71  ' Spalte B: Schüler-Index
    ws.Columns(3).ColumnWidth = 25    ' Spalte C: Schüler-Name
    ' Batch: section columns (4 to 3+sheetCnt)
    ws.Range(ws.Columns(4), ws.Columns(3 + sheetCnt)).ColumnWidth = 11
    ' Batch: sum + points + grade + updown columns (4+sheetCnt to 7+sheetCnt)
    ws.Range(ws.Columns(4 + sheetCnt), ws.Columns(7 + sheetCnt)).ColumnWidth = 8

    ' Abi Zelle
    With ws.Range(ws.Cells(CfgRowStart, CfgColStart), ws.Cells(CfgRowStart, CfgColStart + floor(sheetCnt / 2#)))
        Call setBorder(.Cells, False, True, False, True, False, xlMedium, gClrHeader, True, xlLeft, xlCenter)
        .Font.Bold = True
    End With
    ' Kurs Zelle
    With ws.Range(ws.Cells(CfgRowStart, CfgColStart + floor(sheetCnt / 2#) + 1), ws.Cells(CfgRowStart, CfgColStart + sheetCnt + 4))
        Call setBorder(.Cells, False, False, True, True, False, xlMedium, gClrHeader, True, xlRight, xlCenter)
        .Font.Bold = True
    End With
    ' Bereich Zelle
    With ws.Range(ws.Cells(CfgRowStart + 1, CfgColStart), ws.Cells(CfgRowStart + 2, CfgColStart + sheetCnt + 4))
        Call setBorder(.Cells, False, True, True, False, True, xlMedium, gClrHeader, True, xlHAlignCenterAcrossSelection, xlCenter)
        .Font.Bold = True
        .Font.Size = 12
    End With

    ' Überschrift Namen
    Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1)), False, True, True, True, True, xlMedium, gClrTheme2, True)
    ' Überschrift Summe
    Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + sheetCnt + 2)), False, True, True, True, True, xlMedium, gClrTheme2, False, xlCenter, xlBottom)
    ' Überschrift Punkte / Note
    Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + sheetCnt + 3)), False, True, True, True, True, xlMedium, gClrTheme2, True, xlCenter, xlBottom)

    ' Schüler-Zeilen mit alternierenden Hintergrundfarben
    Dim iPupil As Integer
    Dim rowClr As Long
    For iPupil = 0 To gNumOfPupils - 1
        If iPupil Mod 2 = 0 Then
            rowClr = gClrTheme2
        Else
            rowClr = gClrTheme2a
        End If
        ' Namen
        Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + iPupil, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + iPupil, CfgColStart + CfgColOffsetFirstEx - 1)), False, True, True, True, True, xlThin, rowClr, False)
        ' Punkte-Bereich
        Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + iPupil, CfgColStart + CfgColOffsetFirstEx), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + iPupil, CfgColStart + CfgColOffsetFirstEx + sheetCnt - 1)), False, True, True, True, True, xlThin, rowClr, False, xlCenter, xlCenter)
        ' Summe-Bereich
        Call setBorder(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + iPupil, CfgColStart + sheetCnt + 2), False, True, True, True, True, xlMedium, rowClr, False, xlCenter, xlCenter)
        ' Punkte / Noten Bereich
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + iPupil, CfgColStart + sheetCnt + 3), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + iPupil, CfgColStart + sheetCnt + 4))
            Call setBorder(.Cells, False, True, True, True, True, xlThin, rowClr, False, xlCenter, xlCenter)
            Call setBorder(.Cells, False, True, True, True, True, xlMedium, rowClr, True, xlCenter, xlCenter)
        End With
    Next iPupil

    ' Prozentualer Punkteschnitt
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + sheetCnt + 4))
        Call setBorder(.Cells, False, True, True, True, True, xlMedium, gClrTheme2, True, xlCenter, xlCenter)
        .style = "Percent"
    End With

    ' Border anpassen
    Call setBorder(ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2)), False, True, True, True, True, xlMedium, 0, True)

End Function

Private Function FillGradePage()

    Dim i As Integer
    Dim u As Integer
    Dim r As Long
    Dim colSect As String

    ' Count actual sheets
    Dim sheetCnt As Integer
    sheetCnt = CountSegmentSheets()

    Dim ws As Worksheet
    Set ws = Worksheets(WbNameGradeSheet)

    ' Pre-compute column letters used repeatedly below
    Dim colFirst As String      ' first section/points column on grade sheet
    Dim colLast As String       ' last section column on grade sheet
    Dim colSum As String        ' per-pupil total sum column
    Dim colPoints As String     ' VLOOKUP points result column
    Dim colGrades As String     ' VLOOKUP grade result column
    Dim colUpDown As String     ' VLOOKUP up/down hidden column
    Dim colName As String       ' first name in config
    Dim colNameLast As String   ' last name in config
    Dim colNameIdx As String    ' pupil index column on grade sheet (CfgColStart+1)

    colFirst = ColLetter(CfgColStart + CfgColOffsetFirstEx)
    colLast = ColLetter(CfgColStart + CfgColOffsetFirstEx + sheetCnt - 1)
    colSum = ColLetter(CfgColStart + CfgColOffsetFirstEx + sheetCnt)
    colPoints = ColLetter(CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1)
    colGrades = ColLetter(CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2)
    colUpDown = ColLetter(CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3)
    colName = ColLetter(Range(CfgFirstPupi).Column + 1)
    colNameLast = ColLetter(Range(CfgFirstPupi).Column + 2)
    colNameIdx = ColLetter(CfgColStart + 1)

    Dim cfgPupiFirstRow As Long
    cfgPupiFirstRow = Range(CfgFirstPupi).row

    '------------------------------------
    ' Header-Text
    '------------------------------------
    ws.Cells(CfgRowStart, CfgColStart).Formula = "='" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiTitle).Column) & CStr(Range(CfgAbiTitle).row) & "&"" ""&" & "YEAR('" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiDate).Column) & CStr(Range(CfgAbiDate).row) & ")"
    ws.Cells(CfgRowStart, CfgColStart + sheetCnt + 4).Formula = "=""Kurs ""&'" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiClass).Column) & CStr(Range(CfgAbiClass).row)
    ws.Cells(CfgRowStart + 1, CfgColStart).Value = WbNameGradeSheet

    '------------------------------------
    ' Überschrift Name und Punkte
    '------------------------------------
    ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1).Value = "Name"

    Dim arrHdrVal() As Variant   ' section names (row 1 of header)
    Dim arrHdrFml() As Variant   ' max-points formulas (row 2 of header)
    ReDim arrHdrVal(1 To 1, 1 To sheetCnt)
    ReDim arrHdrFml(1 To 1, 1 To sheetCnt)
    For i = 0 To sheetCnt - 1
        arrHdrVal(1, i + 1) = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, i * 2).Value
        arrHdrFml(1, i + 1) = "='" & WbNameConfig & "'!" & ColLetter(Range(CfgExerCount).offset(0, (i * 2) + 1).Column) & CStr(Range(CfgExerCount).row)
    Next i
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + sheetCnt - 1)).Value = arrHdrVal
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt - 1)).Formula = arrHdrFml

    ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Value = "Summe"
    ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Formula = "=SUM(" & colFirst & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ":" & ColLetter(CfgColStart + CfgColOffsetFirstEx + CfgMaxSheets) & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ")"
    ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Value = "Punkte"
    ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2).Value = "Note"
    ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Value = "uiiii"

    '------------------------------------
    ' Namen
    '------------------------------------
    Dim arrIdx() As Variant
    Dim arrNames() As Variant
    ReDim arrIdx(1 To gNumOfPupils, 1 To 1)
    ReDim arrNames(1 To gNumOfPupils, 1 To 1)
    For i = 0 To gNumOfPupils - 1
        arrIdx(i + 1, 1) = Worksheets(WbNameConfig).Range(CfgFirstPupi).offset(i, 0).Value
        arrNames(i + 1, 1) = "='" & WbNameConfig & "'!" & colName & CStr(cfgPupiFirstRow + i) & "&"", ""&'" & WbNameConfig & "'!" & colNameLast & CStr(cfgPupiFirstRow + i)
    Next i
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart)).Value = arrIdx
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + 1), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + 1)).Formula = arrNames

    '------------------------------------
    ' Summen und Prozentuales Zeug
    '------------------------------------
    Dim arrPupilSum() As Variant
    Dim arrPupilPts() As Variant
    Dim arrPupilGrd() As Variant
    ReDim arrPupilSum(1 To gNumOfPupils, 1 To 1)
    ReDim arrPupilPts(1 To gNumOfPupils, 1 To 1)
    ReDim arrPupilGrd(1 To gNumOfPupils, 1 To 1)
    For i = 0 To gNumOfPupils - 1
        r = CfgRowStart + CfgRowOffsetFirstPupil + i
        arrPupilSum(i + 1, 1) = "=SUM(" & colFirst & CStr(r) & ":" & colLast & CStr(r) & ")"
        arrPupilPts(i + 1, 1) = "=VLOOKUP(VALUE(" & colSum & CStr(r) & ")," & CStr(WbNameGradeKey) & CfgVLookUpPoints & ")"
        arrPupilGrd(i + 1, 1) = "=VLOOKUP(VALUE(" & colSum & CStr(r) & ")," & CStr(WbNameGradeKey) & CfgVLookUpGrades & ")"
    Next i
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + sheetCnt), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt)).Formula = arrPupilSum
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1)).Formula = arrPupilPts
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2)).Formula = arrPupilGrd

    ' Per-section percentage average row
    Dim arrAvgPct() As Variant
    ReDim arrAvgPct(1 To 1, 1 To sheetCnt)
    For i = 0 To sheetCnt - 1
        colSect = ColLetter(CfgColStart + CfgColOffsetFirstEx + i)
        arrAvgPct(1, i + 1) = "=SUM(" & colSect & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":" & colSect & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1) & ")/(" & CStr(gNumOfPupils) & "*" & colSect & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ")"
    Next i
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + CfgColOffsetFirstEx), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + CfgColOffsetFirstEx + sheetCnt - 1)).Formula = arrAvgPct

    '------------------------------------
    ' Punkte aus Teilaufgaben
    '------------------------------------
    FillGradePage_WriteSectionColumns ws, sheetCnt, colNameIdx

    '------------------------------------
    ' Durchschnitt
    '------------------------------------
    Dim avgRow As Long
    avgRow = CfgRowStart + CfgRowOffsetFirstEx + gNumOfPupils + 2
    ' Punkte
    ws.Cells(avgRow, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Value = Chr(248)
    ws.Cells(avgRow, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Formula = "=AVERAGE(" & colPoints & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":" & colPoints & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1) & ")"
    ws.Cells(avgRow, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).NumberFormat = "0.00"
    ' Noten
    ws.Cells(avgRow, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2).FormulaArray = "=AVERAGE(SUBSTITUTE(SUBSTITUTE(" & colGrades & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":" & colGrades & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1) & ",""+"",""""),""-"","""")*1)"
    ws.Cells(avgRow, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2).NumberFormat = "0.00"

    '------------------------------------
    ' Up/Down Color
    '------------------------------------
    Dim arrUpDown() As Variant
    ReDim arrUpDown(1 To gNumOfPupils, 1 To 1)
    For i = 0 To gNumOfPupils - 1
        arrUpDown(i + 1, 1) = "=VLOOKUP(VALUE(" & colSum & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ")," & CStr(WbNameGradeKey) & CfgVLookUpUpDown & ")"
    Next i
    ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3), _
             ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3)).Formula = arrUpDown
    ws.Columns(CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Hidden = vbTrue

    ws.Range("A1").Select


End Function

Public Function UpdateUpDownColors()

    ' Initialize stuff
    Call Init

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = Worksheets(WbNameGradeSheet)
    ws.Unprotect Password:=WbPw

    Dim i As Integer

    ' Count actual sheets
    Dim sheetCnt As Integer
    sheetCnt = CountSegmentSheets()

    ' Durch die Spalte durchgehen und Farbe setzen
    If ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Value = "uiiii" Then
        Dim colPts1 As Long, colPts2 As Long, colUD As Long
        colPts1 = CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1
        colPts2 = CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2
        colUD = CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3
        For i = 0 To gNumOfPupils - 1
            Dim pupilRow As Long
            pupilRow = CfgRowStart + CfgRowOffsetFirstPupil + i
            Dim upDownText As String
            upDownText = ws.Cells(pupilRow, colUD).Text
            Dim clr As Long
            If StrComp(upDownText, "--") = 0 Then
                clr = gClrMinus2
            ElseIf StrComp(upDownText, "-") = 0 Then
                clr = gClrMinus1
            ElseIf StrComp(upDownText, "+") = 0 Then
                clr = gClrPlus1
            ElseIf StrComp(upDownText, "++") = 0 Then
                clr = gClrPlus2
            Else
                If i Mod 2 = 0 Then
                    clr = gClrTheme2
                Else
                    clr = gClrTheme2a
                End If
            End If
            ws.Range(ws.Cells(pupilRow, colPts1), ws.Cells(pupilRow, colPts2)).Interior.color = clr
        Next i
    Else
        MsgBox ("Corrupt!")
    End If
    ws.Range("A1").Select

    If DevMode <> 1 Then
        ws.Protect Password:=WbPw
        ws.EnableSelection = xlUnlockedCells
    End If

    Application.ScreenUpdating = True

End Function

' Protects all sheets except the print sheet.
' Skipped silently when something is on the clipboard (CutCopyMode) to avoid
' clearing the user's copy/cut selection.
Public Sub LockSheets()
    If DevMode <> 1 And Application.CutCopyMode = False Then
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> WbNamePrintSheet Then
                ws.Protect Password:=WbPw, DrawingObjects:=True, Contents:=True, Scenarios:=False
                ws.EnableSelection = xlUnlockedCells
            End If
        Next ws
    End If
End Sub

' Unprotects all sheets.
Public Sub UnlockSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=WbPw
    Next ws
End Sub

Public Function CheckForSelEx()

    Dim tblIdx As Integer
    Dim SelCfg As String
    Dim found As Boolean
    
    found = False
    For tblIdx = 0 To CfgMaxSheets
        SelCfg = Worksheets(WbNameConfig).Range(CfgSelEx).offset(0, tblIdx * 2).MergeArea.Cells(1, 1).Text
        If StrComp(SelCfg, "Ja") = 0 Then
            found = True
            Exit For
        End If
    Next tblIdx
    
    CheckForSelEx = found
    
End Function


Private Sub FillGradePage_WriteSectionColumns(ByVal ws As Worksheet, ByVal sheetCnt As Integer, ByVal colNameIdx As String)
    ' Writes per-pupil VLOOKUP formulas that pull each section's scored points
    ' from the respective segment sheet into the grade-page section columns.
    Dim u As Integer, i As Integer
    Dim sectName As String
    Dim numSubEx As Integer
    Dim arrVlookup() As Variant

    For u = 0 To sheetCnt - 1
        ' Anzahl der Teilaufgaben steht in Config =SVERWEIS(C5;'Infini B'!$C$7:$R$29;16;0)
        If Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, u * 2).Value <> "" Then
            sectName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, u * 2).Value
            numSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).offset(0, u * 2).Value

            ReDim arrVlookup(1 To gNumOfPupils, 1 To 1)
            For i = 0 To gNumOfPupils - 1
                arrVlookup(i + 1, 1) = "=VLOOKUP(" & colNameIdx & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ",'" & sectName & "'!PupilBlock," & CfgColOffsetFirstEx + numSubEx & ",0)"
            Next i
            ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + u), _
                     ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + u)).Formula = arrVlookup
        End If
    Next u
End Sub

Private Function CountSegmentSheets() As Integer
    Dim cnt As Integer, i As Integer
    cnt = 0
    For i = 0 To CfgMaxSheets
        If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, i * 2).Value) Then
            cnt = cnt + 1
        End If
    Next i
    CountSegmentSheets = cnt
End Function

Public Function ColLetter(colNum As Long) As String
    ColLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

Public Sub FillSegmentPages_WriteSubPercentagFormulas(ByVal ws As Worksheet, ByVal numOfSubEx As Integer)

    Dim firstRow As Long, lastRow As Long
    Dim colLtr As String
    Dim rngD As String, rngC As String
    Dim factorCell As String
    Dim funStr As String
    Dim i As Long

    Dim arrFormulas() As Variant
    ReDim arrFormulas(1 To 1, 1 To numOfSubEx)

    ' Pre-calc rows — span the full physical block (pupils + ZK/DK rows)
    firstRow = CfgRowStart + CfgRowOffsetFirstPupil
    lastRow = firstRow + gNumOfPupils - 1

    For i = 0 To numOfSubEx - 1

        colLtr = ColLetter(CfgColStart + CfgColOffsetFirstEx + i)

        rngD = colLtr & firstRow & ":" & colLtr & lastRow
        rngC = "C" & firstRow & ":C" & lastRow
        factorCell = colLtr & (CfgRowStart + CfgRowOffsetFirstEx + 1)

        funStr = "=IF(COUNTIFS(" & rngD & ",""<>""," & rngC & ",""<>ZK""," & rngC & ",""<>DK"")<>0," & _
                 "SUMIFS(" & rngD & "," & rngC & ",""<>ZK""," & rngC & ",""<>DK"")/" & _
                 "(COUNTIFS(" & rngD & ",""<>""," & rngC & ",""<>ZK""," & rngC & ",""<>DK"")*" & factorCell & ")," & _
                 "0)"

        arrFormulas(1, i + 1) = funStr

    Next i

    ' Batch write — row after the full physical block
    Dim pctRow As Long
    pctRow = firstRow + gNumOfPupils
    ws.Range( _
        ws.Cells(pctRow, CfgColStart + CfgColOffsetFirstEx), _
        ws.Cells(pctRow, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1) _
    ).Formula = arrFormulas

End Sub


