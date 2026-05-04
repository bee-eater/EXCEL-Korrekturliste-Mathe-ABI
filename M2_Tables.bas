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
            Worksheets(WbNameConfig).Range(CfgFirstPupi).Offset(0, 1).Select
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
            If StrComp(ws.name, Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Text) = 0 Then
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
        'Request = MsgBox("Ganz sicher?? Es ist wirklich alles weg!", vbCritical + vbOKCancel, "Ahhhhhhhh...")
        'If Request = vbCancel Then
        '    Exit Function
        'End If
        ' Not exited -> sure
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
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, actSheet * 2).Value
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
        Worksheets.Add(Before:=Worksheets(WbNameConfig)).name = actSheetName
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
        numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, actSheet * 2).Value
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
            .Select
            Call setBorder(False, True, False, True, False, xlMedium, gClrHeader, True, xlLeft, xlCenter)
            .Font.Bold = True
        End With
        ' Kurs Zelle
        With ws.Range(ws.Cells(CfgRowStart, CfgColStart + CInt(floor(CDbl(span) / 2)) + 1), ws.Cells(CfgRowStart, CfgColStart + span))
            .Select
            Call setBorder(False, False, True, True, False, xlMedium, gClrHeader, True, xlRight, xlCenter)
            .Font.Bold = True
        End With
        ' Bereich Zelle
        With ws.Range(ws.Cells(CfgRowStart + 1, CfgColStart), ws.Cells(CfgRowStart + 2, CfgColStart + span))
            .Select
            Call setBorder(False, True, True, False, True, xlMedium, gClrHeader, True, xlHAlignCenterAcrossSelection, xlCenter)
            .Font.Bold = True
            .Font.Size = 12
        End With

        ' Überschrift Namen
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1))
            .Select
            Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, True)
        End With
        ' Überschrift Aufgaben / Punkte
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + span))
            .Select
            Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, False, xlCenter, xlBottom)
        End With

        ' Namen
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx - 1))
            .Select
            Call setBorder(False, True, True, True, True, xlThin, gClrTheme1, False)
        End With
        ' Punkte-Bereich
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1))
            .Select
            Call setBorder(False, True, True, True, True, xlThin, RGB(255, 255, 255), False, xlCenter, xlCenter)
            .Locked = False
        End With
        For i = 0 To numOfSubEx - 1
            ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + i), _
                     ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + i)).Select
            setUpperLimit (CStr(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Address))
        Next i
        ' Summe-Bereich
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + span), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + span))
            .Select
            Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, False, xlCenter, xlCenter)
        End With

        ' Prozentualer Punkteschnitt
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + span))
            .Select
            Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, True, xlCenter, xlCenter)
            .NumberFormat = "0%"
        End With
        ' Border anpassen
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1))
            .Select
            Call setBorder(False, True, True, True, True, xlMedium, 0, True)
        End With

        ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx).Select

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
    ' Pre-compute column letters used repeatedly below (unqualified Cells/Range
    ' intentionally refers to the active sheet, preserving original behaviour)
    Dim colSectEx As String     ' config column for exercise names (row captions)
    Dim colSectPts As String    ' config column for max points
    Dim colSumFirst As String   ' first sub-exercise column
    Dim colSumLast As String    ' last sub-exercise column
    Dim colPupiFirst As String  ' config first-name column
    Dim colPupiLast As String   ' config last-name column
    Dim cfgSectBaseRow As Long  ' first data row in config for exercise names
    Dim cfgPupiFirstRow As Long ' first pupil row in config
    Dim arrHdr1() As Variant
    Dim arrHdr2() As Variant
    Dim arrIdx() As Variant
    Dim arrNames() As Variant
    Dim arrSums() As Variant

    For actSheet = 0 To CfgMaxSheets

        ' Set name for further processing
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, actSheet * 2).Value
        If actSheetName = "" Then
            Exit Function
        End If
        Set ws = Worksheets(actSheetName)
        numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, actSheet * 2).Value
        span = numOfSubEx + 2 ' Anzahl der Teilaufgaben + 3 Spalten (Index,Name,Summe)

        colSectEx = ColLetter(Range(CfgFirstSect).Column + actSheet * 2)
        colSectPts = ColLetter(Range(CfgFirstSect).Column + (actSheet * 2) + 1)
        colSumFirst = ColLetter(CfgColStart + CfgColOffsetFirstEx)
        colSumLast = ColLetter(CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)
        colPupiFirst = ColLetter(Range(CfgFirstPupi).Column + 1)
        colPupiLast = ColLetter(Range(CfgFirstPupi).Column + 2)
        cfgSectBaseRow = 2 + Range(CfgFirstSect).row
        cfgPupiFirstRow = Range(CfgFirstPupi).row

        '------------------------------------
        ' Header-Text
        '------------------------------------
        ws.Cells(CfgRowStart, CfgColStart).Formula = "='" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiTitle).Column) & CStr(Range(CfgAbiTitle).row) & "&"" ""&" & "YEAR('" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiDate).Column) & CStr(Range(CfgAbiDate).row) & ")"
        ws.Cells(CfgRowStart, CfgColStart + span).Formula = "=""Kurs ""&'" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiClass).Column) & CStr(Range(CfgAbiClass).row)
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
        ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx).Formula = "='" & WbNameConfig & "'!" & ColLetter(Range(CfgExerCount).Offset(0, (actSheet * 2) + 1).Column) & CStr(Range(CfgExerCount).row)

        '------------------------------------
        ' Namen
        '------------------------------------
        ReDim arrIdx(1 To gNumOfPupils, 1 To 1)
        ReDim arrNames(1 To gNumOfPupils, 1 To 1)
        For i = 0 To gNumOfPupils - 1
            arrIdx(i + 1, 1) = Worksheets(WbNameConfig).Range(CfgFirstPupi).Offset(i, 0).Value
            arrNames(i + 1, 1) = "='" & WbNameConfig & "'!" & colPupiFirst & CStr(cfgPupiFirstRow + i) & "&"", ""&'" & WbNameConfig & "'!" & colPupiLast & CStr(cfgPupiFirstRow + i)
        Next i
        ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), _
                 ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart)).Value = arrIdx
        ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + 1), _
                 ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + 1)).Formula = arrNames

        '------------------------------------
        ' Summen
        '------------------------------------
        ReDim arrSums(1 To gNumOfPupils, 1 To 1)
        For i = 0 To gNumOfPupils - 1
            arrSums(i + 1, 1) = "=SUM(" & colSumFirst & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ":" & colSumLast & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ")"
        Next i
        ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + numOfSubEx), _
                 ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx)).Formula = arrSums

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
    Worksheets.Add(Before:=Worksheets(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, 0).Value)).name = WbNameGradeSheet
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
        .Select
        Call setBorder(False, True, False, True, False, xlMedium, gClrHeader, True, xlLeft, xlCenter)
        .Font.Bold = True
    End With
    ' Kurs Zelle
    With ws.Range(ws.Cells(CfgRowStart, CfgColStart + floor(sheetCnt / 2#) + 1), ws.Cells(CfgRowStart, CfgColStart + sheetCnt + 4))
        .Select
        Call setBorder(False, False, True, True, False, xlMedium, gClrHeader, True, xlRight, xlCenter)
        .Font.Bold = True
    End With
    ' Bereich Zelle
    With ws.Range(ws.Cells(CfgRowStart + 1, CfgColStart), ws.Cells(CfgRowStart + 2, CfgColStart + sheetCnt + 4))
        .Select
        Call setBorder(False, True, True, False, True, xlMedium, gClrHeader, True, xlHAlignCenterAcrossSelection, xlCenter)
        .Font.Bold = True
        .Font.Size = 12
    End With

    ' Überschrift Namen
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1))
        .Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, True)
    End With
    ' Überschrift Summe
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + sheetCnt + 2))
        .Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, False, xlCenter, xlBottom)
    End With
    ' Überschrift Punkte / Note
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2), ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + sheetCnt + 3))
        .Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, True, xlCenter, xlBottom)
    End With

    ' Namen
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx - 1))
        .Select
        Call setBorder(False, True, True, True, True, xlThin, gClrTheme2, False)
    End With
    ' Punkte-Bereich (the original per-column loop only selected each column with no side-effect – removed)
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt - 1))
        .Select
        Call setBorder(False, True, True, True, True, xlThin, gClrTheme2, False, xlCenter, xlCenter)
    End With
    ' Summe-Bereich
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + sheetCnt + 2), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + sheetCnt + 2))
        .Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, False, xlCenter, xlCenter)
    End With
    ' Punkte / Noten Bereich
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + sheetCnt + 3), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + sheetCnt + 4))
        .Select
        Call setBorder(False, True, True, True, True, xlThin, gClrTheme2, False, xlCenter, xlCenter)
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, True, xlCenter, xlCenter)
    End With

    ' Prozentualer Punkteschnitt
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + sheetCnt + 4))
        .Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, True, xlCenter, xlCenter)
        .style = "Percent"
    End With

    ' Border anpassen
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2))
        .Select
        Call setBorder(False, True, True, True, True, xlMedium, 0, True)
    End With

    ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx).Select

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
        arrHdrVal(1, i + 1) = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value
        arrHdrFml(1, i + 1) = "='" & WbNameConfig & "'!" & ColLetter(Range(CfgExerCount).Offset(0, (i * 2) + 1).Column) & CStr(Range(CfgExerCount).row)
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
        arrIdx(i + 1, 1) = Worksheets(WbNameConfig).Range(CfgFirstPupi).Offset(i, 0).Value
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
                clr = gClrTheme1
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

Public Function LockSheets()

    If DevMode <> 1 Then
        Dim i As Integer
        Dim sheetName As String
        For i = 0 To CfgMaxSheets
            sheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value
            If WSExists(sheetName) Then
                Worksheets(sheetName).Protect Password:=WbPw
                Worksheets(sheetName).EnableSelection = xlUnlockedCells
            End If
        Next i

        If WSExists(WbNameGradeSheet) Then
            Worksheets(WbNameGradeSheet).Protect Password:=WbPw
            Worksheets(WbNameGradeSheet).EnableSelection = xlUnlockedCells
        End If

        Worksheets(WbNameConfig).Protect Password:=WbPw
        Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
        Worksheets(WbNameGradeKey).Protect Password:=WbPw
        Worksheets(WbNameGradeKey).EnableSelection = xlUnlockedCells
    End If

End Function

Public Function UnlockSheets()

    If DevMode <> 1 Then
        Dim i As Integer
        Dim sheetName As String
        For i = 0 To CfgMaxSheets
            sheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value
            If WSExists(sheetName) Then
                Worksheets(sheetName).Unprotect Password:=WbPw
                Worksheets(sheetName).EnableSelection = xlUnlockedCells
            End If
        Next i

        If WSExists(WbNameGradeSheet) Then
            Worksheets(WbNameGradeSheet).Unprotect Password:=WbPw
            Worksheets(WbNameGradeSheet).EnableSelection = xlUnlockedCells
        End If

        Worksheets(WbNameConfig).Unprotect Password:=WbPw
        Worksheets(WbNameConfig).EnableSelection = xlUnlockedCells
        Worksheets(WbNameGradeKey).Unprotect Password:=WbPw
        Worksheets(WbNameGradeKey).EnableSelection = xlUnlockedCells
    End If

End Function

Public Function CheckForSelEx()

    Dim tblIdx As Integer
    Dim SelCfg As String
    Dim found As Boolean
    
    found = False
    For tblIdx = 0 To CfgMaxSheets
        SelCfg = Worksheets(WbNameConfig).Range(CfgSelEx).Offset(0, tblIdx * 2).MergeArea.Cells(1, 1).Text
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
    Dim colSectLast As String
    Dim arrVlookup() As Variant

    For u = 0 To sheetCnt - 1
        ' Anzahl der Teilaufgaben steht in Config =SVERWEIS(C5;'Infini B'!$C$7:$R$29;16;0)
        If Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, u * 2).Value <> "" Then
            sectName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, u * 2).Value
            numSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value
            colSectLast = ColLetter(CfgColStart + CfgColOffsetFirstEx + numSubEx)

            ReDim arrVlookup(1 To gNumOfPupils, 1 To 1)
            For i = 0 To gNumOfPupils - 1
                arrVlookup(i + 1, 1) = "=VLOOKUP(" & colNameIdx & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ",'" & sectName & "'!" & colNameIdx & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":$" & colSectLast & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1) & "," & CfgColOffsetFirstEx + numSubEx & ",0)"
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
        If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value) Then
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

    ' Pre-calc rows
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

    ' Batch write
    ws.Range( _
        ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + CfgColOffsetFirstEx), _
        ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1) _
    ).Formula = arrFormulas

End Sub


