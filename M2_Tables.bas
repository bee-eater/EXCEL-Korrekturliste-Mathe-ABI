Attribute VB_Name = "M2_Tables"
Option Explicit

Public Function CreateTables()

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
        
        ' Get back to config page
        Worksheets(WbNameConfig).Activate
        Worksheets(WbNameConfig).Range("A1").Select
        
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
            If StrComp(ws.Name, Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Text) = 0 Then
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
        ' Create new sheet
        Worksheets.Add(Before:=Worksheets(WbNameConfig)).Name = actSheetName
        Worksheets(actSheetName).Tab.color = gClrTabSections
        
        '------------------------------------
        ' Set background color
        '------------------------------------
        Worksheets(actSheetName).Cells.Interior.color = RGB(240, 240, 240)
        Worksheets(actSheetName).Cells.Locked = True
        
        '------------------------------------
        ' Set row heights
        '------------------------------------
        Worksheets(actSheetName).Rows("1:100").RowHeight = 18
        
        '------------------------------------
        ' Set column widths
        '------------------------------------
        ' Set column widths for sub exercises and for sum column
        Worksheets(actSheetName).Columns(1).ColumnWidth = 2.71  ' Spalte A bleibt leer
        Worksheets(actSheetName).Columns(CfgColStart).ColumnWidth = 2.71  ' Spalte B: Schüler-Index
        Worksheets(actSheetName).Columns(CfgColStart + 1).ColumnWidth = 22  ' Spalte C: Schüler-Name
        
        Dim numOfSubEx As Integer
        Dim subExIdx, subEx As Integer
        numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, actSheet * 2).Value
        If numOfSubEx > 0 Then
            ' Set column widths for sub exercises
            For subEx = 0 To numOfSubEx - 1
                Worksheets(actSheetName).Columns(CfgColStart + CfgColOffsetFirstEx + subEx).ColumnWidth = 8
            Next subEx
            ' Set column width for sum column
            Worksheets(actSheetName).Columns(CfgColStart + CfgColOffsetFirstEx + numOfSubEx).ColumnWidth = 8
        End If
        
        '------------------------------------
        ' Create frames and formatting
        '------------------------------------
        ' Range(cells(i,1),cells(k,4)).select
                
        Dim span As Integer
        span = numOfSubEx + 2 ' Anzahl der Teilaufgaben + 3 Spalten (Index,Name,Summe)
                
        ' Abi Zelle
        Worksheets(actSheetName).Range(Cells(CfgRowStart, CfgColStart), Cells(CfgRowStart, CfgColStart + CInt(floor(CDbl(span) / 2)))).Select
        Call setBorder(False, True, False, True, False, xlMedium, gClrHeader, True, xlLeft, xlCenter)
        Selection.Font.Bold = True
        ' Kurs Zelle
        Worksheets(actSheetName).Range(Cells(CfgRowStart, CfgColStart + CInt(floor(CDbl(span) / 2)) + 1), Cells(CfgRowStart, CfgColStart + span)).Select
        Call setBorder(False, False, True, True, False, xlMedium, gClrHeader, True, xlRight, xlCenter)
        Selection.Font.Bold = True
        ' Bereich Zelle
        Worksheets(actSheetName).Range(Cells(CfgRowStart + 1, CfgColStart), Cells(CfgRowStart + 2, CfgColStart + span)).Select
        Call setBorder(False, True, True, False, True, xlMedium, gClrHeader, True, xlHAlignCenterAcrossSelection, xlCenter)
        Selection.Font.Bold = True
        Selection.Font.Size = 12

        ' Überschrift Namen
        Worksheets(actSheetName).Range(Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1)).Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, True)
        ' Überschrift Aufgaben / Punkte
        Worksheets(actSheetName).Range(Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + span)).Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, False, xlCenter, xlBottom)
        
        ' Namen
        Worksheets(actSheetName).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx - 1)).Select
        Call setBorder(False, True, True, True, True, xlThin, gClrTheme1, False)
        ' Punkte-Bereich
        Worksheets(actSheetName).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)).Select
        Call setBorder(False, True, True, True, True, xlThin, RGB(255, 255, 255), False, xlCenter, xlCenter)
        Selection.Locked = False
        For i = 0 To numOfSubEx - 1
            Worksheets(actSheetName).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + i), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + i)).Select
            setUpperLimit (CStr(Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Address))
        Next i
        ' Summe-Bereich
        Worksheets(actSheetName).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + span), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + span)).Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, False, xlCenter, xlCenter)
        
        ' Prozentualer Punkteschnitt
        Worksheets(actSheetName).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + span)).Select
        Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, True, xlCenter, xlCenter)
        Selection.NumberFormat = "0%"
        ' Border anpassen
        Worksheets(actSheetName).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)).Select
        Call setBorder(False, True, True, True, True, xlMedium, 0, True)
        
        Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx).Select
    
    Next actSheet
    
End Function

Private Function FillSegmentPages()

    '----------------------------------------
    ' Create Worksheets
    '----------------------------------------
    Dim actSheet As Integer
    Dim actSheetName As String
    Dim numOfSubEx, span As Integer
    Dim i As Integer
    
    For actSheet = 0 To CfgMaxSheets
    
        ' Set name for further processing
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, actSheet * 2).Value
        If actSheetName = "" Then
            Exit Function
        End If
        numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, actSheet * 2).Value
        span = numOfSubEx + 2 ' Anzahl der Teilaufgaben + 3 Spalten (Index,Name,Summe)
        
        '------------------------------------
        ' Header-Text
        '------------------------------------
        Worksheets(actSheetName).Cells(CfgRowStart, CfgColStart).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiTitle).Column).Address, "$")(1) & CStr(Range(CfgAbiTitle).row) & "&"" ""&" & "YEAR('" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiDate).Column).Address, "$")(1) & CStr(Range(CfgAbiDate).row) & ")"
        Worksheets(actSheetName).Cells(CfgRowStart, CfgColStart + span).Formula = "=""Kurs ""&'" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiClass).Column).Address, "$")(1) & CStr(Range(CfgAbiClass).row)
        Worksheets(actSheetName).Cells(CfgRowStart + 1, CfgColStart).Value = actSheetName
        
        '------------------------------------
        ' Überschrift Name und Punkte
        '------------------------------------
        Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1).Value = "Name"
        For i = 0 To numOfSubEx - 1
            Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + i).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstSect).Column + actSheet * 2).Address, "$")(1) & CStr(2 + Range(CfgFirstSect).row + i)
            Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstSect).Column + (actSheet * 2) + 1).Address, "$")(1) & CStr(2 + Range(CfgFirstSect).row + i)
        Next i
        Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + numOfSubEx).Value = "Summe"
        Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx).Formula = "=SUM(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ")"
        '------------------------------------
        ' Namen
        '------------------------------------
        For i = 0 To gNumOfPupils - 1
            Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart).Value = Worksheets(WbNameConfig).Range(CfgFirstPupi).Offset(i, 0).Value
            Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + 1).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstPupi).Column + 1).Address, "$")(1) & CStr(Range(CfgFirstPupi).row + i) & "&"", ""&'" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstPupi).Column + 2).Address, "$")(1) & CStr(Range(CfgFirstPupi).row + i)
        Next i
        
        '------------------------------------
        ' Summen und Prozentuales Zeug
        '------------------------------------
        For i = 0 To gNumOfPupils - 1
            Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + numOfSubEx).Formula = "=SUM(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ")"
        Next i
        For i = 0 To numOfSubEx - 1
            Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + CfgColOffsetFirstEx + i).Formula = "=SUM(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + i).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + i).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1) & ")/(" & CStr(gNumOfPupils) & "*" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + i).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ")"
        Next i
        
    Next actSheet
    
End Function

Private Function PaintGradePage()

    Dim i, u As Integer

    '------------------------------------
    ' Delete old sheet if exists
    '------------------------------------
    If WSExists(WbNameGradeSheet) Then
        Worksheets(WbNameGradeSheet).Delete
    End If
    ' Create new sheet
    Worksheets.Add(Before:=Worksheets(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, 0).Value)).Name = WbNameGradeSheet
    Worksheets(WbNameGradeSheet).Tab.color = gClrTabGrades
    
    ' Count actual sheets
    Dim sheetCnt As Integer
    sheetCnt = 0
    For i = 0 To CfgMaxSheets
        If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value) Then
            sheetCnt = sheetCnt + 1
        End If
    Next i
    
    '------------------------------------
    ' Set background color
    '------------------------------------
    Worksheets(WbNameGradeSheet).Cells.Interior.color = RGB(240, 240, 240)
    Worksheets(WbNameGradeSheet).Cells.Locked = True
    
    '------------------------------------
    ' Set row heights
    '------------------------------------
    Worksheets(WbNameGradeSheet).Rows("1:100").RowHeight = 18
    
    '------------------------------------
    ' Set column widths
    '------------------------------------
    ' Set column widths for sub exercises and for sum column
    Worksheets(WbNameGradeSheet).Columns(1).ColumnWidth = 2.71  ' Spalte A bleibt leer
    Worksheets(WbNameGradeSheet).Columns(2).ColumnWidth = 2.71  ' Spalte B: Schüler-Index
    Worksheets(WbNameGradeSheet).Columns(3).ColumnWidth = 22  ' Spalte C: Schüler-Name
    For i = 0 To sheetCnt - 1
        Worksheets(WbNameGradeSheet).Columns(4 + i).ColumnWidth = 11
    Next i
    For i = sheetCnt To sheetCnt + 3
        Worksheets(WbNameGradeSheet).Columns(4 + i).ColumnWidth = 8
    Next i
    
    ' Abi Zelle
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart, CfgColStart), Cells(CfgRowStart, CfgColStart + floor(sheetCnt / 2#))).Select
    Call setBorder(False, True, False, True, False, xlMedium, gClrHeader, True, xlLeft, xlCenter)
    Selection.Font.Bold = True
    ' Kurs Zelle
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart, CfgColStart + floor(sheetCnt / 2#) + 1), Cells(CfgRowStart, CfgColStart + sheetCnt + 4)).Select
    Call setBorder(False, False, True, True, False, xlMedium, gClrHeader, True, xlRight, xlCenter)
    Selection.Font.Bold = True
    ' Bereich Zelle
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + 1, CfgColStart), Cells(CfgRowStart + 2, CfgColStart + sheetCnt + 4)).Select
    Call setBorder(False, True, True, False, True, xlMedium, gClrHeader, True, xlHAlignCenterAcrossSelection, xlCenter)
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    
    ' Überschrift Namen
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1)).Select
    Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, True)
    ' Überschrift Summe
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + sheetCnt + 2)).Select
    Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, False, xlCenter, xlBottom)
    ' Überschrift Punkte / Note
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2), Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + sheetCnt + 3)).Select
    Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, True, xlCenter, xlBottom)
    
    ' Namen
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx - 1)).Select
    Call setBorder(False, True, True, True, True, xlThin, gClrTheme2, False)
    ' Punkte-Bereich
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt - 1)).Select
    Call setBorder(False, True, True, True, True, xlThin, gClrTheme2, False, xlCenter, xlCenter)
    For i = 0 To sheetCnt - 1
        Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + i), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + i)).Select
    Next i
    ' Summe-Bereich
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + sheetCnt + 2), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + sheetCnt + 2)).Select
    Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, False, xlCenter, xlCenter)
    ' Punkte / Noten Bereich
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + sheetCnt + 3), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + sheetCnt + 4)).Select
    Call setBorder(False, True, True, True, True, xlThin, gClrTheme2, False, xlCenter, xlCenter)
    Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, True, xlCenter, xlCenter)
        
    ' Prozentualer Punkteschnitt
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + sheetCnt + 4)).Select
    Call setBorder(False, True, True, True, True, xlMedium, gClrTheme2, True, xlCenter, xlCenter)
    Selection.style = "Percent"
    ' Border anpassen
    Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2)).Select
    Call setBorder(False, True, True, True, True, xlMedium, 0, True)
        
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx).Select

End Function

Private Function FillGradePage()

    Dim i As Integer

    ' Count actual sheets
    Dim sheetCnt As Integer
    sheetCnt = 0
    For i = 0 To CfgMaxSheets
        If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value) Then
            sheetCnt = sheetCnt + 1
        End If
    Next i
    
    '------------------------------------
    ' Header-Text
    '------------------------------------
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart, CfgColStart).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiTitle).Column).Address, "$")(1) & CStr(Range(CfgAbiTitle).row) & "&"" ""&" & "YEAR('" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiDate).Column).Address, "$")(1) & CStr(Range(CfgAbiDate).row) & ")"
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart, CfgColStart + sheetCnt + 4).Formula = "=""Kurs ""&'" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiClass).Column).Address, "$")(1) & CStr(Range(CfgAbiClass).row)
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + 1, CfgColStart).Value = WbNameGradeSheet
    
    '------------------------------------
    ' Überschrift Name und Punkte
    '------------------------------------
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1).Value = "Name"
    For i = 0 To sheetCnt - 1
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + i).Value = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgExerCount).Offset(0, (i * 2) + 1).Column).Address, "$")(1) & CStr(Range(CfgExerCount).row)
    Next i
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Value = "Summe"
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Formula = "=SUM(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + 5).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ")"
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Value = "Punkte"
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2).Value = "Note"
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Value = "uiiii"
    
    '------------------------------------
    ' Namen
    '------------------------------------
    For i = 0 To gNumOfPupils - 1
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart).Value = Worksheets(WbNameConfig).Range(CfgFirstPupi).Offset(i, 0).Value
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + 1).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstPupi).Column + 1).Address, "$")(1) & CStr(Range(CfgFirstPupi).row + i) & "&"", ""&'" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstPupi).Column + 2).Address, "$")(1) & CStr(Range(CfgFirstPupi).row + i)
    Next i
    
    '------------------------------------
    ' Summen und Prozentuales Zeug
    '------------------------------------
    For i = 0 To gNumOfPupils - 1
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 0).Formula = "=SUM(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + sheetCnt - 1).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ")"
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Formula = "=VLOOKUP(VALUE(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ")," & CStr(WbNameGradeKey) & CfgVLookUpPoints & ")"
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2).Formula = "=VLOOKUP(VALUE(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ")," & CStr(WbNameGradeKey) & CfgVLookUpGrades & ")"
    Next i
    For i = 0 To sheetCnt - 1
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils, CfgColStart + CfgColOffsetFirstEx + i).Formula = "=SUM(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + i).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + i).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1) & ")/(" & CStr(gNumOfPupils) & "*" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + i).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstEx + 1) & ")"
    Next i
    
    '------------------------------------
    ' Punkte aus Teilaufgaben
    '------------------------------------
    Dim u As Integer
    For u = 0 To sheetCnt - 1
        ' Anzahl der Teilaufgaben steht in Config =SVERWEIS(C5;'Infini B'!$C$7:$R$29;16;0)
        If Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, u * 2).Value <> "" Then
            For i = 0 To gNumOfPupils - 1
                Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + u).Formula = "=VLOOKUP(" & Split(Cells(1, CfgColStart + 1).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ",'" & Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, u * 2).Value & "'!" & Split(Cells(1, CfgColStart + 1).Address, "$")(1) & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":$" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1) & "," & CfgColOffsetFirstEx + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, u * 2).Value & ",0)"
            Next i
        End If
    Next u
    
    '------------------------------------
    ' Durchschnitt
    '------------------------------------
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + gNumOfPupils + 2, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Value = Chr(248)
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + gNumOfPupils + 2, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Formula = "=AVERAGE(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1) & ")"
    Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + gNumOfPupils + 2, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).NumberFormat = "0.00"
    
    '------------------------------------
    ' Up/Down Color
    '------------------------------------
    For i = 0 To gNumOfPupils - 1
        Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Formula = "=VLOOKUP(VALUE(" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + sheetCnt).Address, "$")(1) & CStr(CfgRowStart + CfgRowOffsetFirstPupil + i) & ")," & CStr(WbNameGradeKey) & CfgVLookUpUpDown & ")"
    Next i
    Worksheets(WbNameGradeSheet).Columns(CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Hidden = vbTrue
    
    Worksheets(WbNameGradeSheet).Range("A1").Select
    
    
End Function

Public Function UpdateUpDownColors()

    ' Initialize stuff
    Call Init
    
    Application.ScreenUpdating = False
    
    Worksheets(WbNameGradeSheet).Unprotect Password:=WbPw
    
    Dim i As Integer
        
    ' Count actual sheets
    Dim sheetCnt As Integer
    sheetCnt = 0
    For i = 0 To CfgMaxSheets
        If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value) Then
            sheetCnt = sheetCnt + 1
        End If
    Next i
    
    ' Durch die Spalte durchgehen und Farbe setzen
    If Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Value = "uiiii" Then
        For i = 0 To gNumOfPupils - 1
            Worksheets(WbNameGradeSheet).Range(Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1), Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 2)).Select
            If StrComp(Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Text, "--") = 0 Then
                With Selection
                    .Interior.color = gClrMinus2
                End With
            ElseIf StrComp(Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Text, "-") = 0 Then
                With Selection
                    .Interior.color = gClrMinus1
                End With
            ElseIf StrComp(Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Text, "+") = 0 Then
                With Selection
                    .Interior.color = gClrPlus1
                End With
            ElseIf StrComp(Worksheets(WbNameGradeSheet).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 3).Text, "++") = 0 Then
                With Selection
                    .Interior.color = gClrPlus2
                End With
            Else
                With Selection
                    .Interior.color = gClrTheme1
                End With
            End If
        Next i
    Else
        MsgBox ("Corrupt!")
    End If
    Worksheets(WbNameGradeSheet).Range("A1").Select
    
    If DevMode <> 1 Then
        Worksheets(WbNameGradeSheet).Protect Password:=WbPw
        Worksheets(WbNameGradeSheet).EnableSelection = xlUnlockedCells
    End If
    
    Application.ScreenUpdating = True
        
End Function

Private Function LockSheets()
    
    If DevMode <> 1 Then
        Dim i As Integer
        For i = 0 To CfgMaxSheets
            If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value) Then
                Worksheets(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value).Protect Password:=WbPw
                Worksheets(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value).EnableSelection = xlUnlockedCells
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
