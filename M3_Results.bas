Attribute VB_Name = "M3_Results"

Option Explicit

' -- Module-level context for PaintPrintPage helpers (initialised once per run) --
Private m_wsPrint       As Worksheet
Private m_wsCfg         As Worksheet
Private m_clAbiTitle    As String, m_clAbiDate    As String
Private m_clAbiTeacher  As String, m_clAbiClass   As String
Private m_clPupiFirst   As String, m_clPupiLast   As String
Private m_clCfgStart1   As String, m_clPrintName  As String
Private m_rowAbiTitle   As Long, m_rowAbiDate     As Long
Private m_rowAbiTeacher As Long, m_rowAbiClass    As Long
Private m_pupiFirstRow  As Long, m_firstSectRow As Long
Private m_exerCntRow    As Long

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
            Dim i As Integer, sheetCnt As Integer
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
            Dim wsPrint As Worksheet
            Set wsPrint = Worksheets(WbNamePrintSheet)
            Dim blockSizeMain As Integer
            blockSizeMain = (4 * (sheetCnt + 1)) + 2
            With wsPrint.PageSetup
                .Orientation = xlLandscape
                .FitToPagesWide = 1
                .PrintArea = "A1:Q" & CStr(5 + (sheetCnt * 4) + (gNumOfPupils - 1) * blockSizeMain + 29) '29 Zellen für das Chart
                .LeftMargin = Application.CentimetersToPoints(1)
                .RightMargin = Application.CentimetersToPoints(1)
                .TopMargin = Application.CentimetersToPoints(1)
                .BottomMargin = Application.CentimetersToPoints(1)
                .CenterHorizontally = True
            End With
            For i = 1 To gNumOfPupils
                wsPrint.HPageBreaks.Add Before:=wsPrint.Cells(1 + i * blockSizeMain, 1)
            Next i
            
            '------------------------------------
            ' Druckdialog öffnen
            '------------------------------------
            Dim doPrint As Integer
            doPrint = MsgBox("Möchten Sie drucken?", vbQuestion + vbOKCancel, "Erst mal gucken oder ...")
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

    Dim i As Integer, u As Integer
    Dim pupilRow As Integer, secRow As Integer
    Set m_wsPrint = Worksheets(WbNamePrintSheet)
    Set m_wsCfg = Worksheets(WbNameConfig)

    Dim blockSize As Integer
    blockSize = (4 * (sheetCnt + 1)) + 2

    '------------------------------------
    ' Set row heights and column widths
    '------------------------------------
    m_wsPrint.Rows("1:1000").RowHeight = 15
    m_wsPrint.Columns(1).ColumnWidth = 16.71
    m_wsPrint.Range(m_wsPrint.Columns(2), m_wsPrint.Columns(2 + CfgMaxExercisesPerSection)).ColumnWidth = 5.57

    '------------------------------------
    ' Pre-compute column letters and rows (once)
    '------------------------------------
    m_clAbiTitle = ColLetter(Range(CfgAbiTitle).Column)
    m_clAbiDate = ColLetter(Range(CfgAbiDate).Column)
    m_clAbiTeacher = ColLetter(Range(CfgAbiTeacher).Column)
    m_clAbiClass = ColLetter(Range(CfgAbiClass).Column)
    m_clPupiFirst = ColLetter(Range(CfgFirstPupi).Column + 1)
    m_clPupiLast = ColLetter(Range(CfgFirstPupi).Column + 2)
    m_clCfgStart1 = ColLetter(CfgColStart + 1)
    m_clPrintName = ColLetter(CfgPrintNameCol)
    m_rowAbiTitle = Range(CfgAbiTitle).row
    m_rowAbiDate = Range(CfgAbiDate).row
    m_rowAbiTeacher = Range(CfgAbiTeacher).row
    m_rowAbiClass = Range(CfgAbiClass).row
    m_pupiFirstRow = Range(CfgFirstPupi).row
    m_firstSectRow = Range(CfgFirstSect).row
    m_exerCntRow = Range(CfgExerCount).row

    '------------------------------------
    ' Write per-pupil blocks
    '------------------------------------
    For i = 0 To gNumOfPupils - 1
        pupilRow = 1 + i * blockSize
        Call WriteHeader(pupilRow, i, sheetCnt)
        For u = 0 To sheetCnt
            secRow = pupilRow + 2 + u * 4
            If u < sheetCnt Then
                Call WriteSection(secRow, pupilRow, i, u)
            Else
                Call WriteGesamt(secRow, pupilRow, sheetCnt)
            End If
        Next u
    Next i

    ' Create Chart
    Dim chartRow As Integer
    chartRow = gNumOfPupils * blockSize + 2
    Call AddGradeDistribution(WbNamePrintSheet, chartRow, 1)

End Function

' Writes the bold header row for one pupil (Abi title, name, course).
Private Sub WriteHeader(pupilRow As Integer, pupilIdx As Integer, sheetCnt As Integer)

    With m_wsPrint.Range(m_wsPrint.Cells(pupilRow, 1), m_wsPrint.Cells(pupilRow, 17))
        .Select
        Call setBorder(False, True, True, True, True, xlThin, 0, True)
        .Font.Size = 12
        .Font.Bold = True
    End With
    m_wsPrint.Range(m_wsPrint.Cells(pupilRow, CfgPrintNameCol), m_wsPrint.Cells(pupilRow, 12)).HorizontalAlignment = xlCenterAcrossSelection
    ' Abi
    m_wsPrint.Cells(pupilRow, 1).Formula = "=" & CfgRef(m_clAbiTitle, m_rowAbiTitle) & "&"" ""&TEXT(" & CfgRef(m_clAbiDate, m_rowAbiDate) & ",""TT.MM.JJJJ"")"
    ' Name
    m_wsPrint.Cells(pupilRow, CfgPrintNameCol).Formula = "=" & CfgRef(m_clPupiFirst, m_pupiFirstRow + pupilIdx) & "&"", ""&" & CfgRef(m_clPupiLast, m_pupiFirstRow + pupilIdx)
    ' Kurs
    With m_wsPrint.Cells(pupilRow, 17)
        .Formula = "=" & CfgRef(m_clAbiTeacher, m_rowAbiTeacher) & "&"", Kurs ""&" & CfgRef(m_clAbiClass, m_rowAbiClass)
        .Select
        Call setBorder(False, False, True, True, True, xlThin, 0, True, xlRight)
    End With
    ' Notenbereich formatieren
    m_wsPrint.Range(m_wsPrint.Cells(pupilRow + 1, 2), m_wsPrint.Cells(pupilRow + 4 * (sheetCnt + 1), 17)).HorizontalAlignment = xlCenter

End Sub

' Writes the task names, max-BE and achieved-BE rows for one section of one pupil.
Private Sub WriteSection(secRow As Integer, pupilRow As Integer, pupilIdx As Integer, sectIdx As Integer)

    Dim p As Integer, idx As Integer
    Dim sec As String, tsk As String
    Dim clSectU As String, clSectUNext As String, clExCntU As String
    Dim exCnt As Integer, vlookupBase As String
    Dim arrTaskFml() As Variant, arrMaxFml() As Variant, arrVlookup() As Variant

    clSectU = ColLetter(Range(CfgFirstSect).Column + sectIdx * 2)
    clSectUNext = ColLetter(Range(CfgFirstSect).Column + sectIdx * 2 + 1)
    clExCntU = ColLetter(Range(CfgExerCount).Column + sectIdx * 2 + 1)
    exCnt = m_wsCfg.Range(CfgExerCount).Offset(0, sectIdx * 2).Value
    vlookupBase = ",'" & m_wsCfg.Range(CfgFirstSect).Offset(0, sectIdx * 2).Value & "'!PupilBlock,"

    m_wsPrint.Cells(secRow, 1).Font.Bold = True
    m_wsPrint.Cells(secRow, 1).Value = m_wsCfg.Range(CfgFirstSect).Offset(0, sectIdx * 2).Value
    m_wsPrint.Cells(secRow + 1, 1).Value = "max BE"
    m_wsPrint.Cells(secRow + 2, 1).Value = "erreichte BE"

    If StrComp(m_wsCfg.Range(CfgSelEx).Offset(0, sectIdx * 2).MergeArea.Cells(1, 1).Text, "Ja") = 0 Then
        ' Wahlaufgaben: only write selected exercises (sparse)
        idx = 0
        For p = 0 To exCnt - 1
            sec = CStr(m_wsCfg.Range(CfgFirstSect).Offset(0, sectIdx * 2).Value)
            tsk = CStr(m_wsCfg.Range(CfgFirstSect).Offset(p + 2, sectIdx * 2).Value)
            If PupilHasSelEx(CInt(pupilIdx), sec, tsk) Then
                m_wsPrint.Cells(secRow, 2 + idx).Formula = "=" & CfgRef(clSectU, 2 + m_firstSectRow + p)
                m_wsPrint.Cells(secRow + 1, 2 + idx).Formula = "=" & CfgRef(clSectUNext, 2 + m_firstSectRow + p)
                m_wsPrint.Cells(secRow + 2, 2 + idx).Formula = "=VLOOKUP(" & m_clPrintName & CStr(pupilRow) & vlookupBase & p + 2 & ",0)"
                idx = idx + 1
            End If
        Next p
        If idx > 0 Then m_wsPrint.Range(m_wsPrint.Cells(secRow, 2), m_wsPrint.Cells(secRow, 1 + idx)).Font.Bold = True
    Else
        ' Normal exercises: batch-write all columns at once
        idx = exCnt
        ReDim arrTaskFml(1 To 1, 1 To exCnt)
        ReDim arrMaxFml(1 To 1, 1 To exCnt)
        ReDim arrVlookup(1 To 1, 1 To exCnt)
        For p = 0 To exCnt - 1
            arrTaskFml(1, p + 1) = "=" & CfgRef(clSectU, 2 + m_firstSectRow + p)
            arrMaxFml(1, p + 1) = "=" & CfgRef(clSectUNext, 2 + m_firstSectRow + p)
            arrVlookup(1, p + 1) = "=VLOOKUP(" & m_clPrintName & CStr(pupilRow) & vlookupBase & p + 2 & ",0)"
        Next p
        With m_wsPrint
            .Range(.Cells(secRow, 2), .Cells(secRow, 1 + exCnt)).Formula = arrTaskFml
            .Range(.Cells(secRow, 2), .Cells(secRow, 1 + exCnt)).Font.Bold = True
            .Range(.Cells(secRow + 1, 2), .Cells(secRow + 1, 1 + exCnt)).Formula = arrMaxFml
            .Range(.Cells(secRow + 2, 2), .Cells(secRow + 2, 1 + exCnt)).Formula = arrVlookup
        End With
    End If

    ' Sum column
    With m_wsPrint
        .Cells(secRow, 2 + idx).Value = ChrW(931)
        .Cells(secRow, 2 + idx).Font.Bold = True
        .Cells(secRow + 1, 2 + idx).Formula = "=" & CfgRef(clExCntU, m_exerCntRow)
        .Cells(secRow + 1, 2 + idx).Font.Bold = True
        .Cells(secRow + 2, 2 + idx).Formula = "=SUM(B" & CStr(secRow + 2) & ":" & ColLetter(idx + 1) & CStr(secRow + 2) & ")"
        .Cells(secRow + 2, 2 + idx).Font.Bold = True
    End With

End Sub

' Writes the overall "Gesamt" totals row for one pupil.
Private Sub WriteGesamt(secRow As Integer, pupilRow As Integer, sheetCnt As Integer)

    Dim p As Integer, exCntP As Integer
    Dim arrAbbrev() As Variant, arrMaxFml() As Variant

    m_wsPrint.Cells(secRow, 1).Font.Bold = True
    m_wsPrint.Cells(secRow, 1).Value = "Gesamt"
    m_wsPrint.Cells(secRow + 1, 1).Value = "max BE"
    m_wsPrint.Cells(secRow + 2, 1).Value = "erreichte BE"

    ReDim arrAbbrev(1 To 1, 1 To sheetCnt)
    ReDim arrMaxFml(1 To 1, 1 To sheetCnt)
    For p = 0 To sheetCnt - 1
        arrAbbrev(1, p + 1) = SectAbbrev(m_wsCfg.Range(CfgFirstSect).Offset(0, p * 2).Text)
        arrMaxFml(1, p + 1) = "=" & CfgRef(ColLetter(Range(CfgExerCount).Column + p * 2 + 1), m_exerCntRow)
        exCntP = m_wsCfg.Range(CfgExerCount).Offset(0, p * 2).Value
        m_wsPrint.Cells(secRow + 2, 2 + p).Formula = "=VLOOKUP(" & m_clPrintName & CStr(pupilRow) & ",'" & _
            m_wsCfg.Range(CfgFirstSect).Offset(0, p * 2).Value & "'!PupilBlock," & exCntP + 2 & ",0)"
    Next p
    With m_wsPrint
        .Range(.Cells(secRow, 2), .Cells(secRow, 1 + sheetCnt)).Value = arrAbbrev
        .Range(.Cells(secRow, 2), .Cells(secRow, 1 + sheetCnt)).Font.Bold = True
        .Range(.Cells(secRow + 1, 2), .Cells(secRow + 1, 1 + sheetCnt)).Formula = arrMaxFml
        ' Totals column
        .Cells(secRow, 2 + sheetCnt).Value = ChrW(931)
        .Cells(secRow, 2 + sheetCnt).Font.Bold = True
        .Cells(secRow + 1, 2 + sheetCnt).Formula = "=SUM(B" & CStr(secRow + 1) & ":" & ColLetter(sheetCnt + 1) & CStr(secRow + 1) & ")"
        .Cells(secRow + 1, 2 + sheetCnt).Font.Bold = True
        .Cells(secRow + 2, 2 + sheetCnt).Value = "=SUM(B" & CStr(secRow + 2) & ":" & ColLetter(sheetCnt + 1) & CStr(secRow + 2) & ")"
        .Cells(secRow + 2, 2 + sheetCnt).Font.Bold = True
    End With

    ' Notenpunkte
    With m_wsPrint.Range(m_wsPrint.Cells(secRow + 1, 16), m_wsPrint.Cells(secRow + 2, 17))
        .Select
        Call setBorder(False, True, True, True, True, xlThin, RGB(255, 255, 255), True, xlHAlignCenterAcrossSelection, xlBottom)
    End With
    m_wsPrint.Cells(secRow + 1, 16).Value = "NP"
    m_wsPrint.Cells(secRow + 1, 16).Font.Bold = True
    m_wsPrint.Cells(secRow + 2, 16).Formula = "=VLOOKUP(" & ColLetter(2 + sheetCnt) & CStr(secRow + 2) & "," & WbNameGradeKey & CfgVLookUpPoints & ")"
    m_wsPrint.Cells(secRow + 2, 16).Font.Bold = True

End Sub

' Returns a formula reference to a cell in WbNameConfig: 'SheetName'!ColRow
Private Function CfgRef(col As String, row As Long) As String
    CfgRef = "'" & WbNameConfig & "'!" & col & CStr(row)
End Function

' Returns a short section abbreviation for the Gesamt header row.
Private Function SectAbbrev(sectName As String) As String
    If InStr(sectName, " ") > 0 Then
        SectAbbrev = left(sectName, 3) & right(sectName, 1)
    Else
        SectAbbrev = left(sectName, 4)
    End If
End Function




