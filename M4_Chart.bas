Attribute VB_Name = "M4_Chart"
Public Function AddGradeDistribution(ws As String, row As Integer, col As Integer)

    '------------------------------------
    ' Collect chart data
    '------------------------------------
    Dim i As Integer
    ' Count actual sheets
    Dim sheetCnt As Integer
    sheetCnt = 0
    For i = 0 To CfgMaxSheets
        If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value) Then
            sheetCnt = sheetCnt + 1
        End If
    Next i
    For i = 0 To 15
        ThisWorkbook.Worksheets(ws).Cells(row, col).Offset(0, i).Value = i
        ThisWorkbook.Worksheets(ws).Cells(row, col).Offset(1, i).Formula = "=COUNTIF('" & WbNameGradeSheet & "'!$" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Address, "$")(1) & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil) & ":$" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + sheetCnt + 1).Address, "$")(1) & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils) & "," & Split(Cells(1, col + i).Address, "$")(1) & CStr(row) & ")"
    Next i
    ThisWorkbook.Worksheets(ws).Cells(row, col).Offset(1, 16).Formula = "=MAX($" & Split(Cells(1, col).Address, "$")(1) & "$" & CStr(row + 1) & ":$" & Split(Cells(1, col + 15).Address, "$")(1) & "$" & CStr(row + 1) & ")"
    Application.CalculateFull
    
    '------------------------------------
    ' Create chart
    '------------------------------------
    'Border
    Worksheets(WbNamePrintSheet).Range(Cells(row - 1, 1), Cells(row - 1, 17)).Select
    Call setBorder(False, True, True, True, True, xlThin, 0, True)
    With Selection.Font
        .Size = 12
        .Bold = True
    End With
    ' Notenverteilung " & Format(Worksheets(WbNameGradeSheet).Range("K30").Value, "0.00")
    Worksheets(WbNamePrintSheet).Cells(row - 1, CfgPrintNameCol).Formula = "=""Notenverteilung - ""& CHAR(216) & "" "" & TEXT('" & WbNameGradeSheet & "'!$" & Split(Cells(1, CfgColStart + CfgColOffsetFirstEx + gSheetCnt + 1).Address, "$")(1) & "$" & CStr(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils) & ",""0,00"")"
    Worksheets(WbNamePrintSheet).Range(Cells(row - 1, CfgPrintNameCol), Cells(row - 1, CfgPrintNameCol + 6)).Select
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
    ' Abi
    Worksheets(WbNamePrintSheet).Cells(row - 1, col).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiTitle).Column).Address, "$")(1) & CStr(Range(CfgAbiTitle).row) & "&"" ""&" & "TEXT('" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiDate).Column).Address, "$")(1) & CStr(Range(CfgAbiDate).row) & ",""TT.MM.JJJJ"")"
    ' Kurs
    Worksheets(WbNamePrintSheet).Cells(row - 1, 17).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiTeacher).Column).Address, "$")(1) & CStr(Range(CfgAbiTeacher).row) & "&"", Kurs ""&'" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgAbiClass).Column).Address, "$")(1) & CStr(Range(CfgAbiClass).row)
    Worksheets(WbNamePrintSheet).Cells(row - 1, 17).Select
    Call setBorder(False, False, True, True, True, xlThin, 0, True, xlRight)
    
    '------------------------------------
    ' Create chart
    '------------------------------------
    Dim graphO As ChartObject
    Set graphO = Worksheets(ws).ChartObjects.Add(ThisWorkbook.Worksheets(ws).Cells(row, col).left, ThisWorkbook.Worksheets(ws).Cells(row, col).top, ThisWorkbook.Worksheets(ws).Range("R1").left, 400)
    Dim graph As Chart
    Set graph = graphO.Chart
    
    graphO.Name = CfgNameChart
    
    With graph
        '--------------------------------
        ' Chart type
        '--------------------------------
        .ChartType = xlColumnClustered
        '--------------------------------
        ' Data sources
        '--------------------------------
        .SetSourceData Source:=Worksheets(ws).Range(Cells(row + 1, col), Cells(row + 1, col + 15))
        .SeriesCollection(1).XValues = Worksheets(ws).Range(Cells(row, col), Cells(row, col + 15))
        '--------------------------------
        ' Format axis
        '--------------------------------
        .Axes(xlValue).MaximumScale = CInt(ThisWorkbook.Worksheets(ws).Cells(row, col).Offset(1, 16).Value) + 1
        .Axes(xlValue).MajorUnit = 1
        .Axes(xlValue).MinorUnit = 1
        .Axes(xlValue).Format.Line.Visible = msoFalse
        .Axes(xlValue).HasTitle = msoTrue
        .Axes(xlValue).AxisTitle.Caption = "Anzahl der Schüler"
        With .Axes(xlValue).MajorGridlines.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.150000006
            .Transparency = 0
        End With
        .Axes(xlCategory).HasTitle = msoTrue
        .Axes(xlCategory).AxisTitle.Caption = "Notenpunkte"
        
        '--------------------------------
        ' Format bars
        '--------------------------------
        .ChartGroups(1).Overlap = 0
        .ChartGroups(1).GapWidth = 100
        .SetElement (msoElementDataLabelInsideEnd)
        .FullSeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        '--------------------------------
        ' Format and create title
        '--------------------------------
        .HasTitle = False
        '.ChartTitle.Text = "Notenverteilung - " & Chr(248) & " " & Format(Worksheets(WbNameGradeSheet).Range("K30").Value, "0.00")
        '--------------------------------
        ' Remove legend
        '--------------------------------
        .Legend.Delete
    End With
    ' Update graph
    graph.Refresh
    
End Function
