VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call CheckForUpdate(Version)
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    '----------------------------------------
    ' Grade sheet
    '----------------------------------------
    ' Update colors on grade sheet
    If Sh.Name = WbNameGradeSheet Then
        Call UpdateUpDownColors
    ElseIf Sh.Name <> WbNameConfig And Sh.Name <> WbNamePrintSheet And Sh.Name <> WbNameGradeKey And Sh.Name <> WbNameTestDaten And Sh.Name <> WbNameSelExConfig Then
        Call UpdateZKDKMismatchHighlight
    End If
    
    '----------------------------------------
    ' Print sheet
    '----------------------------------------
    ' Update max y-ax of chart
    If Sh.Name = WbNamePrintSheet Then
        Call Init
        Worksheets(WbNamePrintSheet).Unprotect Password:=WbPw
        ' Hinter dem Chart stehen die Notenpunkte und ihre jeweilige Anzahl als Verweis.
        ' Ändert sich die Anzahl der Noten, so muss der max Datenpunt der Achse neu gesetzt werden
        ' Dafür wird als erste die Adresse der Zelle neu gebildet
        Dim row As Integer
        row = gNumOfPupils * (4 * (gSheetCnt + 1) + 2) + 2
        ' Aktuell fest Spalte 1! Und anschließend der Wert gesetzt
        Worksheets(WbNamePrintSheet).ChartObjects(CfgNameChart).Chart.Axes(xlValue).MaximumScale = CInt(ThisWorkbook.Worksheets(WbNamePrintSheet).Cells(row, 1).offset(1, 16).Value) + 1
        If DevMode <> 1 Then
            Worksheets(WbNamePrintSheet).Protect Password:=WbPw
        End If
    End If
    
    '----------------------------------------
    ' Lock/Unlock all Worksheets
    '----------------------------------------
    If Sh.Name = WbNameConfig Then
        Dim ws As Worksheet
        If DevMode = 1 And ThisWorkbook.Worksheets(WbNameConfig).ProtectContents Then
            For Each ws In ThisWorkbook.Worksheets
                ws.Unprotect Password:=WbPw
            Next ws
        ElseIf DevMode <> 1 And Not ThisWorkbook.Worksheets(WbNameConfig).ProtectContents Then
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name <> WbNamePrintSheet Then
                    ws.Protect Password:=WbPw, DrawingObjects:=True, Contents:=True, Scenarios:=False
                    ws.EnableSelection = xlUnlockedCells
                End If
            Next ws
        End If
    End If
    '----------------------------------------
    ' Always recalculate to update names
    '----------------------------------------
    Application.CalculateFull
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    Application.ScreenUpdating = False
    If Sh.Name = WbNameGradeSheet Then
        Call UpdateUpDownColors
    ElseIf Sh.Name <> WbNameConfig And Sh.Name <> WbNamePrintSheet And Sh.Name <> WbNameGradeKey And Sh.Name <> WbNameTestDaten And Sh.Name <> WbNameSelExConfig Then
        Call UpdateZKDKMismatchHighlight(Sh)
        Call NavigateAfterEntry(Sh, Target)
    End If
    Application.ScreenUpdating = True
       
End Sub

