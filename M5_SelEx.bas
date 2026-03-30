Attribute VB_Name = "M5_SelEx"
Option Explicit

Public Function PaintSelXCfgPage()

    Dim i As Integer
    
    '----------------------------------------
    ' Create work sheet for selex config
    '----------------------------------------
    Dim actSheetName As String

    ' Set name for further processing
    actSheetName = WbNameSelExConfig
     
    '------------------------------------
    ' Delete old sheet if exists
    '------------------------------------
    If WSExists(actSheetName) Then
        Worksheets(actSheetName).Delete
    End If
    ' Create new sheet
    Worksheets.Add(Before:=Worksheets(WbNameConfig)).name = actSheetName
    Worksheets(actSheetName).Tab.color = gClrTabConfig
    
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
    Worksheets(actSheetName).Columns(CfgColStart + 1).ColumnWidth = 25  ' Spalte C: Schüler-Name
    
    ' Find all column widths based on selection if this section contains choosable exercises
    Dim tblIdx As Integer
    Dim SelCfg As String
    Dim colOffset As Integer
    
    colOffset = 0
    
    For tblIdx = 0 To CfgMaxSheets
        ' If this section consists of choosable exercises, add them to the new config page
        SelCfg = Worksheets(WbNameConfig).range(CfgSelEx).Offset(0, tblIdx * 2).MergeArea.Cells(1, 1).Text
        If StrComp(SelCfg, "Ja") = 0 Then
            
            ' Set column widths for sub exercises
            Dim numOfSubEx As Integer
            Dim subExIdx, subEx As Integer
            numOfSubEx = Worksheets(WbNameConfig).range(CfgExerCount).Offset(0, tblIdx * 2).Value
            If numOfSubEx > 0 Then
                For subEx = 0 To numOfSubEx - 1
                    Worksheets(actSheetName).Columns(CfgColStart + CfgColOffsetFirstEx + colOffset + subEx).ColumnWidth = 4
                Next subEx
            End If
            
            colOffset = colOffset + numOfSubEx
            
        End If
    Next tblIdx
    
    ' Spacer column to button
    Worksheets(actSheetName).Columns(CfgColStart + CfgColOffsetFirstEx + colOffset).ColumnWidth = 2
    
    
    '------------------------------------
    ' Create frames and formatting
    '------------------------------------
    Dim span As Integer
    span = colOffset + 1 ' Anzahl der Teilaufgaben + 3 Spalten (Index,Name,Summe)
    
    ' Abi Zelle
    Worksheets(actSheetName).range(Cells(CfgRowStart, CfgColStart), Cells(CfgRowStart, CfgColStart + CInt(floor(CDbl(span) / 2)))).Select
    Call setBorder(False, True, False, True, False, xlMedium, gClrHeader, True, xlLeft, xlCenter)
    Selection.Font.Bold = True
    ' Kurs Zelle
    Worksheets(actSheetName).range(Cells(CfgRowStart, CfgColStart + CInt(floor(CDbl(span) / 2)) + 1), Cells(CfgRowStart, CfgColStart + span)).Select
    Call setBorder(False, False, True, True, False, xlMedium, gClrHeader, True, xlRight, xlCenter)
    Selection.Font.Bold = True
    ' Bereich Zelle
    Worksheets(actSheetName).range(Cells(CfgRowStart + 1, CfgColStart), Cells(CfgRowStart + 2, CfgColStart + span)).Select
    Call setBorder(False, True, True, False, True, xlMedium, gClrHeader, True, xlHAlignCenterAcrossSelection, xlCenter)
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    
    ' Überschrift Namen
    Worksheets(actSheetName).range(Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1)).Select
    Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, True)
    
    ' Überschrift Aufgaben / Punkte
    Worksheets(actSheetName).range(Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx), Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + span)).Select
    Call setBorder(False, True, True, True, True, xlMedium, gClrTheme1, False, xlCenter, xlBottom)
    
    ' Namen
    Worksheets(actSheetName).range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx - 1)).Select
    Call setBorder(False, True, True, True, True, xlThin, gClrTheme1, False)
    
    ' Punkte-Bereich
    Worksheets(actSheetName).range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + colOffset - 1)).Select
    Call setBorder(False, True, True, True, True, xlThin, RGB(255, 255, 255), False, xlCenter, xlCenter)
    Selection.Locked = False
    
    ' Erlaube nur "x" als Zelleninhalt, um zu setzen, dass der Schüler diese Aufgabe ausgewählt hat ...
    For i = 0 To colOffset - 1
        With Worksheets(actSheetName).range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + i), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + i))
            .Select
            With .Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="x,"
                .IgnoreBlank = True
                .InCellDropdown = False
            End With
        End With
    Next i
    
    ' Border anpassen
    Worksheets(actSheetName).range(Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + colOffset - 1)).Select
    Call setBorder(False, True, True, True, True, xlMedium, 0, True)
    
    ' Button hinzufügen
    Call AddUpdateButton(Worksheets(actSheetName).Cells(CfgRowStart, CfgColStart + CfgColOffsetFirstEx + colOffset + 1), "SelExUpdate")
    ' Button handler hinzufügen
    Call InjectWorksheet_ButtonHandler(Worksheets(actSheetName))
    
    ' Kommentarfeld für HowTo
    Dim txtCommentFieldWidth, txtCommentFieldHeight As Integer
    txtCommentFieldWidth = 4 ' columns
    txtCommentFieldHeight = 3 ' rows
    Worksheets(actSheetName).range(Cells(CfgRowStart + 4, CfgColStart + CfgColOffsetFirstEx + colOffset + 1), Cells(CfgRowStart + 4 + txtCommentFieldHeight - 1, CfgColStart + CfgColOffsetFirstEx + colOffset + 1 + txtCommentFieldWidth)).Select
    Call setBorder(True, True, True, True, True, xlMedium, 0, True, xlHAlignLeft, xlVAlignCenter)
    Worksheets(actSheetName).Cells(CfgRowStart + 4, CfgColStart + CfgColOffsetFirstEx + colOffset + 1).WrapText = True
    Worksheets(actSheetName).Cells(CfgRowStart + 4, CfgColStart + CfgColOffsetFirstEx + colOffset + 1).Value = "In nebenstehender Tabelle, alle gewählten Aufgaben des Schülers mit ""x"" selektieren. Anschließend Button anklicken!"
    
    
    Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx).Select
    
End Function


Public Function FillSelXCfgPage()

   '----------------------------------------
    ' Create Worksheets
    '----------------------------------------
    Dim actSheetName As String
    Dim numOfSubEx, span As Integer
    Dim i As Integer
    Dim funStr1 As String
    Dim funStr2 As String
    
    
    ' Set name for further processing
    actSheetName = WbNameSelExConfig
    
    ' Find all column widths based on selection if this section contains choosable exercises
    Dim tblIdx As Integer
    Dim SelCfg As String
    Dim colOffset As Integer
    
    colOffset = 0
    
    '------------------------------------
    ' Header-Text
    '------------------------------------
    Worksheets(actSheetName).Cells(CfgRowStart, CfgColStart).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, range(CfgAbiTitle).Column).Address, "$")(1) & CStr(range(CfgAbiTitle).row) & "&"" ""&" & "YEAR('" & WbNameConfig & "'!" & Split(Cells(1, range(CfgAbiDate).Column).Address, "$")(1) & CStr(range(CfgAbiDate).row) & ")"
    Worksheets(actSheetName).Cells(CfgRowStart, CfgColStart + span).Formula = "=""Kurs ""&'" & WbNameConfig & "'!" & Split(Cells(1, range(CfgAbiClass).Column).Address, "$")(1) & CStr(range(CfgAbiClass).row)
    Worksheets(actSheetName).Cells(CfgRowStart + 1, CfgColStart).Value = "Wahlfachkonfiguration"
    
    '------------------------------------
    ' Überschrift Name und Punkte
    '------------------------------------
    Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1).Value = "Name"
    For i = 0 To gNumOfPupils - 1
        Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart).Value = Worksheets(WbNameConfig).range(CfgFirstPupi).Offset(i, 0).Value
        Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil + i, CfgColStart + 1).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, range(CfgFirstPupi).Column + 1).Address, "$")(1) & CStr(range(CfgFirstPupi).row + i) & "&"", ""&'" & WbNameConfig & "'!" & Split(Cells(1, range(CfgFirstPupi).Column + 2).Address, "$")(1) & CStr(range(CfgFirstPupi).row + i)
    Next i
        
    For tblIdx = 0 To CfgMaxSheets
        ' If this section consists of choosable exercises, add them to the new config page
        SelCfg = Worksheets(WbNameConfig).range(CfgSelEx).Offset(0, tblIdx * 2).MergeArea.Cells(1, 1).Text
        If StrComp(SelCfg, "Ja") = 0 Then
            
            numOfSubEx = Worksheets(WbNameConfig).range(CfgExerCount).Offset(0, tblIdx * 2).Value
            For i = 0 To numOfSubEx - 1
                Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + colOffset + CfgColOffsetFirstEx + i).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, range(CfgFirstSect).Column + tblIdx * 2).Address, "$")(1) & CStr(2 + range(CfgFirstSect).row + i)
                ' WS Name anstatt Punktezahl, um die Referenz zu behalten
                'Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + colOffset + CfgColOffsetFirstEx + i).Formula = "='" & WbNameConfig & "'!" & Split(Cells(1, Range(CfgFirstSect).Column + (tblIdx * 2) + 1).Address, "$")(1) & CStr(2 + Range(CfgFirstSect).row + i)
                Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + colOffset + CfgColOffsetFirstEx + i).Value = Worksheets(WbNameConfig).range(CfgFirstSect).Offset(0, tblIdx * 2).Text
                Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + colOffset + CfgColOffsetFirstEx + i).WrapText = True
            Next i
            
            colOffset = colOffset + numOfSubEx
            
        End If
    Next tblIdx
    
End Function


Public Function InjectWorksheet_ButtonHandler(ws As Worksheet)

    Dim codeModule As Object
    Set codeModule = ws.Parent.VBProject _
        .VBComponents(ws.CodeName).codeModule

    On Error Resume Next
    Dim lineNum As Long
    lineNum = codeModule.ProcStartLine("btnSelExUpdate_Click", 0)
    If lineNum > 0 Then
        codeModule.DeleteLines lineNum, codeModule.ProcCountLines("btnSelExUpdate_Click", 0)
    End If
    On Error GoTo 0

    codeModule.AddFromString _
        "Private Sub btnSelExUpdate_Click()" & vbCrLf & _
        "    " & gBtnSelXUpdateMacro & vbCrLf & _
        "End Sub"

End Function


Public Function AddUpdateButton(targetCell As range, onClickMacro As String)

    Dim btn As OLEObject
    
    gBtnSelXUpdateMacro = onClickMacro
    
    Set btn = targetCell.Worksheet.OLEObjects.Add( _
        ClassType:="Forms.CommandButton.1", _
        left:=targetCell.left, _
        top:=targetCell.top, _
        Width:=Application.CentimetersToPoints(3.78), _
        Height:=Application.CentimetersToPoints(1.42))
    
    With btn
        .name = "btnSelExUpdate"
        .Placement = xlFreeFloating
        .PrintObject = False
    End With
    
    With btn.Object
        .Caption = "Blätter aktualisieren"
        .BackColor = &H80FF80
        .BackStyle = 1   ' fmBackStyleOpaque
        .Font.Size = 10
    End With

End Function


Public Function SelExUpdate()

    Call Init

    ' Über alle Sections loopen
    Dim actSheetName As String
    actSheetName = WbNameSelExConfig
    
    ' Abfragen ob wirklich neue Tabellen erstellt werden sollen...
    Dim Request As Integer
    Request = MsgBox("Sicher, dass sie die Tabellen aktualisieren wollen?" & vbCrLf & "Bereits ausgefüllte Punkte können bei der Aktualisierung verloren gehen!", vbExclamation + vbOKCancel, "Sicher?")
    If Request = vbCancel Then
        Exit Function
    End If
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim i As Integer
    Dim u As Integer
    Dim currSht As String
    Dim currTsk As String
    Dim currCol As Integer
    Dim pLine As Integer
    
    ' So lange weiter bis leere Zelle kommt
    i = 0
    Do While True
        If Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Value = "" Then
            Exit Do
        End If
        
        ' Aktuelles Blatt und aktuelle Aufgabe
        currSht = Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Value
        currTsk = Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + i).Value
        
        ' Spalte in Ziel finden
        currCol = 255
        For u = 0 To CfgMaxExercisesPerSection - 1
            If Worksheets(currSht).Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + u).Value = currTsk Then
                currCol = CfgColStart + CfgColOffsetFirstEx + u
                Exit For
            End If
        Next u
        ' If column was not found, don't continue!
        If currCol = 255 Then
            Exit Function
        End If
        
        ' Über jeden Schüler iterieren und auf dem passenden Worksheet
        '  - die nicht gewählten Aufgaben sperren, hintergrund grau, rotes X rein
        '  - die gewählten entsperren, hintergrund normal, Inhalt leeren
        For pLine = 0 To gNumOfPupils - 1
            If Worksheets(actSheetName).Cells(CfgRowStart + CfgRowOffsetFirstPupil + pLine, CfgColStart + CfgColOffsetFirstEx + i).Value = "x" Then
                Call RemoveCross(Worksheets(currSht).Cells(CfgRowStart + CfgRowOffsetFirstPupil + pLine, currCol))
            Else
                Call CrossOutCell(Worksheets(currSht).Cells(CfgRowStart + CfgRowOffsetFirstPupil + pLine, currCol))
            End If
        Next pLine
        
        i = i + 1
    Loop
    
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Function


Public Function CrossOutCell(cell As range)

    Dim ws As Worksheet
    Set ws = cell.Worksheet
    
    Dim l As Double, T As Double, w As Double, h As Double
    l = cell.left
    T = cell.top
    w = cell.Width
    h = cell.Height
    
    ws.Unprotect
    
    ' Set bg
    cell.Interior.color = gClrBg2
    cell.ClearContents

    ' Remove existing cross if present
    On Error Resume Next
    ws.Shapes("Cross_" & cell.Address(False, False)).Delete
    On Error GoTo 0

    ' Draw both diagonal lines
    Dim line1 As Shape, line2 As Shape
    Set line1 = ws.Shapes.AddLine(l, T, l + w, T + h)
    With line1.Line
        .ForeColor.RGB = RGB(255, 150, 150)
        .Weight = 0.75
        .Transparency = 0.4
    End With

    Set line2 = ws.Shapes.AddLine(l + w, T, l, T + h)
    With line2.Line
        .ForeColor.RGB = RGB(255, 150, 150)
        .Weight = 0.75
        .Transparency = 0.4
    End With

    ' Group and name them
    Dim grp As Shape
    Set grp = ws.Shapes.range(Array(line1.name, line2.name)).Group
    grp.name = "Cross_" & cell.Address(False, False)
    grp.Locked = True

    ' Protect drawing objects only (leaves cell editing untouched)
    cell.Locked = True
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=False

End Function

Public Sub RemoveCross(cell As range)
    Dim ws As Worksheet
    Set ws = cell.Worksheet
    ws.Unprotect
    cell.Locked = False
    On Error Resume Next
    cell.Interior.color = gClrBg1
    ws.Shapes("Cross_" & cell.Address(False, False)).Delete
    On Error GoTo 0
    ws.Protect DrawingObjects:=True, Contents:=False, Scenarios:=False
End Sub

Public Function IsSelEx(Section As String)
    Dim i As Integer
    For i = 0 To CfgMaxSheets
        If Worksheets(WbNameConfig).range(CfgFirstSect).Offset(0, i * 2).Text = Section Then
            If StrComp(Worksheets(WbNameConfig).range(CfgSelEx).Offset(0, i * 2).MergeArea.Cells(1, 1).Text, "Ja") = 0 Then
                IsSelEx = True
                Exit Function
            Else
                IsSelEx = False
                Exit Function
            End If
        End If
    Next i
    IsSelEx = False
End Function


Public Function PupilHasSelEx(PupilIndex As Integer, Section As String, Number As String)

    ' Is not SelEx section -> Always return true
    If Not IsSelEx(Section) Then
        PupilHasSelEx = True
        Exit Function
    Else
    
        ' Reihe des Schülers
        Dim pupilRow As Integer
        pupilRow = CfgRowStart + CfgRowOffsetFirstPupil + PupilIndex
    
        Dim i As Integer
        i = 0
        Do While True
            If Worksheets(WbNameSelExConfig).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Value = "" Then
                Exit Do
            End If
            
            Dim currSht, currTsk As String
            currSht = Worksheets(WbNameSelExConfig).Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Value
            currTsk = Worksheets(WbNameSelExConfig).Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + i).Value
            
            If currSht = Section And currTsk = Number Then
                If Worksheets(WbNameSelExConfig).Cells(pupilRow, CfgColStart + CfgColOffsetFirstEx + i).Value = "x" Then
                    PupilHasSelEx = True
                    Exit Function
                Else
                    PupilHasSelEx = False
                    Exit Function
                End If
            End If
            
            i = i + 1
        Loop
        
    End If
    PupilHasSelEx = False

End Function

