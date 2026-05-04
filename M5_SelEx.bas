Attribute VB_Name = "M5_SelEx"
Option Explicit

Public Function PaintSelXCfgPage()

    Dim i As Integer
    Dim ws As Worksheet
    Dim numOfSubEx As Integer
    Dim subEx As Integer

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
    ' Create new sheet and cache reference
    Worksheets.Add(Before:=Worksheets(WbNameConfig)).name = actSheetName
    Set ws = Worksheets(actSheetName)
    ws.Tab.color = gClrTabConfig
    
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
    ' Set column widths for sub exercises and for sum column
    ws.Columns(1).ColumnWidth = 2.71          ' Spalte A bleibt leer
    ws.Columns(CfgColStart).ColumnWidth = 2.71   ' Spalte B: Schüler-Index
    ws.Columns(CfgColStart + 1).ColumnWidth = 25 ' Spalte C: Schüler-Name
    
    ' Find all column widths based on selection if this section contains choosable exercises
    Dim tblIdx As Integer
    Dim SelCfg As String
    Dim colOffset As Integer
    
    colOffset = 0
    
    For tblIdx = 0 To CfgMaxSheets
        ' If this section consists of choosable exercises, add them to the new config page
        SelCfg = Worksheets(WbNameConfig).Range(CfgSelEx).Offset(0, tblIdx * 2).MergeArea.Cells(1, 1).Text
        If StrComp(SelCfg, "Ja") = 0 Then
            
            ' Set column widths for sub exercises
            numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, tblIdx * 2).Value
            If numOfSubEx > 0 Then
                ws.Range(ws.Columns(CfgColStart + CfgColOffsetFirstEx + colOffset), _
                         ws.Columns(CfgColStart + CfgColOffsetFirstEx + colOffset + numOfSubEx - 1)).ColumnWidth = 4
            End If
            
            colOffset = colOffset + numOfSubEx
            
        End If
    Next tblIdx
    
    ' Spacer column to button
    ws.Columns(CfgColStart + CfgColOffsetFirstEx + colOffset).ColumnWidth = 2
    
    
    '------------------------------------
    ' Create frames and formatting
    '------------------------------------
    Dim span As Integer
    span = colOffset + 1 ' Anzahl der Teilaufgaben + 3 Spalten (Index,Name,Summe)
    
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
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + colOffset - 1))
        .Select
        Call setBorder(False, True, True, True, True, xlThin, RGB(255, 255, 255), False, xlCenter, xlCenter)
        .Locked = False
    End With
    
    ' Erlaube nur "x" als Zelleninhalt, um zu setzen, dass der Schüler diese Aufgabe ausgewählt hat ...
    For i = 0 To colOffset - 1
        With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx + i), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + i))
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
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + CfgColOffsetFirstEx + colOffset - 1))
        .Select
        Call setBorder(False, True, True, True, True, xlMedium, 0, True)
    End With

    ' Button hinzufügen
    'Call AddUpdateButton(ws.Cells(CfgRowStart, CfgColStart + CfgColOffsetFirstEx + colOffset + 1), "SelExUpdate")
    ' Button handler hinzufügen
    'Call InjectWorksheet_ButtonHandler(ws)

    ' Kommentarfeld für HowTo
    Dim txtCommentFieldWidth As Integer, txtCommentFieldHeight As Integer
    txtCommentFieldWidth = 4 ' columns
    txtCommentFieldHeight = 3 ' rows
    With ws.Range(ws.Cells(CfgRowStart + 4, CfgColStart + CfgColOffsetFirstEx + colOffset + 1), ws.Cells(CfgRowStart + 4 + txtCommentFieldHeight - 1, CfgColStart + CfgColOffsetFirstEx + colOffset + 1 + txtCommentFieldWidth))
        .Select
        Call setBorder(True, True, True, True, True, xlMedium, 0, True, xlHAlignLeft, xlVAlignCenter)
    End With
    ws.Cells(CfgRowStart + 4, CfgColStart + CfgColOffsetFirstEx + colOffset + 1).WrapText = True
    ws.Cells(CfgRowStart + 4, CfgColStart + CfgColOffsetFirstEx + colOffset + 1).Value = "In nebenstehender Tabelle, alle gewählten Aufgaben des Schülers mit ""x"" selektieren. Anschließend den 'Wahlaufgaben aktualisieren' Button auf der Config-Seite anklicken!"

    ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + CfgColOffsetFirstEx).Select
    
End Function


Public Function FillSelXCfgPage()

    '----------------------------------------
    ' Create Worksheets
    '----------------------------------------
    Dim actSheetName As String
    Dim numOfSubEx As Integer, span As Integer
    Dim i As Integer
    Dim ws As Worksheet
    Dim colPupiFirst As String, colPupiLast As String
    Dim cfgPupiFirstRow As Long
    Dim arrIdx() As Variant, arrNames() As Variant
    Dim arrExFml() As Variant, arrExVal() As Variant
    Dim totalColOffset As Integer, preIdx As Integer

    ' Set name for further processing
    actSheetName = WbNameSelExConfig
    Set ws = Worksheets(actSheetName)

    ' Find all column widths based on selection if this section contains choosable exercises
    Dim tblIdx As Integer
    Dim SelCfg As String
    Dim colOffset As Integer

    ' Pre-pass: compute total colOffset so span is correct before writing headers
    totalColOffset = 0
    For preIdx = 0 To CfgMaxSheets
        If StrComp(Worksheets(WbNameConfig).Range(CfgSelEx).Offset(0, preIdx * 2).MergeArea.Cells(1, 1).Text, "Ja") = 0 Then
            totalColOffset = totalColOffset + Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, preIdx * 2).Value
        End If
    Next preIdx
    span = totalColOffset + 1
    colOffset = 0

    '------------------------------------
    ' Header-Text
    '------------------------------------
    ws.Cells(CfgRowStart, CfgColStart).Formula = "='" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiTitle).Column) & CStr(Range(CfgAbiTitle).row) & "&"" ""&" & "YEAR('" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiDate).Column) & CStr(Range(CfgAbiDate).row) & ")"
    ws.Cells(CfgRowStart, CfgColStart + span).Formula = "=""Kurs ""&'" & WbNameConfig & "'!" & ColLetter(Range(CfgAbiClass).Column) & CStr(Range(CfgAbiClass).row)
    ws.Cells(CfgRowStart + 1, CfgColStart).Value = "Wahlfachkonfiguration"
    
    '------------------------------------
    ' Überschrift Name und Punkte
    '------------------------------------
    ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + 1).Value = "Name"
    colPupiFirst = ColLetter(Range(CfgFirstPupi).Column + 1)
    colPupiLast = ColLetter(Range(CfgFirstPupi).Column + 2)
    cfgPupiFirstRow = Range(CfgFirstPupi).row
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
        
    For tblIdx = 0 To CfgMaxSheets
        ' If this section consists of choosable exercises, add them to the new config page
        SelCfg = Worksheets(WbNameConfig).Range(CfgSelEx).Offset(0, tblIdx * 2).MergeArea.Cells(1, 1).Text
        If StrComp(SelCfg, "Ja") = 0 Then
            
            numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, tblIdx * 2).Value
            ReDim arrExFml(1 To 1, 1 To numOfSubEx)
            ReDim arrExVal(1 To 1, 1 To numOfSubEx)
            For i = 0 To numOfSubEx - 1
                arrExFml(1, i + 1) = "='" & WbNameConfig & "'!" & ColLetter(Range(CfgFirstSect).Column + tblIdx * 2) & CStr(2 + Range(CfgFirstSect).row + i)
                ' WS Name anstatt Punktezahl, um die Referenz zu behalten
                'arrExFml(1, i + 1) = "='" & WbNameConfig & "'!" & ColLetter(Range(CfgFirstSect).Column + (tblIdx * 2) + 1) & CStr(2 + Range(CfgFirstSect).row + i)
                arrExVal(1, i + 1) = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, tblIdx * 2).Text
            Next i
            ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + colOffset + CfgColOffsetFirstEx), _
                     ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + colOffset + CfgColOffsetFirstEx + numOfSubEx - 1)).Formula = arrExFml
            ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + colOffset + CfgColOffsetFirstEx), _
                     ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + colOffset + CfgColOffsetFirstEx + numOfSubEx - 1)).Value = arrExVal
            ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + colOffset + CfgColOffsetFirstEx), _
                     ws.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + colOffset + CfgColOffsetFirstEx + numOfSubEx - 1)).WrapText = True
            
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


Public Function AddUpdateButton(targetCell As Range, onClickMacro As String)

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
    Dim found As Boolean
    Dim wsCfg As Worksheet
    Dim wsSht As Worksheet
    Set wsCfg = Worksheets(actSheetName)

    ' So lange weiter bis leere Zelle kommt
    i = 0
    Do While True
        If wsCfg.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Value = "" Then
            Exit Do
        End If

        ' Aktuelles Blatt und aktuelle Aufgabe
        currSht = wsCfg.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Value
        currTsk = wsCfg.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + i).Value
        Set wsSht = Worksheets(currSht)

        ' Spalte in Ziel finden
        found = False
        currCol = -1
        For u = 0 To CfgMaxExercisesPerSection - 1
            If wsSht.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + u).Value = currTsk Then
                currCol = CfgColStart + CfgColOffsetFirstEx + u
                found = True
                Exit For
            End If
        Next u
        ' If column was not found, don't continue!
        If Not found Then
            Exit Function
        End If

        ' Über jeden Schüler iterieren und auf dem passenden Worksheet
        '  - die nicht gewählten Aufgaben sperren, hintergrund grau, rotes X rein
        '  - die gewählten entsperren, hintergrund normal, Inhalt leeren
        wsSht.Unprotect
        For pLine = 0 To gNumOfPupils - 1
            If wsCfg.Cells(CfgRowStart + CfgRowOffsetFirstPupil + pLine, CfgColStart + CfgColOffsetFirstEx + i).Value = "x" Then
                Call RemoveCross(wsSht.Cells(CfgRowStart + CfgRowOffsetFirstPupil + pLine, currCol), skipProtect:=True)
            Else
                Call CrossOutCell(wsSht.Cells(CfgRowStart + CfgRowOffsetFirstPupil + pLine, currCol), skipProtect:=True)
            End If
        Next pLine
        wsSht.Protect DrawingObjects:=True, Contents:=True, Scenarios:=False

        i = i + 1
    Loop
    
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Function


Public Function CrossOutCell(cell As Range, Optional skipProtect As Boolean = False)

    Dim ws As Worksheet
    Set ws = cell.Worksheet
    
    Dim l As Double, T As Double, w As Double, h As Double
    l = cell.left
    T = cell.top
    w = cell.Width
    h = cell.Height
    
    If Not skipProtect Then ws.Unprotect
    
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
    Set grp = ws.Shapes.Range(Array(line1.name, line2.name)).Group
    grp.name = "Cross_" & cell.Address(False, False)
    grp.Locked = True

    ' Protect drawing objects only (leaves cell editing untouched)
    cell.Locked = True
    If Not skipProtect Then ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=False

End Function

Public Sub RemoveCross(cell As Range, Optional skipProtect As Boolean = False)
    Dim ws As Worksheet
    Set ws = cell.Worksheet
    If Not skipProtect Then ws.Unprotect
    cell.Locked = False
    On Error Resume Next
    cell.Interior.color = gClrBg1
    ws.Shapes("Cross_" & cell.Address(False, False)).Delete
    On Error GoTo 0
    If Not skipProtect Then ws.Protect DrawingObjects:=True, Contents:=False, Scenarios:=False
End Sub

Public Function IsSelEx(Section As String)
    Dim i As Integer
    For i = 0 To CfgMaxSheets
        If Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Text = Section Then
            If StrComp(Worksheets(WbNameConfig).Range(CfgSelEx).Offset(0, i * 2).MergeArea.Cells(1, 1).Text, "Ja") = 0 Then
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

        Dim wsSelEx As Worksheet
        Set wsSelEx = Worksheets(WbNameSelExConfig)
        Dim i As Integer
        Dim currSht As String, currTsk As String
        i = 0
        Do While True
            If wsSelEx.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Value = "" Then
                Exit Do
            End If

            currSht = wsSelEx.Cells(CfgRowStart + CfgRowOffsetFirstEx + 1, CfgColStart + CfgColOffsetFirstEx + i).Value
            currTsk = wsSelEx.Cells(CfgRowStart + CfgRowOffsetFirstEx, CfgColStart + CfgColOffsetFirstEx + i).Value

            If currSht = Section And currTsk = Number Then
                If wsSelEx.Cells(pupilRow, CfgColStart + CfgColOffsetFirstEx + i).Value = "x" Then
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



