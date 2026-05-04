Attribute VB_Name = "M6_ZKDK"

Option Explicit

'-----------------------------------------------------
' ZK / DK STRIDE HELPERS
'-----------------------------------------------------

' Returns how many physical rows each pupil occupies (1 = no ZK/DK, 2 = ZK, 3 = ZK+DK).
Public Function PupilStride() As Integer
    Dim hasZK As Boolean, hasDK As Boolean
    hasZK = Len(Trim(Worksheets(WbNameConfig).Range(CfgZK).Value)) > 0
    hasDK = hasZK And Len(Trim(Worksheets(WbNameConfig).Range(CfgDK).Value)) > 0
    PupilStride = 1 + IIf(hasZK, 1, 0) + IIf(hasDK, 1, 0)
End Function

' Returns the physical sheet row for pupil i (0-based).
Public Function PhysicalPupilRow(pupilIdx As Integer) As Long
    PhysicalPupilRow = CfgRowStart + CfgRowOffsetFirstPupil + pupilIdx * PupilStride()
End Function

'-----------------------------------------------------
' ZK / DK ADDITIONAL ROWS
'-----------------------------------------------------

Public Sub AddZKDKRows(ws As Worksheet, numOfSubEx As Integer, span As Integer)
    ' Adds secondary-corrector rows (ZK, and optionally DK) below each pupil row.
    ' Reads CfgZK / CfgDK from Config.
    ' If ZK rows already exist but DK is now configured and missing, inserts DK only.
    ' Always defines the sheet-scoped "PupilBlock" named range.

    Dim zkName As String, dkName As String
    zkName = Trim(Worksheets(WbNameConfig).Range(CfgZK).Value)
    dkName = Trim(Worksheets(WbNameConfig).Range(CfgDK).Value)

    Dim hasZK As Boolean, hasDK As Boolean
    hasZK = (Len(zkName) > 0)
    hasDK = hasZK And (Len(dkName) > 0)

    ' Detect what is already present on the sheet
    Dim zkPresent As Boolean, dkPresent As Boolean
    Dim checkRow As Long
    zkPresent = False
    dkPresent = False
    For checkRow = CfgRowStart + CfgRowOffsetFirstPupil To CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3
        Dim cellVal As String
        cellVal = ws.Cells(checkRow, CfgColStart + 1).Value
        If cellVal = "ZK" Then zkPresent = True
        If cellVal = "DK" Then dkPresent = True
    Next checkRow

    ' Nothing to do if everything that config requires is already there
    If (Not hasZK) Then GoTo SkipInsert
    If zkPresent And (Not hasDK Or dkPresent) Then GoTo SkipInsert

    Dim extraRows As Integer
    extraRows = 0

    If Not zkPresent Then
        ' -- CASE A: Neither ZK nor DK present ? insert both from scratch ---------
        Dim i As Integer
        For i = gNumOfPupils - 1 To 0 Step -1
            Dim pupilRow As Long
            pupilRow = CfgRowStart + CfgRowOffsetFirstPupil + i
            Dim rowClr As Long
            If i Mod 2 = 0 Then rowClr = gClrTheme2 Else rowClr = gClrTheme2a

            If hasDK Then
                ws.Rows(pupilRow + 1).Insert Shift:=xlDown
                Call FormatZKDKRow(ws, pupilRow + 1, numOfSubEx, span, "DK", rowClr)
            End If
            ws.Rows(pupilRow + 1).Insert Shift:=xlDown
            Call FormatZKDKRow(ws, pupilRow + 1, numOfSubEx, span, "ZK", rowClr, hasDK)
        Next i
        extraRows = gNumOfPupils * (1 + IIf(hasDK, 1, 0))

    ElseIf zkPresent And hasDK And Not dkPresent Then
        ' -- CASE B: ZK present, DK missing but configured ? insert DK only -------
        ' Scan bottom-to-top for ZK rows; insert DK immediately after each one.
        Dim scanLastRow As Long
        scanLastRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 2 - 1
        Dim r As Long
        For r = scanLastRow To CfgRowStart + CfgRowOffsetFirstPupil Step -1
            If ws.Cells(r, CfgColStart + 1).Value = "ZK" Then
                ' Determine color from the pupil row directly above this ZK row
                Dim pupilIdx As Long
                pupilIdx = 0
                Dim rr As Long
                For rr = CfgRowStart + CfgRowOffsetFirstPupil To r - 1
                    If ws.Cells(rr, CfgColStart + 1).Value <> "ZK" And _
                       ws.Cells(rr, CfgColStart + 1).Value <> "DK" And _
                       ws.Cells(rr, CfgColStart + 1).Value <> "" Then
                        pupilIdx = pupilIdx + 1
                    End If
                Next rr
                Dim dkClr As Long
                If (pupilIdx - 1) Mod 2 = 0 Then dkClr = gClrTheme2 Else dkClr = gClrTheme2a
                ' Update ZK row's softBottom now that DK will follow
                Call FormatZKDKRow(ws, r, numOfSubEx, span, "ZK", dkClr, True)
                ' Insert DK below ZK
                ws.Rows(r + 1).Insert Shift:=xlDown
                Call FormatZKDKRow(ws, r + 1, numOfSubEx, span, "DK", dkClr)
            End If
        Next r
        extraRows = gNumOfPupils * 2  ' now stride = 3, total extra = 2 per pupil
    End If

    ' Re-apply outer border around the full block (pupil + ZK/DK rows)
    Dim totalExtraRows As Integer
    totalExtraRows = gNumOfPupils * (IIf(hasDK, 2, 1))
    Dim rngBlock As Range
    Set rngBlock = ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), _
                            ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils + totalExtraRows - 1, CfgColStart + span))
    With rngBlock
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
    ' Re-force sum column left border to xlMedium (overridden by xlInsideVertical above)
    With ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + span), _
                  ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils + totalExtraRows - 1, CfgColStart + span))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).ColorIndex = 1
    End With

SkipInsert:
    ' Always define/redefine sheet-scoped PupilBlock named range
    Dim stride As Integer
    stride = PupilStride()
    Call DefinePupilBlockName(ws, numOfSubEx, gNumOfPupils * stride)

End Sub

' Defines (or redefines) the sheet-scoped "PupilBlock" named range on ws.
' The range spans the index column through the sum column for all pupil+ZK/DK rows.
' Used by grade page and print page VLOOKUPs via 'SheetName'!PupilBlock.
Public Sub DefinePupilBlockName(ws As Worksheet, numOfSubEx As Integer, totalRows As Integer)
    ws.Names.Add _
        Name:="PupilBlock", _
        RefersTo:=ws.Range( _
            ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart + 1), _
            ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + totalRows - 1, CfgColStart + CfgColOffsetFirstEx + numOfSubEx))
End Sub

Public Sub FormatZKDKRow(ws As Worksheet, rowNum As Long, numOfSubEx As Integer, span As Integer, label As String, rowClr As Long, Optional softBottom As Boolean = False)
    ' Formats a single ZK/DK row: height, font, borders, label and sum formula.
    ' Vertical borders stay black; top horizontal border uses inverted alternating
    ' color to appear subtle (within-group), bottom stays black (between groups).
    ' softBottom=True also lightens the bottom edge (used for ZK when DK follows).

    Dim borderClr As Long
    If rowClr = gClrTheme2 Then
        borderClr = gClrTheme2a
    Else
        borderClr = gClrTheme2
    End If

    ws.Rows(rowNum).RowHeight = 13.2
    ws.Rows(rowNum).Font.Size = 8
    ws.Rows(rowNum).Locked = True

    ' Index + Name columns: alternating bg, subtle top border, label right-aligned
    Dim rngLabel As Range
    Set rngLabel = ws.Range(ws.Cells(rowNum, CfgColStart), ws.Cells(rowNum, CfgColStart + CfgColOffsetFirstEx - 1))
    Call ApplyBordersDirect(rngLabel, rowClr, xlThin, borderClr, softBottom)
    rngLabel.HorizontalAlignment = xlRight
    ws.Cells(rowNum, CfgColStart + 1).Value = label

    ' Points cells: white background, subtle top border, unlocked for manual entry
    Dim rngPts As Range
    Set rngPts = ws.Range(ws.Cells(rowNum, CfgColStart + CfgColOffsetFirstEx), ws.Cells(rowNum, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1))
    Call ApplyBordersDirect(rngPts, RGB(255, 255, 255), xlThin, borderClr, softBottom)
    rngPts.HorizontalAlignment = xlCenter
    rngPts.VerticalAlignment = xlCenter
    rngPts.Locked = False

    ' Sum cell: alternating bg + subtle top border + SUM formula
    Dim colFirst As String, colLast As String
    colFirst = ColLetter(CfgColStart + CfgColOffsetFirstEx)
    colLast = ColLetter(CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)
    Dim rngSum As Range
    Set rngSum = ws.Range(ws.Cells(rowNum, CfgColStart + span), ws.Cells(rowNum, CfgColStart + span))
    Call ApplyBordersDirect(rngSum, rowClr, xlMedium, borderClr, softBottom)
    rngSum.HorizontalAlignment = xlCenter
    rngSum.VerticalAlignment = xlCenter
    rngSum.Formula = "=SUM(" & colFirst & rowNum & ":" & colLast & rowNum & ")"
    rngSum.Locked = True

    ' Force outer vertical edges to xlMedium — cannot rely on post-loop block border
    With ws.Cells(rowNum, CfgColStart).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
    With ws.Cells(rowNum, CfgColStart + span).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With

End Sub

' Applies fill color and all four borders directly to rng without .Select.
Private Sub ApplyBordersDirect(rng As Range, fillClr As Long, edgeWeight As Integer, softClr As Long, softBottom As Boolean)
    rng.Interior.color = fillClr
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = edgeWeight
        .ColorIndex = 1
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = edgeWeight
        .ColorIndex = 1
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .color = softClr
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        If softBottom Then
            .Weight = xlHairline
            .color = softClr
        Else
            .Weight = edgeWeight
            .ColorIndex = 1
        End If
    End With
    If rng.Columns.Count > 1 Then
        With rng.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = edgeWeight
            .ColorIndex = 1
        End With
    End If
End Sub

Public Sub RemoveZKDKRows(ws As Worksheet, numOfSubEx As Integer)
    ' Deletes all ZK/DK extra rows identified by "ZK" or "DK" in the name column.
    ' Scans from the bottom upward to keep row indices stable during deletion.
    ' Redefines PupilBlock named range back to pupils-only after removal.

    Dim lastRow As Long
    lastRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3

    Dim r As Long
    For r = lastRow To CfgRowStart + CfgRowOffsetFirstPupil Step -1
        Dim lbl As String
        lbl = ws.Cells(r, CfgColStart + 1).Value
        If lbl = "ZK" Or lbl = "DK" Then
            ws.Rows(r).Delete Shift:=xlUp
        End If
    Next r

    ' Restore named range to pupils-only size
    Call DefinePupilBlockName(ws, numOfSubEx, gNumOfPupils)

    ' Repaint outer border of pupil block — row deletions can clear adjacent cell borders
    Dim span As Integer
    span = numOfSubEx + 2
    Dim rngOuter As Range
    Set rngOuter = ws.Range(ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil, CfgColStart), _
                            ws.Cells(CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils - 1, CfgColStart + span))
    With rngOuter
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
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).ColorIndex = 1
    End With

End Sub

'-----------------------------------------------------
' ZK / DK ROW VISIBILITY
'-----------------------------------------------------

' EK view: hide all ZK and DK rows (show only main pupil rows).
Public Sub ShowEK()
    Call SetZKDKVisibility(hideZK:=True, hideDK:=True, lockMain:=False)
End Sub

' ZK view: hide DK rows only (show main + ZK rows).
Public Sub ShowZK()
    If Not Len(Trim(Worksheets(WbNameConfig).Range(CfgZK).Value)) > 0 Then Exit Sub
    Call SetZKDKVisibility(hideZK:=False, hideDK:=True, lockMain:=True)
End Sub

' DK view: hide ZK rows only (show main + DK rows).
Public Sub ShowDK()
    Dim hasZK As Boolean, hasDK As Boolean
    hasZK = Len(Trim(Worksheets(WbNameConfig).Range(CfgZK).Value)) > 0
    hasDK = hasZK And Len(Trim(Worksheets(WbNameConfig).Range(CfgDK).Value)) > 0
    If Not hasDK Then Exit Sub
    Call SetZKDKVisibility(hideZK:=True, hideDK:=False, lockMain:=True)
End Sub

' All view: show all rows (unhide ZK and DK).
Public Sub ShowAll()
    If Not Len(Trim(Worksheets(WbNameConfig).Range(CfgZK).Value)) > 0 Then Exit Sub
    Call SetZKDKVisibility(hideZK:=False, hideDK:=False, lockMain:=False)
End Sub

Private Sub SetZKDKVisibility(hideZK As Boolean, hideDK As Boolean, lockMain As Boolean)

    Call Init

    Application.ScreenUpdating = False

    Dim actSheet As Integer
    Dim actSheetName As String
    Dim numOfSubEx As Integer
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim lbl As String

    For actSheet = 0 To CfgMaxSheets
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If WSExists(actSheetName) Then
            Set ws = Worksheets(actSheetName)
            numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, actSheet * 2).Value
            ws.Unprotect Password:=WbPw
            lastRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3
            Dim firstPupilRow As Long
            firstPupilRow = CfgRowStart + CfgRowOffsetFirstPupil
            Dim lastMainRow As Long
            lastMainRow = 0
            For r = firstPupilRow To lastRow
                lbl = ws.Cells(r, CfgColStart + 1).Value
                If lbl = "ZK" Then
                    ws.Rows(r).Hidden = hideZK
                ElseIf lbl = "DK" Then
                    ws.Rows(r).Hidden = hideDK
                ElseIf lbl <> "" Then
                    lastMainRow = r
                    ' Main pupil row: ensure own top border is always present (visible when ZK/DK above are hidden)
                    ' First pupil row keeps xlMedium (it is the block's outer top edge); inner rows use xlThin
                    Dim rngMainRow As Range
                    Set rngMainRow = ws.Range(ws.Cells(r, CfgColStart), ws.Cells(r, CfgColStart + numOfSubEx + 2))
                    With rngMainRow.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = IIf(r = firstPupilRow, xlMedium, xlThin)
                        .ColorIndex = 1
                    End With
                    ' Dim+lock or restore+unlock points cells
                    Dim rngPts As Range
                    Set rngPts = ws.Range( _
                        ws.Cells(r, CfgColStart + CfgColOffsetFirstEx), _
                        ws.Cells(r, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1))
                    If lockMain Then
                        rngPts.Font.color = gClrTheme2
                        rngPts.Locked = True
                    Else
                        rngPts.Font.ColorIndex = xlAutomatic
                        rngPts.Locked = False
                    End If
                End If
            Next r
            ' Restore xlMedium on the outer bottom edge of the last main pupil row
            ' (may have been thinned by ZK/DK row hairline borders or row deletion)
            If lastMainRow > 0 Then
                Dim rngLastRow As Range
                Set rngLastRow = ws.Range(ws.Cells(lastMainRow, CfgColStart), ws.Cells(lastMainRow, CfgColStart + numOfSubEx + 2))
                With rngLastRow.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = 1
                End With
            End If
            If DevMode <> 1 Then
                ws.Protect Password:=WbPw
                ws.EnableSelection = xlUnlockedCells
            End If
        End If
    Next actSheet

    Application.ScreenUpdating = True

End Sub

' Adds ZK/DK rows to ALL segment sheets. Safe to call standalone (e.g. from a button).
Public Sub AddAllZKDKRows()

    Call Init

    ' -- Pre-flight: compare config against what is actually on the sheets -----
    Dim cfgHasZK As Boolean, cfgHasDK As Boolean
    cfgHasZK = Len(Trim(Worksheets(WbNameConfig).Range(CfgZK).Value)) > 0
    cfgHasDK = cfgHasZK And Len(Trim(Worksheets(WbNameConfig).Range(CfgDK).Value)) > 0

    ' Scan all sheets once to find out what rows are present
    Dim sheetsHaveZK As Boolean, sheetsHaveDK As Boolean
    sheetsHaveZK = False
    sheetsHaveDK = False
    Dim chkSheet As Integer
    Dim chkName As String
    For chkSheet = 0 To CfgMaxSheets
        chkName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, chkSheet * 2).Value
        If chkName = "" Then Exit For
        If WSExists(chkName) Then
            Dim chkWs As Worksheet
            Set chkWs = Worksheets(chkName)
            Dim chkRow As Long
            For chkRow = CfgRowStart + CfgRowOffsetFirstPupil To CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3
                Dim chkVal As String
                chkVal = chkWs.Cells(chkRow, CfgColStart + 1).Value
                If chkVal = "ZK" Then sheetsHaveZK = True
                If chkVal = "DK" Then sheetsHaveDK = True
                If sheetsHaveZK And sheetsHaveDK Then Exit For
            Next chkRow
        End If
        If sheetsHaveZK And sheetsHaveDK Then Exit For
    Next chkSheet

    ' Case 1: Config has no ZK at all but ZK (and/or DK) rows exist ? offer full removal
    If Not cfgHasZK And sheetsHaveZK Then
        Dim ans1 As VbMsgBoxResult
        ans1 = MsgBox("In der Konfiguration ist kein ZK eingetragen, aber ZK/DK-Zeilen sind auf den Tabellenblättern vorhanden." & vbNewLine & _
                      "Alle ZK/DK-Zeilen jetzt entfernen?", vbQuestion + vbYesNo, "ZK/DK-Zeilen entfernen?")
        If ans1 = vbYes Then Call RemoveAllZKDKRows
        Exit Sub
    End If

    ' Case 2: Config has ZK but no DK, yet DK rows exist ? offer DK-only removal
    If cfgHasZK And Not cfgHasDK And sheetsHaveDK Then
        Dim ans2 As VbMsgBoxResult
        ans2 = MsgBox("In der Konfiguration ist kein DK eingetragen, aber DK-Zeilen sind auf den Tabellenblättern vorhanden." & vbNewLine & _
                      "DK-Zeilen jetzt entfernen?", vbQuestion + vbYesNo, "DK-Zeilen entfernen?")
        If ans2 = vbYes Then
            ' Remove only DK rows on all sheets
            Application.DisplayAlerts = False
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Dim dkSheet As Integer
            Dim dkName As String
            For dkSheet = 0 To CfgMaxSheets
                dkName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, dkSheet * 2).Value
                If dkName = "" Then Exit For
                If WSExists(dkName) Then
                    Dim dkWs As Worksheet
                    Set dkWs = Worksheets(dkName)
                    Dim dkNumEx As Integer
                    dkNumEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, dkSheet * 2).Value
                    dkWs.Unprotect Password:=WbPw
                    ' Delete DK rows bottom-up; re-format ZK rows to remove softBottom
                    Dim dkSpan As Integer
                    dkSpan = dkNumEx + 2
                    Dim dkLastRow As Long
                    dkLastRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3
                    Dim dkR As Long
                    For dkR = dkLastRow To CfgRowStart + CfgRowOffsetFirstPupil Step -1
                        Dim dkLbl As String
                        dkLbl = dkWs.Cells(dkR, CfgColStart + 1).Value
                        If dkLbl = "DK" Then
                            dkWs.Rows(dkR).Delete Shift:=xlUp
                        ElseIf dkLbl = "ZK" Then
                            ' ZK is now the last row in its group — remove soft bottom
                            Dim dkPupilIdx As Long
                            dkPupilIdx = 0
                            Dim dkRR As Long
                            For dkRR = CfgRowStart + CfgRowOffsetFirstPupil To dkR - 1
                                Dim dkRRLbl As String
                                dkRRLbl = dkWs.Cells(dkRR, CfgColStart + 1).Value
                                If dkRRLbl <> "ZK" And dkRRLbl <> "DK" And dkRRLbl <> "" Then
                                    dkPupilIdx = dkPupilIdx + 1
                                End If
                            Next dkRR
                            Dim dkRowClr As Long
                            If (dkPupilIdx - 1) Mod 2 = 0 Then dkRowClr = gClrTheme2 Else dkRowClr = gClrTheme2a
                            Call FormatZKDKRow(dkWs, dkR, dkNumEx, dkSpan, "ZK", dkRowClr, False)
                        End If
                    Next dkR
                    Call DefinePupilBlockName(dkWs, dkNumEx, gNumOfPupils * 2)
                    If DevMode <> 1 Then
                        dkWs.Protect Password:=WbPw
                        dkWs.EnableSelection = xlUnlockedCells
                    End If
                End If
            Next dkSheet
            Application.DisplayAlerts = True
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
        End If
        Exit Sub
    End If
    ' -------------------------------------------------------------------------

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim actSheet As Integer
    Dim actSheetName As String
    Dim numOfSubEx As Integer
    Dim span As Integer
    Dim ws As Worksheet

    For actSheet = 0 To CfgMaxSheets
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If WSExists(actSheetName) Then
            Set ws = Worksheets(actSheetName)
            numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, actSheet * 2).Value
            span = numOfSubEx + 2
            ws.Unprotect Password:=WbPw
            Call AddZKDKRows(ws, numOfSubEx, span)
            ws.Activate
            ActiveWindow.ScrollRow = 1
            ws.Cells(PhysicalPupilRow(0), CfgColStart + CfgColOffsetFirstEx).Select
            If DevMode <> 1 Then
                ws.Protect Password:=WbPw
                ws.EnableSelection = xlUnlockedCells
            End If
        End If
    Next actSheet

    ' Re-apply SelEx crosses to all newly added ZK/DK rows
    If CheckForSelEx() Then Call ApplySelExCrosses

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


' Removes ZK/DK rows from ALL segment sheets. Safe to call standalone (e.g. from a button).
Public Sub RemoveAllZKDKRows()

    Call Init

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Alle ZK/DK-Zeilen von sämtlichen Tabellenblättern entfernen?" & vbNewLine & _
                     "Diese Aktion kann nicht rückgängig gemacht werden.", _
                     vbQuestion + vbYesNo, "ZK/DK-Zeilen entfernen")
    If confirm <> vbYes Then Exit Sub

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim actSheet As Integer
    Dim actSheetName As String
    Dim numOfSubEx As Integer
    Dim ws As Worksheet

    For actSheet = 0 To CfgMaxSheets
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If WSExists(actSheetName) Then
            Set ws = Worksheets(actSheetName)
            numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, actSheet * 2).Value
            ws.Unprotect Password:=WbPw
            Call RemoveZKDKRows(ws, numOfSubEx)
            ws.Activate
            ActiveWindow.ScrollRow = 1
            ws.Cells(PhysicalPupilRow(0), CfgColStart + CfgColOffsetFirstEx).Select
            If DevMode <> 1 Then
                ws.Protect Password:=WbPw
                ws.EnableSelection = xlUnlockedCells
            End If
        End If
    Next actSheet

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

'-----------------------------------------------------
' ZK / DK IMPORT FROM ANOTHER FILE
'-----------------------------------------------------

' Entry point for "Import ZK" button.
Public Sub ImportZK()
    Call ImportZKDKFromFile("ZK")
End Sub

' Entry point for "Import DK" button.
Public Sub ImportDK()
    Call ImportZKDKFromFile("DK")
End Sub

Private Sub ImportZKDKFromFile(targetLabel As String)
    Call Init

    ' 1. File picker
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsm;*.xlsx;*.xls),*.xlsm;*.xlsx;*.xls", _
        Title:="Quelldatei auswählen – " & targetLabel & "-Werte importieren")
    If srcPath = False Then Exit Sub

    ' 2. Open source workbook read-only (reuse if already open)
    Dim srcWb As Workbook
    Dim alreadyOpen As Boolean
    alreadyOpen = False
    Dim owb As Workbook
    For Each owb In Workbooks
        If owb.FullName = CStr(srcPath) Then
            Set srcWb = owb
            alreadyOpen = True
            Exit For
        End If
    Next owb
    If Not alreadyOpen Then
        Set srcWb = Workbooks.Open(Filename:=CStr(srcPath), ReadOnly:=True, UpdateLinks:=False)
    End If
    ' Restore ThisWorkbook as active — Workbooks.Open shifts focus to the new file,
    ' causing all unqualified Worksheets() calls below to resolve against srcWb.
    ThisWorkbook.Activate

    ' 3. Validate — sheets and pupils must match; targetLabel rows must exist in source
    Dim errMsg As String
    errMsg = ValidateImportSource(srcWb, targetLabel)
    If errMsg <> "" Then
        MsgBox "Import abgebrochen – Unstimmigkeiten gefunden:" & vbNewLine & vbNewLine & errMsg, vbCritical, "Import " & targetLabel
        If Not alreadyOpen Then srcWb.Close SaveChanges:=False
        Exit Sub
    End If

    ' 4. Copy values
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim actSheet As Integer
    Dim actSheetName As String
    Dim numOfSubEx As Integer
    Dim ws As Worksheet
    Dim srcWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim lbl As String
    Dim c As Integer
    Dim importCount As Long
    Dim skipCount As Long
    importCount = 0
    skipCount = 0

    For actSheet = 0 To CfgMaxSheets
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If WSExists(actSheetName) Then
            Set ws = Worksheets(actSheetName)
            Set srcWs = srcWb.Worksheets(actSheetName)
            numOfSubEx = Worksheets(WbNameConfig).Range(CfgExerCount).Offset(0, actSheet * 2).Value
            ws.Unprotect Password:=WbPw
            lastRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3
            For r = CfgRowStart + CfgRowOffsetFirstPupil To lastRow
                lbl = ws.Cells(r, CfgColStart + 1).Value
                If lbl = targetLabel Then
                    For c = 0 To numOfSubEx - 1
                        Dim srcVal As Variant
                        srcVal = srcWs.Cells(r, CfgColStart + CfgColOffsetFirstEx + c).Value
                        If IsEmpty(srcVal) Or srcVal = "" Then
                            ' empty source cell ? leave destination untouched
                            skipCount = skipCount + 1
                        ElseIf IsNumeric(srcVal) Then
                            ws.Cells(r, CfgColStart + CfgColOffsetFirstEx + c).Value = CDbl(srcVal)
                            importCount = importCount + 1
                        Else
                            skipCount = skipCount + 1
                        End If
                    Next c
                End If
            Next r
            If DevMode <> 1 Then
                ws.Protect Password:=WbPw
                ws.EnableSelection = xlUnlockedCells
            End If
        End If
    Next actSheet

    Application.Calculation = xlCalculationAutomatic
    If Not alreadyOpen Then srcWb.Close SaveChanges:=False

    ' 5. Summary
    Dim summary As String
    summary = targetLabel & "-Werte erfolgreich importiert." & vbNewLine & importCount & " Zelle(n) übernommen."
    If skipCount > 0 Then
        summary = summary & vbNewLine & skipCount & " Zelle(n) übersprungen (kein numerischer Wert in Quelle)."
    End If
    MsgBox summary, vbInformation, "Import " & targetLabel

End Sub

' Returns "" if the source workbook is compatible for importing targetLabel rows,
' or a bullet-list of problems if not. Aborts early after MAX_ERRORS issues.
Private Function ValidateImportSource(srcWb As Workbook, targetLabel As String) As String
    Const MAX_ERRORS As Integer = 10
    Dim errors As String
    errors = ""

    Dim actSheet As Integer
    Dim actSheetName As String

    For actSheet = 0 To CfgMaxSheets
        actSheetName = ThisWorkbook.Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If Not WSExists(actSheetName) Then GoTo NextValidateSheet

        ' Check sheet exists in source
        Dim srcWs As Worksheet
        Set srcWs = Nothing
        On Error Resume Next
        Set srcWs = srcWb.Worksheets(actSheetName)
        On Error GoTo 0
        If srcWs Is Nothing Then
            errors = errors & Chr(149) & " Sheet '" & actSheetName & "' not found in source workbook." & vbNewLine
            GoTo NextValidateSheet
        End If

        Dim ws As Worksheet
        Set ws = Worksheets(actSheetName)
        Dim lastRow As Long
        lastRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3
        Dim foundLabel As Boolean
        foundLabel = False
        Dim pupilNum As Long
        pupilNum = 0
        Dim r As Long
        Dim lbl As String
        Dim srcLbl As String

        For r = CfgRowStart + CfgRowOffsetFirstPupil To lastRow
            lbl = ws.Cells(r, CfgColStart + 1).Value
            srcLbl = srcWs.Cells(r, CfgColStart + 1).Value
            If lbl = "ZK" Or lbl = "DK" Then
                If lbl = targetLabel Then
                    If srcLbl <> targetLabel Then
                        errors = errors & Chr(149) & " Sheet '" & actSheetName & "': row " & r & _
                                 " — expected '" & targetLabel & "' in source, found '" & srcLbl & "'." & vbNewLine
                    Else
                        foundLabel = True
                    End If
                End If
            ElseIf lbl <> "" Then
                pupilNum = pupilNum + 1
                If srcLbl = "ZK" Or srcLbl = "DK" Then
                    errors = errors & Chr(149) & " Sheet '" & actSheetName & "': pupil #" & pupilNum & _
                             " — source has '" & srcLbl & "' where target has '" & lbl & "'." & vbNewLine
                ElseIf srcLbl <> lbl Then
                    errors = errors & Chr(149) & " Sheet '" & actSheetName & "': pupil #" & pupilNum & _
                             " name mismatch (here: '" & lbl & "', source: '" & srcLbl & "')." & vbNewLine
                End If
            End If
            ' Stop early if too many errors accumulated
            Dim errLineCount As Integer
            errLineCount = (Len(errors) - Len(Replace(errors, vbNewLine, ""))) / Len(vbNewLine)
            If errLineCount >= MAX_ERRORS Then
                errors = errors & "... (further errors suppressed)"
                GoTo ValidateDone
            End If
        Next r

        If Not foundLabel Then
            errors = errors & Chr(149) & " Sheet '" & actSheetName & "': no '" & targetLabel & "' rows found in source workbook." & vbNewLine
        End If

NextValidateSheet:
    Next actSheet

ValidateDone:
    ValidateImportSource = errors
End Function





