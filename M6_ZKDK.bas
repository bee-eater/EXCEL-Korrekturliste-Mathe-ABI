Attribute VB_Name = "M6_ZKDK"
Option Explicit

' Cache for BuildCrossedMatrix — keyed by sheet name so NavigateAfterEntry
' (called on every keypress) avoids re-reading the SelEx config sheet each time.
' Invalidated automatically on sheet name change; call InvalidateCrossedMatrixCache
' from Workbook_SheetChange when Sh.Name = WbNameSelExConfig.
Private m_crossedSheetName As String
Private m_crossedMatrix() As Boolean

'-----------------------------------------------------
' ZK / DK STRIDE HELPERS
'-----------------------------------------------------

' Returns how many physical rows each pupil occupies (1 = no ZK/DK, 2 = ZK, 3 = ZK+DK).
' Reads Config — use only when you need the *intended* stride (e.g. Add/Remove logic).
Public Function PupilStride() As Integer
    Dim hasZK As Boolean, hasDK As Boolean
    hasZK = Len(Trim(Worksheets(WbNameConfig).Range(CfgZK).Value)) > 0
    hasDK = hasZK And Len(Trim(Worksheets(WbNameConfig).Range(CfgDK).Value)) > 0
    PupilStride = 1 + IIf(hasZK, 1, 0) + IIf(hasDK, 1, 0)
End Function

' Returns the *actual* stride on a given sheet by checking whether ZK/DK rows
' have been physically inserted. Always use this instead of PupilStride() whenever
' you are operating on sheet data — ZK/DK may be configured but not yet added.
Public Function SheetStride(ws As Worksheet) As Integer
    Dim firstPupilRow As Long
    firstPupilRow = CfgRowStart + CfgRowOffsetFirstPupil
    ' Check just the row(s) immediately below the first pupil row.
    Dim v1 As String, v2 As String
    v1 = Trim(ws.Cells(firstPupilRow + 1, CfgColStart + 1).Value)
    v2 = Trim(ws.Cells(firstPupilRow + 2, CfgColStart + 1).Value)
    If v1 = "ZK" And v2 = "DK" Then
        SheetStride = 3
    ElseIf v1 = "ZK" Then
        SheetStride = 2
    Else
        SheetStride = 1
    End If
End Function

' Returns the physical sheet row for pupil i (0-based).
' Pass strideVal when calling inside a loop to avoid repeated sheet reads.
Public Function PhysicalPupilRow(pupilIdx As Integer, Optional strideVal As Integer = 0) As Long
    If strideVal = 0 Then strideVal = 1  ' safe default: no ZK/DK assumed
    PhysicalPupilRow = CfgRowStart + CfgRowOffsetFirstPupil + pupilIdx * strideVal
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
            ' Soften the main pupil row's bottom border so it does not produce a thick
            ' divider between the pupil row and the ZK row directly beneath it.
            Dim softClrA As Long
            softClrA = IIf(rowClr = gClrTheme2, gClrTheme2a, gClrTheme2)
            With ws.Range(ws.Cells(pupilRow, CfgColStart), ws.Cells(pupilRow, CfgColStart + span)) _
                     .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .color = softClrA
            End With
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

SkipInsert:
    ' Always re-apply outer border around the full block (pupil + ZK/DK rows).
    ' Runs even when GoTo SkipInsert was taken (e.g. DK-only insertion) so the
    ' border is never left incomplete.
    Dim stride As Integer
    stride = PupilStride()
    Dim totalExtraRows As Integer
    totalExtraRows = gNumOfPupils * (stride - 1)
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

    ' Always define/redefine sheet-scoped PupilBlock named range
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
    ' Collects all ZK/DK rows into a union then deletes them in one operation
    ' (avoids per-row shifting and is significantly faster).

    Dim lastRow As Long
    lastRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3

    ' Unhide all rows in the scan range first — hidden ZK/DK rows must be visible
    ' before deletion otherwise Excel skips them silently.
    ws.Range(ws.Rows(CfgRowStart + CfgRowOffsetFirstPupil), ws.Rows(lastRow)).Hidden = False

    Dim deleteRng As Range
    Dim r As Long
    For r = CfgRowStart + CfgRowOffsetFirstPupil To lastRow
        Dim lbl As String
        lbl = ws.Cells(r, CfgColStart + 1).Value
        If lbl = "ZK" Or lbl = "DK" Then
            If deleteRng Is Nothing Then
                Set deleteRng = ws.Rows(r)
            Else
                Set deleteRng = Union(deleteRng, ws.Rows(r))
            End If
        End If
    Next r

    If Not deleteRng Is Nothing Then
        deleteRng.Delete Shift:=xlUp
    End If

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

    ' show all to update borders and avoid hidden empty rows if config is later changed to add ZK/DK again
    Call ShowAll

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
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If WSExists(actSheetName) Then
            Set ws = Worksheets(actSheetName)
            numOfSubEx = GetNumOfSubEx(actSheetName)
            ws.Unprotect Password:=WbPw
            lastRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * 3
            Dim firstPupilRow As Long
            firstPupilRow = CfgRowStart + CfgRowOffsetFirstPupil

            ' Pre-build crossed matrix once per sheet so the unlock pass below
            ' knows which cells must stay locked (crossed-out SelEx columns).
            Dim sheetIsSelEx As Boolean
            sheetIsSelEx = IsSelEx(actSheetName)
            Dim visCrossedMatrix() As Boolean
            If sheetIsSelEx Then visCrossedMatrix = BuildCrossedMatrix(ws, numOfSubEx)

            ' Track main pupil row count (= 0-based pupil index) separately.
            Dim mainRowIdx As Integer
            mainRowIdx = 0

            For r = firstPupilRow To lastRow
                lbl = ws.Cells(r, CfgColStart + 1).Value
                If lbl = "ZK" Then
                    ws.Rows(r).Hidden = hideZK
                ElseIf lbl = "DK" Then
                    ws.Rows(r).Hidden = hideDK
                ElseIf lbl <> "" Then
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
                        If sheetIsSelEx Then
                            ' Unlock only non-crossed cells; keep crossed cells locked
                            ' so Tab / EnableSelection=xlUnlockedCells skips them.
                            ' visCrossedMatrix(p,c): True = available, False = crossed.
                            Dim vc As Integer
                            For vc = 0 To numOfSubEx - 1
                                ' True = available ? Locked = False; False = crossed ? Locked = True
                                ws.Cells(r, CfgColStart + CfgColOffsetFirstEx + vc).Locked = _
                                    Not visCrossedMatrix(mainRowIdx, vc)
                            Next vc
                        Else
                            rngPts.Locked = False
                        End If
                    End If
                    mainRowIdx = mainRowIdx + 1
                End If
            Next r
            ' Always reapply the medium border on the percentage row so that row/column
            ' hiding can never leave the bottom of the pupil block without a thick edge.
            Dim stride As Integer
            stride = SheetStride(ws)
            Dim pctRow As Long
            pctRow = CfgRowStart + CfgRowOffsetFirstPupil + gNumOfPupils * stride
            Dim span As Integer
            span = numOfSubEx + 2
            With ws.Range(ws.Cells(pctRow, CfgColStart), ws.Cells(pctRow, CfgColStart + span))
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = 1
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = 1
                End With
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = 1
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = 1
                End With
            End With
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
        chkName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, chkSheet * 2).Value
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
                dkName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, dkSheet * 2).Value
                If dkName = "" Then Exit For
                If WSExists(dkName) Then
                    Dim dkWs As Worksheet
                    Set dkWs = Worksheets(dkName)
                    Dim dkNumEx As Integer
                    dkNumEx = GetNumOfSubEx(dkName)
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
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If WSExists(actSheetName) Then
            Set ws = Worksheets(actSheetName)
            numOfSubEx = GetNumOfSubEx(actSheetName)
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

    Worksheets(WbNameConfig).Activate

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
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If WSExists(actSheetName) Then
            Set ws = Worksheets(actSheetName)
            numOfSubEx = GetNumOfSubEx(actSheetName)
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

    Worksheets(WbNameConfig).Activate

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
    Dim lbl As String
    Dim c As Integer
    Dim importCount As Long
    Dim skipCount As Long
    importCount = 0
    skipCount = 0

    For actSheet = 0 To CfgMaxSheets
        actSheetName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, actSheet * 2).Value
        If actSheetName = "" Then Exit For
        If WSExists(actSheetName) Then
            Set ws = Worksheets(actSheetName)
            Set srcWs = srcWb.Worksheets(actSheetName)
            numOfSubEx = GetNumOfSubEx(actSheetName)
            ws.Unprotect Password:=WbPw

            Dim firstRow As Long
            firstRow = CfgRowStart + CfgRowOffsetFirstPupil
            lastRow = firstRow + gNumOfPupils * 3
            Dim blockRows As Long
            blockRows = lastRow - firstRow + 1

            ' Bulk-read label column, source block and destination block into arrays.
            ' All per-cell interaction happens in memory; only one write per sheet.
            Dim dstLabels As Variant
            dstLabels = ws.Range( _
                ws.Cells(firstRow, CfgColStart + 1), _
                ws.Cells(lastRow, CfgColStart + 1)).Value

            Dim srcBlock As Variant
            srcBlock = srcWs.Range( _
                srcWs.Cells(firstRow, CfgColStart + CfgColOffsetFirstEx), _
                srcWs.Cells(lastRow, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)).Value

            Dim dstBlock As Variant
            dstBlock = ws.Range( _
                ws.Cells(firstRow, CfgColStart + CfgColOffsetFirstEx), _
                ws.Cells(lastRow, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)).Value

            Dim rowIdx As Long
            For rowIdx = 1 To blockRows
                lbl = CStr(dstLabels(rowIdx, 1))
                If lbl = targetLabel Then
                    For c = 1 To numOfSubEx
                        Dim srcVal As Variant
                        srcVal = srcBlock(rowIdx, c)
                        If IsEmpty(srcVal) Or srcVal = "" Then
                            ' empty source cell — leave destination untouched
                            skipCount = skipCount + 1
                        ElseIf IsNumeric(srcVal) Then
                            dstBlock(rowIdx, c) = CDbl(srcVal)
                            importCount = importCount + 1
                        Else
                            skipCount = skipCount + 1
                        End If
                    Next c
                End If
            Next rowIdx

            ' Write the entire exercise block back in a single range assignment
            ws.Range( _
                ws.Cells(firstRow, CfgColStart + CfgColOffsetFirstEx), _
                ws.Cells(lastRow, CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)).Value = dstBlock

            If DevMode <> 1 Then
                ws.Protect Password:=WbPw
                ws.EnableSelection = xlUnlockedCells
            End If
        End If
    Next actSheet

    Application.Calculation = xlCalculationAutomatic
    If Not alreadyOpen Then srcWb.Close SaveChanges:=False

    ' 5. Refresh mismatch highlighting after import
    If importCount > 0 Then Call UpdateZKDKMismatchHighlight

    ' 6. Summary
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
        actSheetName = ThisWorkbook.Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, actSheet * 2).Value
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

'-----------------------------------------------------
' ZK / DK MISMATCH HIGHLIGHTING
'-----------------------------------------------------

' Highlights ZK/DK point cells that contain a value differing from the main pupil row.
' Light-red (gClrPlus1) = mismatch; white = matches or empty.
' Safe to call standalone (e.g. from a button) or after import.
Public Sub UpdateZKDKMismatchHighlight(Optional targetWs As Worksheet = Nothing)

    Call Init

    ' Preserve active cell so Tab/Enter navigation is not disrupted
    Dim savedSheet As Worksheet
    Dim savedCell As Range
    Set savedSheet = ActiveSheet
    Set savedCell = ActiveCell

    Application.ScreenUpdating = False
    Application.EnableEvents = False   ' prevent re-entrant SheetChange while writing colors

    Dim errNum As Long
    Dim errDesc As String
    errNum = 0
    On Error GoTo Cleanup

    If Not targetWs Is Nothing Then
        Call ProcessMismatchSheet(targetWs, SheetStride(targetWs))
    Else
        Dim si As Integer
        For si = 0 To CfgMaxSheets
            Dim sName As String
            sName = Worksheets(WbNameConfig).Range(CfgFirstSect).offset(0, si * 2).Value
            If sName = "" Then Exit For
            If WSExists(sName) Then
                Dim siWs As Worksheet
                Set siWs = Worksheets(sName)
                Call ProcessMismatchSheet(siWs, SheetStride(siWs))
            End If
        Next si
    End If

Cleanup:
    errNum = Err.Number
    errDesc = Err.Description
    Application.EnableEvents = True
    ' Restore original sheet/cell before re-enabling screen updates
    ' so NavigateAfterEntry (called after us) can take over without a visible flash
    On Error Resume Next
    savedSheet.Activate
    savedCell.Select
    On Error GoTo 0
    Application.ScreenUpdating = True
    If errNum <> 0 Then
        MsgBox "Fehler in UpdateZKDKMismatchHighlight:" & vbNewLine & _
               "(" & errNum & ") " & errDesc, vbCritical, "Fehler"
    End If

End Sub

Private Sub ProcessMismatchSheet(ws As Worksheet, stride As Integer)
    Dim numOfSubEx As Integer
    numOfSubEx = GetNumOfSubEx(ws.Name)
    If numOfSubEx = 0 Or stride = 1 Then Exit Sub

    ws.Unprotect Password:=WbPw

    ' Build crossed matrix via shared helper (True = column crossed out for that pupil)
    Dim crossedMatrix() As Boolean
    crossedMatrix = BuildCrossedMatrix(ws, numOfSubEx)

    ' Bulk-read the entire pupil+ZK/DK block into a Variant array.
    ' Columns: starting at CfgColStart+1 (label) through last exercise column.
    ' In the array: label = col 1; exercise col c (0-based) = col (CfgColOffsetFirstEx + c).
    Dim firstPupilRow As Long
    firstPupilRow = CfgRowStart + CfgRowOffsetFirstPupil
    Dim dataBlock As Variant
    dataBlock = ws.Range( _
        ws.Cells(firstPupilRow, CfgColStart + 1), _
        ws.Cells(firstPupilRow + gNumOfPupils * stride - 1, _
                 CfgColStart + CfgColOffsetFirstEx + numOfSubEx - 1)).Value

    Dim p As Integer
    For p = 0 To gNumOfPupils - 1
        Dim mainBlockRow As Long
        mainBlockRow = p * stride + 1              ' 1-based row in dataBlock
        Dim mainSheetRow As Long
        mainSheetRow = firstPupilRow + p * stride  ' absolute sheet row for color writes

        Dim c As Integer
        For c = 0 To numOfSubEx - 1
            Dim blockCol As Long
            blockCol = CfgColOffsetFirstEx + c     ' 1-based col in dataBlock
            Dim mainVal As Variant
            mainVal = dataBlock(mainBlockRow, blockCol)

            Dim offset As Integer
            For offset = 1 To stride - 1
                Dim subBlockRow As Long
                subBlockRow = mainBlockRow + offset
                Dim subLbl As String
                subLbl = CStr(dataBlock(subBlockRow, 1))
                If (subLbl = "ZK" Or subLbl = "DK") And crossedMatrix(p, c) Then
                    Dim subVal As Variant
                    subVal = dataBlock(subBlockRow, blockCol)
                    Dim newClr As Long
                    If IsNumeric(subVal) And subVal <> "" Then
                        If IsNumeric(mainVal) And mainVal <> "" Then
                            If CDbl(subVal) > CDbl(mainVal) Then
                                newClr = gClrZKDKDiffGt
                            ElseIf CDbl(subVal) < CDbl(mainVal) Then
                                newClr = gClrZKDKDiffLt
                            Else
                                newClr = RGB(255, 255, 255)
                            End If
                        Else
                            newClr = RGB(255, 255, 255)
                        End If
                    Else
                        newClr = RGB(255, 255, 255)
                    End If
                    ws.Cells(mainSheetRow + offset, CfgColStart + CfgColOffsetFirstEx + c).Interior.color = newClr
                End If
            Next offset
        Next c
    Next p

    If DevMode <> 1 Then
        ws.Protect Password:=WbPw
        ws.EnableSelection = xlUnlockedCells
    End If

End Sub

'-----------------------------------------------------
' POST-ENTRY NAVIGATION
'-----------------------------------------------------

' Moves selection to the next logical cell after a value is entered in an exercise column.
' Call from Workbook_SheetChange AFTER UpdateZKDKMismatchHighlight:
'   Call UpdateZKDKMismatchHighlight(Sh)
'   Call NavigateAfterEntry(Sh, Target)
'
' Behaviour:
'   - Within a row: moves one column to the right.
'   - At the last exercise column: wraps to the first exercise column of the next row
'     that belongs to the same corrector type (ZK -> next ZK row, DK -> next DK row,
'     main pupil row -> next main pupil row).  Wraps around to the first row of that
'     type when the last pupil is reached.
Public Sub NavigateAfterEntry(ws As Worksheet, changedCell As Range)
    ' Only handle single-cell changes in known segment sheets
    If changedCell.Cells.Count > 1 Then Exit Sub

    Dim cellIsEmpty As Boolean
    cellIsEmpty = IsEmpty(changedCell.Value) Or changedCell.Value = ""

    ' Check the per-option config flags (1 = enabled, anything else = disabled)
    If cellIsEmpty Then
        If Worksheets(WbNameConfig).Range(CfgOptNavAfterDel).Value <> 1 Then Exit Sub
    Else
        If Worksheets(WbNameConfig).Range(CfgOptNavAfterIns).Value <> 1 Then Exit Sub
    End If

    Dim numOfSubEx As Integer
    numOfSubEx = GetNumOfSubEx(ws.Name)
    If numOfSubEx = 0 Then Exit Sub

    Dim firstCol As Long
    firstCol = CfgColStart + CfgColOffsetFirstEx
    Dim lastCol As Long
    lastCol = firstCol + numOfSubEx - 1

    ' Only act on cells inside the exercise column range
    Dim changedCol As Long
    changedCol = changedCell.Column
    If changedCol < firstCol Or changedCol > lastCol Then Exit Sub

    ' Classify the changed row (ZK / DK / main pupil row)
    Dim rowType As String
    rowType = GetRowType(ws, changedCell.row)
    If rowType = "" Then Exit Sub

    ' Determine 0-based pupil index from the changed row
    Dim stride As Integer
    stride = SheetStride(ws)
    Dim firstPupilRow As Long
    firstPupilRow = CfgRowStart + CfgRowOffsetFirstPupil
    Dim pupilIdx As Integer
    pupilIdx = CInt((changedCell.row - firstPupilRow) \ stride)

    ' Build crossed matrix once for this sheet
    Dim crossedMatrix() As Boolean
    crossedMatrix = BuildCrossedMatrix(ws, numOfSubEx)

    Dim nextRow As Long
    Dim nextCol As Long
    nextRow = 0
    nextCol = 0

    ' Try to find the next non-crossed column to the right in the same row
    Dim c As Integer
    For c = (changedCol - firstCol + 1) To numOfSubEx - 1
        If crossedMatrix(pupilIdx, c) Then
            nextRow = changedCell.row
            nextCol = firstCol + c
            Exit For
        End If
    Next c

    ' No available column to the right — wrap to the next same-type row
    If nextRow = 0 Then
        Dim maxRow As Long
        maxRow = firstPupilRow + gNumOfPupils * stride - 1
        Dim scanRow As Long
        Dim scanPupil As Integer
        ' Scan forward
        For scanRow = changedCell.row + 1 To maxRow
            If GetRowType(ws, scanRow) = rowType Then
                scanPupil = CInt((scanRow - firstPupilRow) \ stride)
                nextCol = FirstAvailColForPupil(crossedMatrix, scanPupil, numOfSubEx, firstCol)
                If nextCol > 0 Then
                    nextRow = scanRow
                    Exit For
                End If
            End If
        Next scanRow
        ' Wrap around from the first pupil row
        If nextRow = 0 Then
            For scanRow = firstPupilRow To changedCell.row - 1
                If GetRowType(ws, scanRow) = rowType Then
                    scanPupil = CInt((scanRow - firstPupilRow) \ stride)
                    nextCol = FirstAvailColForPupil(crossedMatrix, scanPupil, numOfSubEx, firstCol)
                    If nextCol > 0 Then
                        nextRow = scanRow
                        Exit For
                    End If
                End If
            Next scanRow
        End If
    End If

    If nextRow = 0 Or nextCol = 0 Then Exit Sub

    On Error Resume Next
    ws.Cells(nextRow, nextCol).Select
    On Error GoTo 0
End Sub

' Returns "ZK", "DK", "MAIN" or "" for the corrector-type label in the name column of rowNum.
Private Function GetRowType(ws As Worksheet, rowNum As Long) As String
    Dim lbl As String
    lbl = ws.Cells(rowNum, CfgColStart + 1).Value
    If lbl = "ZK" Then
        GetRowType = "ZK"
    ElseIf lbl = "DK" Then
        GetRowType = "DK"
    ElseIf lbl <> "" Then
        GetRowType = "MAIN"
    Else
        GetRowType = ""
    End If
End Function

' Returns the absolute column number of the first non-crossed exercise column for pupilIdx,
' or -1 if every column is crossed for that pupil.
Private Function FirstAvailColForPupil(crossedMatrix() As Boolean, pupilIdx As Integer, _
                                       numOfSubEx As Integer, firstCol As Long) As Long
    Dim c As Integer
    For c = 0 To numOfSubEx - 1
        If crossedMatrix(pupilIdx, c) Then
            FirstAvailColForPupil = firstCol + c
            Exit Function
        End If
    Next c
    FirstAvailColForPupil = -1
End Function

' Call from Workbook_SheetChange when Sh.Name = WbNameSelExConfig to force a
' fresh matrix build next time a segment sheet is navigated.
Public Sub InvalidateCrossedMatrixCache()
    m_crossedSheetName = ""
End Sub

' Builds and caches a 2D Boolean matrix: True = sub-exercise column c is NOT crossed
' (i.e. available) for pupil p.  Result is keyed by ws.Name; repeated calls for the
' same sheet return the cached copy instantly (no SelEx sheet reads).
Public Function BuildCrossedMatrix(ws As Worksheet, numOfSubEx As Integer) As Boolean()
    ' Return cached result if still valid for this sheet
    If ws.Name = m_crossedSheetName Then
        BuildCrossedMatrix = m_crossedMatrix
        Exit Function
    End If

    Dim matrix() As Boolean
    ReDim matrix(0 To gNumOfPupils - 1, 0 To numOfSubEx - 1)

    ' Default all columns to available (True); only SelEx "x" entries override to False.
    ' This ensures non-SelEx exercise columns within a SelEx sheet are never locked.
    Dim initP As Integer, initC As Integer
    For initP = 0 To gNumOfPupils - 1
        For initC = 0 To numOfSubEx - 1
            matrix(initP, initC) = True
        Next initC
    Next initP

    If WSExists(WbNameSelExConfig) Then
        Dim wsCfg As Worksheet
        Set wsCfg = Worksheets(WbNameSelExConfig)

        ' Bulk-read the SelEx config block in a single call:
        '   Row 1 = task names, Row 2 = sheet names, Rows 3..2+gNumOfPupils = pupil marks.
        Dim maxCols As Integer
        maxCols = (CfgMaxSheets + 1) * CfgMaxExercisesPerSection
        Dim cfgFirstDataRow As Long
        cfgFirstDataRow = CfgRowStart + CfgRowOffsetFirstEx
        Dim cfgFirstDataCol As Long
        cfgFirstDataCol = CfgColStart + CfgColOffsetFirstEx

        Dim cfgBlock As Variant
        cfgBlock = wsCfg.Range( _
            wsCfg.Cells(cfgFirstDataRow, cfgFirstDataCol), _
            wsCfg.Cells(cfgFirstDataRow + 1 + gNumOfPupils, _
                        cfgFirstDataCol + maxCols - 1)).Value

        ' Bulk-read the ws task-header row to match SelEx task names to ws columns
        Dim wsHeader As Variant
        wsHeader = ws.Range( _
            ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, cfgFirstDataCol), _
            ws.Cells(CfgRowStart + CfgRowOffsetFirstEx, cfgFirstDataCol + numOfSubEx - 1)).Value

        Dim ci As Integer
        For ci = 1 To maxCols
            If CStr(cfgBlock(2, ci)) = "" Then Exit For   ' row 2 = sheet-name row
            If CStr(cfgBlock(2, ci)) = ws.Name Then
                Dim cfgTsk As String
                cfgTsk = CStr(cfgBlock(1, ci))             ' row 1 = task-name row
                ' Find the matching exercise column in ws
                Dim su As Integer
                For su = 1 To numOfSubEx
                    If CStr(wsHeader(1, su)) = cfgTsk Then
                        Dim sp As Integer
                        For sp = 1 To gNumOfPupils
                            ' rows 3..2+gNumOfPupils in cfgBlock = pupil data
                            ' "x" = pupil chose this task ? available (True already set)
                            ' anything else = not chosen ? crossed (False)
                            If CStr(cfgBlock(2 + sp, ci)) <> "x" Then
                                matrix(sp - 1, su - 1) = False
                            End If
                        Next sp
                        Exit For
                    End If
                Next su
            End If
        Next ci
    End If

    ' Store in cache
    m_crossedSheetName = ws.Name
    m_crossedMatrix = matrix
    BuildCrossedMatrix = matrix
End Function

