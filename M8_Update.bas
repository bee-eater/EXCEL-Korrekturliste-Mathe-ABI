Attribute VB_Name = "M8_Update"
Option Explicit

'=======================================================================
' UPDATE MECHANISM
'
' The two entry points have OPPOSITE starting points:
'
'  UpdateFromDownload  -- run from the OLD workbook
'    -> fetches GitHub latest release, downloads new wb, opens it,
'      then hands off via Application.Run to new wb's RunUpdate.
'
'  UpdateFromFile  -- run from the NEW workbook
'    -> file dialog picks the OLD workbook to migrate data from;
'      version direction is verified, then RunUpdate is called
'      directly (we are already in the new workbook's project).
'
' RunUpdate(oldFilePath, knownOldVersion)  -- runs in the NEW workbook
'   1. Open (or locate) the old workbook
'   2. Version warning if very old
'   3. SilentClear this (new) workbook
'   4. CopyConfiguration  old -> new
'   5. CreateTables        (direct call -- same project)
'   6. CopyConfigW + SelExUpdate
'   7. MigrateZKDK         (AddAllZKDKRows direct call)
'   8. CopyScores
'   9. ApplyMigrationPatches
'  10. SaveAs next to old file; close old workbook
'=======================================================================

' Minimum version that is considered directly compatible.
' Anything older triggers an extra warning before migration.
Private Const MIN_COMPATIBLE_VERSION As String = "v2.2.0"

'-----------------------------------------------------------------------
' PUBLIC ENTRY POINT 1 — download latest release from GitHub
'-----------------------------------------------------------------------
Public Sub UpdateFromDownload()

    Dim releaseTag As String
    Dim downloadUrl As String
    If Not GetLatestReleaseInfo(releaseTag, downloadUrl) Then Exit Sub

    ' Already up to date?
    If releaseTag = Version Then
        MsgBox "Du verwendest bereits die aktuelle Version " & Version & ".", _
               vbInformation, "Kein Update verf" & Chr(252) & "gbar"
        Exit Sub
    End If

    ' Confirm
    Dim msg As String
    msg = "Update gefunden!" & vbNewLine & vbNewLine & _
          "Aktuelle Version : " & Version & vbNewLine & _
          "Neue Version     : " & releaseTag & vbNewLine & vbNewLine & _
          "Die Datei wird heruntergeladen und die Konfiguration " & Chr(252) & "bertragen." & vbNewLine & _
          "Fortfahren?"
    If MsgBox(msg, vbQuestion + vbYesNo, "Update auf " & releaseTag) <> vbYes Then Exit Sub

    ' Download to temp folder
    Dim tmpDir As String
    tmpDir = Environ("TEMP") & "\KorrekturlisteUpdate\"
    Dim newFilePath As String
    newFilePath = tmpDir & "Korrekturliste_" & releaseTag & ".xlsm"

    Application.StatusBar = "Lade neue Version herunter " & Chr(133)
    If Not DownloadFile(downloadUrl, newFilePath) Then
        Application.StatusBar = False
        MsgBox "Download fehlgeschlagen. Bitte Datei manuell von GitHub herunterladen" & vbNewLine & _
               "und dann ""Update aus Datei"" verwenden.", vbCritical, "Download-Fehler"
        Exit Sub
    End If
    Application.StatusBar = False

    ' Ensure the downloaded file is not blocked by Windows before opening it.
    If Not CheckUnblockedInteractive(newFilePath) Then
        MsgBox "Update abgebrochen. Bitte die Datei manuell entsperren und erneut versuchen.", _
               vbInformation, "Update abgebrochen"
        Exit Sub
    End If

    ' Open the new workbook and hand off ALL migration logic to it.
    ' RunUpdate runs in the new workbook's context; we pass our own full
    ' path and current version so the new code knows where to pull data from.
    Dim oldPath As String
    oldPath = ThisWorkbook.FullName
    Dim oldVer As String
    oldVer = Version

    Dim newWb As Workbook
    Set newWb = Workbooks.Open(fileName:=newFilePath, UpdateLinks:=False)

    Application.Run "'" & newWb.Name & "'!RunUpdate", oldPath, oldVer

End Sub

'-----------------------------------------------------------------------
' PUBLIC ENTRY POINT 2 — run from the NEW workbook, pick the OLD file
'
' The user downloads the new release manually, opens it, then uses this
' button to pick their existing (old) workbook as the data source.
' Direction is verified by comparing version strings before proceeding.
'-----------------------------------------------------------------------
Public Sub UpdateFromFile()

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Alte Korrekturliste ausw" & Chr(228) & "hlen (Quelldatei f" & Chr(252) & "r Migration)"
        .Filters.Clear
        .Filters.Add "Excel-Arbeitsmappe mit Makros", "*.xlsm"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show <> -1 Then Exit Sub          ' user cancelled

        Dim oldFilePath As String
        oldFilePath = .SelectedItems(1)
    End With

    ' Guard: don't pick the same file we are running from
    If LCase(oldFilePath) = LCase(ThisWorkbook.FullName) Then
        MsgBox "Bitte eine andere Datei als die aktuelle ausw" & Chr(228) & "hlen.", vbExclamation, "Gleiche Datei"
        Exit Sub
    End If

    ' Open old workbook so we can read its version
    Dim oldWb As Workbook
    Set oldWb = GetOrOpenWorkbook(oldFilePath)
    If oldWb Is Nothing Then
        MsgBox "Datei konnte nicht ge" & Chr(246) & "ffnet werden:" & vbNewLine & oldFilePath, vbCritical, "Fehler"
        Exit Sub
    End If

    ' Read version from old workbook.
    ' Older releases don't expose a GetVersion() function — fall back to InputBox.
    Dim oldVer As String
    On Error Resume Next
    oldVer = Application.Run("'" & oldWb.Name & "'!GetVersion")
    On Error GoTo 0
    If Trim(oldVer) = "" Then
        oldVer = InputBox( _
            "Die Version der gew" & Chr(228) & "hlten Datei konnte nicht automatisch ermittelt werden." & vbNewLine & _
            "(Die " & Chr(228) & "ltere Version enth" & Chr(228) & "lt keine GetVersion-Funktion.)" & vbNewLine & vbNewLine & _
            "Bitte Versionsnummer der alten Datei eingeben (z.B. v2.1.0)," & vbNewLine & _
            "oder leer lassen, um als ""unbekannt"" fortzufahren.", _
            "Version der alten Datei", "")
        If Trim(oldVer) = "" Then oldVer = ""   ' stays blank → treated as unknown below
    End If

    ' ---- Direction check -----------------------------------------------
    ' The chosen file should be OLDER than (or equal to) this workbook.
    ' Warn and confirm if the picked file appears to be newer.
    ' Skip direction check when version is unknown (empty = old wb has no GetVersion).
    If oldVer <> "" Then
        If IsVersionGreater(oldVer, Version) Then
        Dim dirMsg As String
        dirMsg = "Die gew" & Chr(228) & "hlte Datei hat Version " & oldVer & _
                 " und ist damit NEUER als diese Version (" & Version & ")." & vbNewLine & vbNewLine & _
                 "Normalerweise solltest du die " & Chr(196) & "LTERE Datei w" & Chr(228) & "hlen, " & _
                 "aus der Daten " & Chr(252) & "bernommen werden sollen." & vbNewLine & vbNewLine & _
                 "Bist du sicher, dass du die richtige Datei gew" & Chr(228) & "hlt hast?" & vbNewLine & _
                 "Trotzdem fortfahren?"
        If MsgBox(dirMsg, vbExclamation + vbYesNo, "Falsche Richtung?") <> vbYes Then
            oldWb.Close SaveChanges:=False
            Exit Sub
        End If
    ElseIf oldVer = Version Then
        Dim sameVerMsg As String
        sameVerMsg = "Die gew" & Chr(228) & "hlte Datei hat dieselbe Version (" & Version & ") wie diese Arbeitsmappe." & vbNewLine & vbNewLine & _
                     "Trotzdem fortfahren?"
        If MsgBox(sameVerMsg, vbQuestion + vbYesNo, "Gleiche Version") <> vbYes Then
            oldWb.Close SaveChanges:=False
            Exit Sub
        End If
    End If
    End If
    ' --------------------------------------------------------------------

    ' We ARE the new workbook — call RunUpdate directly (no Application.Run needed)
    RunUpdate oldFilePath, oldVer

End Sub

'-----------------------------------------------------------------------
' SHARED CORE — runs in the NEW workbook.
' oldFilePath      : full path to the old workbook to migrate data from.
' knownOldVersion  : version string passed by the caller (entry points
'                    read it before handing off); empty = read from old wb.
'-----------------------------------------------------------------------
Public Sub RunUpdate(oldFilePath As String, Optional knownOldVersion As String = "")

    Dim newWb As Workbook
    Set newWb = ThisWorkbook   ' we ARE the new workbook

    ' ---- Open (or locate already-open) old workbook --------------------
    Dim oldWb As Workbook
    Set oldWb = GetOrOpenWorkbook(oldFilePath)
    If oldWb Is Nothing Then
        MsgBox "Quelldatei konnte nicht ge" & Chr(246) & "ffnet werden:" & vbNewLine & oldFilePath, _
               vbCritical, "Fehler"
        Exit Sub
    End If

    ' ---- Backup old workbook -------------------------------------------
    ' Save a copy of the old file as <name>.bak.xlsm next to the original.
    Dim bakPath As String
    bakPath = Left(oldFilePath, Len(oldFilePath) - 5) & ".bak.xlsm"
    On Error Resume Next
    Kill bakPath          ' delete existing backup silently
    On Error GoTo 0
    oldWb.SaveCopyAs bakPath

    ' ---- Determine old version -----------------------------------------
    Dim oldVersion As String
    If knownOldVersion <> "" Then
        oldVersion = knownOldVersion
    Else
        ' Try to call GetVersion() in the old workbook (requires Public Function there).
        On Error Resume Next
        oldVersion = Application.Run("'" & oldWb.Name & "'!GetVersion")
        On Error GoTo 0
        ' Leave empty if unreadable — treated as "unknown", no false v0.0.0 warning
    End If

    ' Extra warning when migrating from a very old version.
    ' Skipped when version is unknown (empty) to avoid false alarms.
    If oldVersion <> "" And IsVersionGreater(MIN_COMPATIBLE_VERSION, oldVersion) Then
        Dim warnMsg As String
        warnMsg = "Achtung: Du aktualisierst von Version " & oldVersion & _
                  ", die " & Chr(228) & "lter als " & MIN_COMPATIBLE_VERSION & " ist." & vbNewLine & _
                  "Bitte pr" & Chr(252) & "fe nach dem Update alle Konfigurationswerte sorgf" & Chr(228) & "ltig!" & vbNewLine & vbNewLine & _
                  "Trotzdem fortfahren?"
        If MsgBox(warnMsg, vbExclamation + vbYesNo, "Alte Version erkannt") <> vbYes Then Exit Sub
    End If

    ' ---- Prepare --------------------------------------------------------
    ' Activate the new workbook NOW so that all bare Worksheets() calls
    ' inside CreateTables, SelExUpdate, AddAllZKDKRows etc. resolve to
    ' the new workbook's sheets (bare Worksheets() = ActiveWorkbook, not
    ' ThisWorkbook, so the active wb must be correct before those calls).
    newWb.Activate
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Silent clear of this (new) workbook
    Call SilentClearWorkbook(newWb)

    ' ---- Copy configuration --------------------------------------------
    Call CopyConfiguration(oldWb, newWb)

    ' ---- CreateTables --------------------------------------------------
    ' Re-activate new wb — CopyConfiguration may have switched focus.
    newWb.Activate
    Application.StatusBar = "Erstelle Tabellen in neuer Version " & Chr(133)
    CreateTables
    Application.StatusBar = False

    ' ---- Copy ConfigW + run SelExUpdate --------------------------------
    ' ConfigW must be populated before scores are written so that the
    ' SelEx selection matrix is in place when SelExUpdate recalculates.
    Application.StatusBar = "Kopiere Wahlaufgaben-Konfiguration " & Chr(133)
    Call CopyConfigW(oldWb, newWb)
    SelExUpdate skipDialog:=True
    Application.StatusBar = False

    ' ---- ZK/DK rows ----------------------------------------------------
    Call MigrateZKDK(oldWb)

    ' ---- Copy scored values --------------------------------------------
    Call CopyScores(oldWb, newWb)

    ' ---- Version-specific migration patches ----------------------------
    Call ApplyMigrationPatches(oldVersion, newWb)

    ' ---- Save next to old file, offer to close old ---------------------
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Default save path: old workbook's folder, new workbook's filename
    Dim stem As String
    stem = left(newWb.Name, Len(newWb.Name) - 5)   ' strip .xlsm

    Dim savePath As String
    savePath = oldWb.Path & "\" & stem & ".xlsm"

    newWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled

    ' ---- Refresh version field on Config sheet -------------------------
    ' CheckForUpdate hits the GitHub API and writes the result (✓ / update
    ' available / error) into CfgUpdateInfo ($J$26) of ThisWorkbook (newWb).
    Application.StatusBar = "Pr" & Chr(252) & "fe Versionsstand " & Chr(133)
    CheckForUpdate Version
    Application.StatusBar = False

    Dim finMsg As String
    finMsg = "Update abgeschlossen!" & vbNewLine & vbNewLine & _
             "Neue Version     : " & Version & vbNewLine & _
             "Neue Datei       : " & savePath & vbNewLine & vbNewLine & _
             "Alte Datei (" & oldWb.Name & ") wird geschlossen!"
    MsgBox finMsg, vbInformation, "Update abgeschlossen"
    oldWb.Close SaveChanges:=False

End Sub

' Opens a workbook by full path, or returns the already-open instance.
Private Function GetOrOpenWorkbook(filePath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.FullName = filePath Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    On Error Resume Next
    Set GetOrOpenWorkbook = Workbooks.Open(fileName:=filePath, UpdateLinks:=False)
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------
' GITHUB API — resolve latest tag + asset download URL
'-----------------------------------------------------------------------
Private Function GetLatestReleaseInfo(ByRef outTag As String, ByRef outUrl As String) As Boolean
    GetLatestReleaseInfo = False
    On Error GoTo ApiError

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://api.github.com/repos/bee-eater/EXCEL-Korrekturliste-Mathe-ABI/releases/latest", False
    http.setRequestHeader "User-Agent", "Excel-VBA/1.0"
    http.Send

    If http.Status <> 200 Then
        MsgBox "GitHub API Fehler: " & http.Status, vbCritical, "Update-Pr" & Chr(252) & "fung fehlgeschlagen"
        Exit Function
    End If

    Dim json As Object
    Set json = JsonConverter.ParseJson(http.responseText)
    outTag = json("tag_name")

    ' Find the .xlsm asset
    Dim assets As Object
    Set assets = json("assets")
    Dim asset As Object
    For Each asset In assets
        If InStr(LCase(asset("name")), ".xlsm") > 0 Then
            outUrl = asset("browser_download_url")
            GetLatestReleaseInfo = True
            Exit Function
        End If
    Next asset

    MsgBox "Kein .xlsm-Asset im neuesten Release gefunden.", vbExclamation, "Update"
    Exit Function

ApiError:
    MsgBox "Fehler bei der Update-Pr" & Chr(252) & "fung: " & Err.Description, vbCritical, "Fehler"
End Function

'-----------------------------------------------------------------------
' HTTP BINARY DOWNLOAD
' Uses WinHttp.WinHttpRequest.5.1 which follows HTTPS redirects
' automatically (GitHub asset URLs redirect to an S3/CDN host).
' Falls back to MSXML2.XMLHTTP60 if WinHttp is unavailable.
'-----------------------------------------------------------------------
Private Function DownloadFile(url As String, destPath As String) As Boolean
    DownloadFile = False
    On Error GoTo DlError

    ' Ensure destination directory exists
    Dim destDir As String
    destDir = left(destPath, InStrRev(destPath, "\"))
    If Not CreateDirRecursive(destDir) Then
        MsgBox "Zielverzeichnis konnte nicht erstellt werden:" & vbNewLine & destDir, _
               vbCritical, "Download-Fehler"
        Exit Function
    End If

    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo DlError
    If http Is Nothing Then
        ' Fallback
        Set http = CreateObject("MSXML2.XMLHTTP60")
    End If

    ' WinHttpRequest option: allow redirects (default, but explicit)
    Const WinHttpRequestOption_EnableRedirects As Long = 6
    On Error Resume Next
    http.Option(WinHttpRequestOption_EnableRedirects) = True
    On Error GoTo DlError

    http.Open "GET", url, False   ' synchronous
    http.setRequestHeader "User-Agent", "Excel-VBA/1.0"
    http.Send

    If http.Status <> 200 Then
        MsgBox "HTTP-Fehler " & http.Status & " beim Download." & vbNewLine & url, _
               vbCritical, "Download-Fehler"
        Exit Function
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 1        ' adTypeBinary
    stream.Write http.ResponseBody
    stream.SaveToFile destPath, 2   ' adSaveCreateOverWrite
    stream.Close

    DownloadFile = True
    Exit Function

DlError:
    MsgBox "Download-Fehler: " & Err.Description & vbNewLine & _
           "(URL: " & url & ")", vbCritical, "Fehler"
End Function

' Creates a directory path including all intermediate folders.
' Returns True if the path exists or was created successfully.
Private Function CreateDirRecursive(dirPath As String) As Boolean
    Dim cleanPath As String
    cleanPath = dirPath
    ' Strip trailing backslash
    If right(cleanPath, 1) = "\" Then cleanPath = left(cleanPath, Len(cleanPath) - 1)

    If Len(Dir(cleanPath, vbDirectory)) > 0 Then
        CreateDirRecursive = True
        Exit Function
    End If

    ' Ensure parent exists first
    Dim parent As String
    parent = left(cleanPath, InStrRev(cleanPath, "\"))
    If Len(parent) > 3 Then   ' not a drive root
        If Not CreateDirRecursive(parent) Then Exit Function
    End If

    On Error Resume Next
    MkDir cleanPath
    CreateDirRecursive = (Err.Number = 0)
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------
' SILENT CLEAR — deletes all sheets except Config + Notenspiegel + TestData
' Does NOT show the confirmation MsgBoxes that cmdClearAll_Click shows.
'-----------------------------------------------------------------------
Private Sub SilentClearWorkbook(wb As Workbook)
    Dim ws As Worksheet
    Dim nameConfig As String, nameGradeKey As String, nameTestData As String
    ' Read constant names from the new workbook's M1_Global if available;
    ' fall back to the values we know from our own globals.
    nameConfig = WbNameConfig
    nameGradeKey = WbNameGradeKey
    nameTestData = WbNameTestDaten

    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name <> nameConfig And _
           ws.Name <> nameGradeKey And _
           ws.Name <> nameTestData Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

'-----------------------------------------------------------------------
' COPY CONFIGURATION  (pupils, segment names, exercise counts, SelEx
' flags, date/teacher/class/title, ZK/DK names, nav options)
'-----------------------------------------------------------------------
Private Sub CopyConfiguration(oldWb As Workbook, newWb As Workbook)
    Dim oldCfg As Worksheet
    Dim newCfg As Worksheet
    Set oldCfg = oldWb.Worksheets(WbNameConfig)
    Set newCfg = newWb.Worksheets(WbNameConfig)

    newCfg.Unprotect Password:=WbPw

    ' --- Scalar config cells (layout must be stable across versions) ----
    Dim scalarRanges As Variant
    scalarRanges = Array( _
        CfgAbiTitle, CfgAbiDate, CfgAbiTeacher, CfgAbiClass, _
        CfgZK, CfgDK, CfgNumOfPupi, _
        CfgOptNavAfterIns, CfgOptNavAfterDel _
    )
    Dim rng As Variant
    For Each rng In scalarRanges
        On Error Resume Next
        newCfg.Range(rng).Value = oldCfg.Range(rng).Value
        On Error GoTo 0
    Next rng

    ' --- Pupils (index + first name + last name columns, up to 50 rows) -
    Dim pupiFirstRow As Long
    pupiFirstRow = oldCfg.Range(CfgFirstPupi).row
    Dim pupiCols As Long
    pupiCols = 3   ' index, first, last

    Dim numPupils As Integer
    numPupils = CInt(oldCfg.Range(CfgNumOfPupi).Value)
    If numPupils > 0 Then
        newCfg.Range( _
            newCfg.Cells(pupiFirstRow, oldCfg.Range(CfgFirstPupi).Column), _
            newCfg.Cells(pupiFirstRow + numPupils - 1, oldCfg.Range(CfgFirstPupi).Column + pupiCols - 1) _
        ).Value = oldCfg.Range( _
            oldCfg.Cells(pupiFirstRow, oldCfg.Range(CfgFirstPupi).Column), _
            oldCfg.Cells(pupiFirstRow + numPupils - 1, oldCfg.Range(CfgFirstPupi).Column + pupiCols - 1) _
        ).Value
    End If

    ' --- Segment configuration (name, exercise count, SelEx flag, task names)
    ' CfgFirstSect points to the base cell; each segment occupies 2 columns.
    Dim sectBaseCol As Long
    sectBaseCol = oldCfg.Range(CfgFirstSect).Column
    Dim sectBaseRow As Long
    sectBaseRow = oldCfg.Range(CfgFirstSect).row

    ' Copy entire segment config block in one shot:
    ' 25 rows × (CfgMaxSheets+1)*2 columns covers all section names,
    ' task headers, exercise counts and SelEx flags.
    Dim totalSegCols As Long
    totalSegCols = (CfgMaxSheets + 1) * 2
    On Error Resume Next
    newCfg.Range( _
        newCfg.Cells(sectBaseRow, sectBaseCol), _
        newCfg.Cells(sectBaseRow + 24, sectBaseCol + totalSegCols - 1) _
    ).Value = oldCfg.Range( _
        oldCfg.Cells(sectBaseRow, sectBaseCol), _
        oldCfg.Cells(sectBaseRow + 24, sectBaseCol + totalSegCols - 1) _
    ).Value
    On Error GoTo 0

    If DevMode <> 1 Then
        newCfg.Protect Password:=WbPw
        newCfg.EnableSelection = xlUnlockedCells
    End If
End Sub

'-----------------------------------------------------------------------
' MIGRATE ZK/DK — add rows to new workbook if old one had them
'-----------------------------------------------------------------------
Private Sub MigrateZKDK(oldWb As Workbook)
    ' Check old workbook config and physical sheets
    Dim oldCfg As Worksheet
    Set oldCfg = oldWb.Worksheets(WbNameConfig)

    Dim oldHasZKConfig As Boolean
    Dim oldHasDKConfig As Boolean
    oldHasZKConfig = Len(Trim(oldCfg.Range(CfgZK).Value)) > 0
    oldHasDKConfig = oldHasZKConfig And Len(Trim(oldCfg.Range(CfgDK).Value)) > 0

    ' Check physical existence on old sheets
    Dim oldZKPhysical As Boolean, oldDKPhysical As Boolean
    oldZKPhysical = False
    oldDKPhysical = False
    Dim si As Integer
    For si = 0 To CfgMaxSheets
        Dim sName As String
        sName = oldCfg.Range(CfgFirstSect).offset(0, si * 2).Value
        If sName = "" Then Exit For
        Dim oldWsExists As Boolean
        oldWsExists = False
        Dim ws As Worksheet
        For Each ws In oldWb.Worksheets
            If ws.Name = sName Then oldWsExists = True: Exit For
        Next ws
        If oldWsExists Then
            Dim stride As Integer
            stride = SheetStride(oldWb.Worksheets(sName))
            If stride >= 2 Then oldZKPhysical = True
            If stride >= 3 Then oldDKPhysical = True
        End If
    Next si

    ' Only add if BOTH config AND physical rows existed in old wb
    If oldHasZKConfig And oldZKPhysical Then
        Application.StatusBar = "F" & Chr(252) & "ge ZK/DK-Zeilen in neuer Version ein " & Chr(133)
        AddAllZKDKRows
        Application.StatusBar = False
    End If
End Sub

'-----------------------------------------------------------------------
' COPY SCORES — copies entered point values from all segment sheets.
' Same-stride sheets: one bulk range copy per sheet.
' Different-stride sheets: read old block into Variant array, remap
' pupil rows to new stride, write once — avoids per-row round trips.
'-----------------------------------------------------------------------
Private Sub CopyScores(oldWb As Workbook, newWb As Workbook)
    Dim oldCfg As Worksheet
    Set oldCfg = oldWb.Worksheets(WbNameConfig)

    Dim firstExCol As Long
    firstExCol = CfgColStart + CfgColOffsetFirstEx
    Dim firstDataRow As Long
    firstDataRow = CfgRowStart + CfgRowOffsetFirstPupil
    Dim numPupils As Integer
    numPupils = CInt(oldCfg.Range(CfgNumOfPupi).Value)

    Application.StatusBar = "Kopiere Punktewerte " & Chr(133)

    Dim si As Integer
    For si = 0 To CfgMaxSheets
        Dim sName As String
        sName = oldCfg.Range(CfgFirstSect).offset(0, si * 2).Value
        If sName = "" Then Exit For

        Dim oldWsOk As Boolean, newWsOk As Boolean
        oldWsOk = False: newWsOk = False
        Dim ws As Worksheet
        For Each ws In oldWb.Worksheets
            If ws.Name = sName Then oldWsOk = True: Exit For
        Next ws
        For Each ws In newWb.Worksheets
            If ws.Name = sName Then newWsOk = True: Exit For
        Next ws
        If Not oldWsOk Or Not newWsOk Then GoTo NextSheet

        Dim oldWs As Worksheet, newWs As Worksheet
        Set oldWs = oldWb.Worksheets(sName)
        Set newWs = newWb.Worksheets(sName)
        newWs.Unprotect Password:=WbPw

        Dim numEx As Integer
        numEx = GetNumOfSubExFromWb(newWb, sName)
        If numEx = 0 Then GoTo NextSheet

        Dim oldStride As Integer, newStride As Integer
        oldStride = SheetStride(oldWs)
        newStride = SheetStride(newWs)

        If oldStride = newStride Then
            ' ---- Fast path: strides identical — one bulk copy ----------
            Dim blockRows As Long
            blockRows = numPupils * oldStride
            On Error Resume Next
            newWs.Range( _
                newWs.Cells(firstDataRow, firstExCol), _
                newWs.Cells(firstDataRow + blockRows - 1, firstExCol + numEx - 1) _
            ).Value = oldWs.Range( _
                oldWs.Cells(firstDataRow, firstExCol), _
                oldWs.Cells(firstDataRow + blockRows - 1, firstExCol + numEx - 1) _
            ).Value
            On Error GoTo 0
        Else
            ' ---- Remap path: read all rows, remap stride, write once ---
            Dim srcArr As Variant
            srcArr = oldWs.Range( _
                oldWs.Cells(firstDataRow, firstExCol), _
                oldWs.Cells(firstDataRow + numPupils * oldStride - 1, firstExCol + numEx - 1) _
            ).Value

            Dim dstArr() As Variant
            ReDim dstArr(1 To numPupils * newStride, 1 To numEx)

            Dim p As Integer, c As Integer
            For p = 0 To numPupils - 1
                ' EK row (always present)
                For c = 1 To numEx
                    dstArr(p * newStride + 1, c) = srcArr(p * oldStride + 1, c)
                Next c
                ' ZK row
                If oldStride >= 2 And newStride >= 2 Then
                    For c = 1 To numEx
                        dstArr(p * newStride + 2, c) = srcArr(p * oldStride + 2, c)
                    Next c
                End If
                ' DK row
                If oldStride >= 3 And newStride >= 3 Then
                    For c = 1 To numEx
                        dstArr(p * newStride + 3, c) = srcArr(p * oldStride + 3, c)
                    Next c
                End If
            Next p

            On Error Resume Next
            newWs.Range( _
                newWs.Cells(firstDataRow, firstExCol), _
                newWs.Cells(firstDataRow + numPupils * newStride - 1, firstExCol + numEx - 1) _
            ).Value = dstArr
            On Error GoTo 0
        End If

        If DevMode <> 1 Then
            newWs.Protect Password:=WbPw
            newWs.EnableSelection = xlUnlockedCells
        End If

NextSheet:
    Next si

    Application.StatusBar = False
End Sub

'-----------------------------------------------------------------------
' COPY CONFIGW — copies the SelEx pupil-selection matrix from old -> new.
' Called right after CreateTables so SelExUpdate can recalculate correctly.
'-----------------------------------------------------------------------
Private Sub CopyConfigW(oldWb As Workbook, newWb As Workbook)
    Dim oldCWok As Boolean, newCWok As Boolean
    oldCWok = False: newCWok = False
    Dim ws As Worksheet
    For Each ws In oldWb.Worksheets
        If ws.Name = WbNameSelExConfig Then oldCWok = True: Exit For
    Next ws
    For Each ws In newWb.Worksheets
        If ws.Name = WbNameSelExConfig Then newCWok = True: Exit For
    Next ws
    If Not oldCWok Or Not newCWok Then Exit Sub

    Dim oldCfg As Worksheet
    Set oldCfg = oldWb.Worksheets(WbNameConfig)

    Dim oldCW As Worksheet, newCW As Worksheet
    Set oldCW = oldWb.Worksheets(WbNameSelExConfig)
    Set newCW = newWb.Worksheets(WbNameSelExConfig)
    newCW.Unprotect Password:=WbPw

    Dim cwFirstRow As Long
    cwFirstRow = CfgRowStart + CfgRowOffsetFirstPupil
    Dim numPupilsCW As Integer
    numPupilsCW = CInt(oldCfg.Range(CfgNumOfPupi).Value)
    Dim maxCols As Integer
    maxCols = (CfgMaxSheets + 1) * CfgMaxExercisesPerSection
    Dim cwFirstDataCol As Long
    cwFirstDataCol = CfgColStart + CfgColOffsetFirstEx

    On Error Resume Next
    newCW.Range( _
        newCW.Cells(cwFirstRow, cwFirstDataCol), _
        newCW.Cells(cwFirstRow + numPupilsCW - 1, cwFirstDataCol + maxCols - 1) _
    ).Value = oldCW.Range( _
        oldCW.Cells(cwFirstRow, cwFirstDataCol), _
        oldCW.Cells(cwFirstRow + numPupilsCW - 1, cwFirstDataCol + maxCols - 1) _
    ).Value
    On Error GoTo 0

    If DevMode <> 1 Then
        newCW.Protect Password:=WbPw
        newCW.EnableSelection = xlUnlockedCells
    End If
End Sub

' Like GetNumOfSubEx but looks up from a specific workbook's Config sheet.
Private Function GetNumOfSubExFromWb(wb As Workbook, sheetName As String) As Integer
    Dim i As Integer
    Dim cfg As Worksheet
    Set cfg = wb.Worksheets(WbNameConfig)
    For i = 0 To CfgMaxSheets
        If cfg.Range(CfgFirstSect).offset(0, i * 2).Value = sheetName Then
            GetNumOfSubExFromWb = CInt(cfg.Range(CfgExerCount).offset(0, i * 2).Value)
            Exit Function
        End If
    Next i
    GetNumOfSubExFromWb = 0
End Function

'-----------------------------------------------------------------------
' VERSION-SPECIFIC MIGRATION PATCHES
'
' Add a new Case block here whenever a release introduces a breaking
' layout change that requires extra data manipulation.
' The oldVersion string is the Version constant from the old workbook
' (e.g. "v2.1.0").  IsVersionGreater() is available from M9_Helper.
'
' Example:
'   Case "v2.1.0", "v2.1.1"
'       ' v2.2.0 moved the class cell from G25 to G26 — nothing to do,
'       ' CopyConfiguration already used the new addresses from constants.
'-----------------------------------------------------------------------
Private Sub ApplyMigrationPatches(oldVersion As String, newWb As Workbook)

    ' Normalise: strip leading "v" for numeric comparison
    Dim ov As String
    ov = oldVersion

    ' ---- Patches for versions older than v2.0.0 ------------------------
    If IsVersionGreater("v2.0.0", ov) Then
        ' Pre-v2.0.0: CfgAbiTeacher did not exist — nothing to migrate,
        ' the field is simply left empty in the new workbook.
        ' Add further pre-v2.0.0 fixups here as needed.
    End If

    ' ---- Patches for specific version ranges ---------------------------
    ' Template — uncomment and fill in when a breaking change is released:
    '
    ' If IsVersionGreater("v2.x.0", ov) And Not IsVersionGreater("v2.x.0", "v2.2.0") Then
    '     ' fixup ...
    ' End If

End Sub
