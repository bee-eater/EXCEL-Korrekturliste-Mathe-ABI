Attribute VB_Name = "M0_ExportImport"
Option Explicit

Private Const WbNameConfig = "Config"
Private Const WbNameGradeKey = "Notenspiegel"


' Sourcecode exportieren für Versionsverwaltung
Public Function ExportSourceFiles()
    Dim destPath As String
    destPath = Application.ActiveWorkbook.Path & "\"
    Dim component As VBComponent
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If (component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule) And component.Name <> "JsonConverter" Then
            component.Export destPath & component.Name & ToFileExtension(component.Type)
        ElseIf component.Type = vbext_ct_Document Then
            ' Export ThisWorkbook and the Config sheet code module
            If component.Name = "DieseArbeitsmappe" Then
                component.Export destPath & component.Name & ".bas"
            ElseIf component.Name = WbNameConfig Or component.Name = WbNameGradeKey Then
                component.Export destPath & "Sht_" & component.Name & ".bas"
            End If
        End If
    Next

End Function


' Sourcecode importieren — liest alle .bas/.cls/.frm Dateien aus dem Workbook-Verzeichnis
' und ersetzt die jeweiligen VBComponents. Document-Module (Sheet-Code, DieseArbeitsmappe)
' werden per CodeModule.AddFromFile aktualisiert, da sie nicht neu erstellt werden können.
Public Function ImportSourceFiles()

    If Not EnsureVBAccess() Then Exit Function

    Dim srcPath As String
    srcPath = Application.ActiveWorkbook.Path & "\"

    Dim proj As VBProject
    Set proj = Application.VBE.ActiveVBProject

    ' ---- 1. Standard- und Klassenmodule: entfernen und neu importieren --------
    Dim fileNames(0 To 99) As String
    Dim fileCount As Integer
    fileCount = 0

    ' Collect all importable files in the workbook folder
    Dim f As String
    Dim ext As Variant
    For Each ext In Array("*.bas", "*.cls", "*.frm")
        f = Dir(srcPath & ext)
        Do While f <> ""
            fileNames(fileCount) = f
            fileCount = fileCount + 1
            f = Dir()
        Loop
    Next ext

    Dim i As Integer
    For i = 0 To fileCount - 1
        Dim fileName As String
        fileName = fileNames(i)
        Dim baseName As String
        baseName = left(fileName, InStrRev(fileName, ".") - 1)
        Dim fullPath As String
        fullPath = srcPath & fileName

        ' Skip the export/import module itself — importing it would overwrite the
        ' currently running code mid-execution, causing unpredictable behaviour.
        If baseName = "M0_ExportImport" Then GoTo NextFile

        ' Determine the logical component name:
        ' Files exported as "Sht_<SheetName>.bas" belong to Document modules.
        Dim isShtFile As Boolean
        isShtFile = (left(baseName, 4) = "Sht_")
        Dim isDieseArbeitsmappe As Boolean
        isDieseArbeitsmappe = (baseName = "DieseArbeitsmappe")

        If isShtFile Or isDieseArbeitsmappe Then
            ' --- Document module: update code in-place ---
            ' AddFromFile inserts the full file verbatim, including the "VERSION 1.0 CLASS"
            ' / BEGIN / End / Attribute preamble that the VBE writes on export.
            ' We strip that preamble by reading the file ourselves and inserting only
            ' the lines that come after the last Attribute header line.
            Dim docComponentName As String
            If isShtFile Then
                docComponentName = Mid(baseName, 5)   ' strip "Sht_" prefix
            Else
                docComponentName = baseName           ' "DieseArbeitsmappe"
            End If

            Dim docComp As VBComponent
            On Error Resume Next
            Set docComp = proj.VBComponents(docComponentName)
            On Error GoTo 0
            If Not docComp Is Nothing Then
                ' Read every line of the exported file
                Dim fNum As Integer
                fNum = FreeFile()
                Dim allLines() As String
                Dim lineCount As Long
                lineCount = 0
                ReDim allLines(0 To 4999)
                Open fullPath For Input As #fNum
                Dim oneLine As String
                Do While Not EOF(fNum)
                    Line Input #fNum, oneLine
                    allLines(lineCount) = oneLine
                    lineCount = lineCount + 1
                Loop
                Close #fNum

                ' Skip all preamble lines (VERSION, BEGIN/END block, Attribute lines,
                ' blank lines) and stop at the first Private / Public / Option keyword
                ' or a comment line — whichever comes first.
                Dim firstCodeLine As Long
                firstCodeLine = 0
                Dim ln As Long
                For ln = 0 To lineCount - 1
                    Dim trimmed As String
                    trimmed = Trim(allLines(ln))
                    Dim firstWord As String
                    firstWord = UCase(Split(trimmed & " ", " ")(0))
                    If firstWord = "PRIVATE" Or firstWord = "PUBLIC" Or _
                       firstWord = "OPTION" Or left(trimmed, 1) = "'" Then
                        Exit For   ' found the first real line — stop
                    End If
                    firstCodeLine = ln + 1  ' still preamble — advance watermark
                Next ln

                ' Build the code string from the remaining lines.
                ' Join with vbNewLine between lines only — no trailing newline —
                ' so InsertLines does not add a growing blank line at the end
                ' each time the file is imported.
                Dim codeStr As String
                codeStr = ""
                Dim lastLine As Long
                lastLine = lineCount - 1
                ' Trim any trailing blank lines from the file so the module
                ' does not accumulate an extra empty line per import cycle.
                Do While lastLine >= firstCodeLine And Trim(allLines(lastLine)) = ""
                    lastLine = lastLine - 1
                Loop
                For ln = firstCodeLine To lastLine
                    If ln = firstCodeLine Then
                        codeStr = allLines(ln)
                    Else
                        codeStr = codeStr & vbNewLine & allLines(ln)
                    End If
                Next ln

                With docComp.codeModule
                    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                    If Len(codeStr) > 0 Then
                        .InsertLines 1, codeStr
                    End If
                End With
            End If

        Else
            ' --- Standard / Class / Form module: remove old, import fresh ---------
            Dim existingComp As VBComponent
            On Error Resume Next
            Set existingComp = proj.VBComponents(baseName)
            On Error GoTo 0
            If Not existingComp Is Nothing Then
                proj.VBComponents.Remove existingComp
            End If
            proj.VBComponents.Import fullPath
        End If
NextFile:
    Next i

    MsgBox fileCount & " Datei(en) erfolgreich importiert.", vbInformation, "Import abgeschlossen"

End Function

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    
    ' Dateiendung entsprechend Typ zurückgeben
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select

End Function

Public Function EnsureVBAccess()
    
    If Not IsVBProjectAccessible() Then
        MsgBox "Bitte aktiviere 'Zugriff auf das VBA-Projektmodell vertrauen':" & vbCrLf & _
               "Datei > Optionen > Trust Center > Einstellungen für das Trust Center > Makroeinstellungen", _
               vbExclamation, "Access Required"
        EnsureVBAccess = False
        Exit Function
    End If
    EnsureVBAccess = True
    
End Function

Public Function IsVBProjectAccessible() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject.VBComponents
    IsVBProjectAccessible = (Err.Number = 0)
    On Error GoTo 0
End Function
