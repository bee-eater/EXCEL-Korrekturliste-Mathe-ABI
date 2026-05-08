VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdClearAll_Click()

    ' Namen prüfen bevor es losgeht
    Call CheckSheetNames
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    makeSure = False
    ' Abfragen ob wirklich neue Tabellen erstellt werden sollen...
    Dim Request As Integer
    Request = MsgBox("Sicher dass Sie alles löschen wollen? Es werden ALLE Blätter, abgesehen von Config und Notenspiegel, gelöscht!", vbCritical + vbOKCancel, "Was passiert hier? O.O")
    If Request = vbCancel Then
        Exit Sub
    End If
    Request = MsgBox("Ganz sicher?? Es ist wirklich alles weg!", vbCritical + vbOKCancel, "(x.x)")
    If Request = vbCancel Then
        Exit Sub
    End If
    ' Not exited -> sure
    makeSure = True
    
    ' Alle anderen Blätter löschen
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> WbNameConfig And _
           ws.Name <> WbNameGradeKey And _
           ws.Name <> WbNameTestDaten Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Sub cmdCreateResults_Click()
    ' Namen prüfen bevor es losgeht
    Call CheckSheetNames
    Call CreateResults
End Sub

Private Sub cmdCreateTables_Click()
    ' Namen prüfen bevor es losgeht
    Call CheckSheetNames
    Call CreateTables
End Sub

Private Sub btnSelExUpdate_Click()
    SelExUpdate
End Sub

Private Sub cmdAddZKDK_Click()
    Call AddAllZKDKRows
End Sub

Private Sub cmdDelZKDK_Click()
    Call RemoveAllZKDKRows
End Sub

Private Sub cmdDK_Click()
    Call ShowDK
End Sub

Private Sub cmdEK_Click()
    Call ShowEK
End Sub

Private Sub cmdImportDK_Click()
    Call ImportDK
End Sub

Private Sub cmdImportZK_Click()
    Call ImportZK
End Sub

Private Sub cmdZK_Click()
    Call ShowZK
End Sub

Private Sub cmdAK_Click()
    Call ShowAll
End Sub

Private Sub cmdUpdate_Click()
    Call UpdateFromDownload
End Sub

Private Sub cmdUpdateFile_Click()
    Call UpdateFromFile
End Sub

'----------------------------------------
' EIGENEN NAMEN PRÜFEN
'----------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    Call CheckSheetNames
End Sub

Private Sub Worksheet_Deactivate()
    Call CheckSheetNames
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call CheckSheetNames
End Sub

Private Sub Worksheet_Activate()
    Call CheckSheetNames
    Call RefreshUpdateFileButton
End Sub

Private Function CheckSheetNames()
    Dim result As Integer
    If Me.Name <> WbNameConfig Then
        Me.Name = WbNameConfig
        result = MsgBox("Leider darf das Config-Sheet nicht umbenannt werden! :-(", vbInformation + vbOKOnly, "Böse!")
    End If
End Function

