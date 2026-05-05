VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Notenspiegel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
End Sub

Private Function CheckSheetNames()
    Dim result As Integer
    If Me.Name <> WbNameGradeKey Then
        Me.Name = WbNameGradeKey
        result = MsgBox("Leider darf das Notenschlüssel-Sheet nicht umbenannt werden! :-(", vbInformation + vbOKOnly, "Böse!")
    End If
End Function
