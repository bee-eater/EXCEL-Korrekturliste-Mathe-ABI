Attribute VB_Name = "M8_Progress"
Option Explicit

Public Function ShowProgress(Text As String, Done As Integer, Show As Boolean) As Boolean

    If Not frmProgress.Visible Then
        frmProgress.StartUpPosition = 0
        frmProgress.top = Application.top + (0.5 * Application.Height) - (0.5 * frmProgress.Height)
        frmProgress.left = Application.left + (0.5 * Application.Width) - (0.5 * frmProgress.Width)
    End If
    
    If Show = True Then
        frmProgress.Show vbModeless
        If Done <= frmProgress.progDownload.Max And Done >= frmProgress.progDownload.Min Then
            frmProgress.progDownload.Value = Done
        Else
            frmProgress.progDownload.Value = 100
        End If
        frmProgress.txtDownload.Text = Text
    Else
        frmProgress.Hide
    End If
    
End Function
