Option Explicit

Private Sub cmdOK_Click()
    Unload frmOptions
End Sub

Private Sub optGeneral_Click()
    fraGeneral.ZOrder 0
End Sub
Private Sub optContnets_Click()
    fraContnets.ZOrder 0
End Sub
Private Sub optViewer_Click()
    fraViewer.ZOrder 0
End Sub