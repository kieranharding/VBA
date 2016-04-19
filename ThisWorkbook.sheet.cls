

Private Sub Workbook_AddinInstall()
    ReAddin
End Sub


Private Sub Workbook_AddinUninstall()
    DeAddin
End Sub

Private Sub Workbook_Open()
    KeyboardShortcuts
End Sub