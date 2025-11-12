Private Sub UserForm_Initialize()
    
    help_path = "file:///" & onedrive_documentation_folder() & S.errors.error_file & S.errors.error_section
    
    Me.Width = S.errors.error_form_width
    Me.Height = S.errors.error_form_height
    
    browser.Navigate help_path

End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnRefresh_Click()
    browser.Refresh
End Sub

Private Sub btnPrint_Click()
    browser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub UserForm_Resize()
    ResizeBrowser
End Sub

Private Sub ResizeBrowser()
    With browser
        .Top = 0
        .Left = 0
        .Width = Me.InsideWidth
        .Height = Me.InsideHeight
    End With
End Sub
