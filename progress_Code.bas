Private Sub UserForm_Initialize()

    ' Center of Excel window
    xlCenterX = Application.Left + (Application.Width / 2)
    xlCenterY = Application.Top + (Application.Height / 2)

    ' Offset the form so its center aligns with Excel's center
    Me.Left = xlCenterX - (Me.Width / 2)
    Me.Top = xlCenterY - (Me.Height / 2)

    Me.progress_text.Caption = ""
    Me.BackColor = vbButtonFace
    'MakeFormTopMost (Me.Caption)
    Me.Caption = ""
End Sub

Public Sub start(msg)
    Me.Show vbModeless
    Me.progress_text.Caption = msg & "..."
    Me.BackColor = vbYellow
    DoEvents
End Sub

Public Sub complete()
    Me.BackColor = vbGreen
End Sub

Public Sub error()
    Me.BackColor = vbRed
End Sub

Public Sub finish()
    Me.progress_text.Caption = "Finished"
    Unload Me
End Sub

Public Sub activity(j)
    If j Mod 500 <> 0 Then Exit Sub
    n = j / 500 Mod 4
    txt = Me.progress_text.Caption
    txt = Replace(txt, ".", "")
    txt = txt & String(n, ".")
    Me.progress_text.Caption = txt
    DoEvents
End Sub
