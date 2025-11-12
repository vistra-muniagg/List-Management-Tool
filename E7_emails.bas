Sub CreateEmailFromTemplate_Excel()
    Dim olApp As Object
    Dim olMail As Object
    Dim TemplatePath As String
    
    ' Path to your .oft template
    TemplatePath = "C:\Path\To\YourTemplate.oft"
    
    ' Create or get Outlook Application
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    ' Create mail item from template
    Set olMail = olApp.CreateItemFromTemplate(TemplatePath)
    
    With olMail
        ' Example: pull values from Excel sheet
        .To = Sheets("Sheet1").Range("A1").value
        .CC = Sheets("Sheet1").Range("A2").value
        .subject = "Report for " & Sheets("Sheet1").Range("B1").value
        
        ' Replace placeholders in template body
        .HTMLBody = Replace(.HTMLBody, "{Name}", Sheets("Sheet1").Range("B2").value)
        .HTMLBody = Replace(.HTMLBody, "{Date}", format(Now, "mmmm d, yyyy"))
        
        ' Save as draft
        .Save
    End With
    
    MsgBox "Draft email created and saved in Outlook."
    
    ' Cleanup
    Set olMail = Nothing
    Set olApp = Nothing
End Sub

