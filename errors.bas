Sub print_error_message()

    SharePointURL = "https://txu.sharepoint.com/:w:/r/sites/Muni-Agg/_layouts/15/Doc.aspx?sourcedoc=%7B0634B59A-9056-4861-B2D6-69D1A45E5D6B%7D&file=Helpfile.docx&action=default&mobileredirect=true&wdsle=0"
    
    fullurl = SharePointURL & "#Error" & Err.Number
    Err.HelpFile = fullurl
    Err.HelpContext = Err.Number
    
    errorText = Replace("Runtime Error 'XXXX':", "XXXX", Err.Number) & vbCrLf & Err.Description
    ErrorTitle = "Microsoft Visual Basic for Applications"
    buttonsAndIcon = vbCritical + vbOKOnly + &H4000
    
    response = MessageBox(0, errorText, ErrorTitle, buttonsAndIcon)
    
End Sub

Sub throw_error(n As Integer, Optional context_array As Variant)

    If Err.Number >= 9000 Then Exit Sub
    
    Err.Number = n
    Err.HelpContext = n
    
    '8888 is the safest error
    If n = 8888 Then Exit Sub
    
    If stepnumber < 2 And Not testing Then
        testing = True
        resetsheet
        testing = False
    End If
    
    Load HelpForm
    
    Select Case n
        'file import errors
        Case 9000: d = "Error Importing List"
        Case 9001: d = "Unexpected File Format" & error_context(context_array)
        Case 9002: d = "No Active Customer List Selected"
        Case 9003: d = "No Utility File Selected"
        Case 9004: d = "Unable To Determine Source Of Active Customer List"
        Case 9005: d = "FE List Is Missing 2nd Mailing Address Tab"
        Case 9006: d = "-"
        Case 9007: d = "Account Numbers Not Detected On Active List"
        Case 9008: d = "Account Numbers Not Detected On Utility List"
        Case 9009: d = "Active Customer List Contains Bad Data"
        
        'qc1 errors
        Case 9010: d = "Utility List Contains No Usable Data"
        Case 9011: d = "Utility List Contains Invalid Account Number(s)"
        Case 9012: d = "-"
        Case 9013: d = "-"
        Case 9014: d = "-"
        Case 9015: d = "-"
        Case 9016: d = "Unable To Process Shifted AEP Data" & error_context(context_array)
        Case 9017: d = "Invalid Read Cycle Detected"
        Case 9018: d = "Account(s) have invalid read cycles on both active and utility list"
        Case 9019: d = "-"

        'active list qc errors
        Case 9020: d = "Error 9020"
        Case 9021: d = "Error 9021"
        Case 9022: d = "LP active list contains records from the wrong utility"
        Case 9023: d = "LP Active Customer List Contains Records From Multiple Communities"
        Case 9024: d = "Account(s) On Active Customer List From EH System Are Missing Address Data"
        Case 9025: d = "Percent of active list accounts not on utility list exceeds limit " & error_context(context_array)
        Case 9026: d = "Error 9026"
        Case 9027: d = "Error 9027"
        Case 9028: d = "Error 9028"
        Case 9029: d = "Error 9029"
        
        'mapping errors
        Case 9030: d = "Cannot Import Unvalidated Mapping"
        Case 9031: d = "File Does Not Contain Geocoding Results"
        Case 9032: d = "Unexpected Number Of Accounts Mapped" & error_context(context_array)
        Case 9033: d = "Invalid Mapping Result Entered"
        Case 9034: d = "Percentage Of Mapped Out Accounts Exceeds Limit"
        Case 9035: d = "Invalid Mapping Results Detected"
        Case 9036: d = "Incorrect mapping file format detected"
        Case 9037: d = "Error 9037"
        Case 9038: d = "Error 9038"
        Case 9039: d = "One or more accounts map out but are still marked as Eligible"
        
        'dna + chains errors
        Case 9040: d = "Unable To Find DNA List"
        Case 9041: d = "DNA List Is More Than 4 Weeks Old"
        Case 9042: d = "Unepexcted Number Of Accounts Mapped"
        Case 9043: d = "Error 9043"
        Case 9044: d = "Error 9044"
        Case 9045: d = "Error 9045"
        Case 9046: d = "Error 9046"
        Case 9047: d = "Error 9047"
        Case 9048: d = "Error 9048"
        Case 9049: d = "One or more accounts map out but are still marked as Eligible"
        
        'contracts query errors
        Case 9050: d = "-"
        Case 9051: d = "No file name detected for contracts qu"
        Case 9052: d = "Account(s) not found on gagg list. Possible incorrect contracts query file was uploaded"
        Case 9053: d = "-"
        Case 9054: d = "-"
        Case 9055: d = "-"
        Case 9056: d = "-"
        Case 9057: d = "-"
        Case 9058: d = "-"
        Case 9059: d = "-"
        
        'migration errors
        Case 9060: d = "Unable to find migration folder"
        Case 9061: d = "Unable to find migration file"
        Case 9062: d = "-"
        Case 9063: d = "-"
        Case 9064: d = "Error 9064"
        Case 9065: d = "Error 9065"
        Case 9066: d = "Error 9066"
        Case 9067: d = "Error 9067"
        Case 9068: d = "Error 9068"
        Case 9069: d = "Migration table not found"
        
        'qc2 errors
        Case 9070: d = "LP Output Error"
        Case 9071: d = "One or more account numbers is in the wrong format"
        Case 9072: d = "One or more accounts on the output tabs has an ineligible status"
        Case 9073: d = "Error 9073"
        Case 9074: d = "Mailing Address Is Invlaid"
        Case 9075: d = "LP REN Output Missing Address Data"
        Case 9076: d = "LP NEW Output Missing Address Data"
        Case 9077: d = "LP REN Output Contains Bad Service State Data" & error_context(context_array)
        Case 9078: d = "LP NEW Output Contains Bad Service State Data" & error_context(context_array)
        Case 9079: d = "One or more accounts map out but are still marked as Eligible"
        
        'file creation errors
        Case 9080: d = "Invalid Renewal Source"
        Case 9081: d = "Unepexcted Number Of Accounts Mapped"
        Case 9082: d = "Error 9082"
        Case 9083: d = "Error 9083"
        Case 9084: d = "Error 9084"
        Case 9085: d = "Error 9085"
        Case 9086: d = "Error 9086"
        Case 9087: d = "Error 9087"
        Case 9088: d = "Error 9088"
        Case 9089: d = "Error 9089"
        
        'export errors
        Case 9090: d = "Unable To Find LP-NEW tab"
        Case 9091: d = "Unable To Find LP-REN tab"
        Case 9092: d = "Error 9022"
        Case 9093: d = "Error 9023"
        Case 9094: d = "Error 9024"
        Case 9095: d = "Error 9025"
        Case 9096: d = "Error 9026"
        Case 9097: d = "Error 9027"
        Case 9098: d = "Error 9028"
        Case 9099: d = "Error Exporting Files"
        
        Case 9999: d = "Generic Error Description"
        
    End Select
    
    d = d & error_context(context_array)
    
    Err.Description = d
    
    'Dim helpCell As Range
    'helpFilePath = "C:\Users\400050\Desktop\HelpFile.docx"
    'Set helpCell = Sheets(SHEET_NAME_HOME).Range("N5")
    
    'Sheets(SHEET_NAME_HOME).Hyperlinks.Add Anchor:=Sheets(SHEET_NAME_HOME).Range("N5"), address:=helpFilePath, TextToDisplay:="Error " & Err.Number
    'helpCell.Font.Color = vbRed
    
    'print_error_message
    
    show_error_help (n)
    
    If 0 Then
    
        If Application.UserName = "Rodgers, Andrew" Then
            Call send_error_message_email(1)
        Else
            Call send_error_message_email
        End If
        
    End If
    
    slow_mode
    
End Sub

Sub send_error_message_email(Optional print_message = 0, _
                            Optional force_email = False, _
                            Optional message_text = "", _
                            Optional context = "", _
                            Optional subject = "")

    If context <> "" Then
        message = string_to_html(Application.UserName & vbCrLf & context)
        send_error_message_teams (message)
        Exit Sub
    End If

    folder_path = Replace(ThisWorkbook.path, "https://txu.sharepoint.com/sites/Muni-Agg/Shared Documents/(1) Operations", "")
    
    user_name = name_reverse(Application.UserName)
    
    If Application.UserName = "Rodgers, Andrew" Then user_name = "You"
    
    error_message_text = user_name & " triggered an error" & vbNewLine & _
                        "Error = " & Err.Number & vbNewLine & _
                        "Description = " & Err.Description & vbNewLine & _
                        "EDC = " & EDC_name & vbNewLine & _
                        "Mail Type = " & mail_type & vbNewLine & _
                        "Step = " & stepnumber & vbNewLine & _
                        "File = " & ThisWorkbook.name & vbNewLine & _
                        "Folder= " & folder_path & vbNewLine
    
    If print_message <> 0 And 0 Then
        MsgBox error_message_text, vbCritical
        Exit Sub
    End If
                        
    'error_message_html = name_reverse(Application.UserName) & " triggered an error" & "<br>" & _
                        "<b>Error = </b>" & Err.Number & "<br>" & _
                        "<b>Description = </b>" & Err.Description & "<br>" & _
                        "<b>EDC = </b>" & EDC_name & "<br>" & _
                        "<b>Mail Type = </b>" & mail_type & "<br>" & _
                        "<b>Step = </b>" & step_name(stepnumber) & "<br>" & _
                        "<b>File = </b>" & ThisWorkbook.name & "<br>" & _
                        "<b>Folder= </b>" & folder_path & "<br>"
                        
    error_message_html = "<table border='1' cellpadding='5' cellspacing='0'>" & _
                        "<tr><td>Error Number</td><td>e1</td></tr>" & _
                        "<tr><td>Description</td><td>e2</td></tr>" & _
                        "<tr><td>EDC</td><td>e3</td></tr>" & _
                        "<tr><td>Mail Type</td><td>e4</td></tr>" & _
                        "<tr><td>Step</td><td>e5</td></tr>" & _
                        "<tr><td>File Name</td><td>e6</td></tr>" & _
                        "<tr><td>Folder Path</td><td>e7</td></tr>" & _
                        "<tr><td>User</td><td>e8</td></tr>" & _
                        "<tr><td>Time</td><td>e0</td></tr>" & _
                        "</table>"
                        
                        error_message_html = Replace(error_message_html, "e0", Now())
                        error_message_html = Replace(error_message_html, "e1", Err.Number)
                        error_message_html = Replace(error_message_html, "e2", Err.Description)
                        error_message_html = Replace(error_message_html, "e3", EDC_name)
                        error_message_html = Replace(error_message_html, "e4", mail_type)
                        error_message_html = Replace(error_message_html, "e5", step_name(stepnumber))
                        error_message_html = Replace(error_message_html, "e6", ThisWorkbook.name)
                        error_message_html = Replace(error_message_html, "e7", folder_path)
                        error_message_html = Replace(error_message_html, "e8", name_reverse(Application.UserName))
    
    
    If message_text <> "" Then error_message_html = string_to_html(message_text)
    
    If Not force_email Then
        If user_name <> "You" Then
            send_error_message_teams (error_message_html)
            Exit Sub
        End If
    End If
                        
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim EmailBody As String
    Dim EmailSubject As String
    Dim Recipient As String
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    For Each OutlookAccount In OutlookApp.Session.accounts
        If OutlookAccount.SmtpAddress Like "*@vistracorp.com" Then
            Set OutlookMail.SendUsingAccount = OutlookAccount
            Exit For
        End If
    Next
    
    Recipient = "andrew.rodgers@vistracorp.com" ' Change this to your email address
    If subject = "" Then
        EmailSubject = "Macro Error Report"
    Else
        EmailSubject = subject
    End If
    EmailBody = error_message_html
    
    ' Set email properties
    With OutlookMail
        .To = Recipient
        .subject = EmailSubject
        .Body = EmailBody
        .HTMLBody = error_message_html
        .Display
        If user_name <> "You" Then .Send
    End With

    ' Clean up
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
End Sub

Sub send_error_message_teams(html_message)

    'sends html message via teams
    
    Dim Http As Object
    Dim url As String
    Dim jsonPayload As String

    url = "https://prod-90.westus.logic.azure.com:443/workflows/67c1fecccafc4a718348d5aaded5becb/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=6RsGf9Ug7Ed774BMj2ExsopubhOSkVi4dQ1L6hMVRSE"
    
    jsonPayload = "{""message"": """ & Replace(html_message, """", "\""") & """}"

    Set Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Http.Open "POST", url, False
    Http.setRequestHeader "Content-Type", "application/json"
    Http.Send jsonPayload
    
End Sub

Function string_to_html(str)
    str = Replace(str, "\", "/")
    str = "<table border='1' cellpadding='5' cellspacing='0'><tr><td>" & str & "</td></tr></table>"
    str = Replace(str, vbCrLf, "</td></tr><tr><td>")
    string_to_html = str
End Function

Sub ShowMessageBoxWithHelpButton()
    Const MB_OKCANCELHELP As Long = &H4001 ' OK, Cancel, and Help buttons
    Const IDHELP As Long = 5               ' Return value when Help is clicked
    Dim result As Long

    ' Display the custom message box
    result = MessageBox(0, "Do you need help?", "Custom Help Box", MB_OKCANCELHELP)

    ' Handle the result
    Select Case result
        Case vbOK
            MsgBox "You clicked OK."
        Case vbCancel
            MsgBox "You clicked Cancel."
        Case IDHELP
            ' Open a webpage or perform any other custom action
            shell "cmd.exe /c start https://www.example.com/help", vbNormalFocus
    End Select
End Sub

Sub show_error_help(n)
    
    With HelpForm
        .error_label.Caption = Replace("Runtime Error 'X':", "X", n)
        .error_text_box.text = Err.Description
        .HelpContextID = n
        .StartUpPosition = 3
        .Top = Application.Top + Application.Height / 2 - .Height / 2
        .Left = Application.Left + Application.Width / 2 - .Width / 2
        .Show vbModeless
    End With
    
End Sub

Function error_context(C)
    If IsMissing(C) Then
        error_context = ""
        Exit Function
    End If
    If Not IsArray(C) Then C = Array(C)
    Select Case Err.Number
        Case 9001:
            error_context = "Expected File Format: " & arr2str(EDC.get_file_formats()) & vbCrLf & "Detected File Format: " & C(0)
        Case 9016:
            error_context = "Utility List Data Is Shifted " & C(0) & vbCrLf & "File Name: " & C(1)
        Case 9025:
            error_context = "Limit: " & 100 * MISMATCH_LIMIT_PCT & "%" & vbCrLf & "Detected Mismatch Percent: " & C(0)
        Case 9032:
            error_context = "Expected Count: " & C(0) & vbCrLf & "Detected Count: " & C(1)
        Case 9077, 9078:
            error_context = "Expected State: " & EDC.state & vbCrLf & "Detected State: " & C(0)
    End Select
    error_context = vbCrLf & vbCrLf & error_context
End Function
