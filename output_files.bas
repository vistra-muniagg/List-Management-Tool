Sub LP_template(sheet_name)
    
    Set LP = Sheets.Add(before:=Sheets(1))
    
    LP.name = SN.LP
    
    With LP
    
        .name = sheet_name
        .Cells(1, 1) = "OptOutDate"
        .Cells(1, 2) = "PremiseType"
        .Cells(1, 3) = "CommercialClassType"
        .Cells(1, 4) = "AccountNumber"
        .Cells(1, 5) = "ContractNumber"
        .Cells(1, 6) = "FirstName"
        .Cells(1, 7) = "LastName"
        .Cells(1, 8) = "Email"
        .Cells(1, 9) = "PrimaryPhone"
        .Cells(1, 10) = "ServiceAddress1"
        .Cells(1, 11) = "ServiceAddress2"
        .Cells(1, 12) = "ServiceCity"
        .Cells(1, 13) = "ServiceState"
        .Cells(1, 14) = "ServicePostalCode"
        .Cells(1, 15) = "BillingAddress1"
        .Cells(1, 16) = "BillingAddress2"
        .Cells(1, 17) = "BillingCity"
        .Cells(1, 18) = "BillingState"
        .Cells(1, 19) = "BillingPostalCode"
        .Cells(1, 20) = "BillCycle"
        .Cells(1, 21) = "SuppressOutboundEnrollmentTransaction"
        .Cells(1, 22) = "SuppressUtilityNotification"
        .Cells(1, 23) = "CustomerNameKey"
        .columns(22).NumberFormat = "###-###-####"

        '.Range("A1:X1").Style = "Bad"
        'For Each p In Array("E", "H", "M", "S", "T", "V", "W")
        '    .Cells(1, p).Style = "Good"
        'Next
        
        For Each p In Array("D", "N", "S")
            .columns(p).NumberFormat = "@"
        Next
        
        .columns("A").NumberFormat = "MM/DD/YY"
        
        '.Columns("N").NumberFormat = "00000"
        '.Columns("S").NumberFormat = "00000"
        
        .Rows(1).Font.Bold = True
        
        .columns.AutoFilter
        .columns.AutoFit

    End With
    
End Sub

Sub SF_template()
    
    'add SF sheet
    Sheets.Add before:=Sheets(1)
    
    With Sheets(1)
    
        .name = SHEET_NAME_SF
        .Cells(1, 1) = "OptOutEndDate"
        .Cells(1, 2) = "LDCType"
        .Cells(1, 3) = "AccountOwner"
        .Cells(1, 4) = "MuniAggCustID"
        .Cells(1, 5) = "LDCVendor"
        .Cells(1, 6) = "ContractNumber"
        .Cells(1, 7) = "ServiceTerritory"
        .Cells(1, 8) = "CustomerName2"
        
        .columns("A").NumberFormat = "MM/DD/YY"
        
        .Rows(1).Font.Bold = True
        
        .columns.AutoFilter
        .columns.AutoFit

    End With
    
End Sub

Sub SF_template_2()

    'add SF sheet
    Sheets.Add before:=Sheets(1)
    
    With Sheets(1)
    
        .name = SHEET_NAME_SF
        .Cells(1, 1) = "LDC_Account_Number__c"
        .Cells(1, 2) = "Contract_Number__c"
        .Cells(1, 3) = "Opt_Out_Period_Ends__c"
        .Cells(1, 4) = "LDC_Type__c" 'res/com
        .Cells(1, 5) = "OwnerId" 'don't use kate's id
        .Cells(1, 6) = "Muni_Agg_Customer_Id__c" 'community id from sf
        .Cells(1, 7) = "LDC_Vendor__c" 'full edc name
        .Cells(1, 8) = "Service_Territory__c" 'EDC
        .Cells(1, 9) = "Last_Name__c" 'full name goes here for SF
        .Cells(1, 10) = "Service_Street_1__c" 'full service address here?
        .Cells(1, 11) = "Service_City__c"
        .Cells(1, 12) = "Service_State__c"
        .Cells(1, 13) = "Service_Postal_Code__c"
        .Cells(1, 14) = "Billing_Street__c"
        .Cells(1, 15) = "Billing__c" 'use Mailing_City__c instead?
        .Cells(1, 16) = "Billing_State_Province__c"
        .Cells(1, 17) = "Billing_Zip_Postal_code__c"
        .Cells(1, 18) = "Bill_Cycle__c"
        .Cells(1, 19) = "Phone__c"
        .Cells(1, 20) = "Email__c"
        
        '.Columns("A").NumberFormat = Replace(Space(EDC.account_number_length), " ", "0")
        .columns("C").NumberFormat = "M/D/YYYY"
        .columns("Q").NumberFormat = "0"
        
        .UsedRange.Rows(1).Style = "Bad"
        .Cells(1, "S").Style = "Good"
        .Cells(1, "T").Style = "Good"
        
        .Rows(1).Font.Bold = True
        
        .columns.AutoFilter
        .columns.AutoFit

    End With

End Sub

Sub make_output()

    'If mail_type = "RENEWAL ONLY" Then Exit Sub
    If Not make_new_LP_upload(mail_type) Then Exit Sub
    
    ThisWorkbook.Activate
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    
    If Not testing Then
        With output_form
            contract_id = .contract_box.value
            oo_date = .date_box.value
        End With
        Unload output_form
    Else
        With Sheets(SHEET_NAME_HOME)
            contract_id = Sheets(SHEET_NAME_HOME).Range("N11")
            oo_date = Sheets(SHEET_NAME_HOME).Range("N12")
        End With
    End If
    
    'create dictionary of contract details
    Set d = CreateObject("scripting.dictionary")
    multiple_contracts = contract_id = "MULTIPLE"
    If multiple_contracts Then
        k = 22
        sanity_check_table = True
        With Sheets(SHEET_NAME_HOME)
            Do While .Cells(k, "D") <> ""
                d.Add .Cells(k, "D").value, Array(.Cells(k, "F").value, .Cells(k, "G").value) 'file_name: (CONTRACT_ID,SF_ID)
                If Not .Cells(k, "F") Like "C-00######" Then sanity_check_table = False
                k = k + 1
            Loop
        End With
        
        If Not sanity_check_table Then GoTo no_table_data

    End If
    
    account = Application.match(EDC.account, Sheets(1).UsedRange.Rows(1), 0)
    name = Application.match(EDC.cust_name, Sheets(1).UsedRange.Rows(1), 0)
    s_1 = Application.match(EDC.service1, Sheets(1).UsedRange.Rows(1), 0)
    s_2 = Application.match(EDC.service2, Sheets(1).UsedRange.Rows(1), 0)
    s_city = Application.match(EDC.service_city, Sheets(1).UsedRange.Rows(1), 0)
    s_state = Application.match(EDC.service_state, Sheets(1).UsedRange.Rows(1), 0)
    s_zip = Application.match(EDC.service_zip, Sheets(1).UsedRange.Rows(1), 0)
    m_1 = Application.match(EDC.mail1, Sheets(1).UsedRange.Rows(1), 0)
    m_2 = Application.match(EDC.mail2, Sheets(1).UsedRange.Rows(1), 0)
    m_city = Application.match(EDC.mail_city, Sheets(1).UsedRange.Rows(1), 0)
    m_state = Application.match(EDC.mail_state, Sheets(1).UsedRange.Rows(1), 0)
    m_zip = Application.match(EDC.mail_zip, Sheets(1).UsedRange.Rows(1), 0)
    read_cycle = Application.match(EDC.read_cycle, Sheets(1).UsedRange.Rows(1), 0)
    rate_code = Application.match(EDC.rate_code, Sheets(1).UsedRange.Rows(1), 0)
    email = Application.match(EDC.email, Sheets(1).UsedRange.Rows(1), 0)
    phone = Application.match(EDC.phone, Sheets(1).UsedRange.Rows(1), 0)
    
    'check for important columns
    col = 0
    With EDC
        column_name_array = Array(.account, .cust_name, .service1, .service_city, .service_state, .service_zip, .mail1, .mail_city, .mail_state, .read_cycle, .rate_code)
    End With
    For Each C In Array(account, name, s_1, s_city, s_state, s_zip, m_1, m_city, m_state, read_cycle, rate_code)
        missing_column_name = column_name_array(col)
        If IsError(C) Then
            GoTo missing_column
        End If
        col = col + 1
    Next
    
    'DUKE is special and puts this somewhere else
    If EDC.get_name = "DUKE" Then
        duke_apt = Application.WorksheetFunction.match("APT", Sheets(1).UsedRange.Rows(1), 0)
        duke_floor = Application.WorksheetFunction.match("FLOOR", Sheets(1).UsedRange.Rows(1), 0)
    End If
    
    If IsError(read_cycle) Then read_cycle = 0
    
    If bb Then bb_col = Application.match(EDC.budget_bill, Sheets(1).UsedRange.Rows(1), 0)
    
    numrows = Sheets(1).Cells(1, account).End(xlDown).row
    numcols = Sheets(1).Cells(1, 1).End(xlToRight).Column
    
    r = EDC.get_res_codes
    
    find_gagg
    
    LP_template (SHEET_NAME_LP_NEW)
    
    comed_special_case = EDC.get_name = "COM"
    
    Call add_stat(6, numrows - 1)
    
    'start progress bar
    filter_title = "Making Output Files"
    Call status_bar(0, numrows)
    
    'start timer
    StartTimer
    
    If numrows > 1000000 Then Exit Sub
    
    For i = 2 To numrows
        
        'id RES/COMM
        C = Application.Trim(gagg.Cells(i, rate_code))
        If Not comed_special_case Then
            If UBound(Filter(r, C)) >= 0 Or (EDC_name = "AES" And C Like "*RES*") Then
            'RES
                cust_class = "RESIDENTIAL"
                class_type = ""
            Else
            'COMM
                cust_class = "COMMERCIAL"
                class_type = "SMALL"
            End If
        Else
            If UCase(C) Like "*RESIDENTIAL*" Then
            'RES
                cust_class = "RESIDENTIAL"
                class_type = ""
            ElseIf UCase(C) Like "*COMMERCIAL*" Then
            'COMM
                cust_class = "COMMERCIAL"
                class_type = "SMALL"
            End If
        End If
        
        If multiple_contracts Then
            file_name = gagg.Cells(i, numcols - 1)
            contract_id = d(file_name)(0)
            cust_id = d(file_name)(1)
        End If
        
        With Sheets(SHEET_NAME_LP_NEW)
        
            'populate details
            .Cells(i, 1) = oo_date
            .Cells(i, 2) = cust_class
            .Cells(i, 3) = class_type
            .Cells(i, 4) = gagg.Cells(i, 1)
            .Cells(i, 5) = contract_id
            
            'populate name
            n = gagg.Cells(i, name)
            n = Replace(n, ".", "")
            n1 = InStrRev(n, ",")
            If n1 > 0 Then
                n = name_reverse(n)
            End If
            'n = Replace(n, " & ", "&")
            n = Application.Trim(n)
            
            'sf.Cells(i, "I") = n
            name_parts (n)
            If cust_class = "RESIDENTIAL" Then
                .Cells(i, 7) = first_name & " " & last_name
            Else
                .Cells(i, 7) = UCase(n)
            End If
            .Cells(i, 7) = Application.Trim(.Cells(i, 7))
            
            'email and phone if applicable
            If Not IsError(email) Then
                .Cells(i, 8) = Application.Trim(UCase(gagg.Cells(i, email)))
            End If
            If Not IsError(phone) Then
                If Len(gagg.Cells(i, phone)) = 10 Then .Cells(i, 9) = Application.Trim(format(gagg.Cells(i, phone), "##########"))
            End If
            
            'populate service address
            .Cells(i, 10) = service_address()
            .Cells(i, 12) = service_city()
            .Cells(i, 13) = service_state()
            .Cells(i, 14) = service_zip()
            
            'populate mail address
            .Cells(i, 15) = mail_address()
            .Cells(i, 17) = mail_city()
            .Cells(i, 18) = mail_state()
            .Cells(i, 19) = mail_zip()
            
            'highlight FE changed addresses
            If gagg.Cells(i, m_1 + 1).Interior.color = vbGreen Then .Cells(i, 15).Interior.color = vbGreen
            If Not gagg.Cells(i, m_1 + 1).Comment Is Nothing Then .Cells(i, 15).AddComment gagg.Cells(i, m_1 + 1).Comment.text
            
            'use service address if mail address info is empty
            If .Cells(i, 15) = "" Or .Cells(i, 17) = "" Or .Cells(i, 18) = "" Or .Cells(i, 19) = "" Then
                .Cells(i, 15) = .Cells(i, 10)
                .Cells(i, 17) = .Cells(i, 12)
                .Cells(i, 18) = .Cells(i, 13)
                .Cells(i, 19) = .Cells(i, 14)
            End If
            
            S = .Cells(i, 10)
            m = .Cells(i, 15)
            z1 = .Cells(i, 14)
            z2 = .Cells(i, 19)
            
            'use most complete data for mail address
            'only works with matching zip codes
            If S <> m And z1 = z2 Then
                If S Like m & "*" Then
                    .Cells(i, 15) = trim_junk(S)
                    '.Cells(i, 15).Style = "Neutral"
                End If
                If m Like S & "*" Then
                    .Cells(i, 10) = m
                    '.Cells(i, 10).Style = "Neutral"
                End If
            End If
            
            'populate read cycle
            mrc = gagg.Cells(i, read_cycle)
            If mrc Like "CE*" Then mrc = Mid(mrc, 3)
            If mrc = "" Then mrc = EDC.default_read_cycle
            .Cells(i, 20) = Val(correct_cycle(mrc))
            
            'add FALSE in column U and V
            .Cells(i, 21) = False
            .Cells(i, 22) = False
            
        End With
            
        Call status_bar(i, numrows)
        
    Next
    
    'end timer
    Call add_stat(7, elapsed_time())
    
    'align zip code and read cycle columns to left
    With Sheets(SHEET_NAME_LP_NEW)
        .columns(14).HorizontalAlignment = xlHAlignLeft
        .columns(19).HorizontalAlignment = xlHAlignLeft
        .columns(20).HorizontalAlignment = xlHAlignLeft
        .columns.AutoFit
    End With
    
    Application.StatusBar = False
    Application.ScreenUpdating = 1
    Application.Cursor = xlDefault
    
    'step (10)
    
    NewControlPanel.output_button.BackColor = 65280
    
    'which rows/columns to manually check? which ones can we skip?
    
    reorder_sheets
    
    pin_top_row (1)
    pin_top_row (2)
    
    With Sheets(SHEET_NAME_LP_NEW)
        If Not .AutoFilterMode Then .Range("A1").AutoFilter
        '.UsedRange.Rows(1).Style = "Good"
        .columns.AutoFit
    End With
    
    Unload NewControlPanel
    
    'create budget billing tab for AEP or AES
    export_BB
    
    Sheets(SHEET_NAME_LP_NEW).Activate
    
    Step (10)
    
    Exit Sub
    
missing_column:
    MsgBox "Column not found on Utility List: " & missing_column_name, vbCritical
    Application.Cursor = xlDefault
    Application.ScreenUpdating = 1
    Application.Calculation = xlCalculationAutomatic
    Unload NewControlPanel
    Exit Sub
    
no_table_data:
    MsgBox "Bulk Filtering Table not populated correctly. Fix it and try again", vbCritical
    Application.Cursor = xlDefault
    Application.ScreenUpdating = 1
    Application.Calculation = xlCalculationAutomatic
    Unload NewControlPanel
    Exit Sub
    
End Sub

Sub make_renewal_output()
    
    target_cell = file_name_cell
    
    Set C = Sheets(SHEET_NAME_HOME).Range(target_cell).Offset(0, -1)
    
    While C.text <> "EH" And C.text <> "LP" And C.text <> "SF"
        Set C = C.Offset(1, 0)
        If C.row > 50 Then
            throw_error (9004)
            Exit Sub
        End If
    Wend
    
    list_source = C.text
    
    If list_source = "EH" Then EH_to_LP
    If list_source = "LP" Then LP_to_LP
    If list_source = "SF" Then SF_to_LP
    
    pin_top_row (1)
    
    Sheets(1).columns.AutoFit
    
    If Sheets(1).name = "LP" Then Sheets(1).name = "LP-REN"
    
End Sub
