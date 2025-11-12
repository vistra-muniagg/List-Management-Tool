Sub make_LP()
    LP_template
    create_LP
End Sub

Sub LP_template()

    delete_sheet (SN.LP)
    
    Set LP = Sheets.Add(after:=filter_tab())
    
    With LP
        .Tab.color = C.GREEN_0.InteriorColor
        .name = SN.LP
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
        .Cells(1, 24) = "MailType"
        .Cells(1, 25) = "Community Name"
        
        .columns(1).NumberFormat = "MM/DD/YY"
        .columns(4).NumberFormat = "@"
        .columns(14).NumberFormat = "@"
        .columns(19).NumberFormat = "@"
        .columns(22).NumberFormat = "###-###-####"
        
        .Rows(1).Font.Bold = True
        
        .columns.AutoFilter
        .columns.AutoFit

    End With
    
End Sub

Sub create_LP()
    
    Set LP = LP_tab()
    
    With F.columns
        eligible_arr = filter_col(.eligible_opt_out)
        account_arr = filter_col(.account_number)
        class_arr = filter_col(.customer_class)
        name_arr = filter_col(.customer_name)
        email_arr = filter_col(.email)
        phone_arr = filter_col(.phone)
        s_address_arr = filter_col(.service_address)
        s_city_arr = filter_col(.service_city)
        s_state_arr = filter_col(.service_state)
        s_zip_arr = filter_col(.service_zip)
        m_address_arr = filter_col(.mail_address)
        m_city_arr = filter_col(.mail_city)
        m_state_arr = filter_col(.mail_state)
        m_zip_arr = filter_col(.mail_zip)
        category_arr = filter_col(.mail_category)
        cycle_arr = filter_col(.read_cycle)
    End With
    
    oo_date = get_oo_date()
    contract_id = get_contract_id()
    community_name = get_community_name()
    
    n = UBound(account_arr, 1)
    
    Dim arr As Variant
    ReDim arr(1 To n, 1 To 25)
    
    For j = 1 To 25
        arr(1, j) = LP.Cells(1, j)
    Next
    
    k = 1
    
    For i = 2 To n
        If eligible_arr(i, 1) <> "Y" Then GoTo skip_row:
        k = k + 1
        arr(k, 1) = oo_date
        arr(k, 2) = class_arr(i, 1)
        If arr(k, 2) = "COMMERCIAL" Then arr(k, 3) = "SMALL"
        arr(k, 4) = account_arr(i, 1)
        arr(k, 5) = contract_id
        arr(k, 6) = ""
        arr(k, 7) = name_arr(i, 1)
        arr(k, 8) = email_arr(i, 1)
        arr(k, 9) = phone_arr(i, 1)
        arr(k, 10) = s_address_arr(i, 1)
        arr(k, 11) = ""
        arr(k, 12) = s_city_arr(i, 1)
        arr(k, 13) = s_state_arr(i, 1)
        arr(k, 14) = s_zip_arr(i, 1)
        arr(k, 15) = m_address_arr(i, 1)
        arr(k, 16) = ""
        arr(k, 17) = m_city_arr(i, 1)
        arr(k, 18) = m_state_arr(i, 1)
        arr(k, 19) = m_zip_arr(i, 1)
        arr(k, 20) = cycle_arr(i, 1)
        arr(k, 21) = False
        arr(k, 22) = False
        arr(k, 23) = ""
        arr(k, 24) = category_arr(i, 1)
        arr(k, 25) = community_name
        progress.activity (i)
skip_row:
    Next
    
    LP.Range("A1").Resize(n, 25).value = arr
    
    LP.columns.AutoFit
    
End Sub
