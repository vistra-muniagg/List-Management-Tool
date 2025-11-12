Sub make_mail_list()

    progress.start "Making Mail List"
    
    delete_sheet (SN.mail_list)

    Set mail_list = Sheets.Add(after:=LP_tab())
    
    With mail_list
        .name = SN.mail_list
        .Cells(1, 1) = "Customer Number"
        .Cells(1, 2) = "2D Barcode"
        .Cells(1, 3) = "Customer Name"
        .Cells(1, 4) = "Mailing Address"
        .Cells(1, 5) = "Mailing Address 2"
        .Cells(1, 6) = "City"
        .Cells(1, 7) = "State"
        .Cells(1, 8) = "Zip"
        .Cells(1, 9) = "Service Address"
        .Cells(1, 10) = "Service Address 2"
        .Cells(1, 11) = "Service City"
        .Cells(1, 12) = "Service State"
        .Cells(1, 13) = "Service Zip"
        .Cells(1, 14) = "Community Name"
        .Cells(1, 15) = "Opt-Out Date"
        .columns(1).NumberFormat = "@"
        .columns(15).NumberFormat = "@"
        .Rows(1).Font.Bold = True
    End With
    
    Set ff = filter_tab()
    
    With F.columns
        eligible_arr = filter_col(.eligible_opt_out)
        account_arr = filter_col(.account_number)
        name_arr = filter_col(.customer_name)
        m1_arr = filter_col(.mail_address)
        m_city_arr = filter_col(.mail_city)
        m_state_arr = filter_col(.mail_state)
        m_zip_arr = filter_col(.mail_zip)
        s1_arr = filter_col(.service_address)
        s_city_arr = filter_col(.service_city)
        s_state_arr = filter_col(.service_state)
        s_zip_arr = filter_col(.service_zip)
    End With
    
    oo_date = get_oo_date()
    oo_date_str = expanded_oo_date(oo_date)
    community = get_community_name()
    
    num_rows = UBound(account_arr, 1)
    
    k = 1
    
    Dim arr As Variant
    ReDim arr(1 To num_rows, 1 To 15)
    
    For j = 1 To 15
        arr(k, j) = mail_list.Cells(1, j)
    Next
    
    For i = 2 To num_rows
        If eligible_arr(i, 1) = "Y" Then
            k = k + 1
            arr(k, 1) = account_arr(i, 1)
            arr(k, 2) = ""
            arr(k, 3) = name_arr(i, 1)
            arr(k, 4) = m1_arr(i, 1)
            arr(k, 5) = ""
            arr(k, 6) = m_city_arr(i, 1)
            arr(k, 7) = m_state_arr(i, 1)
            arr(k, 8) = m_zip_arr(i, 1)
            arr(k, 9) = s1_arr(i, 1)
            arr(k, 10) = ""
            arr(k, 11) = s_city_arr(i, 1)
            arr(k, 12) = s_state_arr(i, 1)
            arr(k, 13) = s_zip_arr(i, 1)
            arr(k, 14) = community
            arr(k, 15) = oo_date_str
            progress.activity (k)
        End If
    Next
    
    mail_list.Cells(1, 1).Resize(k, 15).value = arr
    
    progress.finish
    
    reapply_autofilter (mail_list.index)
    
End Sub
