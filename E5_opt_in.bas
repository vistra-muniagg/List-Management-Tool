Sub make_opt_in_list()

    If Not MT.needs_gagg_list Then Exit Sub

    define_opt_in_settings
    
    If Not S.OptIn.make_opt_in Then Exit Sub
    
    If Not MT.make_opt_in_list Then Exit Sub
    
    Set opt_in = Sheets.Add(after:=LP_tab())
    
    delete_sheet (SN.opt_in)
    
    opt_in.name = SN.opt_in
    
    opt_in.columns(1).NumberFormat = "@"
    
    With F.columns
        account_arr = filter_col(.account_number)
        cust_name_arr = filter_col(.customer_name)
        mail_address_arr = filter_col(.mail_address)
        mail_city_arr = filter_col(.mail_city)
        mail_state_arr = filter_col(.mail_state)
        mail_zip_arr = filter_col(.mail_zip)
        service_address_arr = filter_col(.service_address)
        service_city_arr = filter_col(.service_city)
        service_state_arr = filter_col(.service_state)
        service_zip_arr = filter_col(.service_zip)
        status_arr = filter_col(.status)
        active_arr = filter_col(.active_in_LP)
        oo_eligible_arr = filter_col(.eligible_opt_out)
        oi_eligible_arr = filter_col(.opt_in_eligible)
        dna_arr = filter_col(.do_not_agg)
        mapping_arr = filter_col(.mapping_result)
    End With
    
    num_rows = UBound(account_arr)
    
    Dim arr As Variant
    Dim arr_t As Variant
    ReDim arr(1 To num_rows, 1 To S.OptIn.opt_in_num_cols)
    ReDim arr_t(1 To S.OptIn.opt_in_num_cols, 1 To num_rows)
    
    eligible_count = 0
    
    For j = 1 To S.OptIn.opt_in_num_cols
        arr(1, j) = S.OptIn.opt_in_columns(j).header
    Next
    
    For i = 2 To num_rows
        'eligible if not oo eligibe + shopper + not renewal + maps in
        eligible = True
        If oo_eligible_arr(i, 1) = "Y" Then
            eligible = False
            GoTo next_row
        End If
        If active_arr(i, 1) = "Y" Then
            eligible = False
            GoTo next_row
        End If
        If status_arr(i, 1) <> FS.shopper.ineligible_new_status Then
            If S.OptIn.opt_in_keep_dna And dna_arr(i, 1) = "Y" Then
                eligible = True
            ElseIf S.OptIn.opt_in_keep_mapped_out And mapping_arr(i, 1) = "Y" Then
                elgibile = True
            Else
                eligible = False
            End If
            GoTo next_row
        End If
next_row:
        If eligible Then
            oi_eligible_arr(i, 1) = "Y"
            k = k + 1
            arr(k + 1, 1) = account_arr(i, 1)
            arr(k + 1, 2) = cust_name_arr(i, 1)
            arr(k + 1, 3) = mail_address_arr(i, 1)
                'arr(k + 1, 3) = clean_mail_address(arr(k + 1, 3))
            arr(k + 1, 4) = mail_city_arr(i, 1)
            arr(k + 1, 5) = mail_state_arr(i, 1)
            arr(k + 1, 6) = mail_zip_arr(i, 1)
            arr(k + 1, 7) = service_address_arr(i, 1)
            arr(k + 1, 8) = service_city_arr(i, 1)
            arr(k + 1, 9) = service_state_arr(i, 1)
            arr(k + 1, 10) = service_zip_arr(i, 1)
        End If
        progress.activity (i)
    Next
    
    arr_t = Application.WorksheetFunction.Transpose(arr)
    
    ReDim Preserve arr_t(1 To S.OptIn.opt_in_num_cols, 1 To k + 1)
    
    arr = Application.WorksheetFunction.Transpose(arr_t)
    
    opt_in.Range("A1").Resize(k + 1, S.OptIn.opt_in_num_cols).value = arr
    
    reapply_autofilter (opt_in.index)
    
    filter_tab().UsedRange.columns(F.columns.opt_in_eligible.index).value = oi_eligible_arr
    
End Sub
