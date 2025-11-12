Sub test_active()
    'init
    If Not MT.needs_active_list Then Exit Sub
    progress.start ("Processing Active List")
    check_active_matches
    progress.finish
End Sub

Sub process_active()
    If Not (MT.needs_active_list Or MT.needs_supplier_list) Then Exit Sub
    If MT.needs_active_list Then
        progress.start ("Processing Active List")
        check_active_matches
    Else
        progress.start ("Processing Supplier List")
        check_supplier_matches
    End If
    progress.finish
End Sub

Sub check_active_matches()
    Dim mismatch_arr() As Variant
    ReDim mismatch_arr(0 To 0)
    If Not MT.needs_gagg_list Then
        'active_accounts = active_tab.UsedRange.columns(1).value
        'active_accounts = flatten_array(active_accounts)
        active_data = active_tab.UsedRange.value
        mismatch_arr = find_active_mismatches(gagg_accounts, active_data)
        Call add_mismatch_rows(mismatch_arr, 1)
        Call update_checklist(S.QC.audit_checklist, "audit_pipp", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_mercantile", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_national_chains", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_usage", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_shopping", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_arrears", 0)
        If EDC.state <> "OH" Then
            set_step (5)
        Else
            set_step (3)
        End If
        Exit Sub
    End If
    gagg_accounts = filter_col(F.columns.account_number)
    gagg_accounts = flatten_array(gagg_accounts)
    'active_accounts = active_tab.UsedRange.columns(1).value
    'active_accounts = flatten_array(active_accounts)
    active_data = active_tab().UsedRange.value
    category_arr = filter_col(F.columns.mail_category)
    status_arr = filter_col(F.columns.status)
    match_arr = filter_col(F.columns.active_in_LP)
    sas_arr = filter_col(F.columns.sas_id)
    'match_arr = find_active_matches(gagg_accounts, active_accounts, status_arr, match_arr, category_arr)
    Call find_active_matches(gagg_accounts, active_data, status_arr, match_arr, category_arr, sas_arr)
    With filter_tab().UsedRange
        .columns(F.columns.active_in_LP.index).value = match_arr
        .columns(F.columns.status.index).value = status_arr
        .columns(F.columns.mail_category.index).value = category_arr
        .columns(F.columns.sas_id.index).value = sas_arr
    End With
    mismatch_arr = find_active_mismatches(gagg_accounts, active_data)
    Call add_mismatch_rows(mismatch_arr, UBound(gagg_accounts))
End Sub

Sub check_supplier_matches()
    If supplier_tab() Is Nothing Then Exit Sub
    Dim mismatch_arr() As Variant
    ReDim mismatch_arr(0 To 0)
    gagg_accounts = filter_col(F.columns.account_number)
    gagg_accounts = flatten_array(gagg_accounts)
    supplier_data = supplier_tab.UsedRange.value
    category_arr = filter_col(F.columns.mail_category)
    status_arr = filter_col(F.columns.status)
    match_arr = filter_col(F.columns.active_in_LP)
    Call find_supplier_matches(gagg_accounts, supplier_data, status_arr, match_arr, category_arr)
    With filter_tab().UsedRange
        .columns(F.columns.active_in_LP.index).value = match_arr
        .columns(F.columns.status.index).value = status_arr
        .columns(F.columns.mail_category.index).value = category_arr
    End With
    mismatch_arr = find_active_mismatches(gagg_accounts, supplier_data)
    Call add_mismatch_rows(mismatch_arr, UBound(gagg_accounts))
End Sub

Sub find_active_matches(gagg_accounts As Variant, active_data As Variant, ByRef status_arr As Variant, ByRef match_arr As Variant, ByRef category_arr As Variant, ByRef sas_arr As Variant)
    
    Dim dict As Object
    Dim results() As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    active_headers = get_array_row(active_data, 1)
    
    sas_id_col = search_1d_array(F.columns.sas_id.active_col.header, active_headers)
    
    n = UBound(active_data, 1)
    
    For j = 2 To n
        If Not dict.exists(active_data(j, 1)) Then dict.Add active_data(j, 1), active_data(j, sas_id_col)
    Next
    
    For i = 2 To UBound(gagg_accounts)
        If dict.exists(gagg_accounts(i)) Then
            match_arr(i, 1) = "Y"
            category_arr(i, 1) = "REN"
            status_arr(i, 1) = FS.eligible.eligible_ren_status
            sas_arr(i, 1) = dict(gagg_accounts(i))
        Else
            match_arr(i, 1) = "N"
            sas_arr(i, 1) = "-"
        End If
        progress.activity (i)
    Next
    
End Sub

Sub find_supplier_matches(gagg_accounts As Variant, supplier_data As Variant, ByRef status_arr As Variant, ByRef match_arr As Variant, ByRef category_arr As Variant)
    
    Dim dict As Object
    Dim results() As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    supplier_headers = get_array_row(supplier_data, 1)
    
    n = UBound(supplier_data, 1)
    
    For j = 2 To n
        x = CStr(supplier_data(j, 1))
        If Not dict.exists(x) Then dict.Add x, Nothing
    Next
    For i = 2 To UBound(gagg_accounts)
        x = CStr(gagg_accounts(i))
        If dict.exists(x) Then
            match_arr(i, 1) = "Y"
            category_arr(i, 1) = "REN"
            status_arr(i, 1) = FS.mismatch.eligible_new_status
        Else
            match_arr(i, 1) = "N"
        End If
        progress.activity (i)
    Next
    
End Sub

Function find_active_mismatches(gagg_accounts As Variant, active_data As Variant)
    
    Dim dict As Object
    Dim results() As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    n = UBound(active_data, 1)
    
    k = 0
    ReDim results(1 To n)
    
    If MT.needs_gagg_list Then
        For j = 2 To UBound(gagg_accounts)
            If Not dict.exists(gagg_accounts(j)) Then dict.Add gagg_accounts(j), 1
        Next
        For i = 2 To n
            x = CStr(active_data(i, 1))
            If Not dict.exists(x) Then
                k = k + 1
                results(k) = Array(x, i)
            End If
        Next
    Else
        For i = 2 To n
            k = k + 1
            results(k) = Array(active_data(i, 1), i)
        Next
    End If
    
    If k > 0 Then
        ReDim Preserve results(1 To k)
    Else
        results = Array()
    End If
    
    find_active_mismatches = results
    
End Function

Sub add_mismatch_rows(mismatch_arr As Variant, num_rows)

    If UBound(mismatch_arr) = -1 Then Exit Sub
    
    Set ff = filter_tab()
    
    Dim mismatch_data As Variant
    
    If MT.needs_active_list Then
        active_data = Sheets(SN.Active).UsedRange.value
    Else
        active_data = Sheets(SN.Supplier).UsedRange.value
    End If
    active_headers = get_array_row(active_data, 1)
    Call find_mismatch_data_cols(mismatch_arr, active_headers)
    Call update_filter_col_active_indices(active_headers)
    
    For j = 1 To UBound(F.order_array)
        index = 0
        For k = 1 To UBound(active_headers)
            If active_headers(k) = F.order_array(j).active_col.header Then
                F.order_array(j).active_col.index = k
                Exit For
            End If
        Next
    Next
    
    ReDim mismatch_data(1 To UBound(mismatch_arr), 1 To UBound(F.order_array))
    
    For j = 1 To UBound(mismatch_arr)
        row_num = mismatch_arr(j)(1)
        For k = 1 To UBound(F.order_array)
            col_num = F.order_array(k).active_col.index
            If col_num <> 0 Then
                'dp col_num & ": " & F.order_array(k).active_col.header
                x = active_data(row_num, col_num)
                x = Replace$(x, ",", "")
                If Not F.order_array(k).active_col.header <> "SUBACCOUNSERVICEID" Then
                    x = Replace$(x, "-", " ")
                End If
                x = Application.Trim(UCase$(x))
                If F.order_array(k).header Like "*ZIP*" Then x = Left(x, 5)
                If x = "" Then x = F.order_array(k).default_value
                mismatch_data(j, k) = x
            Else
                mismatch_data(j, k) = F.order_array(k).default_mismatch_value
            End If
        Next
    Next
    
    Set r = ff.Cells(num_rows + 1, 1).Resize(UBound(mismatch_arr), UBound(F.order_array))
    
    r.value = mismatch_data
    
    If S.Filter.highlight_mismatch_cells And MT.needs_gagg_list Then
        r.Cells.Interior.color = S.Filter.highlight_mismatch_color.InteriorColor
    End If
    
    If MT.needs_active_list Then
        reapply_autofilter (Sheets(SN.Active).index)
    Else
        reapply_autofilter (Sheets(SN.Supplier).index)
    End If
    
End Sub
 
Sub find_mismatch_data_cols(mismatch_arr As Variant, active_headers As Variant)
    For j = 1 To UBound(A.mismatch_columns)
        A.mismatch_columns(j).index = search_1d_array(A.mismatch_columns(j).header, active_headers)
        'dp A.mismatch_columns(j).header & ": " & A.mismatch_columns(j).index
    Next
End Sub

Sub update_filter_col_active_indices(active_headers As Variant)
    For j = 1 To UBound(F.order_array)
        If F.order_array(j).active_col.header <> "" Then
            F.order_array(j).active_col.index = search_1d_array(F.order_array(j).active_col.header, active_headers)
            'dp F.order_array(j).active_col.header & ": " & F.order_array(j).active_col.index
        End If
    Next
End Sub

Sub populate_category(match_arr)
    status_arr = filter_col(F.columns.status)
    category_arr = filter_col(F.columns.mail_category)
    sas_arr = filfer_col(F.columns.sas_id)
    For i = 2 To UBound(category_arr)
        If match_arr(i, 1) = "Y" Then
            status_arr(i, 1) = FS.renewal.eligible_ren_status
            category_arr(i, 1) = "REN"
            sas_arr(i, 1) = ""
        End If
    Next
    filter_tab().Cells(1, F.columns.status.index).Resize(UBound(status_arr)).value = status_arr
    filter_tab().Cells(1, F.columns.mail_category.index).Resize(UBound(match_arr)).value = category_arr
End Sub
