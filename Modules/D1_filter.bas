Sub filter_list()

    If Not MT.needs_gagg_list Then Exit Sub

    progress.start ("Applying Filters")
    Set ff = filter_tab()
    
    status_arr = ff.UsedRange.columns(F.columns.status.index).value
    active_arr = ff.UsedRange.columns(F.columns.active_in_LP.index).value
    eligible_arr = ff.UsedRange.columns(F.columns.eligible_opt_out.index).value
    
    num_rows = UBound(eligible_arr, 1)
    
    'pipp
    Call apply_bool_filter(ff, F.columns.pipp, active_arr, status_arr, eligible_arr, num_rows, FS.pipp)
    
    'state rules
    If EDC.state = "OH" Then
        Call apply_bool_filter(ff, F.columns.mercantile, active_arr, status_arr, eligible_arr, num_rows, FS.mercantile)
    ElseIf EDC.state = "IL" Then
        Call apply_bool_filter(ff, F.columns.rtp, active_arr, status_arr, eligible_arr, num_rows, FS.rtp)
        Call apply_bool_filter(ff, F.columns.bgs_hold, active_arr, status_arr, eligible_arr, num_rows, FS.bgs_hold)
        Call apply_bool_filter(ff, F.columns.free_service, active_arr, status_arr, eligible_arr, num_rows, FS.free_service)
        Call apply_bool_filter(ff, F.columns.hourly_pricing, active_arr, status_arr, eligible_arr, num_rows, FS.hourly_pricing)
        Call apply_bool_filter(ff, F.columns.community_solar, active_arr, status_arr, eligible_arr, num_rows, FS.community_solar)
    End If
    
    'usage
    Call apply_usage_filter(ff, active_arr, status_arr, eligible_arr, F.columns.estimated_usage, FS.usage)
    
    'shopper
    Call apply_bool_filter(ff, F.columns.shopping, active_arr, status_arr, eligible_arr, num_rows, FS.shopper)
    
    'arrears
    If S.Filter.remove_arrears Then
        Call apply_bool_filter(ff, F.columns.arrears, active_arr, status_arr, eligible_arr, num_rows, FS.arrears)
    End If
    
    'national chains
    Call apply_national_chains_filter(ff, active_arr, status_arr, eligible_arr, F.columns.national_chains, F.columns.mail_city, F.columns.mail_state, F.columns.customer_class, FS.national_chain)
    
    ff.Cells(1, F.columns.status.index).Resize(num_rows).value = status_arr
    ff.Cells(1, F.columns.eligible_opt_out.index).Resize(num_rows).value = eligible_arr
    
    progress.finish
    
    'create contracts query sql with remaining accounts
    create_contracts_sql
    
    If EDC.state = "OH" Then
        Call update_checklist(S.QC.audit_checklist, "audit_pipp", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_mercantile_national", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_usage", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_shopping", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_arrears", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_hourly_pricing", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_solar", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_free_service", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_bgs_hold", 0)
    ElseIf EDC.state = "IL" Then
        Call update_checklist(S.QC.audit_checklist, "audit_pipp", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_mercantile_national", 0)
        Call update_checklist(S.QC.audit_checklist, "audit_usage", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_shopping", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_arrears", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_hourly_pricing", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_solar", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_free_service", 1)
        Call update_checklist(S.QC.audit_checklist, "audit_bgs_hold", 1)
    End If
    
    make_filter_waterfall
    make_cycle_waterfall
    
    Call set_step(3)
    
    If EDC.state <> "OH" Then Call set_step(4)
    
End Sub

Sub apply_bool_filter(filter_tab, data_col As ColumnHeader, ByRef active_arr, ByRef status_arr, ByRef eligible_arr, num_rows, status As FilterStatus)
    data_arr = filter_tab.UsedRange.columns(data_col.index).value
    For i = 2 To num_rows
        If eligible_arr(i, 1) = "Y" Then
            If active_arr(i, 1) = "Y" Then
                If data_col.apply_to_active Then
                    If data_arr(i, 1) = "Y" Then
                        status_arr(i, 1) = status.ineligible_ren_status
                        eligible_arr(i, 1) = "N"
                    End If
                Else
                    If status.eligible_ren_status <> "" Then status_arr(i, 1) = status.eligible_ren_status
                End If
            Else
                If data_arr(i, 1) = "Y" Then
                    status_arr(i, 1) = status.ineligible_new_status
                    eligible_arr(i, 1) = "N"
                End If
            End If
        End If
        progress.activity (i)
    Next
End Sub

Sub apply_usage_filter(filter_tab, active_arr As Variant, ByRef status_arr As Variant, ByRef eligible_arr As Variant, data_col As ColumnHeader, status As FilterStatus)

    data_arr = filter_col(data_col)
    
    If EDC.state = "IL" Then
        cust_class_arr = filter_col(F.columns.customer_class)
    End If
    
    For i = 2 To UBound(data_arr)
        If eligible_arr(i, 1) = "N" Then
            GoTo next_row
        End If
        If EDC.state = "IL" Then
            If cust_class_arr(i, 1) = "RESIDENTIAL" Then GoTo next_row
        End If
        If active_arr(i, 1) = "Y" Then
            'do we care about usage for renewal accounts?
        Else
            If IsNumeric(data_arr(i, 1)) Then
                If Evaluate(data_arr(i, 1) & EDC.usage_limit) Then
                    status_arr(i, 1) = status.ineligible_new_status
                    eligible_arr(i, 1) = "N"
                End If
            Else
                'non numeric usage data
            End If
        End If
        progress.activity (i)
next_row:
    Next
    
    'filter_tab.Cells(1, F.columns.status.index).Resize(UBound(status_arr, 1)).value = status_arr
    'filter_tab.Cells(1, F.columns.eligible_opt_out.index).Resize(UBound(eligible_arr, 1)).value = eligible_arr
    
End Sub

Sub apply_national_chains_filter(filter_tab, active_arr As Variant, ByRef status_arr As Variant, ByRef eligible_arr As Variant, data_col_0 As ColumnHeader, data_col_1 As ColumnHeader, data_col_2 As ColumnHeader, data_col_3 As ColumnHeader, status As FilterStatus)
    
    If EDC.state <> "OH" Then Exit Sub
    
    progress.start ("Removing National Chain Accounts")
    
    data_arr_0 = filter_col(data_col_0)
    data_arr_1 = filter_col(data_col_1)
    data_arr_2 = filter_col(data_col_2)
    data_arr_3 = filter_col(data_col_3)
    
    For i = 2 To UBound(status_arr)
        If eligible_arr(i, 1) = "N" Then
            GoTo not_eligible
        End If
        If active_arr(i, 1) = "Y" Then
            'do we care about national chains for renewal accounts?
        Else
            If data_arr_3(i, 1) <> "RES" Then
                If data_arr_2(i, 1) Like "WA*" Then
                    If data_arr_1(i, 1) = "SPOKANE" Then
                        status_arr(i, 1) = status.ineligible_new_status
                        eligible_arr(i, 1) = "N"
                        data_arr_0(i, 1) = "Y"
                    End If
                End If
            End If
        End If
        progress.activity (i)
not_eligible:
    Next
    
    filter_tab.UsedRange.columns(data_col_0.index).value = data_arr_0
    
    progress.complete
    
End Sub
