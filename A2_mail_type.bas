Public Type MailType
    name As String
    display_name As String
    combine_filter_pivots As Boolean
    combine_mapping_pivots As Boolean
    add_cycle_pivots As Boolean
    cycle_pivot_colors As Boolean
    has_renewal As Boolean
    has_new As Boolean
    has_contracts_query As Boolean
    needs_active_list As Boolean
    needs_gagg_list As Boolean
    needs_supplier_list As Boolean
    make_opt_in_list As Boolean
    make_ren_LP As Boolean
    make_new_LP As Boolean
    make_ren_mail_list As Boolean
    check_migration_data As Boolean
    highlight_mismatches As Boolean
    keep_active_mapped_out As Boolean
    export_opt_in_list As Boolean
    waterfall_title As String
    LP_file_suffix As String
End Type

Sub define_mail_type(mail_type)
    'how to get mail type selection?
    If mail_type = S.UI.mail_type_items(0) Then
        define_mail_type_NEW
    ElseIf mail_type = S.UI.mail_type_items(1) Then
        define_mail_type_REN
    ElseIf mail_type = S.UI.mail_type_items(2) Then
        define_mail_type_REN_ONLY
    ElseIf mail_type = S.UI.mail_type_items(3) Then
        define_mail_type_SWP
    Else
        Exit Sub
    End If
    x = home_tab().Range(S.HOME.mail_type_location)
    If x <> mail_type Then home_tab().Range(S.HOME.mail_type_location) = mail_type
    If Not MT.needs_active_list Then
        imported_active = True
    Else
        imported_active = False
    End If
    If Not MT.needs_gagg_list Then
        imported_gagg = True
    Else
        imported_gagg = False
    End If
    If Not MT.needs_supplier_list Then
        imported_supplier = True
    Else
        imported_supplier = True
    End If
    ribbon_mail_type = MT.name
    'if MT.name <>"" then Call set_step(1)
    If Not utility_tab() Is Nothing Then imported_gagg = True
    If Not active_tab() Is Nothing Then imported_active = True
    'If Not supplier_tab() Is Nothing Then imported_supplier = True
    If Not UI Is Nothing Then
        UI.InvalidateControl ("import_menu")
        UI.InvalidateControl ("EDC_menu")
    End If
End Sub

Sub define_mail_type_NEW()
    With MT
        .name = "NEW"
        .display_name = "New Community"
        .has_renewal = False
        .has_new = True
        .has_contracts_query = True
        .needs_active_list = False
        .needs_gagg_list = True
        .needs_supplier_list = True
        .check_migration_data = False
        .make_opt_in_list = True
        .export_opt_in_list = True
        .waterfall_title = "Utility + Supplier Accounts"
        .cycle_pivot_colors = True
        .keep_active_mapped_out = False
        .LP_file_suffix = " - NEW"
    End With
End Sub

Sub define_mail_type_REN()
    With MT
        .name = "REN"
        .display_name = "Renewal"
        .has_renewal = True
        .has_new = True
        .has_contracts_query = True
        .needs_active_list = True
        .needs_gagg_list = True
        .needs_supplier_list = False
        .check_migration_data = False
        .make_opt_in_list = True
        .export_opt_in_list = True
        .waterfall_title = "Utility + LP Accounts"
        .cycle_pivot_colors = True
        .keep_active_mapped_out = True
        .LP_file_suffix = " - REN_SWP"
    End With
End Sub

Sub define_mail_type_SWP()
    With MT
        .name = "SWP"
        .display_name = "Sweep"
        .has_renewal = False
        .has_new = True
        .has_contracts_query = True
        .needs_active_list = False
        .needs_gagg_list = True
        .needs_supplier_list = False
        .check_migration_data = True
        .make_opt_in_list = True
        .export_opt_in_list = True
        .waterfall_title = "Utility Accounts"
        .cycle_pivot_colors = True
        .keep_active_mapped_out = True
        .LP_file_suffix = " - SWP"
    End With
End Sub

Sub define_mail_type_REN_ONLY()
    With MT
        .name = "REN_ONLY"
        .display_name = "Renewal (No Sweep)"
        .has_renewal = True
        .has_new = False
        .has_contracts_query = False
        .needs_active_list = True
        .needs_gagg_list = False
        .needs_supplier_list = False
        .check_migration_data = False
        .make_opt_in_list = True
        .export_opt_in_list = False
        .waterfall_title = "LP Accounts"
        .cycle_pivot_colors = True
        .keep_active_mapped_out = True
        .LP_file_suffix = " - REN"
    End With
End Sub

Function empty_MT() As MailType
    'returns empty mail type
End Function
