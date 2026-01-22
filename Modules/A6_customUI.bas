'custom ribbon buttons and insturction pop ups

Sub contracts_instructions()
    MsgBox "Run the contracts query from snowflake using the provided SQL code as you normally would"
End Sub

Sub LP_review_instructions()
    MsgBox "Review the name and address data for the eligible accounts before uploading to LP"
End Sub

Sub ribbon_on_load(ribbon As IRibbonUI)
    Set UI = ribbon
    'define_UI_settings
    'define_sheet_names
    'define_HOME_settings
    init
    'ids = S.HOME.mail_type_items
    'items = S.HOME.mail_type_labels
End Sub

Sub test_invalidate()
    If Not UI Is Nothing Then UI.Invalidate
End Sub

Sub set_step(k)
    k = CStr(k)
    If SN.HOME <> "" Then
        If home_tab().Range(S.HOME.step_number_location) = k Then Exit Sub
        home_tab().Range(S.HOME.step_number_location) = k
    End If
    If Not UI Is Nothing Then UI.Invalidate
End Sub

Function get_step()
    get_step = 0
    If SN.HOME <> "" Then
        get_step = home_tab().Range(S.HOME.step_number_location)
        get_step = Val(get_step)
    End If
End Function

Sub ribbon_reset(control As IRibbonControl)
    reset
End Sub

Sub ribbon_save(control As IRibbonControl)
    save_waterfall
End Sub

Sub ribbon_export_files(control As IRibbonControl)
    If Not review_eligible_data() Then
        progress.finish
        Exit Sub
    End If
    make_LP
    ren_drops
    make_mail_list
    make_opt_in_list
    export_files
End Sub

Sub ribbon_mail_type_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (get_step() <= 1)
End Sub

Sub ribbon_mail_type_item_count(control As IRibbonControl, ByRef count)
    'count = UBound(S.UI.mail_type_items) + 1
    count = 4
End Sub

Sub ribbon_get_mail_type_selected_item(control As IRibbonControl, ByRef id)
    id = ribbon_mail_type_id
End Sub

Sub ribbon_get_mail_type_id(control As IRibbonControl, index As Integer, ByRef id)
    arr = Array("NEW", "REN", "REN_ONLY", "SWP")
    id = arr(index)
End Sub

Sub ribbon_mail_type_labels(control As IRibbonControl, index As Integer, ByRef label)
    'label = S.UI.mail_type_labels(index)
    arr = Array("New Community", "Renewal", "Renewal Only", "Sweep")
    label = arr(index)
End Sub

Sub ribbon_mail_type_images(control As IRibbonControl, index As Integer, ByRef image)
    'image = S.UI.mail_type_images(index)
    arr = Array("ResourcePoolRefresh", "AddOrRemoveAttendees", "ResourcePoolMenu", "ResourcesAddMenu")
    image = arr(index)
End Sub

Sub ribbon_set_mail_type(control As IRibbonControl, id As String, index As Integer)
    'define_mail_type (S.UI.mail_type_items(index))
    arr = Array("NEW", "REN", "REN_ONLY", "SWP")
    define_mail_type (arr(index))
End Sub

Sub ribbon_set_EDC(control As IRibbonControl, id As String, index As Integer)
    define_EDC (S.UI.EDC_labels(index))
End Sub

Sub ribbon_EDC_item_count(control As IRibbonControl, ByRef count)
    count = UBound(S.UI.EDC_items) + 1
End Sub

Sub ribbon_get_EDC_selected_item(control As IRibbonControl, ByRef id)
    id = ribbon_EDC_id
End Sub

Sub ribbon_get_EDC_id(control As IRibbonControl, index As Integer, ByRef id)
    arr = Array("AES", "CEI", "CS", "DUKE", "OE", "OP", "TE", "AM", "COM")
    id = arr(index)
End Sub

Sub ribbon_EDC_labels(control As IRibbonControl, index As Integer, ByRef label)
    'label = S.UI.EDC_labels(index)
    arr = Array("AES", "CEI", "CS", "DUKE", "OE", "OP", "TE", "AM", "COM")
    label = arr(index)
End Sub

Sub ribbon_EDC_images(control As IRibbonControl, index As Integer, ByRef image)
    image = S.UI.EDC_images(index)
End Sub

Sub ribbon_EDC_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (get_step() <= 1)
End Sub

Sub ribbon_import_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 1)
    'enabled = True
End Sub

Sub ribbon_import_gagg_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 1) And (MT.needs_gagg_list) And imported_gagg = False
    'enabled = True
End Sub

Sub ribbon_import_active_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 1) And (MT.needs_active_list) And imported_active = False
    'enabled = True
End Sub

Sub ribbon_import_supplier_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 1) And (MT.needs_supplier_list)
    'enabled = False
End Sub

Sub ribbon_filter_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 1)
    enabled = enabled And imported_gagg And imported_active And imported_supplier
End Sub

Sub ribbon_dna_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 3) And EDC.state = "OH"
End Sub

Sub ribbon_contracts_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 4)
End Sub

Sub ribbon_mapping_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 5)
End Sub

Sub ribbon_review_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 6)
End Sub

Sub ribbon_export_enabled(control As IRibbonControl, ByRef enabled)
    enabled = (MT.name <> "") And (EDC.name <> "") And (get_step() = 7)
End Sub

Sub ribbon_get_mail_name(control As IRibbonControl, ByRef label)
    label = MT.name
End Sub

Sub ribbon_get_EDC_name(control As IRibbonControl, ByRef label)
    label = EDC.name
End Sub

Sub ribbon_import_gagg(control As IRibbonControl)
    Call import_gagg_files
    progress.finish
End Sub

Sub ribbon_import_active(control As IRibbonControl)
    Call import_active_list
    progress.finish
End Sub

Sub ribbon_import_supplier(control As IRibbonControl)
    Call import_supplier_list
    progress.finish
End Sub

Sub ribbon_filter_list(control As IRibbonControl)
    If EDC.name = "" Then Exit Sub
    If MT.name = "" Then Exit Sub
    If ribbon_community = "" Or ribbon_community = "(Community Name)" Then GoTo bad_community_name
    Call define_checklists
    Call preprocess
    Call process_active
    Call format_address_data
    Call filter_list
    Call make_filter_waterfall
    Call generate_mapping
    If MT.needs_gagg_list Then contracts_instructions
    progress.finish
    
    Exit Sub
    
bad_community_name:
    MsgBox "Please enter commnuity name"
    
End Sub

Sub ribbon_dna_check(control As IRibbonControl)
    Call test_dna(12)
    progress.finish
End Sub

Sub ribbon_test_contracts(control As IRibbonControl)
    ThisWorkbook.Save
    Call get_contracts_file
    'Call test_migration
End Sub

Sub ribbon_test_mapping(control As IRibbonControl)
    Call remove_other_ineligible
End Sub

Sub ribbon_review_data(control As IRibbonControl)
    If Not prompt_review Then Exit Sub
    LP_review_instructions
End Sub

Sub ribbon_make_LP(control As IRibbonControl)
    make_LP
End Sub

Sub ribbon_contract_change(control As IRibbonControl, str As String)
    ribbon_contract_number = Application.Trim(str)
    set_contract_id (ribbon_contract_number)
End Sub

Sub ribbon_oo_date_change(control As IRibbonControl, str As String)
    ribbon_opt_out_date = Application.Trim(str)
    set_oo_date (ribbon_opt_out_date)
End Sub

Sub ribbon_community_change(control As IRibbonControl, str As String)
    ribbon_community = Application.Trim(str)
    If str = "" Then
        set_community_name ("(Community Name)")
    Else
        set_community_name (ribbon_community)
    End If
    set_step (1)
    If Not UI Is Nothing Then UI.Invalidate
End Sub

Sub set_contract_id(str)
    ribbon_contract_number = str
    home_tab().Range(S.HOME.contract_location) = ribbon_contract_number
    If Not UI Is Nothing Then UI.Invalidate
End Sub

Sub set_oo_date(str)
    ribbon_opt_out_date = str
    home_tab().Range(S.HOME.oo_date_location) = ribbon_opt_out_date
    If Not UI Is Nothing Then UI.Invalidate
End Sub

Sub set_community_name(str)
    ribbon_community = str
    home_tab().Range(S.HOME.community_name_location) = ribbon_community
    If Not UI Is Nothing Then UI.InvalidateControl ("import_menu")
End Sub

Sub ribbon_get_community(control As IRibbonControl, ByRef str)
    str = ribbon_community
End Sub

Sub ribbon_get_contract(control As IRibbonControl, ByRef str)
    str = ribbon_contract_number
End Sub

Sub ribbon_get_oo_date(control As IRibbonControl, ByRef str)
    str = ribbon_opt_out_date
End Sub

Sub ribbon_get_mail_type(control As IRibbonControl, ByRef str)
    x = home_tab().Range(S.HOME.mail_type_location)
    ribbon_mail_type = x
    str = ribbon_mail_type
    define_mail_type (str)
End Sub

Sub ribbon_get_EDC(control As IRibbonControl, ByRef str)
    x = home_tab().Range(S.HOME.edc_location)
    ribbon_EDC = x
    str = ribbon_EDC
    define_EDC (str)
End Sub

Sub refresh_ribbon()
    ThisWorkbook.Activate
    ribbon_mail_type = home_tab().Range(S.HOME.mail_type_location)
    ribbon_EDC = home_tab().Range(S.HOME.edc_location)
    ribbon_contract_number = home_tab().Range(S.HOME.contract_location)
    ribbon_opt_out_date = home_tab().Range(S.HOME.oo_date_location)
    ribbon_community = home_tab().Range(S.HOME.community_name_location)
    define_EDC (ribbon_EDC)
    define_mail_type (ribbon_mail_type)
    If Not UI Is Nothing Then UI.Invalidate
End Sub
