Public Type HomeTabSettings
    peer_review_checklist_range As String
    version_location As String
    user_location As String
    edc_location As String
    build_date_location As String
    mail_type_location As String
    step_number_location As String
    contract_location As String
    oo_date_location As String
    info_range As String
    add_cycle_pivots As Boolean
    large_community_limit As Long
    filter_waterfall_name As String
    filter_waterfall_location As String
    filter_waterfall_caption As String
    mapping_waterfall_name As String
    mapping_waterfall_location As String
    mapping_waterfall_caption As String
    notes_location As String
    file_log_location As String
    file_log_range As String
    audit_checklist_location As String
    qc_checklist_location As String
    cycle_pivot_name As String
    cycle_pivot_caption As String
    cycle_pivot_location As String
    community_name_location As String
    renewal_drop_count_location As String
End Type

Public Type UserInterfaceSettings
    mail_type_labels() As Variant
    mail_type_items() As Variant
    mail_type_images() As Variant
    EDC_labels() As Variant
    EDC_items() As Variant
    EDC_images() As Variant
End Type

Public Type ImportSettings
    max_csv_cols As Long
    max_copy_size As Long
    trim_sheets As Boolean
    FE_address_replace As Boolean
End Type

Public Type FilterSettings
    keep_OH_renewal_mapped_out As Boolean
    keep_IL_renewal_mapped_out As Boolean
    add_read_cycle_comment As Boolean
    check_active_list_cycles As Boolean
    check_gagg_list_read_cycles As Boolean
    add_cycle_pivot_for_large_list As Boolean
    annualize_usage As Boolean
    highlight_mismatch_cells As Boolean
    highlight_mismatch_color As CellColors
    AES_use_both_usage As Boolean
    remove_arrears As Boolean
End Type

Public Type DoNotAggSettings
    file_name As String
    folder_name As String
    max_file_age As Long
    sheet_name As String
    dna_name_col As Long
    dna_address_col As Long
    wildcard_length As Long
    results_layout As Variant
    include_wildcard_search As Boolean
    result_col As Long
    auto_populate_account_match As Boolean
    account_match_col As Long
End Type

Public Type ContractsQuerySettings
    hide_snowflake_query As Boolean
    remove_suffix_for_xdupx As Boolean
    xdupx_guess_wildcard_length As Long
    Levenshtein_match_len As Long
    xdupx_header As String
    LP_cust_name As String
    sas_id_row_header As String
    status_header As String
    status_reason_header As String
    muniagg_status_header As String
    contract_header As String
    intent_contract_header As String
End Type

Public Type MigrationSettings
    SF_migration_folder As String
    EH_migration_folder As String
    migration_log_file As String
    migration_log_sheet As String
End Type

Public Type StatsSettings
    show_stats_tab As Boolean
End Type

Public Type OptInSettings
    make_opt_in As Boolean
    opt_in_keep_dna As Boolean
    opt_in_keep_mapped_out As Boolean
    opt_in_num_cols As Long
    opt_in_columns() As ColumnHeader
End Type

Public Type MappingSettings
    mapping_db_file_name As String
    maps_in_label As String
    maps_out_label As String
    no_result_label As String
    maps_out_active_status As String
    maps_out_new_status As String
    file_source As String
    mapping_col As String
    notes_col As String
    map_this_sheet As String
    mapped_community As String
    mapped_out_retained_label As String
    mapped_out_retained_status As String
    mapping_tool_file As String
    no_results_label As String
    map_out_limit As Long
    no_result_limit As Long
    no_result_searchable_limit As Long
    no_result_unsearchable_limit As Long
    map_out_retained_limit As Long
End Type

Public Type OneDriveSettings
    parent_folder As String
    mailings_folder As String
    list_management_folder As String
    dna_folder As String
    migration_folder As String
    mapping_db_folder As String
    EH_migration_folder As String
    SF_migration_folder As String
    migration_log_file As String
    ops_folder_url As String
    documentation_folder As String
    documentation_file As String
End Type

Public Type QualityControlSettings
    populate_audit_checklist As Boolean
    populate_qc_checklist As Boolean
    audit_checklist As CheckList
    qc_checklist As CheckList
    mail_address_suffixes As Variant
    audit_checklist_title As String
    qc_checklist_title As String
End Type

Public Type ErrorSettings
    error_form_width As Long
    error_form_height As Long
    show_web_form As Boolean
    error_file As String
    error_section As String
End Type

Public Type TestingSettings
    test_folder As String
    test_log_folder As String
    test_file_folder As String
    test_waterfall_folder As String
End Type

Public Type MacroSettings
    UI As UserInterfaceSettings
    Import As ImportSettings
    Filter As FilterSettings
    DNA As DoNotAggSettings
    contracts As ContractsQuerySettings
    migration As MigrationSettings
    Stats As StatsSettings
    mapping As MappingSettings
    OneDrive As OneDriveSettings
    QC As QualityControlSettings
    HOME As HomeTabSettings
    errors As ErrorSettings
    OptIn As OptInSettings
End Type

Sub define_macro_settings()
    define_HOME_settings
    define_UI_settings
    define_import_settings
    define_filter_settings
    define_dna_settings
    define_contracts_settings
    define_stats_settings
    define_mapping_settings
    define_onedrive_settings
    define_QC_settings
    define_error_settings
End Sub

Sub define_HOME_settings()
    With S.HOME
        .peer_review_checklist_range = "M6:N14"
        .version_location = "B2"
        .community_name_location = "B3"
        .mail_type_location = "C4"
        .edc_location = "C5"
        .user_location = "C6"
        .step_number_location = "C7"
        .contract_location = "C8"
        .oo_date_location = "C9"
        .info_range = "C3:C9"
        .large_community_limit = 500
        .add_cycle_pivots = True
        .filter_waterfall_location = "B11"
        .mapping_waterfall_location = "E19"
        .mapping_waterfall_name = "Geocoding"
        .mapping_waterfall_caption = "Geocoding"
        .audit_checklist_location = "Q17"
        .qc_checklist_location = "Q2"
        .notes_location = "I24"
        .file_log_location = "I19"
        .file_log_range = "I20:O22"
        .cycle_pivot_name = "Eligible_Counts"
        .cycle_pivot_caption = "Eligible Cycles"
        .cycle_pivot_location = "V6"
        .renewal_drop_count_location = "W2"
    End With
End Sub

Sub define_UI_settings()
    With S.UI
        .mail_type_labels = Array("New Community", "Renewal", "Renewal Only", "Sweep")
        .mail_type_items = Array("NEW", "REN", "REN_ONLY", "SWP")
        .mail_type_images = Array("ResourcePoolRefresh", "AddOrRemoveAttendees", "ResourcePoolMenu", "ResourcesAddMenu")
        .EDC_labels = Array("AES", "CEI", "CS", "DUKE", "OE", "OP", "TE", "AM", "COM")
        .EDC_items = Array("AES", "CEI", "CS", "DUKE", "OE", "OP", "TE", "AM", "COM")
        .EDC_images = Array("OH_png", "OH_png", "OH_png", "OH_png", "OH_png", "OH_png", "OH_png", "IL_png", "IL_png")
    End With
    imported_gagg = False
    imported_active = False
    imported_supplier = False
End Sub

Sub define_QC_settings()
    With S.QC
        .populate_audit_checklist = True
        .populate_qc_checklist = True
        .audit_checklist_title = "Controls Checklist"
        .qc_checklist_title = "QC Checklist"
        .audit_checklist = new_checklist("AUDIT")
        .qc_checklist = new_checklist("FILTER")
        .audit_checklist.data_range = "R19:R29"
        .qc_checklist.data_range = "R4:R14"
        .mail_address_suffixes = Array(" ST", " RD", " DR", " LN", " AVE", " BLVD", " HWY", _
                                        " PKWY", " CT", " CIR", " PL", " TER", " WAY", _
                                        " LOOP", " TRCE", " CTR")
    End With
End Sub

Sub define_import_settings()
    With S.Import
        .max_copy_size = 5000000
        .max_csv_cols = 120
        .trim_sheets = True
        .FE_address_replace = True
    End With
End Sub

Sub define_filter_settings()
    With S.Filter
        .annualize_usage = True
        .highlight_mismatch_cells = True
        .highlight_mismatch_color = C.YELLOW
        .AES_use_both_usage = True
        .remove_arrears = True
    End With
End Sub

Sub define_dna_settings()
    With S.DNA
        .file_name = "PUCO - Do Not Aggregate List (MM-DD-YY).xlsx"
        .folder_name = "PUCO Do Not Aggregate (DNA) List"
        .max_file_age = 30
        .sheet_name = "PUCO - DNA"
        .wildcard_length = 12
        .dna_address_col = 5
        .dna_name_col = 4
        .results_layout = Array("Account Number", "Name", "DNA Name", "Address", "DNA Address", "DNA EDC", "Date Added", "DNA List Pull", "Match Type", "Match Source", "Actual Match")
        .include_wildcard_search = True
        If .include_wildcard_search Then Call array_insert_col(.results_layout, 5, "Wildcard Search")
        .account_match_col = UBound(.results_layout) - 1
        .result_col = UBound(.results_layout) + 1
        .auto_populate_account_match = True
    End With
End Sub

Sub define_contracts_settings()
    With S.contracts
        .hide_snowflake_query = True
        .remove_suffix_for_xdupx = False
        .xdupx_guess_wildcard_length = 15
        .Levenshtein_match_len = 2
        .xdupx_header = "XDUPX"
        .LP_cust_name = "CUSTOMERNAME"
        .sas_id_row_header = "SUBACCOUNTSERVICEROWID"
        .status_header = "STATUS"
        .status_reason_header = "STATUSREASON"
        .muniagg_status_header = "MUNIAGG_STATUS"
        .contract_header = "EXTERNALCONTRACTID"
        .intent_contract_header = "INTENT_CONTRACT"
    End With
End Sub

Sub define_migration_settings()
    With S.migration
        .SF_migration_folder = "\SF"
        .EH_migration_folder = "\EH"
        .migration_log_file = "migration_log_2.xlsx"
        .migration_log_sheet = "migration_log"
    End With
End Sub

Sub define_stats_settings()
    With S.Stats
        .show_stats_tab = False
    End With
End Sub

Sub define_mapping_settings()
    With S.mapping
        .mapping_db_file_name = "Mapping Database.accdb"
        .file_source = "Mapping"
        .mapped_community = "DISPOSITIONED COMMUNITY"
        .notes_col = "NOTES"
        .mapping_col = "MAPS IN (Y/N)"
        .map_this_sheet = "Map This"
        .mapping_tool_file = "Mapping Tool (*).xlsm"
        .no_results_label = "NO RESULT (Y)"
        .map_out_limit = 5
        .map_out_retained_limit = 5
        .no_result_limit = 5
        .no_result_searchable_limit = 5
        .no_result_unsearchable_limit = 5
    End With
End Sub

Sub define_opt_in_settings()
    With S.OptIn
        .make_opt_in = True
        .opt_in_keep_dna = True
        .opt_in_keep_mapped_out = False
        .opt_in_num_cols = 10
        ReDim .opt_in_columns(1 To .opt_in_num_cols)
        .opt_in_columns(1) = F.columns.account_number
        .opt_in_columns(2) = F.columns.customer_name
        .opt_in_columns(3) = F.columns.mail_address
        .opt_in_columns(4) = F.columns.mail_city
        .opt_in_columns(5) = F.columns.mail_state
        .opt_in_columns(6) = F.columns.mail_zip
        .opt_in_columns(7) = F.columns.service_address
        .opt_in_columns(8) = F.columns.service_city
        .opt_in_columns(9) = F.columns.service_state
        .opt_in_columns(10) = F.columns.service_zip
    End With
End Sub

Sub define_onedrive_settings()
    With S.OneDrive
        .parent_folder = "\OneDrive - Vistra Corp"
        .mailings_folder = "\*Mailings"
        .list_management_folder = "\(6) List Management"
        .dna_folder = "\(6) List Management\(4) PUCO Do Not Aggregate (DNA) List"
        .migration_folder = "\(15) Energy Harbor Migration\(9) Contracts By Offer"
        .mapping_db_folder = "\(6) List Management\(2) Macro Testing\(7) Mapping Database"
        .ops_folder_url = "https://txu.sharepoint.com/sites/Muni-Agg/Shared Documents/(1) Operations/"
        .documentation_folder = "\(6) List Management\(2) Macro Testing\(4) Documentation"
        .documentation_file = "macro_help.htm"
    End With
End Sub

Sub define_error_settings()
    With S.errors
        .error_file = "macro help\macro_help.htm"
        .error_section = "#errors"
        .error_form_height = 400
        .error_form_width = 600
    End With
End Sub
