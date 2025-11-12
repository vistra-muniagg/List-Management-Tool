Public Type Stat
    name As String
    value As String
End Type

Public Type InfoStats
    category As String
    stat_list() As Variant
    version As Stat
    revision_date As Stat
    waterfall_name As Stat
    mail_type As Stat
    EDC As Stat
    contract_id As Stat
    opt_out_date As Stat
    analyst As Stat
    peer_reviewer As Stat
    peer_review_date As Stat
    log_name As Stat
End Type

Public Type FileInfo
    file_name As Stat
    format As Stat
    initial_count As Stat
    deduped_count As Stat
    dupes As Stat
    modified_date As Stat
    timestamp As Stat
End Type

Public Type FileStats
    category As String
    stat_list() As FileInfo
    utility_files() As FileInfo
    active_list As FileInfo
    supplier_list As FileInfo
    dna_list As FileInfo
    contracts_query As FileInfo
    migration_file As FileInfo
    mapping_file As FileInfo
End Type

Public Type FilterStats
    category As String
    stat_list() As Variant
    shoppers As Stat
    renewal_shoppers As Stat
    net_metering As Stat
    pipp As Stat
    mercantile As Stat
    bgs_hold As Stat
    rtp As Stat
    hourly As Stat
    free_service As Stat
    community_solar As Stat
    high_usage As Stat
    arrears As Stat
    spokane As Stat
    non_oh_commercial As Stat
    national_chains As Stat
End Type

Public Type QualityControlStats
    'what are these??
    category As String
    stat_list() As Variant
    qc_setting_1 As Boolean
End Type

Public Type AddressStats
    category As String
    stat_list() As Variant
    fe_replaced As Stat
    ren_service_replaced As Stat
    ren_mail_replaced As Stat
    ren_name_replaced As Stat
End Type

Public Type DoNotAggStats
    category As String
    stat_list() As Variant
    file_age As Stat
    account_matches As Stat
    address_matches As Stat
    total_potential_matches As Stat
    actual_account_matches As Stat
    actual_address_matches As Stat
    actual_matches As Stat
    false_match_char_len As Stat
    actual_match_char_len As Stat
    total_address_char_match_len As Stat
    false_positives As Stat
    guess_correct As Stat
    guess_wrong_match As Stat
    guess_wrong_false_match As Stat
End Type

Public Type ContractsQueryStats
    category As String
    stat_list() As Variant
    existing_contract As Stat
    Active As Stat
    inctive As Stat
    other As Stat
    xdupx_count As Stat
End Type

Public Type MigrationStats
    category As String
    stat_list() As Variant
    account_matches As Stat
End Type

Public Type MappingStats
    category As String
    stat_list() As Variant
    unique_mapped_count As Stat
    total_count As Stat
    time As Stat
    map_in As Stat
    map_out As Stat
    no_result As Stat
    map_out_retained As Stat
    no_results_exceeds_limit As Boolean
    maps_out_exceeds_limit As Boolean
    eligible_before_mapping As Stat
    eligible_after_mapping As Stat
End Type

Public Type UploadFileStats
    category As String
    stat_list() As Variant
    rate_codes_replaced As Stat
    mail_service_mismatch_count As Stat
    mail_service_mismatch_pct As Stat
    exceeds_mismatch_limit As Boolean
End Type

Public Type ExportFileStats
    category As String
    stat_list() As Variant
    LP_new_stats As UploadFileStats
    LP_ren_stats As UploadFileStats
    opt_in_eligible As Stat
    bb_count As Stat
    nm_count As Stat
End Type

Public Type Statistcs
    info_stats As InfoStats
    file_stats As FileStats
    filter_stats As FilterStats
    qc_stats As QualityControlStats
    address_stats As AddressStats
    dna_stats As DoNotAggStats
    contracts_stats As ContractsQueryStats
    migration_stats As MigrationStats
    mapping_stats As MappingStats
    upload_file_stats As UploadFileStats
    export_stats As ExportFileStats
End Type

Sub define_stats()
    create_stats_tab
    format_stats_tab
End Sub

Sub create_stats_tab()
    delete_sheet (SN.Stats)
    Set s1 = Sheets.Add(before:=Sheets(Sheets.count))
    s1.name = SN.Stats
    s1.visible = S.Stats.show_stats_tab
End Sub

Sub format_stats_tab()
    With Sheets(SN.Stats)
        .Cells(1, 1) = "Stat Group"
        .Cells(1, 2) = "Stat Name"
        .Cells(1, 3) = "Value"
        reapply_autofilter (.index)
    End With
End Sub

Sub stats_arr_append(ByRef arr() As Variant, value() As String)
    If arr(1).name = "" Then
        arr(1) = value
    Else
        n = UBound(arr)
        ReDim Preserve arr(LBound(arr) To n + 1)
        arr(n + 1) = value
    End If
End Sub

Sub populate_info_stat_list()

    ReDim Stats.info_stats.stat_list(1 To 1)
    
    With Stats.info_stats
        .category = "INFO"
        Call stats_arr_append(.stat_list, .analyst)
        Call stats_arr_append(.stat_list, .contract_id)
        Call stats_arr_append(.stat_list, .EDC)
        Call stats_arr_append(.stat_list, .log_name)
        Call stats_arr_append(.stat_list, .mail_type)
        Call stats_arr_append(.stat_list, .opt_out_date)
        Call stats_arr_append(.stat_list, .peer_review_date)
        Call stats_arr_append(.stat_list, .peer_reviewer)
        Call stats_arr_append(.stat_list, .revision_date)
        Call stats_arr_append(.stat_list, .version)
        Call stats_arr_append(.stat_list, .waterfall_name)
    End With
    
End Sub

Sub populate_file_stat_list()
    
    ReDim Stats.file_stats.stat_list(1 To 1)
    
    With Stats.file_stats
        .category = "FILES"
        Call stats_arr_append(.stat_list, .active_list)
        Call stats_arr_append(.stat_list, .contracts_query)
        Call stats_arr_append(.stat_list, .dna_list)
        Call stats_arr_append(.stat_list, .mapping_file)
        Call stats_arr_append(.stat_list, .migration_file)
        Call stats_arr_append(.stat_list, .supplier_list)
        Call stats_arr_append(.stat_list, .utility_files)
    End With
    
End Sub

Sub populate_address_stat_list()
    
    ReDim Stats.address_stats.stat_list(1 To 1)
    
    With Stats.address_stats
        .category = "ADDRESS"
        Call stats_arr_append(.stat_list, "")
    End With
    
End Sub

Sub populate_qc_stat_list()
    
    ReDim Stats.qc_stats.stat_list(1 To 1)
    
    With Stats.qc_stats
        .category = "QC"
        Call stats_arr_append(.stat_list, .qc_setting_1)
    End With
    
End Sub

Sub populate_mapping_stat_list()
    
    ReDim Stats.mapping_stats.stat_list(1 To 1)
    
    With Stats.mapping_stats
        .category = "MAPPING"
        Call stats_arr_append(.stat_list, .eligible_after_mapping)
        Call stats_arr_append(.stat_list, .eligible_before_mapping)
        Call stats_arr_append(.stat_list, .map_in)
        Call stats_arr_append(.stat_list, .map_out)
        Call stats_arr_append(.stat_list, .map_out_retained)
        Call stats_arr_append(.stat_list, .maps_out_exceeds_limit)
        Call stats_arr_append(.stat_list, .no_result)
        Call stats_arr_append(.stat_list, .no_results_exceeds_limit)
        Call stats_arr_append(.stat_list, .time)
        Call stats_arr_append(.stat_list, .total_count)
        Call stats_arr_append(.stat_list, .unique_mapped_count)
    End With
    
End Sub

Sub populate_filter_stat_list()
    
    ReDim Stats.filter_stats.stat_list(1 To 1)
    
    With Stats.filter_stats
        .category = "FILTER"
        Call stats_arr_append(.stat_list, .arrears)
        Call stats_arr_append(.stat_list, .bgs_hold)
        Call stats_arr_append(.stat_list, .community_solar)
        Call stats_arr_append(.stat_list, .free_service)
        Call stats_arr_append(.stat_list, .high_usage)
        Call stats_arr_append(.stat_list, .hourly)
        Call stats_arr_append(.stat_list, .mercantile)
        Call stats_arr_append(.stat_list, .national_chains)
        Call stats_arr_append(.stat_list, .net_metering)
        Call stats_arr_append(.stat_list, .non_oh_commercial)
        Call stats_arr_append(.stat_list, .pipp)
        Call stats_arr_append(.stat_list, .renewal_shoppers)
        Call stats_arr_append(.stat_list, .rtp)
        Call stats_arr_append(.stat_list, .shoppers)
        Call stats_arr_append(.stat_list, .spokane)
    End With
    
End Sub

Sub populate_dna_stat_list()
    
    ReDim Stats.dna_stats.stat_list(1 To 1)
    
    With Stats.dna_stats
        .category = "DNA"
        Call stats_arr_append(.stat_list, .file_age)
        Call stats_arr_append(.stat_list, .account_matches)
        Call stats_arr_append(.stat_list, .actual_account_matches)
        Call stats_arr_append(.stat_list, .actual_address_matches)
        Call stats_arr_append(.stat_list, .actual_match_char_len)
        Call stats_arr_append(.stat_list, .actual_matches)
        Call stats_arr_append(.stat_list, .address_matches)
        Call stats_arr_append(.stat_list, .false_match_char_len)
        Call stats_arr_append(.stat_list, .false_positives)
        Call stats_arr_append(.stat_list, .guess_correct)
        Call stats_arr_append(.stat_list, .guess_wrong_false_match)
        Call stats_arr_append(.stat_list, .guess_wrong_match)
        Call stats_arr_append(.stat_list, .total_address_char_match_len)
        Call stats_arr_append(.stat_list, .total_potential_matches)
    End With
End Sub

Sub populate_contracts_stat_list()
    
    ReDim Stats.contracts_stats.stat_list(1 To 1)
    
    With Stats.contracts_stats
        .category = "CONTRACTS"
        Call stats_arr_append(.stat_list, .Active)
        Call stats_arr_append(.stat_list, .existing_contract)
        Call stats_arr_append(.stat_list, .inctive)
        Call stats_arr_append(.stat_list, .other)
        Call stats_arr_append(.stat_list, .xdupx_count)
    End With
    
End Sub

Sub populate_migration_stat_list()
    
    ReDim Stats.migration_stats.stat_list(1 To 1)
    
    With Stats.migration_stats
        .category = "MIGRATION"
        Call stats_arr_append(.stat_list, .account_matches)
    End With
    
End Sub

Sub populate_upload_file_stat_list()
    
    ReDim Stats.upload_file_stats.stat_list(1 To 1)
    
    With Stats.upload_file_stats
        .category = "UPLOAD"
        Call stats_arr_append(.stat_list, .exceeds_mismatch_limit)
        Call stats_arr_append(.stat_list, .mail_service_mismatch_count)
        Call stats_arr_append(.stat_list, .mail_service_mismatch_pct)
        Call stats_arr_append(.stat_list, .rate_codes_replaced)
    End With
    
End Sub

Sub populate_export_stat_list()
    
    ReDim Stats.export_stats.stat_list(1 To 1)
    
    With Stats.export_stats
        .category = "EXPORT"
        Call stats_arr_append(.stat_list, .Active)
        Call stats_arr_append(.stat_list, .existing_contract)
        Call stats_arr_append(.stat_list, .inctive)
        Call stats_arr_append(.stat_list, .other)
        Call stats_arr_append(.stat_list, .xdupx_count)
    End With
    
End Sub

Sub file_arr_append(ByRef arr() As FileInfo, value As FileInfo)
    If arr(1).file_name.value = "" Then
        arr(1) = value
    Else
        n = UBound(arr)
        ReDim Preserve arr(LBound(arr) To n + 1)
        arr(n + 1) = value
    End If
    'call file_arr_append(STATS.file_stats.utility_files,
End Sub
