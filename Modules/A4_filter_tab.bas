Public filter_tab_initialized As Boolean

Public Type FilterStatus
    eligible_new_status As String
    eligible_ren_status As String
    ineligible_new_status As String
    ineligible_ren_status As String
End Type

Public Type ContractsStatus
    eligible_xdupx As String
    eligible_inactive As String
    ineligible_active As String
    ineligible_previous_mail As String
End Type

Public Type MappingStatus
    ineligible_new_status As String
    ineligible_ren_status As String
    mapped_out_retained_status As String
    mapped_out_label As String
    maps_in_label As String
    no_results_label As String
    mapped_out_retained_label As String
End Type

Public Type FilterStatuses
    eligible As FilterStatus
    renewal As FilterStatus
    dupe As FilterStatus
    mismatch As FilterStatus
    shopper As FilterStatus
    pipp As FilterStatus
    mercantile As FilterStatus
    rtp As FilterStatus
    bgs_hold As FilterStatus
    free_service As FilterStatus
    hourly_pricing As FilterStatus
    community_solar As FilterStatus
    arrears As FilterStatus
    usage As FilterStatus
    national_chain As FilterStatus
    dna_OH As FilterStatus
    dna_IL As FilterStatus
    duke_sibling_account As FilterStatus
    migration As FilterStatus
    contracts As ContractsStatus
    mapping As MappingStatus
End Type

Public Type ColumnHeader
    header As String
    index As Long
    cell_color As CellColors
    data_type As String
    data_subtype As String
    default_value As String
    possible_values() As Variant
    source_col As Variant
    condition_value As String
    values As Variant
    column_group As String
    active_col As ActiveColumnHeader
    default_mismatch_value As String
    apply_to_active As Boolean
    data_format As String
End Type

Public Type FilterTabColumns
    account_number As ColumnHeader
    status As ColumnHeader
    eligible_opt_out As ColumnHeader
    mail_category As ColumnHeader
    active_in_LP As ColumnHeader
    mismatch As ColumnHeader
    sas_id As ColumnHeader
    address_source As ColumnHeader
    address_overwritten As ColumnHeader
    mapping_result As ColumnHeader
    before_mapping_eligible As ColumnHeader
    community_mapped_into As ColumnHeader
    mapping_notes As ColumnHeader
    shopping As ColumnHeader
    pipp As ColumnHeader
    mercantile As ColumnHeader
    national_chains As ColumnHeader
    usage_months As ColumnHeader
    actual_usage As ColumnHeader
    estimated_usage As ColumnHeader
    hourly_pricing As ColumnHeader
    rtp As ColumnHeader
    bgs_hold As ColumnHeader
    community_solar As ColumnHeader
    free_service As ColumnHeader
    arrears As ColumnHeader
    do_not_agg As ColumnHeader
    opt_in_eligible As ColumnHeader
    lp_contracts_query As ColumnHeader
    migration_query As ColumnHeader
    customer_name As ColumnHeader
    service_address As ColumnHeader
    service_city As ColumnHeader
    service_state As ColumnHeader
    service_zip As ColumnHeader
    mail_address As ColumnHeader
    mail_city As ColumnHeader
    mail_state As ColumnHeader
    mail_zip As ColumnHeader
    customer_class As ColumnHeader
    read_cycle As ColumnHeader
    phone As ColumnHeader
    email As ColumnHeader
    source_file As ColumnHeader
End Type

Public Type ColumnGroup
    columns() As ColumnHeader
    name As String
    color As CellColors
End Type

Public Type ColumnGroups
    static_cols As ColumnGroup
    OH_filters As ColumnGroup
    IL_filters As ColumnGroup
End Type

Public Type FilterTab
    columns As FilterTabColumns
    order_array() As ColumnHeader
    'group_columns As ColumnGroups
End Type

Sub define_filter_tab()
    'define_statuses
    define_filter_tab_columns
    define_filter_tab_order
    'define_columns_groups
    filter_tab_initialized = True
End Sub

Sub define_filter_tab_columns()
    With F.columns.account_number
        .header = "ACCOUNT NUMBER"
        .cell_color = C.GRAY_1
        .data_type = "Literal"
        .source_col = EDC.account
        .condition_value = ""
        .active_col = A.columns.account_number
        .data_format = "@"
        .column_group = "LP"
    End With
    With F.columns.status
        .header = "STATUS"
        .cell_color = C.GRAY_1
        .data_type = "Generated"
        .data_subtype = "Status"
        .default_value = "Eligible - New Customer"
        .source_col = ""
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = FS.mismatch.eligible_ren_status
        .column_group = "LP"
    End With
    With F.columns.eligible_opt_out
        .header = "ELIGIBLE TO MAIL"
        .cell_color = C.PINK
        .data_type = "Boolean"
        .default_value = "Y"
        .possible_values = Array("Y", "N")
        .source_col = ""
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "Y"
        .column_group = "LP"
    End With
    With F.columns.mail_category
        .header = "MAIL CATEGORY"
        .cell_color = C.PINK
        .data_type = "Generated"
        .data_subtype = "Mail Category"
        .possible_values = Array("NEW", "REN")
        .default_value = .possible_values(0)
        .source_col = ""
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = .possible_values(1)
        .column_group = "LP"
    End With
    With F.columns.active_in_LP
        .header = "ON ACTIVE LIST"
        .cell_color = C.PINK
        .data_type = "Boolean"
        .default_value = "N"
        .source_col = ""
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "Y"
    End With
    With F.columns.mismatch
        .header = "ACTIVE LIST MISMATCH"
        .cell_color = C.PINK
        .data_type = "Boolean"
        .default_value = "N"
        .source_col = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "Y"
    End With
    With F.columns.sas_id
        .header = "SUBACCOUNTSERVICEID"
        .cell_color = C.PINK
        .data_type = "Active"
        .default_value = "-"
        .source_col = ""
        .active_col = A.columns.sas_id
        .default_mismatch_value = "-"
    End With
    With F.columns.address_source
        .header = "ADDRESS DATA SOURCE"
        .cell_color = C.PINK
        .data_type = "Generated"
        .data_subtype = "Address Source"
        .possible_values = Array("GAGG", "VISTRA")
        .default_value = .possible_values(0)
        .source_col = ""
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = .possible_values(1)
    End With
    With F.columns.mapping_result
        .header = "MAPPING RESULT"
        .cell_color = C.BLUE
        .data_type = "Generated"
        .data_subtype = "Mapping"
        .possible_values = Array("Maps In", "Maps Out", "Maps In (No Result)", "Maps Out - Retained")
        .default_value = .possible_values(0)
        .source_col = ""
        .condition_value = ""
        .column_group = "Mapping Data"
        .active_col = A.columns.EMPTYCOL
        .apply_to_active = False
        .default_mismatch_value = "Maps In"
    End With
    With F.columns.community_mapped_into
        .header = "COMMUNITY MAPPED INTO"
        .cell_color = C.BLUE
        .data_type = "Generated"
        .default_value = "-"
        .source_col = ""
        .condition_value = ""
        .column_group = "Mapping Data"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "-"
    End With
    With F.columns.before_mapping_eligible
        .header = "ELIGIBILITY STATUS BEFORE MAPPING"
        .cell_color = C.BLUE
        .data_type = "Boolean"
        .default_value = "Y"
        .source_col = ""
        .condition_value = ""
        .column_group = "Mapping Data"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "Y"
    End With
    With F.columns.mapping_notes
        .header = "MAPPING NOTES"
        .cell_color = C.BLUE
        .data_type = "Generated"
        .default_value = "-"
        .source_col = ""
        .condition_value = ""
        .column_group = "Mapping Data"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "-"
    End With
    With F.columns.shopping
        .header = "SHOPPING"
        .cell_color = C.GOLD
        .data_type = "Boolean"
        .default_value = "-"
        .source_col = EDC.shopper
        .condition_value = EDC.shopper_yes
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
    End With
    With F.columns.pipp
        .header = "PIPP"
        .cell_color = C.ORANGE
        .data_type = "Boolean"
        .default_value = "-"
        .source_col = EDC.pipp
        .condition_value = EDC.pipp_yes
        .column_group = "OH Filters"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
    End With
    With F.columns.mercantile
        .header = "MERCANTILE"
        .cell_color = C.ORANGE
        .data_type = "Boolean"
        .default_value = "-"
        .source_col = EDC.mercantile
        .condition_value = EDC.mercantile_yes
        .column_group = "OH Filters"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
    End With
    With F.columns.usage_months
        .header = "USAGE MONTHS"
        .cell_color = C.GOLD
        .data_type = "Calculated"
        .data_subtype = "Usage"
        .default_value = "-"
        .source_col = EDC.usage
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "0"
    End With
    With F.columns.actual_usage
        .header = "ACTUAL USAGE"
        .cell_color = C.GOLD
        .data_type = "-"
        .data_subtype = "-"
        .default_value = "-"
        .source_col = "-"
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "0"
        .data_format = "0"
    End With
    With F.columns.estimated_usage
        .header = "ESTIMATED ANNUAL USAGE"
        .cell_color = C.GOLD
        .data_type = "-"
        .data_subtype = "-"
        .default_value = "-"
        .source_col = "-"
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "0"
        .data_format = "0"
    End With
    With F.columns.hourly_pricing
        .header = "HOURLY PRICING"
        .cell_color = C.PURPLE
        .data_type = "Boolean"
        .default_value = "-"
        .source_col = EDC.hourly_pricing
        .condition_value = EDC.hourly_pricing_yes
        .column_group = "IL Filters"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
    End With
    With F.columns.rtp
        .header = "RTP"
        .cell_color = C.PURPLE
        .data_type = "Boolean"
        .default_value = "-"
        .source_col = EDC.rtp
        .condition_value = EDC.rtp_yes
        .column_group = "IL Filters"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
    End With
    With F.columns.bgs_hold
        .header = "BGS HOLD"
        .cell_color = C.PURPLE
        .data_type = "Boolean"
        .default_value = "-"
        .source_col = EDC.bgs
        .condition_value = EDC.bgs_yes
        .column_group = "IL Filters"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
    End With
    With F.columns.community_solar
        .header = "COMMUNITY SOLAR"
        .cell_color = C.PURPLE
        .data_type = "Boolean"
        .default_value = "-"
        .source_col = EDC.solar
        .condition_value = EDC.solar_yes
        .column_group = "IL Filters"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
    End With
    With F.columns.free_service
        .header = "FREE SERVICE"
        .cell_color = C.PURPLE
        .data_type = "Boolean"
        .default_value = "-"
        .source_col = EDC.free_service
        .condition_value = EDC.free_service_yes
        .column_group = "IL Filters"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
    End With
    With F.columns.arrears
        .header = "ARREARS"
        .cell_color = C.GOLD
        .data_type = "Boolean"
        .data_subtype = "Arrears"
        .default_value = "-"
        .source_col = EDC.arrears
        .condition_value = EDC.arrears_yes
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = False
        If EDC.ruleset_name = "AES" Then .data_type = "Calculated"
    End With
    With F.columns.do_not_agg
        .header = "DO NOT AGG"
        .cell_color = C.GOLD
        .data_type = "Boolean"
        .default_value = "-"
        .column_group = "OH Filters"
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "N"
        .apply_to_active = True
    End With
    With F.columns.national_chains
        .header = "NATIONAL CHAINS"
        .cell_color = C.GOLD
        .data_type = "Generated"
        .data_subtype = "National Chains"
        .default_value = "-"
        .source_col = ""
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .apply_to_active = False
        .default_mismatch_value = "N"
    End With
    With F.columns.opt_in_eligible
        .header = "OPT IN ELIGIBLE"
        .cell_color = C.GOLD
        .data_type = "Calculated"
        .data_subtype = "Opt In"
        .default_value = "N"
        .source_col = ""
        .condition_value = ""
        .default_mismatch_value = "N"
        .active_col = A.columns.EMPTYCOL
    End With
    With F.columns.lp_contracts_query
        .header = "LP CONTRACTS QUERY STATUS"
        .cell_color = C.BLUE_2
        .data_type = "Generated"
        .default_value = "-"
        .source_col = ""
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "-"
    End With
    With F.columns.migration_query
        .header = "MIGRATION CONTRACTS QUERY RESULT"
        .cell_color = C.BLUE_2
        .data_type = "Generated"
        .default_value = "-"
        .source_col = ""
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "-"
    End With
    With F.columns.customer_name
        .header = "CUSTOMER NAME"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Customer Name"
        .default_value = ""
        .source_col = EDC.customer_name
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.customer_name
    End With
    With F.columns.service_address
        .header = "SERVICE ADDRESS"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Service Address"
        .default_value = ""
        .source_col = ""
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.service_address
    End With
    With F.columns.service_city
        .header = "SERVICE CITY"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Service City"
        .default_value = ""
        .source_col = ""
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.service_city
    End With
    With F.columns.service_state
        .header = "SERVICE STATE"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Service State"
        .default_value = ""
        .source_col = ""
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.service_state
    End With
    With F.columns.service_zip
        .header = "SERVICE ZIP"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Service Zip"
        .default_value = ""
        .source_col = ""
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.service_zip
        .data_format = "@"
    End With
    With F.columns.mail_address
        .header = "MAIL ADDRESS"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Mail Address"
        .default_value = ""
        .source_col = ""
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.mail_address
    End With
    With F.columns.mail_city
        .header = "MAIL CITY"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Mail City"
        .default_value = ""
        .source_col = ""
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.mail_city
    End With
    With F.columns.mail_state
        .header = "MAIL STATE"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Mail State"
        .default_value = ""
        .source_col = ""
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.mail_state
    End With
    With F.columns.mail_zip
        .header = "MAIL ZIP"
        .cell_color = C.GREEN_1
        .data_type = "Generated"
        .data_subtype = "Mail Zip"
        .default_value = ""
        .source_col = ""
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.mail_zip
        .data_format = "@"
    End With
    With F.columns.phone
        .header = "PHONE"
        .cell_color = C.GREEN_1
        .data_type = "Literal"
        .default_value = ""
        .source_col = EDC.phone
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.phone
        .data_format = "0000000000"
    End With
    With F.columns.email
        .header = "EMAIL"
        .cell_color = C.GREEN_1
        .data_type = "Literal"
        .default_value = ""
        .source_col = EDC.email
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.email
    End With
    With F.columns.customer_class
        .header = "CLASS (RES/COMM)"
        .cell_color = C.GREEN_1
        .data_type = "Calculated"
        .data_subtype = "Class"
        .possible_values = Array("RESIDENTIAL", "COMMERCIAL")
        .default_value = .possible_values(0)
        .source_col = EDC.rate_code
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.customer_class
    End With
    With F.columns.read_cycle
        .header = "READ CYCLE"
        .cell_color = C.GREEN_1
        .data_type = "Literal"
        .default_value = EDC.default_read_cycle
        .source_col = EDC.read_cycle
        .condition_value = ""
        .column_group = "LP"
        .active_col = A.columns.read_cycle
        .default_mismatch_value = "1"
    End With
    With F.columns.source_file
        .header = "SOURCE FILE"
        .cell_color = C.GRAY_1
        .data_type = "Literal"
        .default_value = "add file name to gagg+active list"
        .source_col = "SOURCE FILE"
        .condition_value = ""
        .active_col = A.columns.EMPTYCOL
        .default_mismatch_value = "add active list file name"
    End With
End Sub

Sub define_filter_tab_order()

    ReDim F.order_array(1 To 1)
    
    With F.columns
    
        Call filter_arr_append(F.order_array, .account_number)
        Call filter_arr_append(F.order_array, .status)
        Call filter_arr_append(F.order_array, .eligible_opt_out)
        Call filter_arr_append(F.order_array, .mail_category)
        Call filter_arr_append(F.order_array, .active_in_LP)
        Call filter_arr_append(F.order_array, .mismatch)
        Call filter_arr_append(F.order_array, .sas_id)
        Call filter_arr_append(F.order_array, .address_source)
        'Call filter_arr_append(F.order_array, .address_overwritten)
        Call filter_arr_append(F.order_array, .mapping_result)
        Call filter_arr_append(F.order_array, .community_mapped_into)
        Call filter_arr_append(F.order_array, .before_mapping_eligible)
        Call filter_arr_append(F.order_array, .mapping_notes)
        Call filter_arr_append(F.order_array, .hourly_pricing)
        Call filter_arr_append(F.order_array, .rtp)
        Call filter_arr_append(F.order_array, .bgs_hold)
        Call filter_arr_append(F.order_array, .community_solar)
        Call filter_arr_append(F.order_array, .free_service)
        Call filter_arr_append(F.order_array, .pipp)
        Call filter_arr_append(F.order_array, .mercantile)
        Call filter_arr_append(F.order_array, .usage_months)
        Call filter_arr_append(F.order_array, .actual_usage)
        Call filter_arr_append(F.order_array, .estimated_usage)
        Call filter_arr_append(F.order_array, .shopping)
        Call filter_arr_append(F.order_array, .arrears)
        Call filter_arr_append(F.order_array, .national_chains)
        Call filter_arr_append(F.order_array, .do_not_agg)
        Call filter_arr_append(F.order_array, .opt_in_eligible)
        Call filter_arr_append(F.order_array, .lp_contracts_query)
        Call filter_arr_append(F.order_array, .migration_query)
        Call filter_arr_append(F.order_array, .customer_name)
        Call filter_arr_append(F.order_array, .service_address)
        Call filter_arr_append(F.order_array, .service_city)
        Call filter_arr_append(F.order_array, .service_state)
        Call filter_arr_append(F.order_array, .service_zip)
        Call filter_arr_append(F.order_array, .mail_address)
        Call filter_arr_append(F.order_array, .mail_city)
        Call filter_arr_append(F.order_array, .mail_state)
        Call filter_arr_append(F.order_array, .mail_zip)
        Call filter_arr_append(F.order_array, .phone)
        Call filter_arr_append(F.order_array, .email)
        Call filter_arr_append(F.order_array, .customer_class)
        Call filter_arr_append(F.order_array, .read_cycle)
        'Call filter_arr_append(F.order_array, .source_file)
    
    End With
    
End Sub

Sub define_statuses()
    With FS.eligible
        .eligible_new_status = "Eligible - New Customer"
        .eligible_ren_status = "Eligible - Renewal Customer"
    End With
    With FS.renewal
        .eligible_ren_status = "Eligible - Renewal Account"
    End With
    With FS.dupe
        .ineligible_new_status = "Ineligible - Duplicate"
        .ineligible_ren_status = "Ineligible - Duplicate"
    End With
    With FS.mismatch
        .eligible_new_status = "Eligible - Supplier Match"
        .eligible_ren_status = "Eligible - Not On Utility List"
        .ineligible_new_status = "inelgible but on supplier list?"
        .ineligible_ren_status = "ineligible but on renewal list?"
    End With
    With FS.shopper
        .ineligible_new_status = "Shopper"
    End With
    With FS.pipp
        .ineligible_new_status = "PIPP"
    End With
    With FS.mercantile
        .ineligible_new_status = "Mercantile (OH)"
    End With
    With FS.rtp
        .ineligible_new_status = "Real Time Pricing (IL)"
    End With
    With FS.bgs_hold
        .ineligible_new_status = "BGS Hold (IL)"
    End With
    With FS.free_service
        .ineligible_new_status = "Free Service (IL)"
    End With
    With FS.hourly_pricing
        .ineligible_new_status = "Hourly Pricing (IL)"
    End With
    With FS.community_solar
        .ineligible_new_status = "Community Solar (IL)"
    End With
    With FS.arrears
        .ineligible_new_status = "In Arrears"
    End With
    With FS.usage
        .ineligible_new_status = "Usage > kWh Limit"
    End With
    With FS.national_chain
        .ineligible_new_status = "National Chain (Spokane)"
    End With
    With FS.dna_OH
        .ineligible_ren_status = "PUCO DNA"
        .ineligible_new_status = "PUCO DNA"
    End With
    With FS.dna_IL
        .ineligible_ren_status = "Do Not Aggregate (IL)"
        .ineligible_new_status = "Do Not Aggregate (IL)"
    End With
    With FS.duke_sibling_account
        .ineligible_ren_status = "Sibling Account (DUKE)"
        .ineligible_new_status = "Sibling Account (DUKE)"
    End With
    With FS.mapping
        .ineligible_new_status = "Maps Out"
        .ineligible_ren_status = "Maps Out"
        .mapped_out_retained_status = "Eligible - Maps Out (Retained)"
        .mapped_out_label = "Maps Out"
        .maps_in_label = "Maps In"
        .no_results_label = "Maps In (No Result)"
        .mapped_out_retained_label = "Maps Out (Retained)"
    End With
    With FS.contracts
        .eligible_inactive = FS.eligible.eligible_new_status
        .eligible_xdupx = "Eligible - Recycled AEP Account"
        .ineligible_active = "LP Status (Active)"
        .ineligible_previous_mail = "LP Status (Opt Out+)"
    End With
    With FS.migration
        .ineligible_new_status = "Ineligible - Account In Legacy System"
    End With
End Sub
