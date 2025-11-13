# List Management Tool Technical Documentation
Generated on 2025-11-13 19:53:49

## Executive Summary

  Executive Summary
  This document consolidates the uploaded BAS-style automation scripts, accompanying process notes, and related artifacts found in the session. It is intended as a developer-ready reference for on-boarding, architecture understanding, and quick-look data access for the remote Python-based extraction workflow.

## Architecture

  - Centralized remote file system (RFS) hosting all modules and artifacts.
  - Python-based data-extraction layer using:
    - PyMuPDF (fitz) for PDF text extraction
    - python-docx for Word content extraction
    - pandas/openpyxl for Excel/CSV data sampling
  - Lightweight Markdown-based documentation surface exported as README.md
  - Module-map driven by file metadata; descriptive prose derived from header comments when present
  - TODOs for missing source-detail derivations and integration points

## Module Map

Module Map
| Module / File | Type | Description | Source Path |
|---|---|---|---|
| A0_settings.bas | .bas | Public Type HomeTabSettings | /mnt/data/A0_settings.bas |
| A1_EDC.bas | .bas | system names validation info gagg list import | /mnt/data/A1_EDC.bas |
| A2_mail_type.bas | .bas | how to get mail type selection? | /mnt/data/A2_mail_type.bas |
| A4_filter_tab.bas | .bas | Public filter_tab_initialized As Boolean | /mnt/data/A4_filter_tab.bas |
| A9_init.bas | .bas | If all_initialized Then Exit Sub progress.start ("Initializing") | /mnt/data/A9_init.bas |
| C1_process_active.bas | .bas | init active_accounts = active_tab.UsedRange.columns(1).value active_accounts = flatten_array(active_accounts) | /mnt/data/C1_process_active.bas |
| C1_process_active.docx | .docx | Module from C1_process_active.docx | /mnt/data/C1_process_active.docx |
| C1_process_active.html | .html | Module from C1_process_active.html | /mnt/data/C1_process_active.html |
| C1_process_active.json | .json | Module from C1_process_active.json | /mnt/data/C1_process_active.json |
| C1_process_active.md | .md | Module from C1_process_active.md | /mnt/data/C1_process_active.md |
| C1_process_active.pdf | .pdf | Module from C1_process_active.pdf | /mnt/data/C1_process_active.pdf |
| C1_process_active.png | .png | Module from C1_process_active.png | /mnt/data/C1_process_active.png |
| C1_process_active.pptx | .pptx | Module from C1_process_active.pptx | /mnt/data/C1_process_active.pptx |
| C1_process_active.txt | .txt | Module from C1_process_active.txt | /mnt/data/C1_process_active.txt |
| C1_process_active.xlsx | .xlsx | Module from C1_process_active.xlsx | /mnt/data/C1_process_active.xlsx |
| D1-D6_workflow.md | .md | Module from D1-D6_workflow.md | /mnt/data/D1-D6_workflow.md |
| D1_filter.bas | .bas | pipp state rules usage | /mnt/data/D1_filter.bas |
| D2_dna.bas | .bas | Sub dna_test() | /mnt/data/D2_dna.bas |
| D3_contracts.bas | .bas | init get_contracts_file | /mnt/data/D3_contracts.bas |
| D4_migration.bas | .bas | get current contract | /mnt/data/D4_migration.bas |
| D5_mapping.bas | .bas | init | /mnt/data/D5_mapping.bas |
| D6_misc.bas | .bas | premise_mismatch_accounts | /mnt/data/D6_misc.bas |
| Globals_and_Init.md | .md | Module from Globals_and_Init.md | /mnt/data/Globals_and_Init.md |


## Detailed Module References

### Module: A0_settings.bas

- File: /mnt/data/A0_settings.bas
- Type: .bas
- Description: Public Type HomeTabSettings

- Preview snippet:

```
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
```

### Module: A1_EDC.bas

- File: /mnt/data/A1_EDC.bas
- Type: .bas
- Description: system names validation info gagg list import

- Preview snippet:

```
Public Type UtilitySettings
    'system names
    name As String
    display_name As String
    ruleset_name As String
    full_name As String
    CRM_name As String
    SF_name As String
    LP_name As String
    migration_name As String
    DNA_name As String
    mapping_name As String
    'validation info
    state As String
    file_formats As Variant
    account_number_length As Integer
    leading_pattern As String
    zeros_to_add As Integer
    valid_codes As Variant
    res_codes As Variant
    valid_cycles As Variant
    default_read_cycle As Integer
    usage_limit As String
    usage_multiple As Integer
    'gagg list import
    multiselect As Boolean
    import_all_gagg_sheets As Boolean
    'y/n column headers
    shopper As String
    shopper_yes As String
    pipp As String
    pipp_yes As String
    net_meter As String
    net_meter_yes As String
    mercantile As String
    mercantile_yes As String
    arrears As String
    arrears_yes As String
    budget_bill As Str
```

### Module: A2_mail_type.bas

- File: /mnt/data/A2_mail_type.bas
- Type: .bas
- Description: how to get mail type selection?

- Preview snippet:

```
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
    ElseIf mail_type = S.UI.mail_type_items(2) T
```

### Module: A4_filter_tab.bas

- File: /mnt/data/A4_filter_tab.bas
- Type: .bas
- Description: Public filter_tab_initialized As Boolean

- Preview snippet:

```
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
    hourly_pricing As Filt
```

### Module: A9_init.bas

- File: /mnt/data/A9_init.bas
- Type: .bas
- Description: If all_initialized Then Exit Sub progress.start ("Initializing")

- Preview snippet:

```
Private settings_initialized As Boolean
Private EDC_initialized As Boolean
Private mail_type_initialized As Boolean
Private sheet_names_defined As Boolean
Private stats_initialized As Boolean
Public all_initialized As Boolean

Public imported_gagg As Boolean
Public imported_active As Boolean
Public imported_supplier As Boolean

Public all_reviewed As Boolean

Public ribbon_contract_number As String
Public ribbon_opt_out_date As String
Public ribbon_community As String
Public ribbon_EDC As String
Public ribbon_EDC_id As Long
Public ribbon_mail_type As String
Public ribbon_mail_type_id As Long

Public T As TestCase
Public F As FilterTab
Public SN As SheetNames
Public C As MacroColors
Public FS As FilterStatuses
Public S As MacroSettings
Public EDC As UtilitySettings
Public MT As MailType
Public A As ActiveList
Public Stats As Statistcs
Public UI As IRibbonUI

Sub init(Optional k, Optional mail_type)
    'If all_initialized Then Exit Sub
    'progress.start ("Initializing")
    If IsMissi
```

### Module: C1_process_active.bas

- File: /mnt/data/C1_process_active.bas
- Type: .bas
- Description: init active_accounts = active_tab.UsedRange.columns(1).value active_accounts = flatten_array(active_accounts)

- Preview snippet:

```
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
        Call update_checklist(S.QC.audit_checklist, "audit_pipp",
```

### Module: C1_process_active.docx

- File: /mnt/data/C1_process_active.docx
- Type: .docx
- Description: Module from C1_process_active.docx

- Preview snippet:

```
C1_process_active.bas - Documentation
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
        Call update_checklis
```

### Module: C1_process_active.html

- File: /mnt/data/C1_process_active.html
- Type: .html
- Description: Module from C1_process_active.html

- Preview snippet:

```
<html><head><meta charset='utf-8'><title>C1_process_active.bas Documentation</title></head><body>
<h1>C1_process_active.bas Documentation</h1>
<p>No sections extracted.</p>
<pre>Sub test_active()
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
        mi
```

### Module: C1_process_active.json

- File: /mnt/data/C1_process_active.json
- Type: .json
- Description: Module from C1_process_active.json

- Preview snippet:

```
{
  "source": "C1_process_active.bas",
  "summary": "Sub test_active()\n    'init\n    If Not MT.needs_active_list Then Exit Sub\n    progress.start (\"Processing Active List\")\n    check_active_matches\n    progress.finish\nEnd Sub\n\nSub process_active()\n    I",
  "sections": []
}
```

### Module: C1_process_active.md

- File: /mnt/data/C1_process_active.md
- Type: .md
- Description: Module from C1_process_active.md

- Preview snippet:

```
# C1_process_active.bas Documentation

## Overview
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
        Call up
```

### Module: C1_process_active.pdf

- File: /mnt/data/C1_process_active.pdf
- Type: .pdf
- Description: Module from C1_process_active.pdf

- Preview snippet:

```

```

### Module: C1_process_active.png

- File: /mnt/data/C1_process_active.png
- Type: .png
- Description: Module from C1_process_active.png

### Module: C1_process_active.pptx

- File: /mnt/data/C1_process_active.pptx
- Type: .pptx
- Description: Module from C1_process_active.pptx


### Module: C1_process_active.txt

- File: /mnt/data/C1_process_active.txt
- Type: .txt
- Description: Module from C1_process_active.txt

- Preview snippet:

```
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
        Call update_checklist(S.QC.audit_checklist, "audit_pipp",
```

### Module: C1_process_active.xlsx

- File: /mnt/data/C1_process_active.xlsx
- Type: .xlsx
- Description: Module from C1_process_active.xlsx

- Preview snippet:

```
## Sheet: C1_process_active

| Section   | Content (summary)                    |
|:----------|:-------------------------------------|
| Notes     | No sections found in source content. |

## Sheet: Samples

|   A |   B |
|----:|----:|
|  10 | nan |
|  20 | nan |
|  30 | nan |
|  40 | nan |
```

### Module: D1-D6_workflow.md

- File: /mnt/data/D1-D6_workflow.md
- Type: .md
- Description: Module from D1-D6_workflow.md

- Preview snippet:

```
# C1_process_active.bas

Executive summary
- This document describes the C1_process_active.bas module and how it integrates with the project&#39;s filter and mapping infrastructure. It provides an entry-point scaffold, inferred responsibilities, and concrete next steps to complete the doc when the source is available.
- Where the source lacks explicit procedure signatures or examples, TODO placeholders mark required inputs so you can quickly finalize the file.

## Module overview
- Purpose: Orchestrate processing of &quot;active&quot; records (apply active/supplier matching, set statuses, prepare data for downstream steps).
- Scope: Manipulates the Filter Tab, updates status/eligible flags, and calls matching/mismatch builders.
- Principal data structures: FilterTab (F), MacroSettings (S), MailType (MT), SheetNames (SN), FilterStatuses (FS), progress UI.

## Public entry points (scaffold)
- Sub init(Optional k, Optional mail_type)
  - Purpose: Initialize module-level state and dependen
```

### Module: D1_filter.bas

- File: /mnt/data/D1_filter.bas
- Type: .bas
- Description: pipp state rules usage

- Preview snippet:

```
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
        Call apply_bool_filter(ff, F.columns.free_serv
```

### Module: D2_dna.bas

- File: /mnt/data/D2_dna.bas
- Type: .bas
- Description: Sub dna_test()

- Preview snippet:

```
Sub dna_test()
    For k = 10 To 21
        test_dna (k)
    Next
End Sub

Sub test_dna(k)

    If EDC.state <> "OH" Then Exit Sub
    
    progress.start "Searching DNA List"
    
    S.DNA.wildcard_length = k
    
    add_dna_comparison_sheet
    
    Call sort_sheet_col(filter_tab(), 1, "A")
    
    data_arr = filter_data()
    
    dna_file = find_dna_file
    
    If dna_file = "" Then dna_file = find_dna_list
    
    Set conn = ADO_connection_excel(dna_file)
    dna_rows = ADO_row_count(conn, S.DNA.sheet_name)
    dna_data = ADO_data(conn, S.DNA.sheet_name, "A1:J" & dna_rows, 1)
    conn.Close
    
    dna_cols = UBound(dna_data, 2)
    
    account_match = dna_account_match(data_arr, dna_data)
    add_dna_results (account_match)
    
    dna_data = sort_2d_arr(dna_data, 5)
    
    address_match = dna_address_match(dna_data, account_match)
    add_dna_results (address_match)
    
    reapply_autofilter (dna_tab().index)
    
    dedupe_dna
    
    add_dna_formatting
```

### Module: D3_contracts.bas

- File: /mnt/data/D3_contracts.bas
- Type: .bas
- Description: init get_contracts_file

- Preview snippet:

```
Sub test_contracts()
    'init
    S.contracts.hide_snowflake_query = True
    get_contracts_file
    dedupe_contracts
    process_contracts
End Sub

Sub process_contracts()
    'get_contracts_file
    process_contracts_results
End Sub

Sub create_contracts_sql()

    If Not MT.needs_gagg_list Then Exit Sub
    
    delete_sheet SN.Snowflake
    
    Set query_tab = Sheets.Add(before:=home_tab())
    query_tab.name = SN.Snowflake
    
    Set ff = filter_tab()
    
    eligible_col = F.columns.eligible_opt_out.index
    active_col = F.columns.active_in_LP.index
    
    num_rows = Application.CountA(filter_tab().columns(1))
    num_cols = Application.Max(eligible_col, active_col)
    data_arr = ff.Range("A1").Resize(num_rows, num_cols).value
    
    Dim arr
    ReDim arr(0 To num_rows)
    
    arr(0) = "IN"
    
    k = 1
    
    For i = 2 To num_rows
        If data_arr(i, active_col) = "N" Then
            If data_arr(i, eligible_col) = "Y" Then
                arr(k) = ",'" & dat
```

### Module: D4_migration.bas

- File: /mnt/data/D4_migration.bas
- Type: .bas
- Description: get current contract

- Preview snippet:

```
Sub test_migration()
    If Not MT.check_migration_data Then Exit Sub
    MT.check_migration_data = True
    define_migration_settings
    check_legacy_data
End Sub

Sub check_legacy_data()
    
    If Not MT.check_migration_data Then Exit Sub
    
    migration_log_file = onedrive_migration_folder() & "\" & S.migration.migration_log_file
    
    Set conn = ADO_connection_excel(migration_log_file)
    migration_log_rows = ADO_row_count(conn, S.migration.migration_log_sheet)
    migration_log_data = ADO_data(conn, S.migration.migration_log_sheet, "A2:G" & migration_log_rows, 1)
    conn.Close
    
    current_contracts = get_array_col(migration_log_data, 1)
    previous_contracts = get_array_col(migration_log_data, 2)
    system_EDC = get_array_col(migration_log_data, 5)
    previous_system_arr = get_array_col(migration_log_data, 6)
    migration_files = get_array_col(migration_log_data, 7)
    
    'get current contract
    current_contract = "C-00132523"
    
    migration_row = arra
```

### Module: D5_mapping.bas

- File: /mnt/data/D5_mapping.bas
- Type: .bas
- Description: init

- Preview snippet:

```
Sub test_mapping()
    'init
    remove_other_ineligible
End Sub

Sub remove_other_ineligible()
    import_mapping
    If mapping_tab() Is Nothing Then Exit Sub
    If check_mapping() = False Then
        remove_file_input (S.mapping.file_source)
        Call update_checklist(S.QC.qc_checklist, "correct_mapping", -1)
        Exit Sub
    End If
    Call update_checklist(S.QC.qc_checklist, "correct_mapping", 1)
    remove_dna
    process_contracts
    process_mapping
    misc_filter
    set_step (6)
End Sub

Sub import_mapping()

    If Not mapping_tab() Is Nothing Then Exit Sub
    
    If T.mapping_file = "" Then
        file_name = Application.GetOpenFilename("Geocoding Files (*.xlsm), *.xlsm", , "Select Mapping Results File")
    Else
        mapping_folder = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(2) Macro Testing\(1) Testing Files\Test Mapping\"
        file_name = mapping_folder & T.mapping_file
    End If
    
    If VarType(file_name) = 11 Then Exit Sub
```

### Module: D6_misc.bas

- File: /mnt/data/D6_misc.bas
- Type: .bas
- Description: premise_mismatch_accounts

- Preview snippet:

```
Sub misc_filter()
    duke_sibling_accounts
    'premise_mismatch_accounts
    filter_tab().columns.AutoFit
End Sub

Sub duke_sibling_accounts()
    If EDC.display_name <> "DUKE" Then
        Exit Sub
    Else
        find_DUKE_sibling_accounts
    End If
End Sub

Sub find_DUKE_sibling_accounts()
    
    progress.start "Fixing DUKE Sibling Accounts"
    
    Dim arr As Variant
    
    parent_account = ""
    parent_account_len = 12
    subset_size = 4
    
    Set ff = filter_tab()
    
    n = ff.Range("A1").End(xlDown).row
    
    Set data_range = ff.Range(ff.Cells(1, 1), ff.Cells(n, subset_size))
    
    arr = data_range.value
    
    ReDim Preserve arr(1 To n, 1 To subset_size + 1)
    
    arr(1, 5) = "Parent Account"
    
    For i = 2 To n
        parent_account = Left(arr(i, 1), parent_account_len)
        arr(i, subset_size + 1) = parent_account
    Next
    
    For i = 2 To n
        initial_parent_account = arr(i, 5)
        i = fix_DUKE_sibling_accounts(arr, i, n, ini
```

### Module: Globals_and_Init.md

- File: /mnt/data/Globals_and_Init.md
- Type: .md
- Description: Module from Globals_and_Init.md

- Preview snippet:

```
# C1_process_active.bas

Executive summary
- This document describes the C1_process_active.bas module and how it integrates with the project&#39;s filter and mapping infrastructure. It provides an entry-point scaffold, inferred responsibilities, and concrete next steps to complete the doc when the source is available.
- Where the source lacks explicit procedure signatures or examples, TODO placeholders mark required inputs so you can quickly finalize the file.

## Module overview
- Purpose: Orchestrate processing of &quot;active&quot; records (apply active/supplier matching, set statuses, prepare data for downstream steps).
- Scope: Manipulates the Filter Tab, updates status/eligible flags, and calls matching/mismatch builders.
- Principal data structures: FilterTab (F), MacroSettings (S), MailType (MT), SheetNames (SN), FilterStatuses (FS), progress UI.

## Public entry points (scaffold)
- Sub init(Optional k, Optional mail_type)
  - Purpose: Initialize module-level state and dependen
```

Usage Workflows
- Development: Update BAS/MD files in /mnt/data, run this script to regenerate README.md.
- Review: Open the generated README.md for a consolidated view.
- Extend: Add new modules and corresponding documentation blocks as they are introduced.

TODOs and Open Gaps
- [ ] Source-detail extraction for A0-A4 BAS modules where comments are sparse.
- [ ] Integrate previous draft README content (path: TODO_PLACEHOLDER).
- [ ] Add more robust test coverage for PDF/docx/xlsx extraction paths.
- [ ] Validate and normalize module descriptions across file types.
