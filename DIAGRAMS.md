# Diagram Index

- [Overview](#overview)
- [Import File](#import-files)
- [Trim Data](#trim-data)
- [Export Results](#export-results)

## Overview

```mermaid

flowchart TB

classDef classNode fill:#eef6ff,stroke:#1e40af,stroke-width:2px,font-size:12px

%% Nodes
Initialize["<b>Initialize</b><br>-define_macro_settings<br>-define_EDC<br>-define_mail_type<br>"]:::classNode
Import_Files["<b>Import_Files</b><br>-import_gagg_list<br>-import_active_list<br>-import_supplier_list"]:::classNode
Preprocess_List["<b>Preprocess_List</b><br>-create_filter_tab<br>-create_mapping_file<br>-populate_filter_tab"]:::classNode
Filter_List["<b>Filter_List</b><br>-remove_duplicates<br>-pipp<br>-state_rules<br>-usage<br>-shopping<br>-arrears<br>-national_chains"]:::classNode
PUCO_Do_Not_Agg["<b>PUCO_Do_Not_Agg</b><br>-account_number_match<br>-service_address_match<br>-manual_name_comparison"]:::classNode
Contracts_Query["<b>Contracts_Query</b><br>-import_snowflake_file<br>-dedupe_query_results<br>-active_accounts<br>-previous_opt_outs"]:::classNode
Mapping["<b>Mapping</b><br>-import_mapping_results"]:::classNode
Misc_Filters["<b>Misc_Filters</b><br>-DUKE_sibling_accounts<br>-LP_premise_mismatch"]:::classNode
Review_Data["<b>Review_Data</b>"]:::classNode
Export_Files["<b>Export_Files</b>"]:::classNode

subgraph Setup["<b>Setup</b>"]
direction LR
    Initialize --> Import_Files
end

subgraph Initial_Processing["<b>Initial_Processing</b>"]
direction LR
    Preprocess_List --> Filter_List
end

subgraph External_Data_Filters["<b>External_Data_Filters</b>"]
direction LR
    PUCO_Do_Not_Agg --> Contracts_Query --> Mapping --> Misc_Filters
end

subgraph Review["<b>Review</b>"]
direction LR
    Review_Data --> Export_Files
end

Setup ==> Initial_Processing ==> External_Data_Filters ==> Review

```
### Import Files

```mermaid

graph TD

    subgraph Import_GAGG_List
    direction LR
        select_utility_files --> open_file
        open_file --> copy_first_tab
        copy_first_tab --> paste_data
        copy_first_tab -->|"FE only"| copy_mail_addresses
        copy_first_tab -->|"AM only"| copy_second_tab
        copy_mail_addresses --> paste_data
        copy_second_tab --> paste_data
        paste_data -->|"loop for each file"| open_file
    end
```

## Trim Data

```mermaid

    subgraph Trim_Data
    direction LR
        AM-->combine_sheets
        AES-->combine_sheets
        AEP-->combine_sheets
        COM-->combine_sheets
        DUKE-->combine_sheets
        FE-->combine_sheets
    end

    subgraph AEP
    direction LR
        AEP_delete_first_col
        AEP_delete_last_row
        AEP_delete_second_row
        AEP_delete_empty_cols
    end

    subgraph AES
    direction LR
        AES_delete_first_10_rows
    end

    subgraph AM
    direction LR
        AM_delete_first_10_rows
        AM_unmerge_columns
    end

    subgraph COM
    direction LR
        COM_do_nothing
    end

    subgraph DUKE
    direction LR
        DUKE_do_nothing
    end

    subgraph FE
    direction LR
        FE_delete_first_column
        FE_delete_second_row
    end

    subgraph Format_Utility_Data
    direction LR
        remove_tabs_from_headers --> find_account_column
        Format_Account_Numbers --> dedupe_accounts
    end

    subgraph Format_Account_Numbers
    direction LR
        format_as_string -->|"FE"| A["080*<br>add_leading_zeros<br>len=20"]
        format_as_string -->|"OP"| B["001400607*<br>add_leading_zeros<br>len=17"]
        format_as_string -->|"CS"| C["000406210*<br>add_leading_zeros<br>len=17"]
        format_as_string -->|"AES"| D["080*<br>add_leading_zeros<br>len=23"]
        format_as_string -->|"DUKE"| E["[#x12]Z[#x9]<br>add_leading_zeros<br>len=22"]
        format_as_string -->|"AM"| F["*<br>add_leading_zeros<br>len=10"]
        format_as_string -->|"COM"| G["*<br>add_leading_zeros<br>len=10"]
    end

    find_account_column --> format_as_string

```
