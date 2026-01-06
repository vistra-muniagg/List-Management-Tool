# Diagram Index

- [Overview](#overview)
- [Import File](#import-files)
- [Trim Data](#trim-data)

## Overview

```mermaid

%%{init: {
  "theme": "base",
  "themeVariables": {
    "fontSize": "12px",
    "fontFamily": "Segoe UI",
    "primaryColor": "#eef6ff",          %% node fill
    "primaryBorderColor": "#1e40af",    %% node border
    "primaryTextColor": "#0f172a",      %% node text
    "lineColor": "#1e40af"
  }
}}%%

flowchart TB

classDef classNode fill:#eef6ff,stroke:#1e40af,stroke-width:2px,font-size:12px

%% Nodes
Initialize["<b>Initialize</b><br>-define_macro_settings<br>-define_EDC<br>-define_mail_type<br>"]
Import_Files["<b>Import_Files</b><br>-import_gagg_list<br>-import_active_list<br>-import_supplier_list"]
Preprocess_List["<b>Preprocess_List</b><br>-create_filter_tab<br>-create_mapping_file<br>-populate_filter_tab"]
Filter_List["<b>Filter_List</b><br>-remove_duplicates<br>-pipp<br>-state_rules<br>-usage<br>-shopping<br>-arrears<br>-national_chains"]
PUCO_Do_Not_Agg["<b>PUCO_Do_Not_Agg</b><br>-account_number_match<br>-service_address_match<br>-manual_name_comparison"]
Contracts_Query["<b>Contracts_Query</b><br>-import_snowflake_file<br>-dedupe_query_results<br>-active_accounts<br>-previous_opt_outs"]
Mapping["<b>Mapping</b><br>-import_mapping_results"]
Misc_Filters["<b>Misc_Filters</b><br>-DUKE_sibling_accounts<br>-LP_premise_mismatch"]
Review_Data["<b>Review_Data</b>"]
Export_Files["<b>Export_Files</b>"]

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

## Import Files

```mermaid

%%{init: {
  "theme": "base",
  "themeVariables": {
    "fontSize": "12px",
    "fontFamily": "Segoe UI",
    "primaryColor": "#eef6ff",          %% node fill
    "primaryBorderColor": "#1e40af",    %% node border
    "primaryTextColor": "#0f172a",      %% node text
    "lineColor": "#1e40af"
  }
}}%%

graph TD

    subgraph Import_GAGG_List["**Import GAGG List**"]
    direction LR
        select_utility_files --> open_file
        open_file --> copy_first_tab
        copy_first_tab --> paste_data
        copy_first_tab -->|"**FE only**"| copy_mail_addresses
        copy_first_tab -->|"**AM only**"| copy_second_tab
        copy_mail_addresses --> paste_data
        copy_second_tab --> paste_data
        close_file -->|"**loop for each file**"| open_file
        paste_data --> close_file
    end
```

## Trim Data

```mermaid

%%{init: {
  "theme": "base",
  "themeVariables": {
    "fontSize": "12px",
    "fontFamily": "Segoe UI",
    "primaryColor": "#eef6ff",          %% node fill
    "primaryBorderColor": "#1e40af",    %% node border
    "primaryTextColor": "#0f172a",      %% node text
    "lineColor": "#1e40af"
  }
}}%%

graph TD

    subgraph Trim_Data["**Trim Data**"]
    direction TB
        %%AM --> combine_sheets
        %%AES --> combine_sheets
        %%AEP --> combine_sheets
        %%COM --> combine_sheets
        %%DUKE --> combine_sheets
        %%FE --> combine_sheets
        %%AM---AEP---AES---COM---DUKE---FE
        new_tab -->|"**AM**"| trim_AM["AM_delete_first_10_rows<br>AM_unmerge_columns"] --> combine_sheets
        new_tab -->|"**AEP**"| trim_AEP["AEP_delete_first_col<br>AEP_delete_last_row<br>AEP_delete_second_row<br>AEP_delete_empty_cols"] --> combine_sheets
        new_tab -->|"**AES**"| trim_AES["AES_delete_first_10_rows"] --> combine_sheets
        new_tab -->|"**FE**"| trim_FE["FE_delete_first_column<br>FE_delete_second_row"] --> combine_sheets
        %%new_tab -->|"**COM**"| no_action["do nothing"] --> combine_sheets
        %%new_tab -->|"**DUKE**"| no_action["do nothing"] -->combine_sheets
        new_tab -->|"**COM**"| combine_sheets
        new_tab -->|"**DUKE**"| combine_sheets
        

    end

    %%subgraph AEP[" "]
    %%direction LR
    %%    trim_AEP["AEP_delete_first_col<br>AEP_delete_last_row<br>AEP_delete_second_row<br>AEP_delete_empty_cols"]
    %%end

    %%subgraph AES[" "]
    %%direction LR
    %%    trim_AES["AES_delete_first_10_rows"]
    %%end

    %%subgraph AM[" "]
    %%direction LR
    %%    trim_AM["AM_delete_first_10_rows<br>AM_unmerge_columns"]
    %%end

    %%subgraph FE[" "]
    %%direction LR
    %%    trim_FE["FE_delete_first_column<br>FE_delete_second_row"]
    %%end

    subgraph Format_Utility_Data["**Format Utility Data**"]
    direction LR
        remove_tabs_from_headers --> find_account_column
    end

    subgraph Format_Account_Numbers["**Format Account Numbers**"]
    direction TB
        format_as_string -->|"**OE/TE/CEI**"| A["080*<br>len=20"]
        format_as_string -->|"**OP**"| B["001400607*<br>len=17"]
        format_as_string -->|"**CS**"| C["000406210*<br>len=17"]
        format_as_string -->|"**AES**"| D["080*<br>len=23"]
        format_as_string -->|"**DUKE**"| E["[#x12]Z[#x9]<br>len=22"]
        format_as_string -->|"**AM**"| F["*(no pattern)<br>len=10"]
        format_as_string -->|"**COM**"| G["*(no pattern)<br>len=10"]
        %%A-.->B-.->C-.->D-.->E-.->F-.->G
        A--> dedupe_accounts
        B--> dedupe_accounts
        C--> dedupe_accounts
        D--> dedupe_accounts
        E--> dedupe_accounts
        F--> dedupe_accounts
        G--> dedupe_accounts
    end

    Trim_Data ==> Format_Utility_Data
    Format_Utility_Data ==> Format_Account_Numbers

```
