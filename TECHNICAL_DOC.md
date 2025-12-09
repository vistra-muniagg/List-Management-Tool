# List Management Tool Technical Documentation

## Introduction

This will serve as a guide for mantaining and/or repairing the List Management tool desribed in a higher level in the [README](README.md).
This a guide not an exact instruction manual, meant to be a source of documentation and a reference for anyone else attempting to edit the code.

### Background
The tool was initially designed as a more transparent replacement for the black-box system tool originally implemented at Energy Harbor.
It has evolved over time to include more eligibility checks and more complex logic.
The tool's current functions revolve around the use of custom type objects for various system settings, and mailing properties.

> [!NOTE]
> These custom types may be replaced with either dictionaries or class objects at a later date but the idea is the same, regardless of data structure.

### Supported Use Cases


## Module Map

| Name | Description |
|---|---|
|A0_settings | configuration settings for the tool|
|A1_EDC | definitions for the utility type object|
|A2_mail_type | definitions for the mail type object|
|A3_formatting | definitions for sheet names and colors|
|A4_filter_tab | definitions and configuration settings for the columns on the Filter Tab|
|A5_active_list | definitions and configuration for the Active List|
|A6_customUI | functionality for the custom List Management ribbon tab|
|A7_validation | functionality for the audit checklists and code validation|
|A8_logging | functionality for statistics and process logging|
|A9_init | initialization processes|
|B1_import | routines for importing files|
|B2_preprocess | create and populate the Filter Tab|
|B3_format_data | format name and address data|
|B4_waterfall | create various pivot tables|
|C1_check_active | validate the active list makes sense with the given gagg list|
|C2_process_active | process the active list|
|D1_filter | processes for removing ineligible accounts|
|D2_dna | routines for checking PUCO Do Not Agg List|
|D3_contracts | process LandPower contracts data|
|D4_migration | check gagg list against legac system data|
|D5_mapping | import and process mapping data|
|D6_misc | extra specialized processes done after filtering|
|E1_review_data | sanity checks data for anything that might cause an upload error in LandPower|
|E2_LP_upload_tab | create the|
|E3_drop_at_renewal | |
|E4_mail_list | |
|E5_opt_in_list | |
|E6_emails | |
|E7_export | |
|F1_helpers | |
|F2_old_helpers | |
|F3_data_helpers | |
|F4_fso_helpers | |
|F5_array_helpers | |
|G1_snawflake_api | |
|G2_github_api | |
|G3_windows_api | |
|H1_errors | |
|H2_testing | |
|H3_sandbox | |
|H4_JsonConverter | |

## Detailed Module References


### Module: A0_settings


### Module: A1_EDC


### Module: A2_mail_type


### Module: A4_filter_tab


### Module: A9_init


### Module: C1_process_active


### Module: D1_filter


### Module: D2_dna


### Module: D3_contracts


### Module: D4_migration


### Module: D5_mapping


### Module: D6_misc


