# List Management Tool Technical Documentation

## Overview

The List Management Tool determines eligibility for accounts on a list for a specified utility (`EDC`) and mailing type (`MailType`). It uses supplied data files, external datasets, and manual user input. It is tailored for utilities in Ohio (OH) and Illinois (IL) and supports specific mailing types, each with unique logic and eligibility criteria.

## Definitions

- **Waterfall**: The copy of the tool containing data for a mailing, including pivot tables and summary statistics.
- **Filtering**: Removal of ineligible accounts.
- **Status**: Indicates account eligibility.
- **Home Tab**: Main worksheet for configuration, summaries, and pivots.
- **Filter Tab**: Worksheet with the normalized account dataset and status fields.
- **Geocoding**: Converting addresses to geographical points.
- **Mapping**: Using geocoding results to determine eligibility.

## Supported Utilities

- **Ohio (OH):**
  - AES Ohio (`AES`)
  - Ohio Power Company (`OP`)
  - Columbus Southern Power (`CS`)
  - Duke Energy (`DUKE`)
  - Ohio Edison (`OE`)
  - The Illuminating Company (`CEI`)
  - Toledo Edison (`TE`)
- **Illinois (IL):**
  - Ameren Illinois (`AM`)
  - Commonwealth Edison (`COM`)

## Mailing Types

- **Contract Renewal + Resweep (`CR`)**: Enroll active customers and newly eligible accounts for renewal.
- **Resweep (`SWP`)**: Enroll eligible customers on an existing contract.
- **New Community (`NEW`)**: Enroll eligible customers on a new program or upon supplier switch.
- **Renewal Only (`REN_ONLY`)**: Renewal for currently active customers only.

## Workflow Steps

### 0. Configuration
- Select `Utility`, `MailType`, and enter the formal community name.
- Save a new waterfall file for each mailing.

### 1. Import Files
- Required files depend on mailing type (e.g., customer lists, utility lists).
- Files are imported, combined if multiple sheets, and stored in worksheets.

### 2. Normalize Data
- Import data is standardized to consistent columns in `Filter Tab`.
- Unknown or absent fields are marked with `-`.

### 3. Utility Data Exclusions
- State-specific rules are applied to determine eligibility.
- Each account’s status is updated per utility and state requirements.

### 4. External Data Exclusions

#### Do Not Aggregate (DNA)
- Checks against external opt-out lists (e.g., Ohio PUCO "DNA List").
- Account number and address-based matching, with manual user verification.

#### LandPower System Comparison
- SQL queries Vistra systems for account activity, checking for enrollments in other programs.
- Updates status for accounts active elsewhere.

#### Geocoding (Mapping)
- Checks geographic eligibility by mapping service addresses.
- Only eligible accounts within boundaries are included.

### 5. Create Upload File
- After filtering, enter mailing details (community name, contract number, opt-out date).
- Output files for upload to LandPower, printing, etc.
- Peer review required before exporting final files.

### 6. Export Data
- Use peer review checklist before file creation.
- Automates file exports: LandPower upload, mailing list, opt-in/out lists.

## Data Structures and File Organization

- The workbook consists of at least a `HOME` tab and `FILTER` tab.
- Filtered and normalized data populate the `FILTER` sheet.
- Files for upload and mailing are programmatically generated according to mailing type.

## Known Issues, Special Cases, and Extendability

- Complex logic and many special cases for each utility and mailing type.
- [Documented issues and further details can be added in dedicated sections.]

## Setup and Configuration Modules (`A0_settings.bas` – `A9_init.bas`)

The List Management Tool's VBA codebase is organized into a series of "A#_" modules, each dedicated to a category of application settings, workflow logic, or initialization. Below is a technical overview of their purposes, with example code snippets and specific field references.

### A0_settings.bas: Core Settings Types
Defines settings structures used throughout the workbook.

```vbnet name=A0_settings.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A0_settings.bas
Public Type HomeTabSettings
    peer_review_checklist_range As String
    version_location As String
    user_location As String
    edc_location As String
    ...
    notes_location As String
    file_log_location As String
    ...
End Type

Public Type ImportSettings
    max_csv_cols As Long
    max_copy_size As Long
    trim_sheets As Boolean
    FE_address_replace As Boolean
End Type
' ...and many more types
```

Use: Holds essential cell locations, UI rendering values, import flags, file logging, etc.

---

### A1_EDC.bas: Utility Settings and Selection Logic
Defines all knowledge and logic for utility types ("EDC").

```vbnet name=A1_EDC.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A1_EDC.bas
Public Type UtilitySettings
    name As String
    display_name As String
    ruleset_name As String
    ...
    shop             As String ' shopper col header
    shopper_yes      As String ' y/n value
    ...
    account As String
    service As Variant
    mail() As Variant
    ...
End Type

Sub define_EDC(EDC_name)
    Select Case (EDC_name)
        Case "OE": define_EDC_OE
        ...
    End Select
    ...
End Sub
```

Use: Centralizes utility definitions and selection logic for data import and eligibility.

---

### A2_mail_type.bas: Mailing Type Configuration

```vbnet name=A2_mail_type.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A2_mail_type.bas
Public Type MailType
    name As String
    display_name As String
    ...
    needs_gagg_list As Boolean
    make_opt_in_list As Boolean
    ...
End Type

Sub define_mail_type(mail_type)
    If mail_type = S.UI.mail_type_items(0) Then
        define_mail_type_NEW
    ...
    End If
    ...
End Sub
```

Use: Manages mailing type options (NEW, REN, SWP, REN_ONLY), their requirements and UI logic.

---

### A3_formatting.bas: Sheet, Color, and Formatting Settings

```vbnet name=A3_formatting.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A3_formatting.bas
Public Type SheetNames
    HOME As String
    README As String
    ...
End Type

Public Type CellColors
    FontColor As Long
    InteriorColor As Long
End Type
' ... Color setup routines
Sub define_colors()
    With C.GRAY_1
        .InteriorColor = RGB(217, 217, 217)
        .FontColor = vbBlack
    End With
    ...
End Sub
```

Use: Centralizes worksheet naming, cell color definitions, and cell formatting options.

---

### A4_filter_tab.bas: Filter Tab Structure

```vbnet name=A4_filter_tab.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A4_filter_tab.bas
Public Type FilterTabColumns
    account_number As ColumnHeader
    status As ColumnHeader
    eligible_opt_out As ColumnHeader
    ...
End Type
```

Use: Defines all columns and statuses for filter tab logic, including eligibility, address handling, opt-out, etc.

---

### A5_active_list.bas: Active Account Columns

```vbnet name=A5_active_list.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A5_active_list.bas
Public Type ActiveListColumns
    account_number As ActiveColumnHeader
    customer_name As ActiveColumnHeader
    ...
End Type

Sub define_active_cols()
    With A.columns
        .account_number.header = "UTILITYACCOUNTVALUE"
        ...
    End With
End Sub
```

Use: Provides mapping from raw file columns to tool-internal logic for active account lists.

---

### A6_customUI.bas: Ribbon and User Interface

```vbnet name=A6_customUI.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A6_customUI.bas
Sub contracts_instructions()
    MsgBox "Run the contracts query from snowflake using the provided SQL code as you normally would"
End Sub

Sub ribbon_on_load(ribbon As IRibbonUI)
    Set UI = ribbon
    init
End Sub
' ...Ribbon callbacks for selection and export
```

Use: Manages custom Ribbon buttons, instruction dialogs, and ribbon state tracking.

---

### A7_validation.bas: Checklist Structures and Validation

```vbnet name=A7_validation.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A7_validation.bas
Public Type CheckListItem
    name As String
    label As String
    index As Long
End Type

Sub define_checklists()
    With S.QC.audit_checklist
        .location = S.HOME.audit_checklist_location
        ...
    End With
End Sub
```

Use: Handles peer review and QC audit checklist structures, their configuration, and updates during workflow.

---

### A8_logging.bas: Statistics Types

```vbnet name=A8_logging.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A8_logging.bas
Public Type Stat
    name As String
    value As String
End Type

Public Type InfoStats
    category As String
    stat_list() As Variant
    ...
End Type
' ...and other FileStats, FilterStats
```

Use: Tracks needed statistics, file metadata, review info, and operational tallies throughout runs.

---

### A9_init.bas: Initialization Logic

```vbnet name=A9_init.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/A9_init.bas
Sub init(Optional k, Optional mail_type)
    If IsMissing(mail_type) Then mail_type = "REN"
    If Not IsMissing(k) Then Call define_test_case(k, mail_type)
    define_colors
    define_sheet_names
    ...
    all_initialized = True
    progress.complete
End Sub
```

Use: Master initialization that loads configuration, settings, Ribbon tracking, color schemes, defines columns, and sets up the runtime. 

---

## Import, Preprocessing, and Filtering Workflow Modules (`B1_import.bas` – `B2_preprocess.bas`)

These modules organize the List Management Tool’s data intake and filtering logic, progressing from file import, through filtering preparation, to data normalization and conditional population.

### B1_import.bas: Active & Utility Data Import

Handles importing "Active Lists" and "Utility Lists" required for each mailing and community, from designated directories or manual file selection:

```vbnet name=B1_import.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/B1_import.bas
Sub test_import()
    progress.start ("Importing Files")
    active_list_folder = "...\\Test Active Lists\\"
    gagg_list_folder = "...\\Test GAGG Lists\\"
    If MT.needs_active_list Then Call import_active_list(active_list_folder & T.active_list, "snowflake", 1)
    If MT.needs_gagg_list Then Call import_gagg_files(gagg_list_folder & T.gagg_list, "Utility", 1)
    progress.finish
End Sub

Sub import_active_list(Optional file_name, Optional file_source, Optional target_location)
    If imported_active Then Exit Sub
    selected_file = Application.GetOpenFilename("Active Lists (*.csv), *.csv", , "Select Active List", False)
    If selected_file = False Then Exit Sub
    Call import_csv_file(selected_file, file_source, target_location)
    Sheets(target_location).name = SN.Active
    dupes = remove_duplicates(target_location)
    reapply_autofilter (target_location)
    Call sort_sheet_col(Sheets(SN.Active), 1, "A")
    Set k = add_file_input(selected_file, file_source)
    k.Offset(0, 1).value = dupes(0) - 1
    k.value = dupes(1)
    imported_active = True
End Sub

Sub import_gagg_files(Optional file_name, Optional file_source, Optional target_location)
    If imported_gagg Then Exit Sub
    selected_files = Application.GetOpenFilename("Select Utility File(s), *.*", , "Utility File(s)", multiselect:=EDC.multiselect)
    If IsArray(selected_files) Then
        For Each gagg_file In selected_files
            Call import_file(gagg_file, "Utility", 1)
        Next
    Else
        Call import_file(selected_files, "Utility", 1)
    End If
End Sub
```

**Key Operations:**  
- Automated or manual file selection for both active customer lists and utility lists.
- Duplicate removal, sorting, and standardization.
- Tracking and marking imported status for each category.
---

### B2_preprocess.bas: Data Preprocessing & Filter Tab Creation

Sets up and populates the main Filter Tab, transforming imported utility/customer data into a normalized set suitable for eligibility filtering.

```vbnet name=B2_preprocess.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/B2_preprocess.bas
Sub preprocess()
    define_filter_tab
    define_filter_tab_columns
    create_filter_tab          'Creates sheet, sets up header, formatting and autofilter
    populate_filter_tab        'Populates tab with deduped utility/customer data
    If EDC.state <> "IL" Then hide_filter_group ("IL Filters")
    If EDC.state <> "OH" Then hide_filter_group ("OH Filters")
End Sub

Sub create_filter_tab()
    progress.start ("Creating Filter Tab")
    delete_sheet (SN.Filter)
    Set s1 = Sheets.Add(before:=Sheets(SN.HOME))
    s1.name = SN.Filter
    s1.Rows(1).Delete
    For j = 1 To UBound(F.order_array)
        col = F.order_array(j)
        s1.Cells(1, j) = col.header
        s1.Cells(1, j).Interior.color = col.cell_color.InteriorColor
        s1.Cells(1, j).Font.color = col.cell_color.FontColor
        s1.columns(j).NumberFormat = col.data_format
    Next
    reapply_autofilter (s1.index)
    progress.complete
End Sub

Sub populate_filter_tab()
    If Not MT.needs_gagg_list Then Exit Sub
    Set ff = filter_tab()
    Set gagg_list = Sheets(SN.Utility)
    gagg_data = deduped_data_arr(gagg_list)
    gagg_headers = get_array_row(gagg_data, 1)
    num_rows = UBound(gagg_data, 1)
    num_cols = UBound(gagg_data, 2)
    For k = 1 To UBound(F.order_array)
        col = F.order_array(k)
        Select Case col.data_type
            Case "Literal":
                Call populate_literal_filter_col(ff, col, k, gagg_headers, gagg_data, num_rows)
            Case "Generated":
                Call populate_generated_filter_col(ff, col, k, gagg_headers, gagg_data, num_rows)
            Case "Boolean":
                Call populate_bool_filter_col(ff, col, k, gagg_headers, gagg_data, num_rows)
            Case "Calculated":
                Call populate_calculated_filter_col(ff, col, k, gagg_headers, gagg_data, num_rows)
        End Select
    Next
    reapply_autofilter (ff.index)
End Sub
```

**Key Operations:**  
- Creates a new Filter Tab with configured headers, colors, and filters.
- Populates with deduplicated, normalized utility/customer data, using type-specific logic for each column.
- Hides irrelevant filter groups for the non-active operating state.

---

### B3_format_data.bas: Data Formatting and Cleanup

This module provides all string, address, and cleanup functions for customer, service, and mail data imported into the tool. It standardizes names and addresses, handles edge cases and anomalies per utility, and prepares the cleaned data for eligibility assessment and export.

### Name Cleaning and Parsing

Functions for splitting, normalizing, and reordering customer name strings:

```vbnet name=B3_format_data.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/B3_format_data.bas
Function clean_name(n)
    n = Application.Trim(n)
    split_name = name_parts(name_reverse(n))
    clean_name = Trim(split_name(0) & " " & split_name(2))
    clean_name = UCase(clean_name)
End Function

Function name_reverse(n)
    comma = InStr(n, ",")
    ' Reorders names like "Smith, John" and strips suffixes/prefixes (JR, LLC, etc)
    ...
End Function

Function name_parts(namestring) As Variant
    ' Parses string to First, Middle, Last (handles suffixes/prefixes, abbreviates, etc.)
    ...
End Function
```

### Address Parsing, Cleaning, and Formatting

Functions for extracting, cleaning, and standardizing addresses:

```vbnet name=B3_format_data.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/B3_format_data.bas
Function service_address(source_row, gagg_data, service_cols As Variant, comed_special_case As Boolean)
    ' Combines address columns, fixes special edge cases for COMED
    ...
End Function

Function mail_address(source_row, gagg_data, mail_cols As Variant)
    ' Handles PO BOX, DUKE-specific rules, deduplication
    ...
End Function

Sub add_apt_numbers()
    ' Syncs apt/unit numbers between service/mail addresses for rows with matches
    ...
End Sub

Sub replace_empty_mail()
    ' If mail address/city/state/zip is empty, copy from service information
    ...
End Sub

Function clean_mail_address(str)
    str = clean_mail_address_1(str)
    str = clean_mail_address_3(str)
    clean_mail_address = str
End Function
```

### Utility-Specific Logic

Extensive handling of edge cases for each supported utility (AES, AM, COMED, DUKE, AEP):

```vbnet name=B3_format_data.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/B3_format_data.bas
Function AES_service_city(str)
    ' Extracts service city for AES addresses
End Function
Function COMED_service_address(s1)
    ' Detailed edge-case trims and reorder for ComEd addresses, apt/unit logic
End Function
Function DUKE_mail_address(m1, m2)
    ' Cleans, deduplicates, and fixes mail address for Duke accounts
End Function
' ... and many more for other address components
```

### State Abbreviation Tools

Dictionary & functions to map full state names to abbreviations and vice versa:

```vbnet name=B3_format_data.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/B3_format_data.bas
Function state_dict() As Object
    ' Returns dictionary for full-name to abbreviation mapping
End Function
Function state_abbrev_dict() As Object
    ' Returns dictionary for abbreviation to full-name mapping
End Function
```

### Primary Formatting Workflow

Addresses are processed through these steps, which are run at various points in the overall macro:

```vbnet name=B3_format_data.bas url=https://github.com/vistra-muniagg/List-Management-Tool/blob/main/B3_format_data.bas
Sub format_address_data()
    replace_empty_mail
    clean_address_data
    add_apt_numbers
    clean_mail_addresses
    Call update_checklist(S.QC.qc_checklist, "account_number_format", 1)
End Sub

Sub clean_address_data()
    ' Cleans the service and mail address columns with utility-specific logic
    ...
End Sub
```

---

**Usage Notes**
- Called automatically after import/preprocessing for all address columns in the Filter Tab.
- Handles missing data, inconsistent formatting, duplicate apartment/unit info, and corrects utility-specific quirks.
- Successful processing is required before exporting upload files or running eligibility routines.
