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

## Setup Files (`A0_settings.bas` to `A9_init.bas`)

The List Management Tool relies on a series of configuration and initialization modules named `A0_settings.bas` through `A9_init.bas`. These files are implemented as VBA modules and are critical for the runtime environment of the tool, typically within Excel or a similar environment.

### Overview & Organization

- **A0_settings.bas**: Central configuration module (global constants, default parameters, user options).
- **A1_community.bas**: Contains logic and settings for handling community details such as names, IDs, contract numbers, and configurations relevant to each mailing.
- **A2_utilities.bas**: Maps utility identifiers, account formatting conventions, and provides utility-specific operational logic.
- **A3_mailtype.bas**: Encapsulates mailing type recognition, toggles, and type-specific controls (CR, SWP, NEW, REN_ONLY).
- **A4_fieldmap.bas**: Contains logic and lookup definitions for standardizing and normalizing imported data fields.
- **A5_eligibility.bas**: Implements eligibility rules, filtering order, and state/utility-specific criteria.
- **A6_datasources.bas**: Manages external data connection setup and queries (DNA List reference, LandPower, or other databases).
- **A7_address.bas**: Processes address normalization and prepares geocoding/mapping routines.
- **A8_peerreview.bas**: Versioning and logic for peer review steps and checklist handling.
- **A9_init.bas**: Initialization and runtime setup—ensures all modules and configuration files are loaded and ready.

### Example Responsibilities

- **Initialization routines** are typically in `A9_init.bas`, ensuring all configuration and dependency modules (`A0`–`A8`) are loaded at workbook open.
- **Global constants** and shared parameters are usually defined in `A0_settings.bas`.
- **Eligibility logic** (such as filtering rules per state and utility) is implemented in `A5_eligibility.bas`.
- **Community and contract details** (such as opt-out windows or contract numbers) are handled in `A1_community.bas`.
- **Utility file mapping** and specific formatting (e.g. AES vs. COM fields) are managed in `A2_utilities.bas` and `A4_fieldmap.bas`.
- **External data source querying** (SQL for LandPower or .csv for DNA List) routines are in `A6_datasources.bas`.
- **Address normalization** and preparation for mapping are processed in `A7_address.bas`.

### Typical Structure

Each `.bas` file contains one or more VBA modules or classes, organizing procedures and variables relevant to its configuration or operation area. 
Common elements in these files may include:

```vbnet
' A0_settings.bas Example
Public Const DEFAULT_EDC As String = "AES"
Public Const DEFAULT_MAIL_TYPE As String = "CR"
Public UserOptions As Collection

Sub LoadSettings()
    ' Load global settings from workbook or SharePoint config
End Sub
```

```vbnet
' A5_eligibility.bas Example
Function IsAccountEligible(accountInfo As AccountType) As Boolean
    ' ... complex filter and state logic
    If accountInfo.State = "OH" Then
        ' Apply Ohio-specific rules
    End If
End Function
```

### Usage & Maintenance

- **Required for Tool Execution**: All files `A0_settings.bas` through `A9_init.bas` must be present and properly referenced in the VBA project for the tool to function.
- **Update & Documentation**: Each file/module should be independently documented, with comments for all procedures and configuration values.
- **Version Control**: Whenever changes are made, version information should be updated in module headers for audit trail purposes.
- **Customization**: For new utilities, mailing types, or eligibility rules, update the relevant `.bas` file with new mappings or logic.

### Best Practices

- Modularize settings and logic—avoid hardcoding business logic in worksheet code-behind.
- Use shared constants for parameter values to ease maintenance.
- Add header comments to each `.bas` file indicating its role and date/version.
- Cross-reference these modules in your README or separate developer guide when onboarding new maintainers.

---

If you need file header templates or want a breakdown of typical functions/procedures for each module, I can provide examples for your VBA environment!
