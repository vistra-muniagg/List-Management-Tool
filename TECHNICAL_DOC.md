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

---

For a user-level summary and workflow, see the [README.md](https://github.com/vistra-muniagg/List-Management-Tool/blob/main/README.md). This technical guide should be supplemented with instruction on additional sheets (Mapping Tool, special logic tabs) and further breakdowns for eligibility and file formats as needed.
