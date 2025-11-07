# List-Management-Tool

&copy; 2025 Andrew Rodgers

## About

This tool is used to determine eligibility for accounts on a list for a specified utility `EDC` and mailing type `MailType` using data from the file(s) provided, external data, and manual input to determine the list of eligible accounts.
There are many special conditions and logic features built in for each EDC and MailType, which will be summarized in a later section (see [Special Cases](#special-cases)).
This is intended to be a summary suitable for management and new users.
There is a master copy of this tool saved in Sharepoint, and each mailing creates a copy in its resepective folder, containing the relevant data for that community.
Detailed technical documentation can be found [here] <!-- [here] (TECHNICAL_DOC.md) -->

## Definitions
`Waterfall` - the copy of this tool containing the pertinent data for a mailing. Also contains pivot tables and summary statistics about the mailing.\
`Filtering` - the process of removing ineligible accounts from any lists relevant to a mailing.\
`Status` - label desribing the eligibility of an account
`Home Tab` - the worksheet on the waterfall named HOME. Contains pivot tables, summaries, and configuration data for the waterfall.\
`Filter Tab` - the worksheet on the waterfall named FILTER. Contains a standardized dataset for the mailing, containing each account and various fields describing its status 
`Mapping` - the process of geocoding addresses and using the results to determine eligibility

## Utilities
There are currently 9 utilites (7 in OH + 2 in IL) supported.

* OH
  * AES Ohio (formerly Dayton Power and Light) `AES`
  * American Electric Power
    * Ohio Power Company `OP`
    * Columbus Southern Power `CS`
  * Duke Energy `DUKE`
  * FirstEnergy
    * Ohio Edison `OE`
    * The Illuminating Company `CEI`
    * Toledo Edison `TE`
* IL
  * Ameren Illinois `AM`
  * Commonwealth Edison `COM`
    
## Mailing Types
There are currently 4 different types of mailing supported.\

**Contract Renewal + Resweep** `CR`\
  Use a list of currently active customers from an expiring contract along with a utility list to enroll those currently active customers on a new contract, as well as enroll newly eligible customers from the utility list
  
**Resweep** `SWP`\
  Use a utility list to enroll eligible customers on an exiting contract
  
**New Community** `NEW`\
  Use a utility list to enroll eligible customers on a newly created aggregation program. Optionally, the community may not be new to aggregation just switching suppliers to Dynegy. In this case, use the provided list of prior active customers to enroll those customers on the new contract, as well as enroll any newly eligible customers from the utility list.
  
**Renewal Only** `REN_ONLY`\
  Use a list of currently active customers from an expiring contract with no utility list to enroll those currently active customers on a new contract.
  
**Term Renewal (Not Supported)** `TR`\
  (Need description)

## Using The Tool
Use of the tool can be described by a number of different ordered steps.

0. [Configuration](#0-configuration)
1. [Import Files](#1-import-files)
2. [Normalize Data](#2-normalize-data)
3. [Utility Data Exclusions](#3-utility-data-exclusions)
4. [External Data Exclusions](#4-external-data-exclusions)
5. [Export Data](#5-export-data)

### 0. Configuration
The tool has a `List Management` ribbon tab containing the relevant controls for use. 
Firstly, the `Save Waterfall` button prompts the user to save a copy of the tool as a separate waterfall file in the mailing folder for the appropriate community
Additionally, the user needs to use the dropdowns to select the appropriate `Utility` and `Mail Type`
Below those dropdowns is a place to enter the community name. This is the formal name of the community (City of Akron, Summmit County, etc.)

### 1. Import Files
After the `Utility`, `Mail Type`, and `Community Name` have been provided, the `Import File` button will unlock and provide the appropriate options based on `Mail Type'.
When a file is selected, it is automatically opened and the data is added to a separate worksheet.
If the selection inclues more than one worksheet containing utility data, these are combined into a single sheet.
All required files must be imported before proceeding to the next step.
>[!Note]
>These files are aquired through different processes not described here.
* Contract Renewal `CR`
  *  Active Customer List (from LandPower)
  *  Utility List(s)
* Resweep `SWP`
  *  Utility List(s)
* New Community `NEW`
  *  Utility List(s)
  *  Previous Supplier List (Optional)
* Renewal Only `REN_ONLY`
  *  Active Customer List (from LandPower)

### 2. Normalize Data
After all required files have been imported, the utility data (if present) is transferred to the `Filter Tab`, which contains all the rows from the utility file, but the columns are standardized.
This enables us to easily read data from every utility the exact same way. Details about the various columns can be found [here](FILTER_TAB.md).
If a field is not provided on the utility list, it is marked with the standard `-` to denote an empty cell.
If any lists are applcable besides the utility list, those records are checked against the existing list and any mismatches are added, populating any fields that are available on the list.

### 3. Utility Data Exclusions
Depending on the operating state (`OH` or `IL`) there are different rules that govern which accounts are eligible to enroll.
Using the available data on the `Filter Tab`, various fields are checked against the eligibility condition for that state and the `Status` of each account is updated accordingly.
The order those fields are checked in and the conditions for eligibility can be found [here](#filter-order)

### 4. External Data Exclusions
In addition to the rules put forth by the operating state, there are requirements to check outside sources of data to further trim the eligible population after the initial checks.

### Do Not Aggregate `DNA`
In `OH`, there is a list maintained by the PUCO of customers wishing to not be included in any aggregation programs.
Called the `Do Not Aggregate List`, or `DNA List`, the tool will check the eligible accounts against a recent copy of the `DNA List` and provide the user with a sheet of potential matches for manual comparison.
Due to less-than-ideal data on this list, there are 2 kinds of searches carried out in this step.
First is an account number match, meaning that if an account number is present on our eligible list as well as the `DNA List`, it is added to the comparison.
Second is the more wide-reaching address wildcard search.
The first `12` characters of the service address are compared to the first `12` characters of those addresses on the list and any matches are added to the comparison.
This is done to prevent any misspellings from affecting the accuracy of our search (`168 E Market St` vs `168 E Market Street` would both use `168 E Market` as the search term).
This comparison includes the customer name, service address, and account number. The user uses `Y` and `N` to notate which instances are truly matching.
### 5. Export Data



## Special Cases

