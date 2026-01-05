# List-Management-Tool `Draft #1 (11-10-25)`

&copy; 2025 Andrew Rodgers

## About

This tool is used to determine eligibility for accounts on a list for a specified utility `EDC` and mailing type `MailType` using data from the file(s) provided, external data, and manual input to determine the list of eligible accounts.
There are many special conditions and logic features built in for each EDC and MailType, which will be summarized in a later section (see [Special Cases](#special-cases)).
This is intended to be a summary suitable for management and new users.
There is a master copy of this tool saved in Sharepoint, and each mailing creates a copy in its resepective folder, containing the relevant data for that community.
Detailed technical documentation can be found [here] <!-- [here] (TECHNICAL_DOC.md) -->
Process diagrams can be found [here] <!-- (DIAGRAMS.md) -->

## Definitions
`Waterfall` - the copy of this tool containing the pertinent data for a mailing. Also contains pivot tables and summary statistics about the mailing.\
`Filtering` - the process of removing ineligible accounts from any lists relevant to a mailing.\
`Status` - label desribing the eligibility of an account
`Home Tab` - the worksheet on the waterfall named HOME. Contains pivot tables, summaries, and configuration data for the waterfall.\
`Filter Tab` - the worksheet on the waterfall named FILTER. Contains a standardized dataset for the mailing, containing each account and various fields describing its status.\
`Geocoding` - the process of translating an address to a geographic location.\
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
  
**Term Renewal (Not Currently Supported)** `TR`\
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
This enables us to easily read data from every utility the exact same way. Details about the various columns can be found [here](#filter-tab).
If a field is not provided on the utility list, it is marked with the standard `-` to denote an empty cell.
If any lists are applcable besides the utility list, those records are checked against the existing list and any mismatches are added, populating any fields that are available on the list.

### 3. Utility Data Exclusions
Depending on the operating state (`OH` or `IL`) there are different rules that govern which accounts are eligible to enroll.
Using the available data on the `Filter Tab`, various fields are checked against the eligibility condition for that state and the `Status` of each account is updated accordingly.
The order those fields are checked in and the conditions for eligibility can be found [here](#eligibility-filter-order)

### 4. External Data Exclusions
In addition to the rules put forth by the operating state, there are requirements to check outside sources of data to further trim the eligible population after the initial checks.

#### Do Not Aggregate (DNA)
In OH, the PUCO maintains a list of customers wishing to not be included in any aggregation programs, called the `Do Not Aggregate List`, or `DNA List`.
The tool checks the eligible accounts against a recent copy of the `DNA List` and provides the user with a sheet of potential matches for manual comparison.
Due to less-than-ideal data on this list, there are 2 kinds of searches carried out in this step.
First is an account number match, meaning that if an account number is present on our eligible list as well as the `DNA List`, it is added to the comparison.
Second is the more wide-reaching address wildcard search.
The first `12` characters of the service address are compared to the first `12` characters of those addresses on the list and any matches are added to the comparison.
This is done to prevent any misspellings from affecting the accuracy of our search (`168 E Market St` vs `168 E Market Street` would both use `168 E Market` as the search term).
This comparison includes the customer name, service address, and account number.
The user uses `Y` and `N` to notate which instances are truly matching.

#### LandPower System Comparison (Contracts Query)
Due to the time delay between a customer enrolling and the utility file being updated to reflect the new shopping status, it is necessary to query the Vistra systems to ensure the account is not participating in any electric choice programs offered by Vistra.
In order to query all account activity in the Vistra systems, the eligible account numbers are put into a SQL query that returns any enrollemnt activity matching these account numbers. This query is currently run manually by the user, but may be automated in a future version.
>[!NOTE]
>There is a known gap in this step. The query only has access to accounts in `Mass Market` and `Muni-agg` programs, so any activity in another sales channel will not be returned in the query results.

Using only the most recent activity on each account, the current enrollment status is determined. If an account is active on any other Vistra programs, its `Status` is updated accordingly and it is excluded from the current mailing.

#### Geocoding (Mapping)
Legal has mandated that only customers located in the eligible boundaries for the program can be enrolled.
Due to the time required to geocode an entire list, there is another tool (see [Mapping Tool](MAPPING_TOOL.md) specifically used for geocoding.
For each account, the service address is mapped and the customer is determined to be inside the geographic boundaries of the program (`Mapped In`) or outside (`Mapped Out`).
Any accounts that are `Mapped Out` are ineligible for enrollment, and all other accounts are eligible.
Using the results from this mapping, each `Status` is updated accordingly to reflect the geographic eligibility of an account and the geocoded community is added to the `Filter Tab` for tracking.

### 5. Create Upload File
When all eligibility filter have been applied to disposition accounts as either `Eligible` or `Ineligible`, the user enters the mailing details for this list in order to create the upload file for the LandPower system.
These details include Community Name (already entered), Contract Number (C-00XXXXXX), and Opt-Out Date (MM/DD/YY).
The LandPower upload file is based on the template for the Muni-Agg File Processor, so the normalized data columns from the `Filter Tab`, as well as the mailing details, can be used to populate required fields.
The user is prompted to review a subset of these fields for accuracy and must make corrections if the data does not meet the quality standards required for upload to LandPower.
Once data quality is confirmed, the user must save and send the waterfall file to another analyst for `Peer Review`.

### 6. Export Data
There is a [Peer Review Checklist](#peer-review-checklist) for the analyst doing peer review to follow.
When all items on the checklist are complete, the peer reivew analyst can click the `Export Files` button.
This will create the upload file required for LandPower, as well as exporting a specially formatted Mail List, which will be sent to the printer after upload.
In some cases, there will be additional files created.
The possible files created and saved in the community mailing folder include the following,

* Email to Printer
* LandPower Upload Email
* Opt-In Mail List
* Drops at Renewal List

## Special Cases

## Filter Tab Columns
|Column Name|Data Type|Description|
|-------|---------|---------|
|Account Number|string|duh|
|Status|text|label reflecting the specific eligibility of an account|
|Eligible Opt-Out|Y/N|whether or not account is eligible for Opt-Out mailing|
|Mail Category|text|New enrollemnt (NEW) or price change (REN)|
|On GAGG List|Y/N|duh|
|On Active List|Y/N|duh|
|Active List Mismatch|Y/N|active list account not present on gagg list|
|SAS ID|text|subaccountserviceid for renewal accounts|
|LP Premise Mismatch|Y/N|if premise in LP does not match gagg list|
|Address Source|text|most recent source of address, VISTRA for mismatch account, GAGG for all else|
|COMMUNITY|text|commnuity name|
|Mapping Result|text|whether account maps in or out of commnuity|
|Community Mapped Into|text|geographic area account service address maps in to|
|Eligible Before Mapping|Y/N|whether or not account was eligible before mapping results were applied|
|Mapping Notes|text|any notes from the mapping process about the account disposition|
|Hourly Pricing|Y/N|IL only, mirrors the hourly pricing flag from the gagg list|
|RTP|Y/N|IL only, mirrors the real time pricing flag from the gagg list|
|BGS Hold|Y/N|IL only, mirrors the BGS Hold flag from the gagg list|
|Commnuity Solar|Y/N|IL only, mirrors the Community Solar flag from the gagg list|
|Free Service|Y/N|IL only, mirrors the Free Service flag from the gagg list|
|PIPP|Y/N|OH only, mirrors the PIPP flag from the gagg list|
|Mercantile|Y/N|OH only, mirrors the Mercantile flag from the gagg list|
|Usage Months|number|count of monthly usage data present on the gagg list|
|Actual Usage|number|real sum of monthly usage from the gagg list|
|Estimated Usage|number|annualized usage if any months have no usage data present|
|Shopping|Y/N|mirrors the Shopping flag on the gagg list|
|Arrears|Y/N|mirrors the Arrears flag on the gagg list|
|National Chain|Y/N|only Y for a commercial account mailing to Spokane, WA|
|Do Not Agg|Y/N|only Y if the account is present on the PUCO DNA list|
|Opt In Eligible|Y/N|whether or not account is eligible for Opt-In mailing|
|LP Contracts Query|text|describes the status of the account in LP|
|Recent Contract|text|most recent community contract id the account has activity on|
|Contract Start Date|date|most recent start date on the most recent contract id for the account|
|Customer Name|text|clean version of the account name|
|Service Address|text|clean version of the account service addres|
|Service City|text|clean version of the account service city|
|Service State|text|clean version of the account service state|
|Service Zip|text|clean version of the account service zip code|
|Mail Address|text|clean version of the account mailing address|
|Mail City|text|clean version of the account mailing city|
|Mail State|text|clean version of the account mailing state|
|Mail Zip|text|clean version of the account mailing zip code|
|Phone|number|account phone number if present|
|Email|text|account email address if present|
|Customer Class|text|Residential or Commercial based on gagg list rate code or LP premise|
|Read Cycle|number|detected read cycle from the gagg list|
## Eligibility Filter Order

## Peer Review Checklist

## Known Issues
