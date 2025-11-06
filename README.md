# List-Management-Tool

&copy; 2025 Andrew Rodgers

## About

This tool is used to determine eligibility for accounts on a list for a specified utility `EDC` and mailing type `MailType` using data from the file(s) provided, external data, and manual input to determine the list of eligible accounts.
There are many special conditions and logic features built in for each EDC and MailType, which will be summarized in a later section (see [Special Cases](#special-cases)).
This is intended to be a summary suitable for management and new users.
Detailed technical documentation can be found [here] <!-- [here] (TECHNICAL_DOC.md) -->

## Utilities
There are currently 9 utilites (7 in OH + 2 in IL) supported.
* OH
  * AES Ohio (formerly Dayton Power and Light) - AES
  * American Electric Power
    * Ohio Power Company - OP
    * Columbus Southern Power - CS
  * Duke Energy - DUKE
  * FirstEnergy
    * Ohio Edison - OE
    * The Illuminating Company - CEI
    * Toledo Edison - TE
* IL
  * Ameren Illinois - AM
  * Commonwealth Edison - COM
    
## Mailing Types
There are currently 4 different types of mailing supported.\
**Contract Renewal + Resweep** `REN` or `CR`\
  Use a list of currently active customers from an expiring contract along with a utility list to enroll those currently active customers on a new contract, as well as enroll newly eligible customers from the utility list\
**Resweep** `SWP`\
  Use a utility list to enroll eligible customers on an exiting contract\
**New Community** `NEW`\
  Use a utility list to enroll eligible customers on a newly created aggregation program. Optionally, the community may not be new to aggregation just switching suppliers to Dynegy. In this case, use the provided list of prior active customers to enroll those customers on the new contract, as well as enroll any newly eligible customers from the utility list.\
**Renewal Only** `REN_ONLY`\
  Use a list of currently active customers from an expiring contract with no utility list to enroll those currently active customers on a new contract.\





## Special Cases

