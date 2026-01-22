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
    budget_bill As String
    budget_bill_yes As String
    solar As String
    solar_yes As String
    free_service As String
    free_service_yes As String
    hourly_pricing As String
    hourly_pricing_yes As String
    rtp As String
    rtp_yes As String
    bgs As String
    bgs_yes As String
    'other column headers
    account As String
    customer_name As String
    phone As String
    email As String
    rate_code As String
    read_cycle As String
    usage As String
    'address columns
    service As Variant
    service_city As String
    service_state As String
    service_zip As String
    mail() As Variant
    mail_city As String
    mail_state As String
    mail_zip As String
    po_box As String
End Type

Sub define_EDC(EDC_name)
    Select Case (EDC_name)
        Case "OE": define_EDC_OE
        Case "TE": define_EDC_TE
        Case "CEI": define_EDC_CEI
        Case "OP": define_EDC_OP
        Case "CS": define_EDC_CS
        Case "AES": define_EDC_AES
        Case "DUKE": define_EDC_DUKE
        Case "AM": define_EDC_AM
        Case "COM": define_EDC_COM
        Case Else: Exit Sub
    End Select
    x = home_tab().Range(S.HOME.edc_location)
    If x <> EDC_name Then home_tab().Range(S.HOME.edc_location) = EDC_name
    'If MT.name <> "" Then Call set_step(1)
    ribbon_EDC = EDC.name
    If Not UI Is Nothing Then
        UI.InvalidateControl ("import_menu")
        UI.InvalidateControl ("EDC_menu")
    End If
    define_filter_tab_columns
    add_user_name
End Sub

Sub define_EDC_OE()
    With EDC
        'system names
        .name = "OE"
        .ruleset_name = "FE"
        .full_name = "Ohio Edison Company"
        .CRM_name = "Ohio Edison Company"
        .SF_name = "FEOHIO"
        .LP_name = "FirstEnergy Ohio - Ohio Edison"
        .migration_name = "OE"
        .display_name = "FE-OE"
        .DNA_name = ""
        .mapping_name = "FE-OE"
        'import settings
        .multiselect = True
        .import_all_gagg_sheets = True
        'validation
        .state = "OH"
        .file_formats = Array("XLS", "XLSX")
        .account_number_length = 20
        .leading_pattern = "080*"
        .zeros_to_add = 1
        .valid_codes = Array("OE-GSD", "OE-GSF", "OE-RSD", "OE-RSF")
        .res_codes = Array("OE-RSD", "OE-RSF")
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)
        .read_cycle = "RD"
        .default_read_cycle = 1
        'filter headers
        .account = "CUSTOMER NUMBER"
        .customer_name = "CUSTOMER NAME"
        .rate_code = "RATE_CD"
        .shopper = "SHOPPING IND"
        .pipp = "INSTALLMENT PLAN"
        .net_meter = ""
        .mercantile = ""
        .arrears = "IN ARREARS"
        .budget_bill = "BUDGET BILL"
        'filter conditions
        .shopper_yes = "Y"
        .pipp_yes = "Y"
        .net_meter_yes = "Y"
        .mercantile_yes = "Y"
        .arrears_yes = "Y"
        .budget_bill_yes = "Y"
        'IL filters
        .solar = ""
        .free_service = ""
        .hourly_pricing = ""
        .rtp = ""
        .bgs = ""
        'IL filter conditions
        .solar_yes = ""
        .free_service_yes = ""
        .hourly_pricing_yes = ""
        .rtp_yes = ""
        .bgs_yes = ""
        'usage filter
        .usage = "KWH1"
        .usage_limit = ">700000"
        .usage_multiple = 3
        'mailing info
        .service = Array("SERVICE ADDRESS1", "SERVICE ADDRESS2")
        .service_city = "SERV_ADDR5"
        .service_state = "SERV_ADDR5"
        .service_zip = "SERV_ADDR5"
        .mail = Array("MAIL_ADDR_HOUSE_NUM", "MAIL_ADDR_STREET_NAME", "MAILING ADDRESS - PART 2")
        .mail_city = "MAILING CITY"
        .mail_state = "ST"
        .mail_zip = "MAIL ZIP"
        'optional mailing info
        .po_box = ""
        .phone = "PHONE_NUM"
        .email = ""
    End With
End Sub

Sub define_EDC_CEI()
    With EDC
        'system names
        .name = "CEI"
        .ruleset_name = "FE"
        .full_name = "The Cleveland Electric Illuminating Co."
        .CRM_name = "Cleveland Electric Illuminating Company"
        .SF_name = "FEOHIO"
        .LP_name = "FirstEnergy Ohio - Illuminating Company"
        .migration_name = "CEI"
        .display_name = "FE-CEI"
        .mapping_name = "FE-CEI"
        .DNA_name = ""
        'import settings
        .multiselect = True
        .import_all_gagg_sheets = True
        'validation
        .state = "OH"
        .file_formats = Array("XLS", "XLSX")
        .account_number_length = 20
        .leading_pattern = "080*"
        .zeros_to_add = 1
        .valid_codes = Array("CE-GSD", "CE-GSF", "CE-RSD", "CE-RSF")
        .res_codes = Array("CE-RSD", "CE-RSF")
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)
        .read_cycle = "RD"
        .default_read_cycle = 1
        'filter headers
        .account = "CUSTOMER NUMBER"
        .customer_name = "CUSTOMER NAME"
        .rate_code = "RATE_CD"
        .shopper = "SHOPPING IND"
        .pipp = "INSTALLMENT PLAN"
        .net_meter = ""
        .mercantile = ""
        .arrears = "IN ARREARS"
        .budget_bill = "BUDGET BILL"
        'filter conditions
        .shopper_yes = "Y"
        .pipp_yes = "Y"
        .net_meter_yes = "Y"
        .mercantile_yes = "Y"
        .arrears_yes = "Y"
        .budget_bill_yes = "Y"
        'IL filters
        .solar = ""
        .free_service = ""
        .hourly_pricing = ""
        .rtp = ""
        .bgs = ""
        'IL filter conditions
        .solar_yes = ""
        .free_service_yes = ""
        .hourly_pricing_yes = ""
        .rtp_yes = ""
        .bgs_yes = ""
        'usage filter
        .usage = "KWH1"
        .usage_limit = ">700000"
        .usage_multiple = 3
        'mailing info
        .service = Array("SERVICE ADDRESS1", "SERVICE ADDRESS2")
        .service_city = "SERV_ADDR5"
        .service_state = "SERV_ADDR5"
        .service_zip = "SERV_ADDR5"
        .mail = Array("MAIL_ADDR_HOUSE_NUM", "MAILING ADDRESS - PART 2")
        .mail_city = "MAILING CITY"
        .mail_state = "ST"
        .mail_zip = "MAIL ZIP"
        'optional mailing info
        .po_box = ""
        .phone = "PHONE_NUM"
        .email = ""
    End With
End Sub

Sub define_EDC_TE()
    With EDC
        'system names
        .name = "TE"
        .ruleset_name = "FE"
        .full_name = "The Toledo Edison Company"
        .CRM_name = "Toledo Edison Company"
        .SF_name = "FEOHIO"
        .LP_name = "FirstEnergy Ohio - Toledo Edison"
        .migration_name = "TE"
        .display_name = "FE-TE"
        .DNA_name = ""
        'import settings
        .multiselect = True
        .import_all_gagg_sheets = True
        'validation
        .state = "OH"
        .file_formats = Array("XLS", "XLSX")
        .account_number_length = 20
        .leading_pattern = "080*"
        .zeros_to_add = 1
        .valid_codes = Array("TE-GSD", "TE-GSF", "TE-RSD", "TE-RSF")
        .res_codes = Array("TE-RSD", "TE-RSF")
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)
        .read_cycle = "RD"
        .default_read_cycle = 1
        'filter headers
        .account = "CUSTOMER NUMBER"
        .customer_name = "CUSTOMER NAME"
        .rate_code = "RATE_CD"
        .shopper = "SHOPPING IND"
        .pipp = "INSTALLMENT PLAN"
        .net_meter = ""
        .mercantile = ""
        .arrears = "IN ARREARS"
        .budget_bill = "BUDGET BILL"
        'filter conditions
        .shopper_yes = "Y"
        .pipp_yes = "Y"
        .net_meter_yes = "Y"
        .mercantile_yes = "Y"
        .arrears_yes = "Y"
        .budget_bill_yes = "Y"
        'IL filters
        .solar = ""
        .free_service = ""
        .hourly_pricing = ""
        .rtp = ""
        .bgs = ""
        'IL filter conditions
        .solar_yes = ""
        .free_service_yes = ""
        .hourly_pricing_yes = ""
        .rtp_yes = ""
        .bgs_yes = ""
        'usage filter
        .usage = "KWH1"
        .usage_limit = ">700000"
        .usage_multiple = 3
        'mailing info
        .service = Array("SERVICE ADDRESS1", "SERVICE ADDRESS2")
        .service_city = "SERV_ADDR5"
        .service_state = "SERV_ADDR5"
        .service_zip = "SERV_ADDR5"
        .mail = Array("MAIL_ADDR_HOUSE_NUM", "MAILING ADDRESS - PART 2")
        .mail_city = "MAILING CITY"
        .mail_state = "ST"
        .mail_zip = "MAIL ZIP"
        'optional mailing info
        .po_box = ""
        .phone = "PHONE_NUM"
        .email = ""
    End With
End Sub

Sub define_EDC_OP()
    With EDC
        'system names
        .name = "OP"
        .ruleset_name = "AEP"
        .full_name = "Ohio Power Company"
        .CRM_name = "Ohio Power Company"
        .SF_name = "AEPOHIO"
        .LP_name = "AEP Ohio- Ohio Power"
        .migration_name = "OP"
        .display_name = "AEP-OP"
        .DNA_name = ""
        'import settings
        .multiselect = True
        .import_all_gagg_sheets = True
        'validation
        .state = "OH"
        .file_formats = Array("CSV", "XLSX")
        .account_number_length = 17
        .leading_pattern = "001400607*"
        .zeros_to_add = 17
        .valid_codes = Array(15, 208, 211, 212, 215, 820, 830, 840)
        .res_codes = Array(15, 820)
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 30, 31, 32, 33)
        .read_cycle = "MTR-READ-CYCLE"
        .default_read_cycle = 1
        'filter headers
        .account = "SDID-NUMBER"
        .customer_name = "CUST-NAME"
        .rate_code = "TARF-CD"
        .shopper = "CUST-SWITCH-FL"
        .pipp = "PIP-FL"
        .net_meter = "NET-MTR-FL"
        .mercantile = "MERCANTILE-NAT-ACCT-FL"
        .arrears = ""
        .budget_bill = "BDGT-BILL-FL"
        'filter conditions
        .shopper_yes = "Y"
        .pipp_yes = "Y"
        .net_meter_yes = "Y"
        .mercantile_yes = "Y"
        .arrears_yes = ""
        .budget_bill_yes = "Y"
        'IL filters
        .solar = ""
        .free_service = ""
        .hourly_pricing = ""
        .rtp = ""
        .bgs = ""
        'IL filter conditions
        .solar_yes = ""
        .free_service_yes = ""
        .hourly_pricing_yes = ""
        .rtp_yes = ""
        .bgs_yes = ""
        'usage filter
        .usage = "KWH-USAGE"
        .usage_limit = ">700000"
        .usage_multiple = 4
        'mailing info
        .service = Array("SERV-ADDR-1", "SERV-ADDR-2")
        .service_city = "SERV-CITY"
        .service_state = "SERV-STATE"
        .service_zip = "SERV-ZIP"
        .mail = Array("MAIL-ADDR1", "MAIL-ADDR2")
        .mail_city = "MAIL-CITY"
        .mail_state = "MAIL-STATE"
        .mail_zip = "MAIL-ZIP"
        'optional mailing info
        .po_box = ""
        .phone = ""
        .email = ""
    End With
End Sub

Sub define_EDC_CS()
    With EDC
        'system names
        .name = "CS"
        .ruleset_name = "AEP"
        .full_name = "Columbus Southern Power Company"
        .CRM_name = "Columbus Southern Power Company"
        .SF_name = "AEPOHIO"
        .LP_name = "AEP Ohio - Columbus Southern Power"
        .migration_name = "CS"
        .display_name = "AEP-CS"
        .DNA_name = ""
        'import settings
        .multiselect = True
        .import_all_gagg_sheets = True
        'validation
        .state = "OH"
        .file_formats = Array("CSV", "XLSX")
        .account_number_length = 17
        .leading_pattern = "000406210*"
        .zeros_to_add = 17
        .valid_codes = Array(15, 208, 211, 212, 215, 820, 830, 840)
        .res_codes = Array(15, 820)
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 30, 31, 32, 33)
        .read_cycle = "MTR-READ-CYCLE"
        .default_read_cycle = 1
        'filter headers
        .account = "SDID-NUMBER"
        .customer_name = "CUST-NAME"
        .rate_code = "TARF-CD"
        .shopper = "CUST-SWITCH-FL"
        .pipp = "PIP-FL"
        .net_meter = "NET-MTR-FL"
        .mercantile = "MERCANTILE-NAT-ACCT-FL"
        .arrears = ""
        .budget_bill = "BDGT-BILL-FL"
        'filter conditions
        .shopper_yes = "Y"
        .pipp_yes = "Y"
        .net_meter_yes = "Y"
        .mercantile_yes = "Y"
        .arrears_yes = ""
        .budget_bill_yes = "Y"
        'IL filters
        .solar = ""
        .free_service = ""
        .hourly_pricing = ""
        .rtp = ""
        .bgs = ""
        'IL filter conditions
        .solar_yes = ""
        .free_service_yes = ""
        .hourly_pricing_yes = ""
        .rtp_yes = ""
        .bgs_yes = ""
        'usage filter
        .usage = "KWH-USAGE"
        .usage_limit = ">700000"
        .usage_multiple = 3
        'mailing info
        .service = Array("SERV-ADDR-1", "SERV-ADDR-2")
        .service_city = "SERV-CITY"
        .service_state = "SERV-STATE"
        .service_zip = "SERV-ZIP"
        .mail = Array("MAIL-ADDR1", "MAIL-ADDR2")
        .mail_city = "MAIL-CITY"
        .mail_state = "MAIL-STATE"
        .mail_zip = "MAIL-ZIP"
        'optional mailing info
        .po_box = ""
        .phone = ""
        .email = ""
    End With
End Sub

Sub define_EDC_AES()
    With EDC
        'system names
        .name = "AES"
        .ruleset_name = "AES"
        .full_name = "AES - Ohio"
        .CRM_name = "Dayton Power and Light Company"
        .SF_name = "DAYTON"
        .LP_name = "AES Ohio"
        .migration_name = "AES"
        .display_name = "AES"
        .DNA_name = ""
        'import settings
        .multiselect = True
        .import_all_gagg_sheets = True
        'validation
        .state = "OH"
        .file_formats = Array("CSV")
        .account_number_length = 23
        .leading_pattern = "*"
        .zeros_to_add = 0
        .valid_codes = Array(25, 26, 45, 46, 97, 111, 117, 121, 137, 141, 711, 717, 721, 737, 741)
        .res_codes = Array("ORES", "ORES-H")
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22)
        .read_cycle = "BILL UNIT"
        .default_read_cycle = 1
        'filter headers
        .account = "ACCT NUMBER"
        .customer_name = "CUSTOMER SVC NAME 1"
        .rate_code = "RATE"
        .shopper = "SWITCHED"
        .pipp = "PIPP"
        .net_meter = "NET METER"
        .mercantile = "MERCANTILE"
        .arrears = "TOTAL ARREARS"
        .budget_bill = "BUDGET BILL"
        'filter conditions
        .shopper_yes = "Y"
        .pipp_yes = "Y"
        .net_meter_yes = "Y"
        .mercantile_yes = "Y"
        .arrears_yes = ">200"
        .budget_bill_yes = "Y"
        'IL filters
        .solar = ""
        .free_service = ""
        .hourly_pricing = ""
        .rtp = ""
        .bgs = ""
        'IL filter conditions
        .solar_yes = ""
        .free_service_yes = ""
        .hourly_pricing_yes = ""
        .rtp_yes = ""
        .bgs_yes = ""
        'usage filter
        .usage = "KWH 1"
        .usage_limit = ">700000"
        .usage_multiple = 10
        'mailing info
        .service = Array("SERVICE ADDRESS", "SUPPL SERV ADDR")
        .service_city = "CITY ST ZIP"
        .service_state = "CITY ST ZIP"
        .service_zip = "ZIP PLUS 4"
        .mail = Array("MAILING ADDRESS", "SUPPL MAIL ADDR")
        .mail_city = "MAILING CITY ST ZIP"
        .mail_state = "MAILING CITY ST ZIP"
        .mail_zip = "MAIL ZIP PLUS 4"
        'optional mailing info
        .po_box = ""
        .phone = "PHONE NUMBER"
        .email = ""
    End With
    
End Sub

Sub define_EDC_DUKE()
    With EDC
        'system names
        .name = "DUKE"
        .ruleset_name = "DUKE"
        .full_name = "Duke Energy Ohio, Inc."
        .CRM_name = "Duke Energy Ohio, Inc."
        .SF_name = "DEOHIO"
        .LP_name = "Duke Energy Ohio - Electric"
        .migration_name = "DUKE"
        .display_name = "DUKE"
        .DNA_name = ""
        'import settings
        .multiselect = True
        .import_all_gagg_sheets = True
        'validation
        .state = "OH"
        .file_formats = Array("XLSX", "XLSB")
        .account_number_length = 22
        .leading_pattern = "############Z#########"
        .zeros_to_add = 0
        .valid_codes = Array("DM0", "DS4", "RS0", "RS5", "RS6")
        .res_codes = Array("RS0", "RS5", "RS6")
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
        .read_cycle = "Meter Reading Cycle"
        .default_read_cycle = 1
        'filter headers
        .account = "Choice Service ID"
        .customer_name = "Customer"
        .rate_code = "Load Profile Class"
        .shopper = "Supplier Indicator"
        .pipp = "Budget Bill PIPP Indicator"
        .net_meter = "Net Metering Indicator"
        .mercantile = "Mercantile Flag"
        .arrears = ""
        .budget_bill = ""
        'filter conditions
        .shopper_yes = "Y"
        .pipp_yes = "PIP"
        .net_meter_yes = "Y"
        .mercantile_yes = "X"
        .arrears_yes = ""
        .budget_bill_yes = ""
        'IL filters
        .solar = ""
        .free_service = ""
        .hourly_pricing = ""
        .rtp = ""
        .bgs = ""
        'IL filter conditions
        .solar_yes = ""
        .free_service_yes = ""
        .hourly_pricing_yes = ""
        .rtp_yes = ""
        .bgs_yes = ""
        'usage filter
        .usage = "Billed Kwh1"
        .usage_limit = ">700000"
        .usage_multiple = 5
        'mailing info
        .service = Array("House Address", "Service Street Address", "Apt", "Floor")
        .service_city = "Service City"
        .service_state = "Service State"
        .service_zip = "Service Zip Code"
        .mail = Array("Mailing Address 1", "Mailing Address 2")
        .mail_city = "Mailing Address 3"
        .mail_state = "Mailing Address 3"
        .mail_zip = "Mailing Address Zip Code"
        'optional mailing info
        .po_box = ""
        .phone = ""
        .email = "Email Address"
    End With
End Sub

Sub define_EDC_AM()
    With EDC
        'system names
        .name = "AM"
        .ruleset_name = "AM"
        .full_name = "Ameren Illinois Company"
        .CRM_name = "Ameren Illinois Company"
        .SF_name = "Ameren-IL"
        .LP_name = "Ameren Electric IL"
        .migration_name = "AM"
        .display_name = "AM"
        .DNA_name = ""
        'import settings
        .multiselect = False
        .import_all_gagg_sheets = True
        'validation
        .state = "IL"
        .file_formats = Array("XLS", "XLSX")
        .account_number_length = 10
        .leading_pattern = "*"
        .zeros_to_add = 10
        .valid_codes = Array("DS1", "DS2", "DS3", "DS5")
        .res_codes = Array("DS1")
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
        .read_cycle = "Bill Group"
        .default_read_cycle = 1
        'filter headers
        .account = "Account Number"
        .customer_name = "Account Name"
        .rate_code = "DS Classification"
        .shopper = "Supply Type"
        .pipp = ""
        .net_meter = "Ride NM?"
        .mercantile = ""
        .arrears = ""
        .budget_bill = ""
        'filter conditions
        .shopper_yes = "RES"
        .pipp_yes = ""
        .net_meter_yes = "Y"
        .mercantile_yes = ""
        .arrears_yes = ""
        .budget_bill_yes = ""
        'IL filters
        .solar = "Rider QF?"
        .free_service = ""
        .hourly_pricing = "Hourly"
        .rtp = "Supply Type"
        .bgs = "BGS Hold?"
        'IL filter conditions
        .solar_yes = "Y"
        .free_service_yes = ""
        .hourly_pricing_yes = "Y"
        .rtp_yes = "RTP"
        .bgs_yes = "Y"
        'usage filter
        .usage = "Total Usage"
        .usage_limit = ">15000"
        .usage_multiple = 0
        'mailing info
        .service = Array("Premise Address Line 1", "Premise Address Line 2")
        .service_city = "Premise Address City"
        .service_state = "Premise Address State"
        .service_zip = "Premise Address Zip Code"
        .mail = Array("Billing Address Line 1", "Billing Address Line 2", "Billing Address Line 3")
        .mail_city = "Billing Address City"
        .mail_state = "Billing Address State"
        .mail_zip = "Billing Address Zip Code"
        'optional mailing info
        .po_box = ""
        .phone = ""
        .email = ""
    End With
End Sub

Sub define_EDC_COM()
    With EDC
        'system names
        .name = "COM"
        .ruleset_name = "COM"
        .full_name = "Commonwealth Edison Company"
        .CRM_name = "ComEd"
        .SF_name = "ComEd-IL"
        .LP_name = "ComEd"
        .migration_name = "COM"
        .display_name = "COM"
        .DNA_name = ""
        'import settings
        .multiselect = False
        .import_all_gagg_sheets = True
        'validation
        .state = "IL"
        .file_formats = Array("XLSX")
        .account_number_length = 10
        .leading_pattern = "*"
        .zeros_to_add = 10
        .valid_codes = Array("B70", "B71", "B72", "B73", "B90", "B91", "B92", "B93", "R70", "R71", "R72", "R73", "R90", "R91", "R92", "R93", "H70", "H71", "H72", "H73", "H90", "H91", "H92", "H93")
        .res_codes = Array("B70", "B71", "B90", "B91", "R70", "R71", "R90", "R91", "H70", "H71", "H90", "H91")
        'read cycles
        .valid_cycles = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)
        .read_cycle = "Bill Group"
        .default_read_cycle = 1
        'filter headers
        .account = "Electric choice ID"
        .customer_name = "Customer Name"
        .rate_code = "Tariff Rate Nm"
        .shopper = "Status"
        .pipp = ""
        .net_meter = "Net Metering Indicator"
        .mercantile = ""
        .arrears = ""
        .budget_bill = "BB"
        'filter conditions
        .shopper_yes = "*RES Supply*"
        .pipp_yes = ""
        .net_meter_yes = "Y"
        .mercantile_yes = ""
        .arrears_yes = ""
        .budget_bill_yes = "Y"
        'IL filters
        .solar = "Community Supply"
        .free_service = "Free Service Indicator"
        .hourly_pricing = "Hourly"
        .rtp = ""
        .bgs = ""
        'IL filter conditions
        .solar_yes = "Y"
        .free_service_yes = "Y"
        .hourly_pricing_yes = "Y"
        .rtp_yes = ""
        .bgs_yes = ""
        'usage filter
        .usage = ""
        .usage_limit = ">15000"
        .usage_multiple = 0
        'mailing info
        .service = Array("Premise Address")
        .service_city = "Premise Address"
        .service_state = "Premise Address"
        .service_zip = "Premise Zip Code"
        .mail = Array("Mailing Address")
        .mail_city = "Mailing City"
        .mail_state = "Mailing State"
        .mail_zip = "Mailing Zip Code"
        'optional mailing info
        .po_box = ""
        .phone = ""
        .email = ""
    End With
End Sub

Function empty_EDC() As UtilitySettings
    'returns empty EDC
End Function
