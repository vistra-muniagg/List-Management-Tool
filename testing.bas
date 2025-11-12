Public Type TestCase
    active_list As String
    gagg_list As String
    contracts_file As String
    mapping_file As String
    name As String
    results_folder_path As String
    mail_type As String
    community As String
End Type

Sub define_test_case(n, mail_type)
    name_array = Array("OP", "AES", "AM", "COM", "DUKE", "OE")
    active_file_array = Array("OP.CSV", "AES.CSV", "AM.CSV", "COM.CSV", "DUKE.CSV", "OE.CSV")
    gagg_file_array = Array("OP.CSV", "AES.CSV", "AM.XLS", "COM.XLSX", "DUKE.XLSX", "OE.XLSX")
    With T
        .active_list = active_file_array(n)
        .gagg_list = gagg_file_array(n)
        .contracts_file = .active_list
        .mapping_file = name_array(n) & "_" & mail_type & ".XLSM"
        .name = name_array(n)
        .mail_type = mail_type
        .community = "City of Harrison"
        .results_folder_path = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(2) Macro Testing\(3) Testing Results\" & results_folder_name()
    End With
End Sub

Sub test(Optional k, Optional mail_type, Optional export_results = True)
    'ThisWorkbook.Save
    start_time = Timer
    reset
    If IsMissing(k) Then k = 1
    Call init(k, mail_type)
    set_community_name (T.community)
    test_import
    test_pre
    test_active
    format_address_data
    filter_list
    test_dna (12)
    test_contracts
    'test_migration
    test_mapping
    misc_filter
    make_filter_waterfall
    make_geocode_waterfall
    make_cycle_waterfall
    review_eligible_data
    'make_LP
    'ren_drops
    'make_opt_in_list
    'make_mail_list
    'If export_results Then export_test_results
    progress.finish
End Sub

Sub test_all()
    For Each i In Array(4, 5)
        Call test(i, "REN")
        Call test(i, "SWP")
        Call test(i, "REN_ONLY")
    Next
End Sub

Sub reset()
    Application.ScreenUpdating = False
    all_initialized = False
    EDC = empty_EDC()
    MT = empty_MT()
    define_sheet_names
    define_UI_settings
    define_QC_settings
    define_HOME_settings
    With home_tab()
        .Range(S.HOME.info_range) = ""
        .Range(S.HOME.file_log_range) = ""
        .Range(S.HOME.file_log_range).ClearComments
        .Range(S.HOME.community_name_location) = "(Community Name)"
        .Range(S.QC.qc_checklist.data_range) = ""
        .Range(S.QC.audit_checklist.data_range) = ""
        .Range(S.HOME.renewal_drop_count_location) = ""
        .Range(S.HOME.renewal_drop_count_location).Offset(0, -1) = ""
    End With
    Call set_step(0)
    ribbon_community = ""
    ribbon_contract_number = ""
    ribbon_opt_out_date = ""
    For Each ws In ThisWorkbook.Sheets
        If ws.name <> SN.README And ws.name <> SN.HOME Then
            delete_sheet (ws.name)
        End If
    Next
    For Each pvt In home_tab().PivotTables
        pvt.TableRange2.Clear
    Next
    
    imported_gagg = False
    imported_active = False
    imported_supplier = False
    
    home_tab().Activate
    
    Range("A1").Activate
    
    If Not UI Is Nothing Then UI.Invalidate
    
    Application.ScreenUpdating = True
    
End Sub

Function results_folder_name()
    Select Case MT.name
        Case "NEW": results_folder_name = "(1) New Community\"
        Case "SWP": results_folder_name = "(2) Sweep\"
        Case "REN": results_folder_name = "(3) Renewal\"
        Case "REN_ONLY": results_folder_name = "(4) Renewal (No Sweep)\"
    End Select
End Function

Function test_results_file_name()
    test_results_file_name = EDC.name & "_" & MT.name & "_TEST.xlsx"
End Function

Sub export_test_results()
    
    Dim wb As Workbook
    Dim wb_test As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    
    folder_path = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(2) Macro Testing\(3) Testing Results" & "\" & results_folder_name
    file_name = folder_path & test_results_file_name()
    
    progress.start ("Exporting Test Results")
    
    Set wb_test = Workbooks.Add
    
    For Each ws In wb.Worksheets
        ws.Copy after:=wb_test.Sheets(wb_test.Sheets.count)
    Next
    
    Application.DisplayAlerts = False
    Do While wb_test.Sheets.count > wb.Sheets.count
        wb_test.Sheets(1).Delete
    Loop
    Application.DisplayAlerts = True
    
    wb_test.SaveAs fileName:=file_name, FileFormat:=xlOpenXMLWorkbook
    wb_test.Close False
    
    ThisWorkbook.Activate
    
    progress.finish
    
End Sub

Sub test_onedrive()
    init
    
    x1 = onedrive_parent_folder()
    x2 = onedrive_list_management_folder
    x3 = onedrive_mailings_folder()
    x4 = onedrive_dna_folder()
    x5 = onedrive_mapping_db_folder()
    x6 = onedrive_migration_folder()
    x7 = onedrive_documentation_folder()
    
    x = Application.UserName & vbCrLf & _
        "OneDrive Folder: " & x1 & vbCrLf & _
        "List Management: " & x2 & vbCrLf & _
        "Mailings: " & x3 & vbCrLf & _
        "DNA: " & x4 & vbCrLf
        '"Mapping: " & x5 & vbCrLf & _
        '"Migration: " & x6 & vbCrLf & _
        '"Documentation: " & x7 & vbCrLf
        
    test_pass = x1 <> "" And x2 <> "" And x3 <> "" And x4 <> ""
    
    If test_pass Then
        MsgBox "PASS"
    Else
        MsgBox "FAIL"
    End If
    
    msg = string_to_html(x)
    
    send_error_message_teams (msg)
    
End Sub

Sub t4()
    Call test(4, "REN", False)
End Sub
