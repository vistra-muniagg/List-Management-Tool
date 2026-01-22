Private settings_initialized As Boolean
Private EDC_initialized As Boolean
Private mail_type_initialized As Boolean
Private sheet_names_defined As Boolean
Private stats_initialized As Boolean
Public all_initialized As Boolean

Public imported_gagg As Boolean
Public imported_active As Boolean
Public imported_supplier As Boolean

Public all_reviewed As Boolean

Public ribbon_contract_number As String
Public ribbon_opt_out_date As String
Public ribbon_community As String
Public ribbon_EDC As String
Public ribbon_EDC_id As Long
Public ribbon_mail_type As String
Public ribbon_mail_type_id As Long

Public T As TestCase
Public F As FilterTab
Public SN As SheetNames
Public C As MacroColors
Public FS As FilterStatuses
Public S As MacroSettings
Public EDC As UtilitySettings
Public MT As MailType
Public A As ActiveList
Public Stats As Statistcs
Public UI As IRibbonUI

Sub init(Optional k, Optional mail_type)
    'If all_initialized Then Exit Sub
    'progress.start ("Initializing")
    If IsMissing(mail_type) Then mail_type = "REN"
    If Not IsMissing(k) Then Call define_test_case(k, mail_type)
    define_colors
    define_sheet_names
    
    define_macro_settings
    
    If T.name <> "" Then
        define_mail_type (T.mail_type)
        define_EDC (T.name)
    End If
    
    define_statuses
    define_active_cols
    
    define_filter_tab
    
    define_mismatch_cols
    'define_stats
    define_checklists
    'define_log
    
    If Not utility_tab() Is Nothing Then imported_gagg = True
    If Not active_tab() Is Nothing Then imported_active = True
    'If Not supplier_tab() Is Nothing Then imported_supplier = True
    
    refresh_ribbon
    
    all_initialized = True
    progress.complete
End Sub
