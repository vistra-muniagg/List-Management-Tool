Sub export_sheet(export_sheet, file_name)
    If export_sheet Is Nothing Then Exit Sub
    Dim wb As Workbook
    If export_sheet Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    export_sheet.Copy
    ActiveSheet.SaveAs fileName:=ThisWorkbook.path & "/" & file_name & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close False
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
End Sub

Sub export_files()
    progress.start "Exporting Files"
    x1 = get_contract_id()
    x2 = get_community_name()
    x = x1 & " - " & x2
    Call export_sheet(LP_tab(), x & MT.LP_file_suffix)
    Call export_sheet(mail_tab(), x & " Mail List")
    Call export_sheet(drop_tab(), x & " Drops")
    Call export_sheet(opt_in_tab(), x & " Opt-In Mail List")
    If EDC.ruleset_name = "DUKE" Then Call export_sheet(sibling_tab(), x & " DUKE Sibling Accounts")
    filter_tab().UsedRange.columns.Hidden = False
    progress.finish
    set_step (8)
    ThisWorkbook.Save
End Sub
