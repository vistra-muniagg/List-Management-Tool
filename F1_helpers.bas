Sub delete_sheet(sheet_name)
    On Error Resume Next
    Application.DisplayAlerts = 0
    Sheets(sheet_name).Delete
    Application.DisplayAlerts = 1
End Sub

Sub insert_column(sheet_name, target_col, col_name, header_color As CellColors)
    'target_col=0 adds to end
    'target_col=1 adds to beginning
    With Sheets(sheet_name)
        col_count = Application.CountA(Sheets(sheet_name).Rows(1))
        If target_col = 0 Then
            target_col = col_count + 1
        Else
            .columns(target_col).Insert Shift:=xlToRight
        End If
        .Cells(1, target_col) = col_name
        .Cells(1, target_col).Font.Bold = True
        If header_color.InteriorColor <> 0 Then
            .Cells(1, target_col).Interior.color = header_color.InteriorColor
            .Cells(1, target_col).Font.color = header_color.FontColor
        End If
    End With
End Sub

Sub add_autofilter(sheet_name)
    If Sheets(sheet_name).AutoFilterMode Then
        Sheets(sheet_name).AutoFilterMode = False
        Sheets(sheet_name).UsedRange.columns.AutoFilter
        Sheets(sheet_name).UsedRange.columns.AutoFit
    Else
        Sheets(sheet_name).UsedRange.columns.AutoFilter
        Sheets(sheet_name).UsedRange.columns.AutoFit
    End If
End Sub

Sub reapply_autofilter(sheet_num)
    With Sheets(sheet_num)
        If .AutoFilterMode Then .AutoFilterMode = False
        .columns.AutoFilter
        .Rows(1).Font.Bold = True
        .columns.AutoFit
    End With
End Sub

Sub trim_AEP_empty_end_cols(k)
    With Sheets(k)
        num_rows = Application.CountA(.columns(1))
        num_cols = Application.CountA(.Rows(1))
        extra_cols = .UsedRange.columns.count - num_cols
        While extra_cols > 0
            d = .UsedRange.columns(num_cols + extra_cols).value
            For j = 1 To num_rows
                If Application.Trim(d(j, 1)) <> "" Then
                    'detected shifted data
                    Exit Sub
                End If
            Next
            .columns(num_cols + extra_cols).Delete
            extra_cols = extra_cols - 1
        Wend
    End With
End Sub

Sub trim_headers(sheet_num)
    With Sheets(sheet_num)
        Set r = .UsedRange.Rows(1)
        header_row = r.value
        For j = 1 To UBound(header_row, 2)
            header = header_row(1, j)
            header = Application.Trim(header)
            header = Replace(header, vbTab, " ")
            header = Replace(header, vbLf, " ")
            header = Replace(header, vbCr, " ")
            header = Replace(header, vbCrLf, " ")
            header = Application.Trim(header)
            header_row(1, j) = header
        Next
        If EDC.ruleset_name = "AES" Then
            For j = 1 To UBound(header_row, 2)
                If header_row(1, j) = EDC.service_city Then
                    header_row(1, j) = EDC.mail_city
                    Exit For
                End If
            Next
            For j = 1 To UBound(header_row, 2)
                If header_row(1, j) = EDC.service_zip Then
                    header_row(1, j) = EDC.mail_zip
                    Exit For
                End If
            Next
        End If
        r.value = header_row
        r.WrapText = False
        reapply_autofilter (sheet_num)
    End With
End Sub

Sub move_column(source_sheet As Worksheet, source_col, target_sheet As Worksheet, target_col)
    If IsEmpty(source_col) Then Exit Sub
    d = source_sheet.UsedRange.columns(source_col).Value2
    num_rows = Application.CountA(source_sheet.UsedRange.columns(source_col))
    With target_sheet
        .columns(target_col).Insert Shift:=xlToRight
        If target_col = 1 Then .columns(1).NumberFormat = "@"
        .Range(.Cells(1, target_col), .Cells(num_rows, target_col)).Value2 = d
        If target_col = 1 Then .columns(1).NumberFormat = "General"
    End With
    If source_sheet.index = target_sheet.index Then
        source_sheet.columns(source_col + 1).Delete
    End If
End Sub

Sub move_accounts_to_front(sheet_num)
    With Sheets(sheet_num)
        account_col = find_column_header(EDC.account, sheet_num)
        If account_col = 1 Then Exit Sub
        Call move_column(Sheets(sheet_num), account_col, Sheets(sheet_num), 1)
        reapply_autofilter (sheet_num)
    End With
End Sub

Function format_accounts(source_array) As Variant
    num_rows = UBound(source_array)
    For j = 2 To num_rows
        account = CStr(source_array(j, 1))
        z = 0
        While z < EDC.zeros_to_add And Len(account) < EDC.account_number_length
            account = "0" & account
            z = z + 1
        Wend
        source_array(j, 1) = account
    Next
    format_accounts = source_array
End Function

Sub filter_tab_data_literal(filter_tab, source_col, target_col, gagg_data, num_rows)
    source_data = extract_array_col(source_col, gagg_data, True)
    filter_tab.Cells(1, target_col).Offset(1, 0).Resize(num_rows - 1).value = source_data
End Sub

Sub filter_tab_data_conditional(filter_tab, source_col, target_col, gagg_data, num_rows, condition)
    source_data = extract_array_col(source_col, gagg_data, True)
    Dim default_data() As Variant
    ReDim default_data(2 To num_rows, 1 To 1)
    For j = 2 To num_rows
        default_data(j, 1) = default_value
    Next
    filter_tab.Cells(2, target_col).Resize(num_rows - 1).value = default_data
    'unfinished
End Sub

Sub filter_tab_default_value(filter_tab, target_col, num_rows, default_value)
    Dim default_data() As Variant
    ReDim default_data(2 To num_rows, 1 To 1)
    For j = 2 To num_rows
        default_data(j, 1) = default_value
    Next
    filter_tab.Cells(2, target_col).Resize(num_rows - 1).value = default_data
    'unfinished?
End Sub

Function total_usage(target_row, gagg_data, usage_col, usage_multiple) As Variant
    Dim result_arr As Variant
    ReDim result_arr(0 To 2)
    If usage_multiple = 0 Then
        total_usage = gagg_data(target_row, usage_col)
        Exit Function
    End If
    usage = 0
    months = 0
    For j = 1 To 12
        monthly_usage = gagg_data(target_row, usage_col + usage_multiple * (j - 1))
        If Application.Trim(monthly_usage) = "" Then monthly_usage = 0
        If monthly_usage <> 0 Then months = months + 1
        usage = usage + monthly_usage
    Next
    
    usage = Round(usage, 1)
    
    If months <> 12 And months > 0 Then est_usage = Round((usage * 12 / months), 1)
    
    result_arr(0) = months
    result_arr(1) = usage
    result_arr(2) = Application.Max(est_usage, usage)
    
    total_usage = result_arr
    
End Function

Sub compare_replace_addresses(service_data, mail_data)
    'what does this need to do?
End Sub

Function calculate_usage(filter_tab, arr, target_col, source_col, num_rows)
    
    Dim partial_data As Variant
    Dim data_arr As Variant
    
    Set gagg = utility_tab()
    
    gagg_headers = flatten_array(gagg.UsedRange.Rows(1).value)
    
    c1 = F.columns.usage_months.index
    c2 = F.columns.actual_usage.index
    c3 = F.columns.estimated_usage.index
    
    first_usage_col = Application.Min(c1, c2, c3)
    
    usage_months_col = c1 - first_usage_col + 1
    actual_usage_col = c2 - first_usage_col + 1
    estimated_usage_col = c3 - first_usage_col + 1
    
    source_col_num = search_1d_array(source_col, gagg_headers)
    
    ReDim arr(2 To num_rows, 1 To 3)
    
    gagg_rows = Application.CountA(gagg.columns(1))
    gagg_cols = 12 * EDC.usage_multiple
    If EDC.usage_multiple = 0 Then gagg_cols = 1
    
    data_arr = gagg.Cells(1, source_col_num).Resize(gagg_rows, gagg_cols).value
    
    guess_usage = 0
    
    For j = 2 To num_rows
        
        calculated_usage = 0
        usage_months = 0
        estimated_usage = 0
        For m = 1 To gagg_cols Step EDC.usage_multiple
            usage = Val(data_arr(j, m))
            If EDC.ruleset_name = "AES" And S.Filter.AES_use_both_usage Then
                usage = usage + data_arr(j, m + 3)
            End If
            If usage <> 0 Then usage_months = usage_months + 1
            calculated_usage = calculated_usage + usage
            If EDC.usage_multiple = 0 Then Exit For
        Next
        
        If EDC.usage_multiple > 0 Then
            If usage_months <> 0 Then estimated_usage = calculated_usage * 12 / usage_months
            estimated_usage = Round(estimated_usage, 3)
        Else
            estimated_usage = calculated_usage
        End If
        
        If EDC.usage_multiple = 0 Then usage_months = "-"
        
        arr(j, 1) = usage_months
        arr(j, 2) = calculated_usage
        arr(j, 3) = estimated_usage
        
        progress.activity (j)
        
    Next
    
    calculate_usage = arr
    
End Function

Function AES_arrears(filter_tab, gagg_data, target_col, source_col, num_rows)

    Dim partial_data As Variant
    Dim data_arr As Variant
    
    gagg_headers = get_array_row(gagg_data, 1)
    
    source_col_num = search_1d_array(source_col, gagg_headers)
    
    ReDim arr(2 To num_rows, 1 To 1)
    
    For j = 2 To num_rows
        
        in_arrears = Evaluate(gagg_data(j, source_col_num) & EDC.arrears_yes)
        If in_arrears Then
            arr(j, 1) = "Y"
        Else
            arr(j, 1) = "N"
        End If
        
        progress.activity (j)
        
    Next
    
    AES_arrears = arr

End Function

Sub sort_sheet_col(target_sheet, target_col_num, Optional order = "A")
    If target_sheet Is Nothing Then Exit Sub
    If order = "A" Then
        sort_order = xlAscending
    Else
        sort_order = xlDescending
    End If
    target_sheet.UsedRange.columns.Sort key1:=target_sheet.UsedRange.columns(target_col_num), Order1:=sort_order, header:=xlYes
End Sub

Function cust_class(rate_code) As Variant
    rate_code = clean_rate_code(rate_code)
    k = search_1d_array(rate_code, EDC.res_codes)
    If k > -1 Then
        cust_class = F.columns.customer_class.possible_values(0)
    Else
        cust_class = F.columns.customer_class.possible_values(1)
    End If
End Function

Function clean_rate_code(rate_code)
    rate_code = Application.Trim(rate_code)
    If EDC.name = "COM" Then
        rate_code = Mid$(rate_code, 3, 3)
    End If
    clean_rate_code = rate_code
End Function

Function filter_data() As Variant
    filter_data = filter_tab().UsedRange.value
End Function

Function onedrive_parent_folder()

    On Error GoTo pathnotfound

    onedrive_parent_folder = ""
    
    env = Environ("USERPROFILE")
    
    b = S.OneDrive.parent_folder
    
    d = Dir(env & b, vbDirectory)
    
    If d <> "" Then
        onedrive_parent_folder = env & b
    Else
        GoTo pathnotfound
    End If
    
    Exit Function
    
pathnotfound:
    
End Function

Function onedrive_mailings_folder()

    If Application.UserName = "Scally, Meghan" Then
        onedrive_mailings_folder = "C:\Users\NGQP\OneDrive - Vistra Corp\MUNI AGG\(1) Operations\(8) Mailings"
        Exit Function
    End If
    
    If Application.UserName = "Baglia, Chelsea" Then
        onedrive_mailings_folder = "C:\Users\400026\OneDrive - Vistra Corp\(1) Operations\(8) Mailings"
        Exit Function
    End If
    
    If Application.UserName = "Crewson, Kevin" Then
        onedrive_mailings_folder = "C:\Users\48859\OneDrive - Vistra Corp\Shared Documents - Muni-Agg\(1) Operations\(8) Mailings"
        Exit Function
    End If
    
    On Error GoTo pathnotfound

    onedrive_mailings_folder = ""
    
    od = onedrive_parent_folder()
    
    If od = "" Then GoTo pathnotfound
    
    d = Dir(od & S.OneDrive.mailings_folder, vbDirectory)
    
    If d <> "" Then
        onedrive_mailings_folder = od & "\" & d
    Else
        GoTo pathnotfound
    End If
    
    Exit Function
    
pathnotfound:
    
End Function

Function onedrive_dna_folder()

    If Application.UserName = "Scally, Meghan" Then
        onedrive_dna_folder = "C:\Users\NGQP\OneDrive - Vistra Corp\MUNI AGG\(1) Operations\(6) List Management\(4) PUCO Do Not Aggregate (DNA) List"
        Exit Function
    End If

    On Error GoTo pathnotfound

    onedrive_dna_folder = ""
    
    od = onedrive_parent_folder()
    
    If od = "" Then GoTo pathnotfound
    
    d = Dir(od & S.OneDrive.dna_folder, vbDirectory)
    
    If d <> "" Then
        onedrive_dna_folder = od & S.OneDrive.dna_folder
    Else
        GoTo pathnotfound
    End If
    
    Exit Function
    
pathnotfound:

    onedrive_dna_folder = onedrive_dna_folder_old()
    
End Function

Function onedrive_list_management_folder()

    If Application.UserName = "Scally, Meghan" Then
        onedrive_list_management_folder = "C:\Users\NGQP\OneDrive - Vistra Corp\MUNI AGG\(1) Operations\(6) List Management"
        Exit Function
    End If
    
    If Application.UserName = "Baglia, Chelsea" Then
        onedrive_list_management_folder = "C:\Users\400026\OneDrive - Vistra Corp\(1) Operations\(6) List Management"
        Exit Function
    End If
    
    If Application.UserName = "Crewson, Kevin" Then
        onedrive_list_management_folder = "C:\Users\48859\OneDrive - Vistra Corp\Shared Documents - Muni-Agg\(1) Operations\(6) List Management"
        Exit Function
    End If

    On Error GoTo pathnotfound
    
    onedrive_list_management_folder = ""
    
    od = onedrive_parent_folder()
    
    If od = "" Then GoTo pathnotfound
    
    d = Dir(od & S.OneDrive.list_management_folder, vbDirectory)
    
    If d <> "" Then
        onedrive_list_management_folder = od & "\" & d
    Else
        GoTo pathnotfound
    End If
    
    Exit Function
    
pathnotfound:
    
End Function

Function onedrive_mapping_db_folder()

    If Application.UserName = "Scally, Meghan" Then
        onedrive_mapping_db_folder = "C:\Users\NGQP\OneDrive - Vistra Corp\MUNI AGG\(1) Operations\(6) List Management\(7) Mapping Database"
        Exit Function
    End If

    On Error GoTo pathnotfound

    onedrive_mapping_db_folder = ""
    
    od = onedrive_list_management_folder()
    
    If od = "" Then GoTo pathnotfound
    
    d = Dir(od & S.OneDrive.mapping_db_folder, vbDirectory)
    
    If d <> "" Then
        onedrive_mapping_db_folder = od & S.OneDrive.mapping_db_folder & "\"
    Else
        GoTo pathnotfound
    End If
    
    Exit Function
    
pathnotfound:
    
End Function

Function onedrive_migration_folder()

    On Error GoTo pathnotfound

    onedrive_migration_folder = ""
    
    od = onedrive_parent_folder()
    
    If od = "" Then GoTo pathnotfound
    
    d = Dir(od & S.OneDrive.migration_folder, vbDirectory)
    
    If d <> "" Then
        onedrive_migration_folder = od & S.OneDrive.migration_folder & "\"
    Else
        GoTo pathnotfound
    End If
    
    Exit Function
    
pathnotfound:
    
End Function

Function onedrive_documentation_folder()

    On Error GoTo pathnotfound

    onedrive_documentation_folder = ""
    
    od = onedrive_parent_folder()
    
    If od = "" Then GoTo pathnotfound
    
    d = Dir(od & S.OneDrive.documentation_folder, vbDirectory)
    
    If d <> "" Then
        onedrive_documentation_folder = od & S.OneDrive.documentation_folder & "\"
    Else
        GoTo pathnotfound
    End If
    
    Exit Function
    
pathnotfound:
    
End Function

Sub array_insert_col(ByRef arr_1d As Variant, index, value As String)
    n = UBound(arr_1d)
    ReDim Preserve arr_1d(0 To n + 1)
    For k = n + 1 To index Step -1
        arr_1d(k) = arr_1d(k - 1)
    Next
    arr_1d(index) = value
End Sub

Function is_1d_arr(arr)
    Err.Number = 0
    is_1d_arr = False
    On Error Resume Next
    L2 = LBound(arr, 2)
    If Err.Number <> 0 Then is_1d_arr = True
End Function

Function is_2d_arr(arr)
    Err.Number = 0
    is_2d_arr = True
    On Error Resume Next
    L2 = LBound(arr, 2)
    If Err.Number <> 0 Then is_2d_arr = False
End Function

Function filter_col(col As ColumnHeader)
    filter_col = filter_tab().UsedRange.columns(col.index).value
End Function

Function filter_row(row)
    filter_col = filter_tab().UsedRange.Rows(row).value
End Function

Function active_column(col As ActiveColumnHeader)
    active_column = active_tab().UsedRange.columns(col.index).value
End Function

Function clean_file_name(file_name)
    clean_file_name = Mid(file_name, InStrRev(file_name, "\") + 1)
End Function

Function searchable_address(str)
    searchable_address = str Like "[1-9]*"
End Function

Sub save_waterfall()
    ChDir onedrive_mailings_folder()
    file_name = Application.GetSaveAsFilename(InitialFileName:=CurDir, filefilter:="Waterfall Files (*.xlsm), *.xlsm", title:="Save Waterfall In Community Folder")
    If Not VarType(file_name) = 11 Then
        ThisWorkbook.SaveAs fileName:=file_name, FileFormat:=xlOpenXMLWorkbookMacroEnabled, addtomru:=False
    End If
End Sub

Function find_folder(start_path, target)
    
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    Set folder = fso.GetFolder(start_path)
    
    On Error Resume Next
    
    For Each subfolder In folder.SubFolders
        'If subfolder < target Then
        dp subfolder.name
        If UCase(subfolder.name) Like UCase(target) Then
            find_folder = subfolder.path
            Exit Function
        Else
            find_folder = find_folder(subfolder.path, target)
            If find_folder <> "" Then Exit Function
        End If
    Next
End Function

Sub test_ff()
    dp find_folder("C:\Users\400050\OneDrive - Vistra Corp", "(6) List Management")
End Sub

Function get_community_name()
    get_community_name = home_tab().Range(S.HOME.community_name_location)
End Function

Function get_contract_id()
    get_contract_id = home_tab().Range(S.HOME.contract_location)
End Function

Function get_oo_date()
    get_oo_date = home_tab().Range(S.HOME.oo_date_location)
End Function

Function parent_folder(str)
    If str = "" Then
        parent_folder = ""
        Exit Function
    End If
    x1 = InStrRev(str, "\")
    str2 = Left$(str, x1 - 1)
    x2 = InStrRev(str2, "\")
    parent_folder = Left$(str2, x2)
End Function

Sub add_user_name()
    x = name_reverse(Application.UserName)
    k = Split(x, " ")
    If k(0) Like "I*" Then k(0) = "Iris"
    x = home_tab().Range(S.HOME.user_location)
    If x <> k(0) And x = "" Then home_tab().Range(S.HOME.user_location) = k(0)
End Sub

Function expanded_oo_date(str)
    expanded_oo_date = format(str, "MMMM D, YYYY")
End Function

Function format_phone(str)
    format_phone = str
End Function

Sub make_dir(path)
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
End Sub

