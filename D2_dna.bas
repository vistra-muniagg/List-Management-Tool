Sub dna_test()
    For k = 10 To 21
        test_dna (k)
    Next
End Sub

Sub test_dna(k)

    If EDC.state <> "OH" Then Exit Sub
    
    progress.start "Searching DNA List"
    
    S.DNA.wildcard_length = k
    
    add_dna_comparison_sheet
    
    Call sort_sheet_col(filter_tab(), 1, "A")
    
    data_arr = filter_data()
    
    dna_file = find_dna_file
    
    If dna_file = "" Then dna_file = find_dna_list
    
    Set conn = ADO_connection_excel(dna_file)
    dna_rows = ADO_row_count(conn, S.DNA.sheet_name)
    dna_data = ADO_data(conn, S.DNA.sheet_name, "A1:J" & dna_rows, 1)
    conn.Close
    
    dna_cols = UBound(dna_data, 2)
    
    account_match = dna_account_match(data_arr, dna_data)
    add_dna_results (account_match)
    
    dna_data = sort_2d_arr(dna_data, 5)
    
    address_match = dna_address_match(dna_data, account_match)
    add_dna_results (address_match)
    
    reapply_autofilter (dna_tab().index)
    
    dedupe_dna
    
    add_dna_formatting
    
    dna_tab().Activate
    
    set_step (4)
    
    progress.finish
    
End Sub

Function find_dna_file()

    find_dna_file = ""
    
    env = Environ("USERPROFILE")
    
    dna_folder = onedrive_dna_folder()
    
    'dna_folder = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(4) PUCO Do Not Aggregate (DNA) List"
    
    start_date = format(Date, "M-D-YY")
    
    search_date = start_date
    
    dna_file_name = S.DNA.file_name
    
    search_file_name = Replace(dna_file_name, "MM-DD-YY", search_date)
    
    found_dna_list = False
    
    days_checked = 0
    
    Do While Not found_dna_list And days_checked <= S.DNA.max_file_age
        find_dna_file = Dir(dna_folder & "\" & search_file_name)
        If find_dna_file <> "" Then
            find_dna_file = dna_folder & "\" & find_dna_file
            Exit Do
        Else
            search_date = previous_day(search_date)
            search_file_name = Replace(dna_file_name, "MM-DD-YY", search_date)
            days_checked = days_checked + 1
        End If
    Loop
    
End Function

Function get_dna_data(dna_file) As Variant

    On Error GoTo no_file
    
    Set open_dna_file = Nothing
    
    Set dna_wb = Workbooks.Open(dna_file, ReadOnly:=True, addtomru:=False)
    Set dna_list = dna_wb.Sheets(1)
    
    dna_rows = Application.CountA(dna_list.columns(1))
    dna_cols = Application.CountA(dna_list.Rows(1))
    
    Set sort_range = dna_list.Range("A1").Resize(dna_rows, dna_cols)
    
    Set sort_col = dna_list.columns(1)
    
    sort_range.Sort key1:=sort_col, Order1:=xlAscending, header:=xlYes
    
    'dna_data = dna_list.Range("A1").Offset(1, 0).Resize(dna_rows - 1, dna_cols).value
    dna_data = dna_list.Range("A1").Resize(dna_rows, dna_cols).value
    
    get_dna_data = dna_data
    
    dna_list.Parent.Close False
    
no_file:
    Exit Function
    
End Function

Sub add_dna_comparison_sheet(Optional k)
    delete_sheet (SN.DNA)
    Set dna_sheet = Sheets.Add(after:=filter_tab())
    dna_sheet.name = SN.DNA
    If Not IsMissing(k) Then dna_sheet.name = SN.DNA & "-" & k
    col_names = S.DNA.results_layout
    dna_sheet.Range("A1").Resize(1, UBound(col_names) + 1).value = col_names
    dna_sheet.Rows(1).Font.Bold = True
    dna_sheet.columns(1).NumberFormat = "@"
    reapply_autofilter (dna_tab().index)
End Sub

Sub add_dna_results(match_arr)
    
    n1 = UBound(match_arr, 1)
    n2 = UBound(match_arr, 2)
    
    If IsEmpty(match_arr(1, 1)) Then Exit Sub
    
    Set dna_sheet = dna_tab()
    
    k = UBound(S.DNA.results_layout)
    If S.DNA.include_wildcard_search Then k1 = 1
    
    With dna_sheet
        start_row = Application.CountA(dna_tab().columns(1)) + 1
        Set start_cell = .Cells(start_row, 1)
        start_cell.Resize(n1, k + k1).value = match_arr
    End With
    
End Sub

Function dna_account_match(data_arr, dna_list)
    
    Dim results_arr As Variant
    Dim temp_arr As Variant
    
    n = UBound(data_arr, 1)
    
    results_cols = UBound(S.DNA.results_layout) + 1
    
    ReDim results_arr(1 To n, 1 To results_cols)
    
    result_row = 1
    
    dna_rows = UBound(dna_list, 1)
    dna_cols = UBound(dna_list, 2)
    
    search_arr = get_array_col(dna_list, 1)
    
    'array_binary_search(target, search_array, search_start, search_end)
    
    'search_arr = transpose_ADO(search_arr)
    
    search_start = 1
    search_end = UBound(search_arr)
    
    name_col = F.columns.customer_name.index
    address_col = F.columns.service_address.index
    
    If S.DNA.include_wildcard_search Then k1 = 1
    
    If S.DNA.auto_populate_account_match And EDC.ruleset_name <> "AEP" Then
        assume_match = "Y"
        match_source = "Automatic"
    Else
        assume_match = ""
        match_source = "User"
    End If
    
    eligible_col = F.columns.eligible_opt_out.index
    
    For i = 2 To n
        If data_arr(i, eligible_col) = "Y" Then
            target = data_arr(i, 1)
            If EDC.ruleset_name = "DUKE" Then target = Left$(target, 12)
            match_row = array_binary_search(target, search_arr, search_start, search_end)
            If match_row > 0 Then
                search_start = match_row
                
                results_arr(result_row, 1) = data_arr(i, 1)
                results_arr(result_row, 2) = data_arr(i, name_col)
                results_arr(result_row, 3) = clean_name(UCase$(dna_list(match_row, 4)))
                results_arr(result_row, 4) = data_arr(i, address_col)
                results_arr(result_row, 5) = UCase$(dna_list(match_row, 5))
                If k1 > 0 Then results_arr(result_row, 5 + k1) = ""
                results_arr(result_row, 6 + k1) = UCase$(dna_list(match_row, 3))
                results_arr(result_row, 7 + k1) = dna_list(match_row, 2)
                results_arr(result_row, 8 + k1) = dna_list(match_row, 9)
                results_arr(result_row, 9 + k1) = "Account"
                results_arr(result_row, 10 + k1) = match_source
                results_arr(result_row, 11 + k1) = assume_match
                
                result_row = result_row + 1
                
            End If
        End If
        progress.activity (i)
    Next
    
    temp_arr = Application.Transpose(results_arr)
    
    ReDim Preserve temp_arr(1 To results_cols, 1 To result_row)
    
    results_arr = Application.Transpose(temp_arr)
    
    If result_row = 1 Then ReDim results_arr(1 To 1, 1 To results_cols)
    
    dna_account_match = results_arr
    
End Function

Function dna_address_match(dna_list, account_match_arr)
    
    Dim results_arr As Variant
    Dim temp_arr As Variant
    
    Call sort_sheet_col(filter_tab(), F.columns.service_address.index, "A")
    
    data_arr = filter_data()
    
    n = UBound(data_arr, 1)
    
    results_cols = UBound(S.DNA.results_layout) + 1
    
    ReDim results_arr(1 To n, 1 To results_cols)
    
    result_row = 1
    
    dna_rows = UBound(dna_list, 1)
    dna_cols = UBound(dna_list, 2)
    
    search_arr = get_array_col(dna_list, 12)
    
    'array_binary_search(target, search_array, search_start, search_end)
    
    'search_arr = transpose_ADO(search_arr)
    
    search_start = 1
    search_end = UBound(search_arr)
    
    name_col = F.columns.customer_name.index
    address_col = F.columns.service_address.index
    
    If S.DNA.include_wildcard_search Then k1 = 1
    
    account_match_len = UBound(account_match_arr, 1)
    
    elgible_col = F.columns.eligible_opt_out.index
    
    For i = 2 To n
        If data_arr(i, elgible_col) = "Y" Then
            target_account = data_arr(i, 1)
            target_str = data_arr(i, F.columns.service_address.index)
            wildcard_target = Left(target_str, S.DNA.wildcard_length)
            match_row = array_binary_search(wildcard_target, search_arr, search_start, search_end)
            If match_row > 0 Then
            
                For j = 2 To account_match_len
                    If account_match_arr(j, 1) = target_account Then GoTo next_row
                Next
            
                While search_arr(match_row) = wildcard_target
                    
                    search_start = match_row
                    
                    results_arr(result_row, 1) = data_arr(i, 1)
                    results_arr(result_row, 2) = data_arr(i, name_col)
                    results_arr(result_row, 3) = clean_name(UCase$(dna_list(match_row, 4)))
                    results_arr(result_row, 4) = data_arr(i, address_col)
                    results_arr(result_row, 5) = UCase$(dna_list(match_row, 5))
                    If k1 > 0 Then results_arr(result_row, 5 + k1) = wildcard_target
                    results_arr(result_row, 6 + k1) = UCase$(dna_list(match_row, 3))
                    results_arr(result_row, 7 + k1) = dna_list(match_row, 2)
                    results_arr(result_row, 8 + k1) = dna_list(match_row, 9)
                    results_arr(result_row, 9 + k1) = "Address"
                    results_arr(result_row, 10 + k1) = "User"
                    If T.name <> "" Then results_arr(result_row, 11 + k1) = dna_guess(results_arr(result_row, 2), results_arr(result_row, 3))
                    
                    match_row = match_row + 1
                    result_row = result_row + 1
                    
                Wend
            End If
        End If
        progress.activity (i)
next_row:
    Next
    
    temp_arr = Application.Transpose(results_arr)
    
    ReDim Preserve temp_arr(1 To results_cols, 1 To result_row)
    
    results_arr = Application.Transpose(temp_arr)
    
    If result_row = 1 Then ReDim results_arr(1 To 1, 1 To results_cols)
    
    dna_address_match = results_arr
    
    'dp "Address Match Length: " & S.DNA.wildcard_length & vbTab & "Address Matches: " & result_row - 1 & vbNewLine
    'readme_tab().Cells(S.DNA.wildcard_length + 14, 9) = result_row - 1
    
End Function

Sub dedupe_dna()
    Call sort_sheet_col(dna_tab(), 1, "A")
    With dna_tab()
        .columns(1).CurrentRegion.RemoveDuplicates columns:=1, header:=xlYes
        old_row_count = .UsedRange.Rows.count
        new_row_count = Application.CountA(.columns(1))
        Do While old_row_count > new_row_count
            .Rows(old_row_count).Delete
            old_row_count = old_row_count - 1
        Loop
    End With
End Sub

Sub remove_dna()

    Call update_checklist(S.QC.audit_checklist, "audit_dna", 0)
    
    If EDC.state <> "OH" Then Exit Sub
    If dna_tab() Is Nothing Then Exit Sub
    
    Set ff = filter_tab()
    
    Call sort_sheet_col(ff, 1, "A")
    Call sort_sheet_col(dna_tab(), 1, "A")
    With dna_tab()
        dna_table = .UsedRange.value
        n = UBound(dna_table, 1)
        For j = 2 To n
            If dna_table(j, S.DNA.result_col) = "" Then
                GoTo incomplete_dna
                Exit For
            End If
        Next
    End With
    
    arr = ff.UsedRange.columns(1).value
    search_arr = flatten_array(arr)
    
    search_start = 2
    
    For j = 2 To n
        If UCase$(dna_table(j, S.DNA.result_col)) = "Y" Then
            target = dna_table(j, 1)
            row_num = array_binary_search(target, search_arr, search_start, UBound(arr))
            search_start = row_num - 1
            If ff.Cells(row_num, F.columns.eligible_opt_out.index) = "N" Then
                GoTo next_row
            End If
            ff.Cells(row_num, F.columns.eligible_opt_out.index) = "N"
            If ff.Cells(row_num, F.columns.active_in_LP.index) = "Y" Then
                label = FS.dna_OH.ineligible_ren_status
            Else
                label = FS.dna_OH.ineligible_new_status
            End If
            
            ff.Cells(row_num, F.columns.do_not_agg.index) = "Y"
            ff.Cells(row_num, F.columns.status.index) = label
            
        End If
next_row:
    Next
    
    Call update_checklist(S.QC.audit_checklist, "audit_dna", 1)
    
    ThisWorkbook.RefreshAll
    
incomplete_dna:
    'do something for dna not completely filled out
    Exit Sub
End Sub

Function dna_guess(x1, x2)
    dna_guess = "N"
    x = same_name(x1, x2)
    If x Then dna_guess = "Y"
End Function

Function dna_double_check(dna_tab, dna_count) As Variant()
    probable_match_count = 0
    probable_miss_count = 0
    dna_guess_count = 0
    For j = 2 To dna_count + 1
        A = UCase(dna_tab.Cells(j, "L"))
        b = dna_tab.Cells(j, "M")
        If A = b Then dna_guess_count = dna_guess_count + 1
        If b = "Y" Then
            probable_match_count = probable_match_count + 1
            If A = "N" Then
                probable_miss_count = probable_miss_count + 1
            End If
        End If
    Next
    dna_guess_pct = dna_guess_count / dna_count
    dna_double_check = Array(dna_guess_pct, probable_miss_count)
End Function

Sub add_dna_formatting()
    dna_col_letter = "L"
    With dna_tab()
        .columns.AutoFit
        Set r = .UsedRange.columns(dna_col_letter)
        r.FormatConditions.Delete
        'Add first rule. red if Y
        r.FormatConditions.Add Type:=xlExpression, Formula1:="=UPPER(" & dna_col_letter & "1)=""Y"""
        r.FormatConditions(1).Font.ColorIndex = 9
        r.FormatConditions(1).Interior.ColorIndex = 3
        r.FormatConditions(1).Interior.TintAndShade = r.FormatConditions(1).Interior.TintAndShade + 0.75
        'Add second rule. green if N
        r.FormatConditions.Add Type:=xlExpression, Formula1:="=UPPER(" & dna_col_letter & "1)=""N"""
        r.FormatConditions(2).Font.ColorIndex = 51
        r.FormatConditions(2).Interior.ColorIndex = 35
        r.FormatConditions(2).Interior.TintAndShade = r.FormatConditions(2).Interior.TintAndShade
        'make all empty (not y/n) cells in column yellow
        r.Style = "Neutral"
        'format first row
        .Range("A1:L1").Font.Bold = True
        '.columns("E:G").Hidden = True
        reapply_autofilter (.index)
        .columns("F:K").Hidden = True
        .Activate
    End With
End Sub
