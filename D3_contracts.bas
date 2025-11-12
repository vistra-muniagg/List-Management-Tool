Sub test_contracts()
    'init
    S.contracts.hide_snowflake_query = True
    get_contracts_file
    dedupe_contracts
    process_contracts
End Sub

Sub process_contracts()
    'get_contracts_file
    process_contracts_results
End Sub

Sub create_contracts_sql()

    If Not MT.needs_gagg_list Then Exit Sub
    
    delete_sheet SN.Snowflake
    
    Set query_tab = Sheets.Add(before:=home_tab())
    query_tab.name = SN.Snowflake
    
    Set ff = filter_tab()
    
    eligible_col = F.columns.eligible_opt_out.index
    active_col = F.columns.active_in_LP.index
    
    num_rows = Application.CountA(filter_tab().columns(1))
    num_cols = Application.Max(eligible_col, active_col)
    data_arr = ff.Range("A1").Resize(num_rows, num_cols).value
    
    Dim arr
    ReDim arr(0 To num_rows)
    
    arr(0) = "IN"
    
    k = 1
    
    For i = 2 To num_rows
        If data_arr(i, active_col) = "N" Then
            If data_arr(i, eligible_col) = "Y" Then
                arr(k) = ",'" & data_arr(i, 1) & "'"
                k = k + 1
            End If
        End If
        progress.activity (i)
    Next
    
    arr(1) = Replace(arr(1), ",", "(")
    arr(k) = ")"
    
    ReDim Preserve arr(0 To k + 1)
    
    query_tab.Rows(1).WrapText = False
    query_tab.columns(1).AutoFit
    
    query_tab.Range("A1").Resize(k + 1, 1).value = Application.Transpose(arr)
    
    'add button to process query
    'Set r = query_tab.Range("C1:D3")
    'Set btn = ActiveSheet.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
    'With btn
    '    .OnAction = "gcc"
    '    .Caption = "Process Contracts Query"
    'End With
    
    'If S.contracts.hide_snowflake_query Then query_tab.visible = False
    
End Sub

Sub get_contracts_file()
    
    If Not MT.needs_gagg_list Then Exit Sub
    
    delete_sheet SN.contracts
    
    If T.contracts_file = "" Then
        file_name = Application.GetOpenFilename("Snowflake Files (*.csv), *.csv", , "Select Contracts Query Results")
    Else
        contracts_folder = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(2) Macro Testing\(1) Testing Files\Test Contracts Queries\"
        file_name = contracts_folder & T.contracts_file
    End If
    
    If VarType(file_name) = 11 Then
        msg = MsgBox("Was the contracts query empty?", vbQuestion + vbYesNo)
        If msg = vbNo Then
            Exit Sub
        Else
            set_step (5)
            If Not UI Is Nothing Then UI.Invalidate
            Exit Sub
        End If
    End If
    
    before_sheet = readme_tab().index
    
    Call import_csv_file(file_name, "Snowflake", before_sheet)
    
    reapply_autofilter (before_sheet)
    
    Sheets(before_sheet).name = SN.contracts
    
    dedupe_contracts
    
    If S.contracts.hide_snowflake_query Then Sheets(SN.Snowflake).visible = False
    
    set_step (5)
    
End Sub

Sub dedupe_contracts()

    If Not MT.needs_gagg_list Then Exit Sub
    
    Set query_tab = contracts_tab()
    
    If query_tab Is Nothing Then Exit Sub
    
    sas_id_row = Application.match(S.contracts.sas_id_row_header, query_tab.Rows(1), 0)
    
    Call sort_sheet_col(query_tab, sas_id_row, "D")
    
    query_tab.columns(sas_id_row).CurrentRegion.RemoveDuplicates columns:=1, header:=xlYes
    
End Sub

Sub process_XDUPX()

    If Not MT.needs_gagg_list Then Exit Sub

    If EDC.ruleset_name <> "AEP" Then Exit Sub
    
    Set ff = filter_tab()
    Set query_tab = contracts_tab()
    
    If query_tab Is Nothing Then Exit Sub
    
    xdupx_col = Application.match(S.contracts.xdupx_header, query_tab.Rows(1), 0)
    LP_name_col = Application.match(S.contracts.LP_cust_name, query_tab.Rows(1), 0)
    filter_name_col = F.columns.customer_name.index
    contracts_rows = Application.CountA(query_tab.columns(1))
    num_rows = Application.CountA(ff.columns(1))
    
    Call sort_sheet_col(ff, 1, "A")
    Call sort_sheet_col(query_tab, 1, "A")
    
    a1 = flatten_array(query_tab.UsedRange.columns(1).value)
    a2 = flatten_array(ff.UsedRange.columns(1).value)
    
    n1_arr = query_tab.UsedRange.columns(LP_name_col).value
    n2_arr = ff.UsedRange.columns(filter_name_col).value
    
    search_start = 2
    
    xdupx_arr = query_tab.UsedRange.columns(xdupx_col).value
    
    For j = 2 To contracts_rows
        n1 = n1_arr(j, 1)
        k = array_binary_search(a1(j), a2, search_start, num_rows)
        If k <= 0 Then
            'not on filter tab
            Exit Sub
        End If
        search_start = Application.Max(2, k)
        n2 = n2_arr(k, 1)
        
        xdupx_arr(j, 1) = Not same_name(n1, n2)
        
        progress.activity (j)
        
    Next
    
    query_tab.Cells(1, xdupx_col).Resize(contracts_rows).value = xdupx_arr
    
End Sub

Sub process_contracts_results()

    If contracts_tab() Is Nothing Then Exit Sub
    
    process_XDUPX
    
    Dim status_arr As Variant
    Dim active_arr As Variant
    Dim eligible_arr As Variant
    Dim LP_status_arr As Variant
    Dim status_reason_arr As Variant
    Dim contract_arr As Variant
    Dim xdupx_arr As Variant
    Dim account_arr As Variant
    Dim query_account_arr As Variant
    
    Set ff = filter_tab()
    Set query_tab = contracts_tab()
    
    If query_tab Is Nothing Then Exit Sub
    
    query_tab.columns(1).NumberFormat = "@"
    
    contracts_rows = Application.CountA(query_tab.columns(1))
    
    num_rows = Application.CountA(ff.columns(1))
    
    Call sort_sheet_col(ff, 1, "A")
    Call sort_sheet_col(query_tab, 1, "A")
    
    account_arr = flatten_array(ff.UsedRange.columns(1).value)
    query_account_arr = query_tab.UsedRange.columns(1).value
    
    LP_status_col = Application.match(S.contracts.status_header, query_tab.Rows(1), 0)
    status_reason_col = Application.match(S.contracts.status_reason_header, query_tab.Rows(1), 0)
    contract_col = Application.match(S.contracts.contract_header, query_tab.Rows(1), 0)
    intent_col = Application.match(S.contracts.intent_contract_header, query_tab.Rows(1), 0)
    xdupx_col = Application.match(S.contracts.xdupx_header, query_tab.Rows(1), 0)
    muniagg_status_col = Application.match(S.contracts.muniagg_status_header, query_tab.Rows(1), 0)
    
    status_arr = ff.UsedRange.columns(F.columns.status.index).value
    active_arr = ff.UsedRange.columns(F.columns.active_in_LP.index).value
    eligible_arr = ff.UsedRange.columns(F.columns.eligible_opt_out.index).value
    LP_status_arr = query_tab.UsedRange.columns(LP_status_col).value
    status_reason_arr = query_tab.UsedRange.columns(status_reason_col).value
    muniagg_status_arr = query_tab.UsedRange.columns(status_reason_col).value
    contract_arr = query_tab.UsedRange.columns(muniagg_status_col).value
    If Not IsError(xdupx_col) Then xdupx_arr = query_tab.UsedRange.columns(xdupx_col).value
    
    'ineligible if active on anything
        'active
        'inactive processing
        'inactive pending activation
    'if xdupx and not active then eligible - xdupx
    'ineligible if opt out on current contract
        'status doesnt matter?
        'if externalcontractid=current then ineligible
    'else eligible
    
    search_start = 2
    
    For j = 2 To contracts_rows
        If Len(query_account_arr(j, 1)) < EDC.account_number_length Then
            query_account_arr(j, 1) = format(query_account_arr(j, 1), String(EDC.account_number_length, "0"))
        End If
        If LP_status_arr(j, 1) = "ACTIVE" Then
            'active
            status = FS.contracts.ineligible_active
            eligible = "N"
        ElseIf status_reason_arr(j, 1) = "DROP_PENDING" Then
            'inactive processing
            status = FS.contracts.ineligible_active
            eligible = "N"
        ElseIf status_reason_arr(j, 1) = "PROCESSING" Then
            'inactive processing
            status = FS.contracts.ineligible_active
            eligible = "N"
        ElseIf status_reason_arr(j, 1) = "PENDING_ACTIVATION" Then
            'inactive pending activation
            status = FS.contracts.ineligible_active
            eligible = "N"
        ElseIf EDC.ruleset_name = "AEP" And xdupx_arr(j, 1) = True Then
            'eligible xdupx
            status = FS.contracts.eligible_xdupx
            eligible = "Y"
        ElseIf MT.check_migration_data And contract_arr(j, 1) = current_contract Then
            'ineligible previously on contract
            status = FS.contracts.ineligible_previous_mail
            eligible = "N"
        Else
            'eligible inactive
            status = FS.contracts.eligible_inactive
            eligible = "Y"
        End If
        k = array_binary_search(query_account_arr(j, 1), account_arr, search_start, num_rows)
        If k > 0 Then
            If eligible_arr(k, 1) = "Y" Then
                search_start = k
                'apply statuses
                status_arr(k, 1) = status
                eligible_arr(k, 1) = eligible
            End If
        End If
        progress.activity (j)
    Next
    
    ff.Cells(1, F.columns.status.index).Resize(num_rows).value = status_arr
    ff.Cells(1, F.columns.eligible_opt_out.index).Resize(num_rows).value = eligible_arr
    query_tab.Cells(1, 1).Resize(contracts_rows).value = query_account_arr
    
    ThisWorkbook.RefreshAll
    
End Sub

Function same_name(n1, n2) As Boolean

    If n1 = "" Or n2 = "" Then
        same_name = False
        Exit Function
    End If
    
    Dim name_arr() As String
    
    name_arr1 = name_parts(n1)
    name_arr2 = name_parts(n2)
    
    first1 = name_arr1(0)
    first2 = name_arr2(0)
    
    middle1 = name_arr1(1)
    middle2 = name_arr2(1)
    
    last1 = name_arr1(2)
    last2 = name_arr2(2)
    
    If S.contracts.remove_suffix_for_xdupx Then
        last1 = no_suffix(last1)
        last2 = no_suffix(last2)
    End If
    
    same_name = False
    
    firstlast1 = first1 & last1
    firstlast2 = first2 & last2
    
    k = S.contracts.xdupx_guess_wildcard_length
    
    If firstlast1 = firstlast2 Then
        same_name = True
        Exit Function
    ElseIf last1 = last2 Or Left$(n1, k) = Left$(n2, k) Then
        same_name = True
        Exit Function
    ElseIf Levenshtein(firstlast1, firstlast2) <= S.contracts.Levenshtein_match_len Then
        same_name = True
    End If
    
End Function

Function Levenshtein(s1, s2) As Integer
    
    Dim matrix() As Integer
    
    ' Get string lengths
    len1 = Len(s1)
    len2 = Len(s2)
    
    ' Edge case: if one of the strings is empty, return the length of the other
    If len1 = 0 Then
        Levenshtein = len2
        Exit Function
    ElseIf len2 = 0 Then
        Levenshtein = len1
        Exit Function
    End If
    
    ' Resize matrix
    ReDim matrix(0 To len1, 0 To len2)
    
    ' Initialize first row and column
    For i = 0 To len1
        matrix(i, 0) = i
    Next i
    For j = 0 To len2
        matrix(0, j) = j
    Next
    
    ' Fill the matrix
    For i = 1 To len1
        For j = 1 To len2
            ' Determine cost (0 if characters are the same, 1 if different)
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            ' Take the minimum of: deletion, insertion, or substitution
            matrix(i, j) = Application.Min(matrix(i - 1, j) + 1, _
                                           matrix(i, j - 1) + 1, _
                                           matrix(i - 1, j - 1) + cost)
        Next
    Next
    
    ' Return the Levenshtein Distance from bottom-right of matrix
    Levenshtein = matrix(len1, len2)
End Function
