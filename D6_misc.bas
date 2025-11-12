Sub misc_filter()
    duke_sibling_accounts
    'premise_mismatch_accounts
    filter_tab().columns.AutoFit
End Sub

Sub duke_sibling_accounts()
    If EDC.display_name <> "DUKE" Then
        Exit Sub
    Else
        find_DUKE_sibling_accounts
    End If
End Sub

Sub find_DUKE_sibling_accounts()
    
    progress.start "Fixing DUKE Sibling Accounts"
    
    Dim arr As Variant
    
    parent_account = ""
    parent_account_len = 12
    subset_size = 4
    
    Set ff = filter_tab()
    
    n = ff.Range("A1").End(xlDown).row
    
    Set data_range = ff.Range(ff.Cells(1, 1), ff.Cells(n, subset_size))
    
    arr = data_range.value
    
    ReDim Preserve arr(1 To n, 1 To subset_size + 1)
    
    arr(1, 5) = "Parent Account"
    
    For i = 2 To n
        parent_account = Left(arr(i, 1), parent_account_len)
        arr(i, subset_size + 1) = parent_account
    Next
    
    For i = 2 To n
        initial_parent_account = arr(i, 5)
        i = fix_DUKE_sibling_accounts(arr, i, n, initial_parent_account)
        progress.activity (i)
    Next
    
    arr = trim_arr_end_column(arr)
    
    Set r = ff.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2))
    
    r.value = arr
    
    track_sibling_accounts (parent_account_len)
    
    progress.finish
    
    ThisWorkbook.RefreshAll
    
End Sub

Function fix_DUKE_sibling_accounts(ByRef arr As Variant, index, n, parent_account)

    ineligible_sibling = False
    
restart_loop:

    j = index

    Do While j <= n
        If arr(j, 5) <> parent_account Then
            j = j + 1
            Exit Do
        End If
        If Not ineligible_sibling And arr(j, 3) = "N" Then
            ineligible_sibling = True
            j = index
            GoTo restart_loop
        End If
        If ineligible_sibling And arr(j, 3) = "Y" Then
            If arr(j, 4) = F.columns.mail_category.possible_values(1) Then
                arr(j, 2) = FS.duke_sibling_account.ineligible_ren_status
                arr(j, 3) = "N"
            Else
                arr(j, 2) = FS.duke_sibling_account.ineligible_new_status
                arr(j, 3) = "N"
            End If
        End If
        
        j = j + 1
        
    Loop
    
    fix_DUKE_sibling_accounts = j - 1

End Function

Sub track_sibling_accounts(parent_account_len)

    parent_account = ""
    
    Set ff = filter_tab()
    
    With F.columns
        account_arr = filter_col(.account_number)
        status_arr = filter_col(.status)
    End With
    
    data_arr = ff.UsedRange.value
    
    n1 = UBound(account_arr, 1)
    n2 = UBound(data_arr, 2)
    
    ren_sibling_label = FS.duke_sibling_account.ineligible_ren_status
    new_sibling_label = FS.duke_sibling_account.ineligible_new_status
    
    Dim arr As Variant
    ReDim arr(1 To n1, 1 To n2)
    
    data_row = 1
    
    For j = 1 To n2
        arr(1, j) = data_arr(1, j)
    Next
    
    For i = 2 To n1
        parent_account = ""
        If status_arr(i, 1) = ren_sibling_label Or status_arr(i, 1) = new_sibling_label Then
            parent_account = Left$(account_arr(i, 1), parent_account_len)
            For j = i - 1 To 2 Step -1
                If Left$(account_arr(j, 1), parent_account_len) <> parent_account Then Exit For
                data_row = data_row + 1
                For k = 1 To n2
                    arr(data_row, k) = data_arr(j, k)
                Next
            Next
            For j = i To n1
                If Left$(account_arr(j, 1), parent_account_len) <> parent_account Then Exit For
                data_row = data_row + 1
                For k = 1 To n2
                    arr(data_row, k) = data_arr(j, k)
                Next
            Next
        End If
        progress.activity (i)
    Next
    
    If data_row = 1 Then Exit Sub
    
    Set duke_sibling_tab = Sheets.Add(after:=dna_tab())
    
    With duke_sibling_tab
        delete_sheet (SN.duke_siblings)
        .name = SN.duke_siblings
        .columns(1).NumberFormat = "@"
        .Range("A1").Resize(k, UBound(F.order_array)).value = arr
        .Rows(1).Font.Bold = True
        .columns.AutoFilter
        .columns.AutoFit
        For Each cell In .UsedRange.columns(1).Cells
            If cell.row > 1 Then
                cell.Characters(start:=1, Length:=12).Font.color = C.RED.InteriorColor
                cell.Characters(start:=1, Length:=12).Font.Bold = True
            End If
        Next
    End With
    
End Sub

Function premise_mismatch_accounts()
    Dim d As Object
    Dim results As Object
    Set aa = active_tab()
    arr = aa.UsedRange.value
    active_headers = get_array_row(arr, 1)
    For j = 1 To UBound(active_headers)
        If active_headers(j) = A.columns.customer_class.header Then
            Exit For
        End If
    Next
    arr1 = get_array_col(arr, 1)
    arr2 = get_array_col(arr, j)
    arr3 = filter_col(F.columns.account_number)
    arr4 = filter_col(F.columns.customer_class)
    Set d = CreateObject("scripting.dictionary")
    Set results = CreateObject("scripting.dictionary")
    For k = 2 To UBound(arr1, 1)
        d.Add arr1(k), arr2(k)
    Next
    For k = 2 To UBound(arr3, 1)
        If d.exists(arr3(k, 1)) Then
            gagg_premise = arr4(k, 1)
            x = d(arr3(k, 1))
            If gagg_premise <> x Then
                results.Add k, x
            End If
        End If
    Next
    Set premise_mismatch_accounts = results
End Function

Sub track_premise_mismatches()

    If Not MT.has_renewal Then Exit Sub
    
    Set aa = premise_mismatch_accounts()
    
    If aa.count = 0 Then Exit Sub
    
    k = aa.count
    
    Set premise_tab = Sheets.Add(after:=dna_tab())
    
    arr = filter_data()
    
    For j = 2 To UBound(arr, 1)
        If aa.exists(arr(j, 1)) Then
            
        End If
    Next
    
    With premise_tab
        delete_sheet (SN.premise_mismatch)
        .name = SN.premise_mismatch
        .columns(1).NumberFormat = "@"
        .Range("A1").Resize(aa.count + 1, UBound(F.order_array)).value = arr
        .Rows(1).Font.Bold = True
        .columns.AutoFilter
        .columns.AutoFit
    End With
    
End Sub

