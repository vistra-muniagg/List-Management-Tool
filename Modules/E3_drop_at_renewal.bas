Sub ren_drops()
    
    If Not MT.has_renewal Then Exit Sub
    Set ff = filter_tab()
    
    With F.columns
        category_arr = filter_col(.mail_category)
        eligible_arr = filter_col(.eligible_opt_out)
    End With
    
    n = UBound(category_arr, 1)
    
    data_arr = ff.UsedRange.value
    
    Dim arr As Variant
    ReDim arr(1 To n, 1 To UBound(F.order_array))
    
    ren_label = F.columns.mail_category.possible_values(1)
    
    k = 1
    
    For j = 1 To UBound(F.order_array)
        arr(k, j) = data_arr(1, j)
    Next
    
    For i = 2 To n
        If category_arr(i, 1) <> ren_label Then GoTo skip_row
        If eligible_arr(i, 1) <> "N" Then GoTo skip_row
        k = k + 1
        For j = 1 To UBound(F.order_array)
            arr(k, j) = data_arr(i, j)
        Next
skip_row:
    Next
    
    If k = 1 Then Exit Sub
    
    Set drops = Sheets.Add(after:=filter_tab())
    
    With drops
        .name = SN.ren_drops
        .columns(1).NumberFormat = "@"
        .Range("A1").Resize(k, UBound(F.order_array)).value = arr
        .Rows(1).Font.Bold = True
        .columns.AutoFilter
        .columns.AutoFit
    End With
    
    home_tab().Range(S.HOME.renewal_drop_count_location).Offset(0, -1) = "Renewal Drop Count"
    home_tab().Range(S.HOME.renewal_drop_count_location) = k - 1
    
End Sub
