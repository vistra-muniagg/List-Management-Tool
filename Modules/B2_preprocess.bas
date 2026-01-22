Sub test_pre()
    preprocess
    If EDC.state <> "IL" Then hide_filter_group ("IL Filters")
    If EDC.state <> "OH" Then hide_filter_group ("OH Filters")
    'hide_filter_group ("Mapping Data")
    'hide_filter_group ("Mail Data")
End Sub

Sub preprocess()
    define_filter_tab
    define_filter_tab_columns
    create_filter_tab
    populate_filter_tab
    If EDC.state <> "IL" Then hide_filter_group ("IL Filters")
    If EDC.state <> "OH" Then hide_filter_group ("OH Filters")
End Sub

Sub create_filter_tab()

    progress.start ("Creating Filter Tab")

    Dim col As ColumnHeader
    
    delete_sheet (SN.Filter)
    Set s1 = Sheets.Add(before:=Sheets(SN.HOME))
    
    With s1
        .name = SN.Filter
        .Rows(1).Delete
        .Tab.color = C.GREEN_0.InteriorColor
        For j = 1 To UBound(F.order_array)
            col = F.order_array(j)
            .Cells(1, j) = col.header
            .Cells(1, j).Interior.color = col.cell_color.InteriorColor
            .Cells(1, j).Font.color = col.cell_color.FontColor
            .columns(j).NumberFormat = col.data_format
        Next
    End With
    
    reapply_autofilter (s1.index)
    
    progress.complete
    
End Sub

Sub populate_filter_tab()

    If Not MT.needs_gagg_list Then Exit Sub

    Set ff = filter_tab()

    On Error Resume Next

    Dim col As ColumnHeader
    
    Set gagg_list = Sheets(SN.Utility)
    
    gagg_data = deduped_data_arr(gagg_list)
    gagg_headers = get_array_row(gagg_data, 1)
    
    num_rows = UBound(gagg_data, 1)
    num_cols = UBound(gagg_data, 2)
    
    For k = 1 To UBound(F.order_array)
        col = F.order_array(k)
        'dp col.header
        progress.start ("Calculating " & col.header)
        Select Case col.data_type
            Case "Literal":
                Call populate_literal_filter_col(ff, col, k, gagg_headers, gagg_data, num_rows)
            Case "Generated":
                Call populate_generated_filter_col(ff, col, k, gagg_headers, gagg_data, num_rows)
            Case "Boolean":
                Call populate_bool_filter_col(ff, col, k, gagg_headers, gagg_data, num_rows)
            Case "Calculated":
                Call populate_calculated_filter_col(ff, col, k, gagg_headers, gagg_data, num_rows)
        End Select
        progress.complete
    Next
    
    reapply_autofilter (ff.index)
    
End Sub

Sub populate_literal_filter_col(filter_sheet, filter_col As ColumnHeader, target_col, gagg_headers, gagg_data, num_rows)

    source_col_num = search_1d_array(filter_col.source_col, gagg_headers)
    
    If source_col_num = -1 Then Exit Sub
    
    Dim arr() As Variant
    ReDim arr(2 To num_rows, 1 To 1)
    
    Set data_range = filter_sheet.Cells(2, target_col).Resize(num_rows - 1)
    
    If filter_col.index = F.columns.read_cycle.index Then
        fix_COM_cycles = True
    Else
        fix_COM_cycles = False
    End If
    
    For j = 2 To num_rows
        If source_col_num = 0 Then
            arr(j, 1) = filter_col.default_value
        Else
            arr(j, 1) = Application.Trim(UCase(gagg_data(j, source_col_num)))
            If arr(j, 1) = "" Then arr(j, 1) = filter_col.default_value
        End If
        If arr(j, 1) = "-" Then arr(j, 1) = ""
        If fix_COM_cycles Then
            If arr(j, 1) Like "CE#*" Then
                arr(j, 1) = Val(Mid$(arr(j, 1), 3))
            End If
        End If
        progress.activity (j)
    Next
    
    data_range.value = arr
        
End Sub

Sub populate_bool_filter_col(filter_tab, filter_col As ColumnHeader, target_col, gagg_headers, gagg_data, num_rows)
    source_col_num = search_1d_array(filter_col.source_col, gagg_headers)
    If source_col_num = -1 Then
        Call filter_tab_default_value(filter_tab, target_col, num_rows, filter_col.default_value)
        Exit Sub
    End If
    Dim col_data As Variant
    ReDim col_data(2 To num_rows, 1 To 1)
    With filter_tab
        Select Case filter_col.data_type
            Case "Boolean":
                condition = filter_col.condition_value
                For j = 2 To num_rows
                    value = gagg_data(j, source_col_num)
                    col_data(j, 1) = filter_col_bool(value, condition)
                Next
                .Cells(2, target_col).Resize(num_rows - 1).value = col_data
            Case "Literal":
                For j = 2 To num_rows
                    value = Application.Trim(gagg_data(j, source_col_num))
                    col_data(j, 1) = value
                Next
                .Cells(2, target_col).Resize(num_rows - 1).value = col_data
            Case Else:
                Exit Sub
        End Select
    End With
End Sub

Sub populate_generated_filter_col(filter_tab, filter_col As ColumnHeader, target_col, gagg_headers, gagg_data, num_rows)

    On Error Resume Next
    
    source_col_num = search_1d_array(filter_col.source_col, gagg_headers)
    
    Dim service_cols() As Variant
    Dim mail_cols() As Variant
    Dim arr() As Variant
    Dim is_comed As Boolean
    
    If EDC.ruleset_name = "COM" Then
        is_comed = True
    Else
        is_comed = False
    End If
    
    ReDim arr(2 To num_rows, 1 To 1)
    
    Set data_range = filter_tab.Cells(2, target_col).Resize(num_rows - 1)
    
    Select Case filter_col.data_subtype
        Case "Status":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
            Next
            GoTo populate_rng
        Case "Mail Category":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
            Next
            GoTo populate_rng
        Case "Address Source":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
            Next
            GoTo populate_rng
        Case "Mapping":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
            Next
            GoTo populate_rng
        Case "Opt In":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
            Next
            GoTo populate_rng
        Case "Customer Name":
            For j = 2 To num_rows
                arr(j, 1) = clean_name(gagg_data(j, source_col_num))
            Next
            GoTo populate_rng
        Case "Service Address":
            ReDim service_cols(0 To UBound(EDC.service))
            For j = 0 To UBound(EDC.service)
                service_cols(j) = search_headers(EDC.service(j), gagg_data)
            Next
            For j = 2 To num_rows
                arr(j, 1) = service_address(j, gagg_data, service_cols, is_comed)
                progress.activity (j)
            Next
            GoTo populate_rng
        Case "Service City":
            service_col = search_headers(EDC.service_city, gagg_data)
            For j = 2 To num_rows
                arr(j, 1) = service_city(j, gagg_data, service_col, is_comed)
                progress.activity (j)
            Next
            GoTo populate_rng
        Case "Service State":
            service_col = search_headers(EDC.service_state, gagg_data)
            For j = 2 To num_rows
                arr(j, 1) = service_state(j, gagg_data, service_col, is_comed)
                progress.activity (j)
            Next
            GoTo populate_rng
        Case "Service Zip":
            zip_col = search_headers(EDC.service_zip, gagg_data)
            For j = 2 To num_rows
                arr(j, 1) = service_zip(j, gagg_data, zip_col, is_comed)
                progress.activity (j)
            Next
            GoTo populate_rng
        Case "Mail Address":
            ReDim mail_cols(0 To UBound(EDC.mail))
            For j = 0 To UBound(EDC.mail)
                mail_cols(j) = search_headers(EDC.mail(j), gagg_data)
            Next
            For j = 2 To num_rows
                arr(j, 1) = mail_address(j, gagg_data, mail_cols)
                'arr(j, 1) = clean_mail_address(arr(j, 1))
                progress.activity (j)
            Next
            GoTo populate_rng
        Case "Mail City":
            mail_col = search_headers(EDC.mail_city, gagg_data)
            For j = 2 To num_rows
                arr(j, 1) = mail_city(j, gagg_data, mail_col)
                progress.activity (j)
            Next
            GoTo populate_rng
        Case "Mail State":
            mail_col = search_headers(EDC.mail_state, gagg_data)
            For j = 2 To num_rows
                arr(j, 1) = mail_state(j, gagg_data, mail_col)
                progress.activity (j)
            Next
            GoTo populate_rng
        Case "Mail Zip":
            zip_col = search_headers(EDC.mail_zip, gagg_data)
            For j = 2 To num_rows
                arr(j, 1) = mail_zip(j, gagg_data, zip_col)
                progress.activity (j)
            Next
            GoTo populate_rng
        Case "National Chains":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
                progress.activity (j)
            Next
            GoTo populate_rng
        Case Else:
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
                progress.activity (j)
            Next
            GoTo populate_rng
    End Select
populate_rng:
    data_range.value = arr
End Sub

Sub populate_calculated_filter_col(filter_tab, filter_col As ColumnHeader, target_col, gagg_headers, gagg_data, num_rows)
    
    Set data_range = filter_tab.Cells(2, target_col).Resize(num_rows - 1)
    
    If filter_col.source_col = "" Then
        GoTo empty_source
        Exit Sub
    End If
    
    source_col_num = search_1d_array(filter_col.source_col, gagg_headers)
    
    Dim arr() As Variant
    ReDim arr(2 To num_rows, 1 To 1)
    
    Select Case filter_col.data_subtype
        Case "Usage":
            progress.start ("Calculating Usage")
            Set data_range = data_range.Resize(num_rows - 1, 3)
            data_range.value = calculate_usage(filter_tab, arr, target_col, filter_col.source_col, num_rows)
        Case "Status":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
                progress.activity (j)
            Next
            data_range.value = arr
        Case "Class":
            For j = 2 To num_rows
                arr(j, 1) = cust_class(gagg_data(j, source_col_num))
                progress.activity (j)
            Next
            data_range.value = arr
        Case "Mail Category":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
                progress.activity (j)
            Next
            data_range.value = arr
        Case "Address Source":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
                progress.activity (j)
            Next
            data_range.value = arr
        Case "Opt In":
            For j = 2 To num_rows
                arr(j, 1) = filter_col.default_value
                progress.activity (j)
            Next
            data_range.value = arr
        Case "Arrears":
            data_range.value = AES_arrears(filter_tab, gagg_data, target_col, filter_col.source_col, num_rows)
        Case Else:
            Exit Sub
    End Select
    
    Exit Sub
    
empty_source:

    num_cols = 1
    default_value = filter_col.default_value
    If filter_col.data_subtype = "Usage" Then
        Set data_range = data_range.Resize(num_rows - 1, 3)
        num_cols = 3
    End If
    ReDim arr(2 To num_rows, 1 To num_cols)
    For j = 2 To num_rows
        For k = 1 To num_cols
            arr(j, k) = default_value
        Next
        progress.activity (j)
    Next
    data_range.value = arr
    
End Sub

Function filter_col_bool(value, yes_value)
    If value Like yes_value Then
        filter_col_bool = "Y"
    Else
        filter_col_bool = "N"
    End If
End Function

Sub hide_filter_group(group_name)
    Set ff = filter_tab()
    For j = 1 To UBound(F.order_array)
        If F.order_array(j).column_group = group_name Then filter_tab.columns(j).Hidden = True
    Next
End Sub

Sub show_filter_group(group_name)
    Set ff = filter_tab()
    For j = 1 To UBound(F.order_array)
        If F.order_array(j).column_group <> group_name Then filter_tab.columns(j).Hidden = True
    Next
End Sub
