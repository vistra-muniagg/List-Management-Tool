Sub test_import()
    progress.start ("Importing Files")
    active_list_folder = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(2) Macro Testing\(1) Testing Files\Test Active Lists\"
    gagg_list_folder = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(2) Macro Testing\(1) Testing Files\Test GAGG Lists\"
    If MT.needs_active_list Then Call import_active_list(active_list_folder & T.active_list, "snowflake", 1)
    Set hdict = CreateObject("Scripting.Dictionary")
    'Call import_file(gagg_list_folder & T.gagg_list, "source", 1)
    If MT.needs_gagg_list Then Call import_gagg_files(gagg_list_folder & T.gagg_list, "Utility", 1)
    progress.finish
End Sub

Sub import_active_list(Optional file_name, Optional file_source, Optional target_location)
    'If Not all_initialized Then init
    If imported_active Then Exit Sub
    If IsMissing(file_name) Then
        selected_file = Application.GetOpenFilename("Active Lists (*.csv), *.csv", , "Select Active List", False)
    Else
        selected_file = file_name
    End If
    'progress.start ("Importing Active List")
    'delete_sheet (SN.Active)
    If selected_file = False Then Exit Sub
    target_location = 1
    file_source = "LandPower"
    Call import_csv_file(selected_file, file_source, target_location)
    Sheets(target_location).name = SN.Active
    dupes = remove_duplicates(target_location)
    reapply_autofilter (target_location)
    Call sort_sheet_col(Sheets(SN.Active), 1, "A")
    Set k = add_file_input(selected_file, file_source)
    k.Offset(0, 1).value = dupes(0) - 1
    k.value = dupes(1)
    'If dupes(1) > 0 Then k.AddComment dupes(1) & " Duplicates"
    'progress.complete
    imported_active = True
    If Not UI Is Nothing Then
        UI.InvalidateControl ("import_menu")
        UI.InvalidateControl ("filter_button")
    End If
End Sub

Sub import_gagg_files(Optional file_name, Optional file_source, Optional target_location)
    'If Not all_initialized Then init
    If imported_gagg Then Exit Sub
    'select files
    If IsMissing(file_name) Then
        selected_files = Application.GetOpenFilename("Select Utility File(s), *.*", , "Utility File(s)", multiselect:=EDC.multiselect)
    Else
        selected_files = file_name
    End If
        
    'if no files then exit?
    
    If IsArray(selected_files) Then
        If IsError(UBound(selected_files)) Then Exit Sub
    Else
        If selected_files = "" Or selected_files = False Then Exit Sub
    End If
    
    progress.start ("Importing Utility List(s)")
    If IsArray(selected_files) Then
        For Each gagg_file In selected_files
            file_name = gagg_file
            Call import_file(gagg_file, "Utility", 1)
            comment_text = comment_text & clean_file_name(gagg_file) & vbCrLf
        Next
    Else
        file_name = selected_files
        Call import_file(selected_files, "Utility", 1)
    End If
    n = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.name Like "Sheet*" Then n = ws.index
    Next
    Call combine_sheets(n, 1)
    format_accounts_col
    dupes = remove_duplicates(1)
    If IsMissing(file_source) Then file_source = "Utility"
    Set k = add_file_input(file_name, file_source)
    If IsArray(selected_files) Then
        If UBound(selected_files) > 1 Then
            k.value = "(Multiple)"
            k.Offset(0, -2).AddComment comment_text
        Else
            k.value = clean_file_name(file_name)
        End If
    Else
        k.value = clean_file_name(file_name)
    End If
    k.Offset(0, 1).value = dupes(0) - 1
    k.value = dupes(1)
    'If dupes(1) > 0 Then k.AddComment dupes(1) & " Duplicates"
    progress.finish
    'ReDim Stats.file_stats.utility_files(1 To 1)
    'Set dict = CreateObject("Scripting.Dictionary")
    imported_gagg = True
    If Not UI Is Nothing Then
        UI.InvalidateControl ("import_menu")
        UI.InvalidateControl ("filter_button")
    End If
End Sub

Sub import_supplier_list(Optional file_name, Optional file_source, Optional target_location)
    'If Not all_initialized Then init
    If imported_supplier Then Exit Sub
    If IsMissing(file_name) Then
        selected_file = Application.GetOpenFilename("Previous Supplier Lists (*.*), *.*", , "Select Previous Supplier List", False)
    Else
        selected_file = file_name
    End If
    If selected_file = False Then Exit Sub
    selected_file = UCase(selected_file)
    target_location = 1
    file_source = "Previous Supplier"
    If selected_file Like "*.CSV" Then
        Call import_csv_file(selected_file, file_source, target_location)
    Else
        Call import_excel_file(selected_file, file_source, target_location, False)
    End If
    Sheets(target_location).name = SN.Supplier
    dupes = remove_duplicates(target_location)
    reapply_autofilter (target_location)
    Call sort_sheet_col(Sheets(SN.Supplier), 1, "A")
    Set k = add_file_input(selected_file, file_source)
    k.Offset(0, 1).value = dupes(0) - 1
    k.value = dupes(1)
    imported_supplier = True
    If Not UI Is Nothing Then
        UI.InvalidateControl ("import_menu")
        UI.InvalidateControl ("filter_button")
    End If
End Sub

Sub import_file(file_name, file_source, target_location)
    file_name = UCase(file_name)
    ext = Mid(file_name, InStrRev(file_name, ".") + 1)
    If ext = "CSV" Then
        Call import_csv_file(file_name, file_source, target_location)
    ElseIf ext Like "XLS*" Then
        Call import_excel_file(file_name, file_source, target_location, EDC.import_all_gagg_sheets)
    End If
    ChDir parent_folder(file_name)
    'dict.Add dict.count, Array(file_source, file_name, file_count)
End Sub

Sub import_excel_file(file_name, file_source, target_location, import_all)
    
    Set w = Workbooks.Open(file_name, ReadOnly:=True, addtomru:=False)
    'w.Windows(1).visible = True
    
    With ThisWorkbook
    
        Set s1 = .Sheets(target_location)
        
        k = 1
        Do While k <= w.Sheets.count
            If EDC.ruleset_name = "FE" And k > 2 Then Exit Do
            Set ws = w.Sheets(k)
            Set source_data = ws.UsedRange
            n = source_data.Rows.count * source_data.columns.count
            If n > S.Import.max_copy_size Then
                Set s1 = .Sheets.Add(before:=.Sheets(target_location))
                source_data.Copy
                If file_source = S.mapping.file_source Then s1.columns(1).NumberFormat = "@"
                s1.Range("A1").PasteSpecial xlPasteValues
            Else
                Set import_sheet = .Sheets.Add(before:=s1)
                If file_source = S.mapping.file_source Then import_sheet.columns(1).NumberFormat = "@"
                d = ws.UsedRange.Value2
                num_r = UBound(d, 1)
                num_c = UBound(d, 2)
                account_col = find_column_header(EDC.account, ws.index)
                If account_col > 0 Then
                    import_sheet.columns(account_col).NumberFormat = "@"
                End If
                If EDC.ruleset_name = "FE" And k > 1 And S.Import.FE_address_replace Then
                    import_sheet.columns(1).NumberFormat = "@"
                End If
                import_sheet.Range("A1").Resize(num_r, num_c).value = d
            End If
            If Not import_all Then Exit Do
            k = k + 1
        Loop
        
        w.Close False
        
        If EDC.ruleset_name = "FE" And S.Import.FE_address_replace And Sheets(2).name Like "Sheet*" Then
            Sheets(2).name = "FE Mail"
        End If
    
    End With
    
End Sub

Sub import_csv_file(file_name, file_source, target_location)

    If VarType(file_name) = 11 Then Exit Sub
    
    max_cols = S.Import.max_csv_cols
    
    ReDim columnDataTypes(1 To max_cols)
    
    For i = 1 To max_cols
        columnDataTypes(i) = 1 ' GENERAL
        If i = 1 Then columnDataTypes(i) = 2 ' TEXT
    Next
    
    Set s1 = Sheets.Add(before:=Sheets(target_location))
    
    s1.Activate
    
    With s1.QueryTables.Add(Connection:="TEXT;" & file_name, Destination:=s1.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = columnDataTypes
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    s1.Rows(1).Font.Bold = True
    s1.columns.AutoFilter
    s1.columns.AutoFit

End Sub

Sub combine_sheets(k, target_location, Optional trim_sheets As Boolean)
    
    'combine first k sheets and put them at target location
    
    delete_sheet (SN.Utility)
    
    For j = 1 To k
        trim_sheet (j)
        'trim_headers (j)
        move_accounts_to_front (j)
        reapply_autofilter (j)
    Next
    
    If k <= 1 Then
        Sheets(1).name = SN.Utility
        Exit Sub
    End If
    
    If k >= Sheets.count Then Exit Sub
    
    If target_location < 1 Then Exit Sub
    
    Set s1 = Sheets.Add(after:=Sheets(Sheets.count))
    
    col_count = Sheets(1).UsedRange.columns.count
    
    Set headers = Sheets(1).UsedRange.Rows(1)
    
    s1.Range("A1").Resize(, col_count).value = headers.value
    
    s1.Rows(1).WrapText = False
    
    row_count = 1
    
    n = 0
    Do While n < k
        add_rows = Sheets(n + 1).UsedRange.Rows.count - 1
        Set r = Sheets(n + 1).UsedRange.Offset(1, 0).Resize(add_rows)
        s1.Cells(row_count + 1, 1).Resize(add_rows, col_count).value = r.value
        row_count = row_count + add_rows
        n = n + 1
    Loop
    
    trim_headers (s1.index)
    
    For n = 1 To k
        delete_sheet (1)
    Next
    
    If target_location > Sheets.count Then target_location = Sheets.count
    
    If target_locaton = 1 Then
        s1.Move before:=Sheets(1)
    ElseIf target_location = Sheets.count Then
        s1.Move after:=Sheets(Sheets.count)
    Else
        s1.Move before:=Sheets(target_location)
    End If
    
    s1.name = SN.Utility

End Sub

Sub trim_sheet(k)
        
    'trim off junk data
    'trim headers
    'rename headers
    'delete empty columns at end
    
    Select Case EDC.ruleset_name
        Case "AEP":
            With Sheets(k)
                num_rows = Application.CountA(.columns(1))
                num_cols = Application.CountA(.Rows(1))
                .columns(1).Delete
                .Rows(num_rows).Delete
                .Rows(2).Delete
                trim_AEP_empty_end_cols (k)
            End With
        Case "AES":
            With Sheets(k)
                Set cell_range = .Range("A1:A10")
                trim_rows = 0
                For Each cell In cell_range
                    If Application.Trim(cell.value) = "" Then
                        trim_rows = cell.row
                        Exit For
                    End If
                Next
                If trim_rows > 0 Then
                    Do While trim_rows > 0
                        .Rows(1).Delete
                        trim_rows = trim_rows - 1
                    Loop
                End If
            End With
        Case "AM":
            With Sheets(k)
                Set cell_range = .Range("A1:A10")
                trim_rows = 0
                For Each cell In cell_range
                    If Application.Trim(cell.value) Like "Please*" Then
                        trim_rows = cell.row
                        Exit For
                    End If
                Next
                trim_rows = trim_rows + 1
                Do While trim_rows > 0
                    .Rows(1).Delete
                    trim_rows = trim_rows - 1
                Loop
                'unmerge columns
                col_count = Application.CountA(.Rows(1))
                Set right_cell = .Cells(1, col_count)
                col = right_cell.End(xlToLeft).Column
                While col > 1
                    If .Cells(1, col - 1) = "" Then
                        .columns(col - 1).Delete
                        col = right_cell.End(xlToLeft).Column
                    End If
                Wend
            End With
        Case "DUKE":
        
        Case "FE":
            With Sheets(k)
                If IsEmpty(.Range("A2")) Then .columns(1).Delete
                If IsEmpty(.Range("A2")) Then .Rows(2).Delete
            End With
        
    End Select
    
End Sub

Function remove_duplicates(sheet_num) As Variant
    
    Dim arr
    ReDim arr(0 To 1)
    
    With Sheets(sheet_num)
        If .name = SN.Filter Then GoTo dedupe_filter
        original_count = Application.CountA(.columns(1))
        .UsedRange.RemoveDuplicates columns:=1, header:=xlYes
        deduped_count = Application.CountA(.columns(1))
        If original_count > deduped_count Then
            'for k=deduped to
        End If
        Call sort_sheet_col(Sheets(sheet_num), 1, "A")
        reapply_autofilter (sheet_num)
        arr(0) = deduped_count
        arr(1) = original_count - deduped_count
        remove_duplicates = arr
        Exit Function
    
dedupe_filter:
        progress.start ("Removing Duplicates")
        original_count = Application.CountA(.columns(1))
        deduped_count = Application.CountA(.columns(1))
        If original_count > deduped_count Then
            eligible_arr = filter_col(F.columns.eligible_opt_out)
            arr = filter_col(F.columns.account_number)
            status_arr = filter_col(F.columns.status)
            For j = 2 To num_row
                
            Next
        End If
        progress.complete
        arr(0) = deduped_count
        arr(1) = original_count - deduped_count
        remove_duplicates = arr
    End With
    
End Function

Sub format_accounts_col()
    With Sheets(SN.Utility)
        account_col_data = .UsedRange.columns(1).value
        .columns(1).NumberFormat = "@"
        .UsedRange.columns(1).value = format_accounts(account_col_data)
    End With
End Sub
