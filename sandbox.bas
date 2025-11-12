Sub dp(S)
    Debug.Print S
End Sub

Function flatten_array(d2) As Variant

    On Error GoTo already_flat
    
    Dim d1() As Variant
    
    max_dim = WorksheetFunction.Max(UBound(d2, 1), UBound(d2, 2))
    
    ReDim d1(1 To max_dim)
    
    vertical_array = max_dim = UBound(d2, 1)
    horizontal_array = max_dim = UBound(d2, 2)
    
    For i = 1 To max_dim
        If vertical_array Then
            d1(i) = d2(i, 1)
        Else
            d1(i) = d2(1, i)
        End If
    Next
    
    flatten_array = d1
    
    Exit Function
    
already_flat:
    flatten_array = d2
    
End Function

Function array_binary_search(target, search_array, search_start, search_end)

    'searches for a target value in an ordered 1D array
    
    array_binary_search = 0
    
    If target < search_array(search_start) Then
        Exit Function
    ElseIf target > search_array(search_end) Then
        Exit Function
    ElseIf search_start >= search_end Then
        Exit Function
    ElseIf search_end - search_start = 1 Then
        If search_array(search_start) = target Then
            array_binary_search = search_start
        ElseIf search_array(search_end) = target Then
            array_binary_search = search_end
        Else
            Exit Function
        End If
    Else
        search_middle = Int((search_start + search_end) / 2)
        If target = search_array(search_middle) Then
            array_binary_search = search_middle
            Exit Function
        ElseIf target < search_array(search_middle) Then
            array_binary_search = array_binary_search(target, search_array, search_start, search_middle)
        ElseIf target > search_array(search_middle) Then
            array_binary_search = array_binary_search(target, search_array, search_middle, search_end)
        End If
    End If

End Function

Function find_column_header(target_string, sheet_num)
    search_range = Sheets(sheet_num).UsedRange.Rows(1).value
    For j = 1 To UBound(search_range, 2)
        If search_range(1, j) Like target_string & "*" Then
            find_column_header = j
            Exit Function
        End If
    Next
End Function

Function search_1d_array(target_val, arr As Variant) As Variant
    If IsArray(target_val) Then
        search_1d_array = Array()
        If IsEmpty(target_val) Then Exit Function
        n = UBound(arr)
        Dim return_arr() As Variant
        ReDim return_arr(0 To UBound(target_val))
        For j = 0 To UBound(target_val)
            For k = 1 To n
                If arr(k) = target_val(j) Or UCase(Application.Trim(arr(k))) = target_val(j) Then
                    'search_1d_array = j
                    return_arr(j) = k
                    Exit For
                End If
            Next
        Next
        search_1d_array = return_arr
    Else
        search_1d_array = -1
        If target_val = "" Then Exit Function
        For j = LBound(arr) To UBound(arr)
            If arr(j) = target_val Or UCase(Application.Trim(arr(j))) = target_val Then
                search_1d_array = j
                Exit For
            End If
        Next
    End If
End Function

Function get_array_row(arr_2d As Variant, row_number) As Variant
    
    Dim result() As Variant
    ReDim result(1 To UBound(arr_2d, 2))
    
    For j = LBound(arr_2d, 2) To UBound(arr_2d, 2)
        result(j) = arr_2d(row_number, j)
    Next

    get_array_row = result
    
End Function

Function get_array_col(arr_2d As Variant, col_number) As Variant
    
    Dim result() As Variant
    ReDim result(1 To UBound(arr_2d, 1))

    For j = LBound(arr_2d, 1) To UBound(arr_2d, 1)
        result(j) = arr_2d(j, col_number)
    Next

    get_array_col = result
    
End Function
Function GetFillRGB(cell As Range) As String
    Dim colorVal As Long
    Dim r As Long, g As Long, b As Long

    colorVal = cell.Interior.color
    r = colorVal Mod 256
    g = (colorVal \ 256) Mod 256
    b = (colorVal \ 65536) Mod 256

    GetFillRGB = "RGB(" & r & ", " & g & ", " & b & ")"
End Function

Sub rename_sheet(ws, count)
    
End Sub

Function search_headers(target, d2 As Variant)
    search_headers = 0
    For k = 1 To UBound(d2, 2)
        If d2(1, k) = target Then
            search_headers = k
            Exit Function
        End If
    Next
End Function

Function is_arr_init(arr As Variant) As Boolean
    On Error Resume Next
    is_arr_init = Not IsEmpty(arr) And Not IsError(LBound(arr)) And LBound(arr) <= UBound(arr)
    On Error GoTo 0
End Function

Function find_filter_tab_col(target As ColumnHeader)
    On Error Resume Next
    For j = 1 To UBound(F.order_array)
        If F.order_array(j).header = target.header Then
            find_filter_tab_col = j
            Exit Function
        End If
    Next
End Function

Sub filter_arr_append(ByRef arr() As ColumnHeader, value As ColumnHeader)
    If arr(1).header = "" Then
        arr(1) = value
        value.index = 1
    Else
        n = UBound(arr)
        ReDim Preserve arr(LBound(arr) To n + 1)
        arr(n + 1) = value
        value.index = n + 1
    End If
End Sub

Sub active_arr_append(ByRef arr() As ActiveColumnHeader, value As ActiveColumnHeader)
    If arr(1).header = "" Then
        arr(1) = value
    Else
        n = UBound(arr)
        ReDim Preserve arr(LBound(arr) To n + 1)
        arr(n + 1) = value
    End If
End Sub

Function trim_arr_end_column(arr As Variant) As Variant
    
    Dim result() As Variant
    
    row_count = UBound(arr, 1)
    col_count = UBound(arr, 2)
    
    ReDim result(1 To row_count, 1 To col_count - 1)
    
    For row = 1 To row_count
        For col = 1 To col_count - 1
            result(row, col) = arr(row, col)
        Next
    Next
    
    trim_arr_end_column = result
    
End Function

Function deduped_data_arr(ws) As Variant
    arr = ws.UsedRange.Value2
    n1 = UBound(arr, 1)
    n2 = UBound(arr, 2)
    Dim clean_arr As Variant
    clean_arr = Application.Transpose(arr)
    For j = 1 To n1
        If arr(j, 1) = "" Then
            k = j - 1
            Exit For
        End If
    Next
    If k > 0 Then
        ReDim Preserve clean_arr(1 To n2, 1 To k)
        deduped_data_arr = Application.Transpose(clean_arr)
    Else
        deduped_data_arr = arr
    End If
End Function

Function add_file_input(file, file_source) As Range
    Set h = home_tab()
    k = 1
    While h.Range(S.HOME.file_log_location).Offset(k, 0) <> ""
        k = k + 1
        If k > 100 Then Exit Function
    Wend
    Set cell = h.Range(S.HOME.file_log_location).Offset(k, 0)
    cell.Offset(0, 0) = file_source
    cell.Offset(0, 1) = clean_file_name(file)
    cell.Offset(0, 2 + cell.Offset(0, 1).MergeArea.Cells.count - 1) = DateValue(FileDateTime(file))
    Set add_file_input = cell.Offset(0, 3 + cell.Offset(0, 1).MergeArea.Cells.count - 1)
End Function

Function remove_file_input(file_source)
    Set h = home_tab()
    k = 1
    While h.Range(S.HOME.file_log_location).Offset(k, 0) <> ""
        k = k + 1
        If k > 100 Then Exit Function
    Wend
    Set cell = h.Range(S.HOME.file_log_location).Offset(k - 1, 0)
    If cell = file_source Then
        cell.Offset(0, 0) = ""
        cell.Offset(0, 1) = ""
        cell.Offset(0, 2 + cell.Offset(0, 1).MergeArea.Cells.count - 1) = ""
        cell.Offset(0, 3 + cell.Offset(0, 1).MergeArea.Cells.count - 1) = ""
        cell.Offset(0, 4 + cell.Offset(0, 1).MergeArea.Cells.count - 1) = ""
    End If
End Function

Function previous_day(day)
    m = Split(day, "-")(0)
    d = Split(day, "-")(1)
    y = Split(day, "-")(2)
    dd = d - 1
    If dd = 0 Then
        dd = 31
        mm = m - 1
        YY = y
        If mm = 0 Then
            mm = 12
            YY = y - 1
        End If
    Else
        mm = m
        YY = y
    End If
    previous_day = mm & "-" & dd & "-" & YY
End Function

Function ADO_connection_excel(file_path) As Object
    
    'file_path = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(4) PUCO Do Not Aggregate (DNA) List\Do Not Aggregate List Lookup.accdb"
    
    On Error GoTo no_file
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & file_path & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"";"
    
    Set ADO_connection_excel = conn
    
    Exit Function
    
no_file:
    Set ADO_connection_excel = Nothing
    
End Function

Function ADO_row_count(ByRef conn, sheet_name) As Long

    If conn Is Nothing Then Exit Function
    
    Dim rs As Object
    
    Set rs = CreateObject("ADODB.Recordset")
    
    Sql = "SELECT COUNT(F1) AS RowCount FROM [" & sheet_name & "$A:A] WHERE F1 IS NOT NULL"
    
    rs.Open Sql, conn, 1, 1
    
    If Not rs.EOF Then
        ADO_row_count = rs.Fields("RowCount").value
    Else
        ADO_row_count = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Function ADO_data(ByRef conn, sheet_name, cell_range, sort_col) As Variant

    If conn Is Nothing Then Exit Function
    
    Dim rs As Object
    
    Set rs = CreateObject("ADODB.Recordset")
    
    'dna_name_col = 4
    'dna_address_col = 5
    
    sort_str = "F" & sort_col
    
    If sheet_name Like S.DNA.sheet_name & "*" Then
        Sql = "SELECT F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,UCASE(F4),LEFT(UCASE(F5)," & S.DNA.wildcard_length & ") " & _
            "FROM [" & sheet_name & "$" & cell_range & "] " & _
            "WHERE F1 IS NOT NULL ORDER BY " & sort_str & " ASC"
    Else
        Sql = "SELECT * FROM [" & sheet_name & "$" & cell_range & "] " & _
            "WHERE F1 IS NOT NULL ORDER BY " & sort_str & " ASC"
    End If
    
    rs.Open Sql, conn, 1, 1
    
    If rs.EOF Then
        Set ADO_data = Nothing
    Else
        arr = rs.GetRows()
        ADO_data = transpose_ADO(arr)
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Function transpose_ADO(arr) As Variant

    d1 = UBound(arr, 1)
    d2 = UBound(arr, 2)
    
    Dim new_arr() As Variant
    ReDim new_arr(1 To d2 + 1, 1 To d1 + 1)
    
    For row_num = 1 To d2 + 1
        For col_num = 1 To d1 + 1
            new_arr(row_num, col_num) = arr(col_num - 1, row_num - 1)
        Next
    Next
    
    transpose_ADO = new_arr
    
End Function

Function sort_2d_arr(arr, col) As Variant

    Set ws = Sheets.Add(before:=Sheets(1))
    
    ws.Cells(1, 1).Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(1, col), ws.Cells(UBound(arr, 1), col)), _
            SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2))
        .header = xlNo
        .Apply
    End With
    
    sort_2d_arr = ws.UsedRange.value
    
    delete_sheet (1)
    
End Function

Sub qc_checklist()
    Set d = CreateObject("Scripting.Dictionary")
    d.Add ChrW(&H2714), RGB(0, 175, 0) 'green check
    d.Add ChrW(&H2718), RGB(255, 0, 0) 'red x
    d.Add ChrW(&H25C9), RGB(225, 200, 0) 'yellow circle
    With home_tab()
        For i = 31 To 41
            arr = GetRandomDictEntry(d)
            .Cells(i, "M") = arr(0)
            .Cells(i, "M").Font.color = arr(1)
            .Cells(i, "M").HorizontalAlignment = xlCenter
        Next
    End With
End Sub

Sub help()
    Load help_doc_form
    help_doc_form.Show vbModeless
End Sub
