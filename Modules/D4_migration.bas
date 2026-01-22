Sub test_migration()
    If Not MT.check_migration_data Then Exit Sub
    MT.check_migration_data = True
    define_migration_settings
    check_legacy_data
End Sub

Sub check_legacy_data()
    
    If Not MT.check_migration_data Then Exit Sub
    
    migration_log_file = onedrive_migration_folder() & "\" & S.migration.migration_log_file
    
    Set conn = ADO_connection_excel(migration_log_file)
    migration_log_rows = ADO_row_count(conn, S.migration.migration_log_sheet)
    migration_log_data = ADO_data(conn, S.migration.migration_log_sheet, "A2:G" & migration_log_rows, 1)
    conn.Close
    
    current_contracts = get_array_col(migration_log_data, 1)
    previous_contracts = get_array_col(migration_log_data, 2)
    system_EDC = get_array_col(migration_log_data, 5)
    previous_system_arr = get_array_col(migration_log_data, 6)
    migration_files = get_array_col(migration_log_data, 7)
    
    'get current contract
    current_contract = "C-00132523"
    
    migration_row = array_binary_search(current_contract, current_contracts, 1, UBound(current_contracts))

    If migration_row <= 0 Then
        Exit Sub
    End If
    
    migration_log_arr = get_array_row(migration_log_data, migration_row)
    
    previous_system_folder = migration_log_arr(6)
    
    migration_folder = onedrive_migration_folder()
    migration_folder = migration_folder & previous_system_folder
    migration_folder = migration_folder & "\Files By EDC\" & EDC.migration_name
    
    If previous_system_folder = "EH" Then
        k = 2
    Else
        k = 1
    End If
    
    migration_file = migration_folder & "\" & migration_log_arr(7) & ".xlsx"
    
    Set conn = ADO_connection_excel(migration_file)
    
    If conn Is Nothing Then Exit Sub
    
    row_count = ADO_row_count(conn, "Sheet1")
    community_data = ADO_data(conn, "Sheet1", "A2:B" & row_count, k)
    conn.Close
    
    num_rows = Application.CountA(filter_tab().columns(1))
    
    migration_account_arr = get_array_col(community_data, k)
    
    With F.columns
        account_arr = filter_col(.account_number)
        status_arr = filter_col(.status)
        migration_arr = filter_col(.migration_query)
    End With
    
    search_start = 1
    
    For i = 2 To num_rows
        migration_match = array_binary_search(account_arr(i, 1), migration_account_arr, search_start, row_count - 1)
        If migration_match > 0 Then
            search_start = i
            status_arr(i, 1) = FS.migration.ineligible_new_status
            migration_arr(i, 1) = migration_log_arr(3)
        End If
        progress.activity (i)
    Next
    
    filter_tab().Cells(1, F.columns.status.index).Resize(num_rows).value = status_arr
    filter_tab().Cells(1, F.columns.migration_query.index).Resize(num_rows).value = migration_arr
    
End Sub
