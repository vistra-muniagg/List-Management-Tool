Sub test_mapping()
    'init
    remove_other_ineligible
End Sub

Sub remove_other_ineligible()
    import_mapping
    If mapping_tab() Is Nothing Then Exit Sub
    If check_mapping() = False Then
        remove_file_input (S.mapping.file_source)
        Call update_checklist(S.QC.qc_checklist, "correct_mapping", -1)
        Exit Sub
    End If
    Call update_checklist(S.QC.qc_checklist, "correct_mapping", 1)
    remove_dna
    process_contracts
    process_mapping
    misc_filter
    set_step (6)
End Sub

Sub import_mapping()

    If Not mapping_tab() Is Nothing Then Exit Sub
    
    If T.mapping_file = "" Then
        file_name = Application.GetOpenFilename("Geocoding Files (*.xlsm), *.xlsm", , "Select Mapping Results File")
    Else
        mapping_folder = "C:\Users\400050\OneDrive - Vistra Corp\(6) List Management\(2) Macro Testing\(1) Testing Files\Test Mapping\"
        file_name = mapping_folder & T.mapping_file
    End If
    
    If VarType(file_name) = 11 Then Exit Sub
    
    before_sheet = home_tab().index
    
    Call import_excel_file(file_name, S.mapping.file_source, before_sheet, False)
    
    dupes = remove_duplicates(before_sheet)
    
    Set k = add_file_input(file_name, S.mapping.file_source)
    
    k.Offset(0, 1).value = dupes(0) - 1
    
    k.value = 0 'should be no dupes in mapping
    
    reapply_autofilter (before_sheet)
    
    Sheets(before_sheet).name = SN.mapping
       
End Sub

Function check_mapping() As Variant
    
    Set m = mapping_tab()
    Set ff = filter_tab()
    
    mapped_rows = Application.CountA(m.columns(1))
    num_rows = Application.CountA(ff.columns(1))
    
    If mapped_rows <> num_rows Then
        check_mapping = False
        GoTo bad_mapping_file
    End If
    
    mapping_col = Application.match(S.mapping.mapping_col, m.UsedRange.Rows(1), 0)
    mapping_arr = m.UsedRange.columns(mapping_col).value
    
    maps_out_count = 0
    no_results_count = 0
    
    For j = 2 To mapped_rows
        x = UCase(mapping_arr(j, 1))
        If x <> "N" And x <> "Y" And x <> S.mapping.no_results_label Then
            check_mapping = False
            delete_sheet (SN.mapping)
            Exit Function
        End If
        If x = "N" Then
            maps_out_count = maps_out_count + 1
        ElseIf x = S.mapping.no_results_label Then
            no_results_count = no_results_count + 1
        End If
    Next
    
    map_out_pct = Round(100 * maps_out_count / (mapped_rows - 1), 2)
    no_result_pct = Round(100 * no_results_count / (mapped_rows - 1), 2)
    
    If T.name = "" Then
        If map_out_pct > S.mapping.map_out_limit Then
            msg = MsgBox("Percentage of Mapped Out accounts exceeds " & S.mapping.map_out_limit & "%. Did you check with Kevin before continuing?" & _
                            vbCrLf & vbCrLf & "Mapped Out Percent = " & map_out_pct & " %", vbExclamation + vbYesNo)
            If msg = vbNo Then
                check_mapping = False
                delete_sheet (SN.mapping)
                Exit Function
            End If
        End If
        If no_result_pct > S.mapping.no_result_limit Then
            msg = MsgBox("Percentage of No Result accounts exceeds " & S.mapping.no_result_limit & "%. Did you check with Kevin before continuing?" & _
                            vbCrLf & vbCrLf & "No Result Percent = " & no_result_pct & " %", vbExclamation + vbYesNo)
            If msg = vbNo Then
                check_mapping = False
                delete_sheet (SN.mapping)
                Exit Function
            End If
        End If
    End If
    
    check_mapping = True
    
    Exit Function
    
bad_mapping_file:
    MsgBox "Number of mapped accounts does not match expected count", vbCritical
    delete_sheet (SN.mapping)
    check_mapping = False
    
End Function

Sub process_mapping()

    Set m = mapping_tab()
    Set ff = filter_tab()
    
    Call sort_sheet_col(m, 1, "A")
    Call sort_sheet_col(ff, 1, "A")
    
    prior_status_col = F.columns.before_mapping_eligible.index
    
    mapping_col = Application.match(S.mapping.mapping_col, m.UsedRange.Rows(1), 0)
    community_col = Application.match(S.mapping.mapped_community, m.UsedRange.Rows(1), 0)
    notes_col = Application.match(S.mapping.notes_col, m.UsedRange.Rows(1), 0)
    
    mapping_arr = m.UsedRange.columns(mapping_col).value
    community_arr = m.UsedRange.columns(community_col).value
    notes_arr = m.UsedRange.columns(notes_col).value
    
    With F.columns
        mapped_community_arr = filter_col(.community_mapped_into)
        status_arr = filter_col(.status)
        before_mapping_arr = status_arr
        mapping_result_arr = filter_col(.mapping_result)
        mapping_notes_arr = filter_col(.mapping_notes)
        active_arr = filter_col(.active_in_LP)
        eligible_arr = filter_col(.eligible_opt_out)
    End With
    
    num_rows = Application.CountA(ff.columns(1))
    mapping_rows = Application.CountA(m.columns(1))
    
    For j = 2 To mapping_rows
    
        mapping_notes_arr(j, 1) = notes_arr(j, 1)
        mapped_community_arr(j, 1) = community_arr(j, 1)
        'before_mapping_arr(j, 1) = status_arr(j, 1)
        
        If mapping_arr(j, 1) = "" Then
            'empty result
            'exit?
        ElseIf UCase(mapping_arr(j, 1)) = "Y" Then
            mapping_arr(j, 1) = FS.mapping.maps_in_label
        ElseIf UCase(mapping_arr(j, 1)) = "N" Then
            If active_arr(j, 1) = "Y" Then
                If MT.keep_active_mapped_out Then
                    'status_arr(j, 1) = FS.mapping.mapped_out_retained_status
                    status_arr(j, 1) = FS.eligible.eligible_ren_status
                    mapping_arr(j, 1) = FS.mapping.mapped_out_retained_label
                Else
                    status_arr(j, 1) = FS.mapping.ineligible_ren_status
                    mapping_arr(j, 1) = FS.mapping.mapped_out_label
                    eligible_arr(j, 1) = "N"
                End If
            Else
                status_arr(j, 1) = FS.mapping.ineligible_new_status
                mapping_arr(j, 1) = FS.mapping.mapped_out_label
                eligible_arr(j, 1) = "N"
            End If
        Else
            mapping_arr(j, 1) = FS.mapping.no_results_label
        End If
        
        mapping_result_arr(j, 1) = mapping_arr(j, 1)
        
        progress.activity (j)
        
    Next
    
    before_mapping_arr(1, 1) = F.columns.before_mapping_eligible.header
    
    ff.Cells(1, F.columns.status.index).Resize(num_rows).value = status_arr
    ff.Cells(1, F.columns.eligible_opt_out.index).Resize(num_rows).value = eligible_arr
    ff.Cells(1, F.columns.community_mapped_into.index).Resize(num_rows).value = mapped_community_arr
    ff.Cells(1, F.columns.before_mapping_eligible.index).Resize(num_rows).value = before_mapping_arr
    ff.Cells(1, F.columns.mapping_result.index).Resize(num_rows).value = mapping_result_arr
    ff.Cells(1, F.columns.mapping_notes.index).Resize(num_rows).value = mapping_notes_arr
    
    m.Cells(1, mapping_col).Resize(mapping_rows).value = mapping_arr
    
    Call update_checklist(S.QC.audit_checklist, "audit_mapping", 1)
    
    make_geocode_waterfall
    
End Sub

Sub create_map_this()
    delete_sheet ("Map This")
    Dim cols() As ColumnHeader
    ReDim cols(1 To 6)
    Set map_this = Sheets.Add(after:=Sheets(Sheets.count))
    map_this.name = S.mapping.map_this_sheet
    With F.columns
        cols(1) = .account_number
        cols(2) = .service_address
        cols(3) = .service_city
        cols(4) = .service_state
        cols(5) = .service_zip
    End With
    num_rows = Application.CountA(filter_tab().columns(1))
    For k = 1 To UBound(cols) - 1
        col_data = filter_col(cols(k))
        map_this.Cells(1, k).Resize(num_rows).value = col_data
        map_this.Cells(1, k) = cols(k).header
    Next
    reapply_autofilter (Sheets.count)
    map_this.Tab.color = C.GREEN_1.InteriorColor
End Sub

Sub generate_mapping()
    progress.start ("Generating Mapping")
    folder_path = onedrive_list_management_folder()
    Set ff = filter_tab()
    mapping_file = Dir(folder_path & "\" & S.mapping.mapping_tool_file)
    actual_path = Replace(ThisWorkbook.path, S.OneDrive.ops_folder_url, onedrive_parent_folder())
    If mapping_file = "" Then
        'idk what to do here
    End If
    a1 = filter_col(F.columns.account_number)
    s2 = filter_col(F.columns.service_address)
    s3 = filter_col(F.columns.service_city)
    s4 = filter_col(F.columns.service_state)
    s5 = filter_col(F.columns.service_zip)
    community_name = get_community_name()
    s20 = filter_col(F.columns.address_source)
    num_rows = UBound(a1, 1)
    
    Dim mapping_data As Variant
    ReDim mapping_data(1 To num_rows, 1 To 20)
    
    community_name = folder_community_name(community_name)
    
    For i = 2 To num_rows
        mapping_data(i, 1) = a1(i, 1)
        mapping_data(i, 2) = s2(i, 1)
        mapping_data(i, 3) = s3(i, 1)
        mapping_data(i, 4) = s4(i, 1)
        mapping_data(i, 5) = s5(i, 1)
        mapping_data(i, 6) = EDC.display_name
        mapping_data(i, 13) = community_name
        mapping_data(i, 20) = s20(i, 1)
        progress.activity (i)
    Next
    Application.ScreenUpdating = False
    
    full_path = ThisWorkbook.path & "/" & community_name & " " & MT.name & " Mapping.xlsm"
    
    ' Check if file already exists
    On Error GoTo generate_mapping
    If Dir(full_path) <> "" Then
        If MsgBox("Warning: Mapping file already exists." & vbCrLf & vbCrLf & _
                "Do you want to overwrite it?", vbYesNo + vbExclamation) = vbNo Then
            GoTo exitsub
        End If
    End If
    
generate_mapping:
    Set wb = ThisWorkbook
    wb.Activate
    Set w1 = Workbooks.Open(folder_path & "\" & mapping_file, ReadOnly:=True, addtomru:=False)
    With w1
        .Windows(1).visible = True
        mapping_headers = .Sheets(1).UsedRange.Rows(1).value
        .Sheets(1).columns(1).NumberFormat = "@"
        .Sheets(1).Range("A1").Resize(num_rows, 20).value = mapping_data
        .Sheets(1).Range("A1").Resize(1, 20).value = mapping_headers
        .SaveAs fileName:=full_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        .Close False
    End With
exitsub:
    progress.complete
    Application.ScreenUpdating = True
End Sub

Function folder_community_name(str)
    'str = UCase(str)
    If str Like "Village of *" Then str = Replace$(str, "Village of ", "") & " (V)"
    If str Like "City of *" Then str = Replace$(str, "City of ", "") & " (C)"
    folder_community_name = str
End Function
