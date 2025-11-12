Function prompt_review()
    prompt_review = False
    If ribbon_contract_number = "" Then GoTo bad_contract_id
    If ribbon_opt_out_date = "" Then GoTo bad_oo_date
    If Not ribbon_contract_number Like "C-00######" Then GoTo bad_contract_id
    Set ff = filter_tab()
    ff.Activate
    Call sort_sheet_col(ff, F.columns.eligible_opt_out.index, "D")
    ff.UsedRange.AutoFilter Field:=F.columns.eligible_opt_out.index, Criteria1:="Y"
    show_filter_group ("LP")
    reorder_tabs
    ff.Activate
    prompt_review = True
    set_step (7)
    'review_instructions
    Exit Function
    
bad_contract_id:
bad_oo_date:
    MsgBox "Populate the Contract ID and Opt Out Date Fields in the ribbon above with valid data"
End Function

Sub review_contract_data()
    
    If Not get_contract_id() Like "C-00######" Then GoTo bad_contract
    If Not get_opt_out_date() Like "##/##/##" Then GoTo bad_date
    
    all_reviewed = True
    
    Exit Sub
    
bad_contract:
    all_reviewed = False
    Exit Sub
bad_date:
    all_reviewed = False
    Exit Sub
End Sub

Function review_eligible_data()

    Call update_checklist(S.QC.qc_checklist, "all_files_present", 1)
    
    If MT.needs_active_list And active_tab() Is Nothing Then
        Call update_checklist(S.QC.qc_checklist, "all_files_present", -1)
    End If
    
    If MT.needs_gagg_list And utility_tab() Is Nothing Then
        Call update_checklist(S.QC.qc_checklist, "all_files_present", -1)
    End If
    
    If mapping_tab() Is Nothing Then
        Call update_checklist(S.QC.qc_checklist, "all_files_present", -1)
    End If
    
    Set ff = filter_tab()
    
    With ff
        num_rows = .UsedRange.Rows.count
        num_cols = .UsedRange.columns.count
        Set arr_range = .Cells(1, 1).Resize(num_rows, num_cols)
        arr = arr_range.value
    End With
    
    Set states = state_abbrev_dict()
    
    With F.columns
        eligible_col = .eligible_opt_out.index
        name_col = .customer_name.index
        s1 = .service_address.index
        s_city = .service_city.index
        s_state = .service_state.index
        s_zip = .service_zip.index
        m1 = .mail_address.index
        m_city = .mail_city.index
        m_state = .mail_state.index
        m_zip = .mail_zip.index
        'customer_class = .customer_class.index
        read_cycle = .read_cycle.index
    End With
    
    address_match = 0
    
    progress.start ("Checking Output Data")
    
    For j = 2 To num_rows
        If arr(j, eligible_col) = "Y" Then
            'arr(j, s1) = clean_service_address(arr(j, s1))
            If Not share_apt_number(arr(j, s1), arr(j, m1)) Then GoTo missing_apt_number
            If arr(j, s1) = arr(j, m1) Then address_match = address_match + 1
            If arr(j, s_city) = "" Then GoTo bad_service_city
            If arr(j, s_state) <> EDC.state Then GoTo bad_service_state
            arr(j, s_zip) = format(Split(arr(j, s_zip), "-")(0), "00000")
            If Not IsNumeric(arr(j, s_zip)) Then GoTo bad_service_zip
            If arr(j, m_city) = "" Then GoTo bad_mail_city
            If Not states.exists(Left(arr(j, m_state), 2)) Then GoTo bad_mail_state
            arr(j, m_zip) = format(Split(arr(j, m_zip), "-")(0), "00000")
            If Not IsNumeric(arr(j, m_zip)) Then GoTo bad_mail_zip
            If Not IsNumeric(arr(j, read_cycle)) Then GoTo bad_read_cycle
        End If
    Next
    
    If address_match / (num_rows - 1) Then
        'GoTo mail_service_mismatch
    End If
    
    arr_range.AutoFilter Field:=F.columns.eligible_opt_out.index, Criteria1:="Y"
    
    progress.finish
    
    all_reviewed = True
    
    Call update_checklist(S.QC.qc_checklist, "apt_numbers", 1)
    Call update_checklist(S.QC.qc_checklist, "valid_states", 1)
    Call update_checklist(S.QC.qc_checklist, "valid_zips", 1)
    
    If Application.CountA(home_tab().Range(S.HOME.peer_review_checklist_range)) <> 14 Then GoTo bad_peer_review
    
    review_eligible_data = True
    
    Exit Function
    
bad_peer_review:
    review_eligible_data = False
    MsgBox "Complete the peer review checklist first", vbCritical
    Exit Function
mail_service_mismatch:
    Call update_checklist(S.QC.qc_checklist, "apt_numbers", -1)
    review_eligible_data = False
    MsgBox "Missing Apt Number in row " & j, vbCritical
    Exit Function
missing_apt_number:
    Call update_checklist(S.QC.qc_checklist, "apt_numbers", -1)
    review_eligible_data = False
    MsgBox "Missing Apt Number in row " & j, vbCritical
    Exit Function
bad_service_city:
    review_eligible_data = False
    MsgBox "Bad Service City in row " & j, vbCritical
    Exit Function
bad_service_state:
    Call update_checklist(S.QC.qc_checklist, "valid_states", -1)
    review_eligible_data = False
    MsgBox "Bad Service State in row " & j, vbCritical
    Exit Function
bad_service_zip:
    Call update_checklist(S.QC.qc_checklist, "valid_zips", -1)
    review_eligible_data = False
    MsgBox "Bad Service Zip in row " & j, vbCritical
    Exit Function
bad_mail_city:
    review_eligible_data = False
    MsgBox "Bad Mail City in row " & j, vbCritical
    Exit Function
bad_mail_state:
    Call update_checklist(S.QC.qc_checklist, "valid_states", -1)
    review_eligible_data = False
    MsgBox "Bad Mail State in row " & j, vbCritical
    Exit Function
bad_mail_zip:
    Call update_checklist(S.QC.qc_checklist, "valid_zips", -1)
    review_eligible_data = False
    MsgBox "Bad Mail Zip in row " & j, vbCritical
    Exit Function
bad_read_cycle:
    review_eligible_data = False
    MsgBox "Bad Read Cycle in row " & j, vbCritical
    Exit Function

End Function

Sub reorder_tabs()
    'move sheets
    Call move_sheet(home_tab(), 1)
    Call move_sheet(filter_tab(), 2)
    Call move_sheet(LP_tab(), 3)
    Call move_sheet(mapping_tab(), 4)
    Call move_sheet(dna_tab(), 5)
End Sub

Sub move_sheet(ws, index)
    If ws Is Nothing Then Exit Sub
    ws.Move before:=Sheets(index)
End Sub

Function share_apt_number(x1, x2)
    share_apt_number = True
    If x1 = x2 Then
        share_apt_number = True
    ElseIf x1 Like x2 & "*" Then
        str_diff = Mid$(x1, Len(x2) + 2)
        If str_diff Like "APT*" Then share_apt_number = False
    ElseIf x2 Like x1 & "*" Then
        str_diff = Mid$(x2, Len(x1) + 2)
        If str_diff Like "APT*" Then share_apt_number = False
    Else
        share_apt_number = True
    End If
End Function
