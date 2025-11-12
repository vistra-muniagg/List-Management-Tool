Function keep_renewal_mapped_out(m)
    If m Like "TR*" Then
        keep_renewal_mapped_out = True
    Else
        keep_renewal_mapped_out = False
    End If
End Function

Function no_gagg_list(m)
    If m Like "*NO*" Then
        no_gagg_list = True
    Else
        no_gagg_list = False
    End If
End Function

Function has_renewal_list(m)
    If m Like "*CR*" Or m Like "*TR*" Then
        has_renewal_list = True
    Else
        has_renewal_list = False
    End If
End Function

Function previous_supplier_list(m)
    If m = "NEW" Then
        previous_supplier_list = True
    Else
        previous_supplier_list = False
    End If
End Function

Function create_opt_in_list(m)
    If m = "NEW" Or m Like "*CR*" Then
        create_opt_in_list = True
    Else
        create_opt_in_list = False
    End If
End Function

Function save_opt_in_list(m)
    If m = "NEW" Or m = "CR+SWP" Then
        save_opt_in_list = True
    Else
        save_opt_in_list = False
    End If
End Function

Function make_new_LP_upload(m)
    If m Like "*NO*" Then
        make_new_LP_upload = False
    Else
        make_new_LP_upload = True
    End If
End Function

Function check_existing_contracts(m)
    If m = "SWP" Or m Like "TR*" Then
        check_existing_contracts = True
    Else
        check_existing_contracts = False
    End If
End Function

Function add_2d_barcode(m)
    If m Like "TR*" Then
        add_2d_barcode = True
    Else
        add_2d_barcode = False
    End If
End Function

Function renewal_file_name(m)
    'If m Like "CR*" Then
    '    renewal_file_name = "CR-REN"
    'ElseIf m Like "TR*" Then
    '    renewal_file_name = "TR-REN"
    'End If
End Function

Function swp_file_name(m)
    If m Like "CR*" Or m Like "TR*" Then
        renewal_file_name = "CR-SWP"
    End If
End Function

Function overwrite_renewal_address(m)
    If m = "CR+SWP" Or m = "TR+SWP" Then
        overwrite_renewal_address = True
    Else
        overwrite_renewal_address = False
    End If
End Function

Function has_mismatches(m)
    If m = "CR+SWP" Or m = "TR+SWP" Or m = "NEW" Then
        has_mismatches = True
    Else
        has_mismatches = False
    End If
End Function
