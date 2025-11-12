Public Type ActiveColumnHeader
    header As String
    index As Long
    cell_color As CellColors
End Type

Public Type ActiveListColumns
    account_number As ActiveColumnHeader
    customer_name As ActiveColumnHeader
    sas_id As ActiveColumnHeader
    customer_class As ActiveColumnHeader
    service_address As ActiveColumnHeader
    service_city As ActiveColumnHeader
    service_state As ActiveColumnHeader
    service_zip As ActiveColumnHeader
    mail_address As ActiveColumnHeader
    mail_city As ActiveColumnHeader
    mail_state As ActiveColumnHeader
    mail_zip As ActiveColumnHeader
    read_cycle As ActiveColumnHeader
    phone As ActiveColumnHeader
    email As ActiveColumnHeader
    EMPTYCOL As ActiveColumnHeader
End Type

Public Type ActiveList
    columns As ActiveListColumns
    mismatch_columns() As ActiveColumnHeader
End Type

Sub define_active_cols()
    With A.columns
        .account_number.header = "UTILITYACCOUNTVALUE"
        .customer_name.header = "CUSTOMERNAME"
        .sas_id.header = "SUBACCOUNTSERVICEID"
        .customer_class.header = "PREMISETYPE"
        .read_cycle.header = "LDCMETERCYCLE"
        .service_address.header = "SERVICEADDRESSLINE1"
        .service_city.header = "SERVICECITY"
        .service_state.header = "SERVICESTATE"
        .service_zip.header = "SERVICEPOSTALCODE"
        .mail_address.header = "BILLINGADDRESSLINE1"
        .mail_city.header = "BILLINGCITY"
        .mail_state.header = "BILLINGSTATE"
        .mail_zip.header = "BILLINGPOSTALCODE"
        .phone.header = "PHONENUMBER"
        .email.header = "EMAIL"
        .EMPTYCOL.header = ""
        .EMPTYCOL.index = 0
    End With
End Sub

Sub define_mismatch_cols()

    ReDim A.mismatch_columns(1 To 1)
    
    With A.columns
    
        Call active_arr_append(A.mismatch_columns, .account_number)
        Call active_arr_append(A.mismatch_columns, .sas_id)
        Call active_arr_append(A.mismatch_columns, .customer_class)
        Call active_arr_append(A.mismatch_columns, .read_cycle)
        Call active_arr_append(A.mismatch_columns, .customer_name)
        Call active_arr_append(A.mismatch_columns, .service_address)
        Call active_arr_append(A.mismatch_columns, .service_city)
        Call active_arr_append(A.mismatch_columns, .service_state)
        Call active_arr_append(A.mismatch_columns, .service_zip)
        Call active_arr_append(A.mismatch_columns, .mail_address)
        Call active_arr_append(A.mismatch_columns, .mail_city)
        Call active_arr_append(A.mismatch_columns, .mail_state)
        Call active_arr_append(A.mismatch_columns, .mail_zip)
        Call active_arr_append(A.mismatch_columns, .phone)
        Call active_arr_append(A.mismatch_columns, .email)
    
    End With
    
End Sub
