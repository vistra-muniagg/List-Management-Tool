''''''''''''''''''''''''
''use $ to get strings''
''''''''''''''''''''''''

Function clean_name(n)
    n = Application.Trim(n)
    split_name = name_parts(name_reverse(n))
    clean_name = Trim(split_name(0) & " " & split_name(2))
    clean_name = UCase(clean_name)
End Function

Function name_reverse(n)
    comma = InStr(n, ",")
    If n Like "*JR" Or n Like "*SR" Or n Like "* II" Or n Like "* III" Then
        name_reverse = Replace(n, ",", "")
    ElseIf n Like "*LLC" Or n Like "*INC" Or n Like "* LP" Or n Like "*LPA" Or n Like "* DDS" Then
        name_reverse = Replace(n, ",", "")
    ElseIf InStrRev(n, " ") > 0 Then
        If comma > 0 Then
            n1 = Right$(n, Len(n) - comma)
            n2 = Left$(n, comma - 1)
            name_reverse = n1 & " " & n2
        Else
            name_reverse = n
        End If
    ElseIf comma > 0 Then
        n1 = Right$(n, Len(n) - comma)
        n2 = Left$(n, comma - 1)
        name_reverse = n1 & " " & n2
    Else
        name_reverse = n
    End If
    name_reverse = Application.Trim$(name_reverse)
    
End Function

Function name_parts(namestring) As Variant
    
    Dim arr() As String
    ReDim arr(0 To 2)
    
    On Error Resume Next

    namestring = Application.Trim(UCase(namestring))
    n = namestring
    
    If n = "" Then Exit Function
    
    'remove comma and period
    n = Replace(n, ",", "")
    n = Replace(n, ".", "")
    
    'remove and store prefix
    For Each p In Array("MR", "MRS", "MS", "DR")
        If n Like p & " *" Then
            pf = p & " "
            n = Right$(n, Len(n) - Len(p) - 1)
            Exit For
        End If
    Next
    
    'remove and store suffix
    For Each suffix In Array("JR", "SR", "III", "II")
        If n Like suffix & " *" Then
            n = Mid$(n, Len(suffix) + 2) & " " & suffix
        End If
        If n Like "* " & suffix Then
            sf = suffix
            n = Left$(n, Len(n) - Len(suffix) - 1)
            Exit For
        End If
    Next
    
    k = Split(n, " ")
    If UBound(k) = 0 And sf = "" Then
        'only populate last name if name is one word
        arr(0) = ""
        arr(1) = ""
        arr(2) = UCase$(k(0))
        GoTo cleaned_name
    ElseIf UBound(k) = 0 And sf <> "" Then
        'some DUKE names look like "LAST JR"
        arr(0) = UCase$(k(0))
        arr(1) = ""
        arr(2) = UCase$(Trim(sf))
        GoTo cleaned_name
    End If
    
    j = 1
    
    'do not include prefix
    pf = ""
    
    'first name
    arr(0) = pf & k(0)
    If k(1) = "&" Then
        'arr(0) = k(0) & k(1) & k(2)
        j = 2
    End If
    first_name = UCase$(arr(0))
    
    'middle name
    If Len(k(1)) = 1 Or k(1) Like "[A-Z]." Then arr(1) = k(1)
    arr(1) = UCase$(arr(1))
    
    'last name
    arr(2) = Trim$(Right$(n, Len(n) - Len(arr(0)) - Len(arr(1)) - j) & " " & sf)
    last_name = UCase$(arr(2))
    
    arr(1) = Replace(arr(1), ".", "")
    
cleaned_name:
    
    name_parts = arr
    
End Function

Function no_suffix(n)
    n = UCase(n)
    If Not n Like "* *" Then
        no_suffix = n
        Exit Function
    End If
    suffixes = Array("JR", "SR", "II", "III")
    For Each S In suffixes
        If n Like "* " & S Then
            no_suffix = Trim(Left$(n, InStrRev(n, " ") - 1))
            If no_suffix Like "*," Then
                no_suffix = Left$(no_suffix, Len(no_suffix) - 1)
            End If
            Exit Function
        End If
    Next
    no_suffix = n
End Function

Function service_address(source_row, gagg_data, service_cols As Variant, comed_special_case As Boolean)
    For j = 0 To UBound(service_cols)
        x = gagg_data(source_row, service_cols(j))
        If comed_special_case Then
            x = Left$(x, InStrRev(x, "  "))
        End If
        If j > 1 And Not IsNumeric(Left$(x, 1)) Then x = ""
        service_address = service_address & " " & x
    Next
    service_address = Replace$(service_address, "-", " ")
    service_address = Application.Trim(UCase(service_address))
End Function

Function service_city(source_row, gagg_data, service_city_col, is_comed As Boolean)
    x = gagg_data(source_row, service_city_col)
    If is_comed Then
        x = Mid$(x, InStrRev(x, "  "))
    End If
    service_city = Application.Trim(UCase(x))
End Function

Function service_state(source_row, gagg_data, service_state_col, is_comed As Boolean)
    x = gagg_data(source_row, service_state_col)
    If is_comed Then
        x = Mid$(x, InStrRev(x, "  "))
    End If
    service_state = Application.Trim(UCase(x))
End Function

Function service_zip(source_row, gagg_data, service_zip_col, is_comed As Boolean)
    x = CStr(gagg_data(source_row, service_zip_col))
    service_zip = format(Left$(x, 5), "00000")
    'service_zip = Application.Trim(UCase(service_zip))
End Function

Function mail_address(source_row, gagg_data, mail_cols As Variant)
    For j = 0 To UBound(mail_cols)
        x = gagg_data(source_row, mail_cols(j))
        If EDC.ruleset_name = "DUKE" Then
            If Not IsNumeric(Left$(x, 1)) Then
                If UCase$(x) Like "*MISC *" Then
                    x = ""
                ElseIf x Like "PO BOX*" Then
                    mail_address = x
                    Exit Function
                Else
                    x = ""
                End If
            End If
        End If
        If x <> "" Then
            If Left$(mail_address, 10) & "*" Like Left$(x, 1) & "*" Then x = ""
        End If
        mail_address = mail_address & " " & x
    Next
    mail_address = Application.Trim(UCase(mail_address))
End Function

Function mail_city(source_row, gagg_data, mail_city_col)
    mail_city = gagg_data(source_row, mail_city_col)
    mail_city = Application.Trim(UCase(mail_city))
    'If mail_city Like "SPOKANE*" Then mail_city = "SPOKANE"
End Function

Function mail_state(source_row, gagg_data, mail_state_col)
    mail_state = gagg_data(source_row, mail_state_col)
    mail_state = Application.Trim(UCase(mail_state))
    'If mail_state Like "SPOKANE*" Then mail_state = "WA"
End Function

Function mail_zip(source_row, gagg_data, mail_zip_col)
    mail_zip = gagg_data(source_row, mail_zip_col)
    mail_zip = format(Left$(mail_zip, 5), "00000")
    'mail_zip = Application.Trim(UCase(mail_zip))
End Function

Sub add_apt_numbers()
    
    progress.start ("Checking APT Numbers")
    
    Set ff = filter_tab()
    num_rows = Application.CountA(ff.columns(1))
    
    With F.columns
        service_arr = filter_col(.service_address)
        mail_arr = filter_col(.mail_address)
    End With
    
    'get apt number from service address
    For j = 2 To num_rows
        If service_arr(j, 1) = mail_arr(j, 1) Then GoTo next_row_1
        If service_arr(j, 1) Like mail_arr(j, 1) & "*" Then
            mail_arr(j, 1) = service_arr(j, 1)
        End If
next_row_1:
        progress.activity (j)
    Next
    
    'get apt number from mail address
    For j = 2 To num_rows
        If service_arr(j, 1) = mail_arr(j, 1) Then GoTo next_row_2
        If mail_arr(j, 1) Like service_arr(j, 1) & "*" Then
            service_arr(j, 1) = mail_arr(j, 1)
        End If
next_row_2:
        progress.activity (j)
    Next
    
    ff.UsedRange.columns(F.columns.service_address.index).value = service_arr
    ff.UsedRange.columns(F.columns.mail_address.index).value = mail_arr
    
    progress.complete
    
End Sub

Sub replace_empty_mail()

    progress.start ("Fixing Empty Mail Addresses")

    Set ff = filter_tab()
    num_rows = Application.CountA(ff.columns(1))
    With F.columns
        s_1 = filter_col(.service_address)
        s_2 = filter_col(.service_city)
        s_3 = filter_col(.service_state)
        s_4 = filter_col(.service_zip)
        m_1 = filter_col(.mail_address)
        m_2 = filter_col(.mail_city)
        m_3 = filter_col(.mail_state)
        m_4 = filter_col(.mail_zip)
    End With
    For i = 2 To num_rows
        If m_1(i, 1) = "" Then
            m_1(i, 1) = s_1(i, 1)
            m_2(i, 1) = s_2(i, 1)
            m_3(i, 1) = s_3(i, 1)
            m_4(i, 1) = s_4(i, 1)
        End If
        progress.activity (i)
    Next
    ff.UsedRange.columns(F.columns.mail_address.index).value = m_1
    ff.UsedRange.columns(F.columns.mail_city.index).value = m_2
    ff.UsedRange.columns(F.columns.mail_state.index).value = m_3
    ff.UsedRange.columns(F.columns.mail_zip.index).value = m_4
    
    progress.finish
    
End Sub

Function clean_mail_address(str)
    str = clean_mail_address_1(str)
    'str = clean_mail_address_2(str)
    str = clean_mail_address_3(str)
    clean_mail_address = str
End Function

Function clean_mail_address_1(str)
    str = Replace$(str, ",", "")
    str = Replace$(str, "-", " ")
    str = Replace$(str, ".", "")
    str = Replace$(str, ";", "")
    str = Replace$(str, "P O BOX", "PO BOX")
    str = Application.Trim$(str)
    j = InStr(1, str, "PO BOX")
    If j > 0 Then
        x = Mid$(str, j)
        If j - 1 > 0 Then x = x & " " & Left(str, j - 2)
        k = Split(x, " ")
        If UBound(k) >= 3 Then x = k(0) & " " & k(1) & " " & k(2)
        clean_mail_address_1 = x
        Exit Function
    End If
    For j = 1 To Len(str)
        If Mid$(str, j, 1) Like "[0-9]" Then
            clean_mail_address_1 = Mid$(str, j)
            Exit Function
        End If
        progress.activity (j)
    Next
    str = Replace$(str, ",", "")
    str = Replace$(str, ".", "")
    str = Replace$(str, ":", "")
    clean_mail_address_1 = str
End Function

Function clean_mail_address_2(str)
    If str = "" Then
        clean_mail_address_2 = ""
        Exit Function
    End If
    suffixes = Array(" ST", " RD", " DR", " LN", " AVE", " BLVD", " HWY", _
                     " PKWY", " CT", " CIR", " PL", " TER", " WAY", _
                     " LOOP", " TRCE", " CTR")
                     
    directions = Array(" N", " S", " E", "W", " NE", " NW", " SE", " SW", "")
    j = 0
    For Each suffix In suffixes
        For Each compass In directions
            k = InStrRev(UCase$(str & " "), suffix & compass & " ")
            If k > j Then
                j = k + Len(suffix)
                Exit For
            End If
        Next
    Next
    If j > 0 Then
        clean_mail_address_2 = Left$(str, j) & mail_suffix(str, j + 1)
    Else
        clean_mail_address_2 = str
    End If
    clean_mail_address_2 = Application.Trim(clean_mail_address_2)
End Function

Function clean_mail_address_3(str)
    If str = "" Then
        clean_mail_address_3 = ""
        Exit Function
    End If
    If EDC.ruleset_name = "DUKE" Then
        str = Replace$(str, " MISC: ", "")
        str = Replace$(str, " MISC ", "")
        str = Application.Trim(str)
    End If
    If EDC.ruleset_name = "FE" Then
        x = InStrRev(str, "BLK")
        If x > 0 Then
            str = Left$(str, x)
            str = Application.Trim(str)
        End If
    End If
    
    While Left$(str, 1) = "0" And Len(str) > 1
        str = Mid$(str, 2)
    Wend
    
    clean_mail_address_3 = str
End Function

Function mail_suffix(str, n)
    suffix = UCase$(Trim$(Mid$(str, n)))
    If suffix Like "APT *" Or suffix Like "UNIT *" Or suffix Like "STE *" Or suffix Like "SUITE *" Then
        mail_suffix = " " & Trim$(Mid$(str, n))
    Else
        mail_suffix = ""
    End If
End Function

Function split_city_state_zip(str) As Variant
    If str = "" Then
        split_city_state_zip = Array("-", "-", "-", "-")
        Exit Function
    End If
    Dim arr As Variant
    ReDim arr(1 To 3)
    'city st zip
    k = Split(str, " ")
    arr(3) = k(UBound(k))
    arr(2) = Application.Trim(k(UBound(k) - 1))
    x = InStrRev(str, arr(2))
    arr(1) = Application.Trim(Left$(str, x - 2))
    arr(1) = Replace$(arr(1), ",", "")
    arr(1) = Replace$(arr(1), "-", " ")
    split_city_state_zip = arr
End Function

Function split_city_state(str) As Variant
    If str = "" Then
        split_city_state = Array("-", "-", "-")
        Exit Function
    End If
    Dim arr As Variant
    ReDim arr(1 To 2)
    'city st (zip?)
    If str Like "* #####" Then str = Left$(str, Len(str) - 6)
    If EDC.ruleset_name = "DUKE" Then str = abbreviate_states(str)
    k = Split(str, " ")
    If EDC.ruleset_name = "COM" Then k = Split(str, ",")
    If UBound(k) = 0 Then
        split_city_state = Array()
        Exit Function
    End If
    arr(2) = Application.Trim(k(UBound(k)))
    x = InStrRev(str, arr(2))
    arr(1) = Application.Trim(Replace(Left$(str, x - 2), "-", " "))
    split_city_state = arr
End Function

Sub format_address_data()
    replace_empty_mail
    clean_address_data
    add_apt_numbers
    clean_mail_addresses
    Call update_checklist(S.QC.qc_checklist, "account_number_format", 1)
End Sub

Sub clean_address_data()

    If Not MT.needs_gagg_list Then Exit Sub
    
    Set ff = filter_tab()
    mismatch_arr = filter_col(F.columns.mismatch)
    progress.start ("Cleaning Service Addresses")
    If EDC.service_city = EDC.service_state Then
        If EDC.service_state = EDC.service_zip Then
            arr1 = filter_col(F.columns.service_city)
            arr2 = filter_col(F.columns.service_state)
            arr3 = filter_col(F.columns.service_zip)
            num_rows = UBound(arr1, 1)
            For j = 2 To num_rows
                If mismatch_arr(j, 1) <> "Y" Then
                    k = split_city_state_zip(arr1(j, 1))
                    arr1(j, 1) = k(1)
                    arr2(j, 1) = k(2)
                    arr3(j, 1) = k(3)
                Else
                    arr1(j, 1) = arr1(j, 1)
                    arr2(j, 1) = arr2(j, 1)
                    arr3(j, 1) = arr3(j, 1)
                End If
                progress.activity (j)
            Next
            ff.UsedRange.columns(F.columns.service_city.index).value = arr1
            ff.UsedRange.columns(F.columns.service_state.index).value = arr2
            ff.UsedRange.columns(F.columns.service_zip.index).value = arr3
        Else
            arr1 = filter_col(F.columns.service_city)
            arr2 = filter_col(F.columns.service_state)
            num_rows = UBound(arr1, 1)
            For j = 2 To num_rows
                If mismatch_arr(j, 1) <> "Y" Then
                    k = split_city_state(arr1(j, 1))
                    arr1(j, 1) = k(1)
                    arr2(j, 1) = k(2)
                Else
                    arr1(j, 1) = arr1(j, 1)
                    arr2(j, 1) = arr2(j, 1)
                End If
                progress.activity (j)
            Next
            ff.UsedRange.columns(F.columns.service_city.index).value = arr1
            ff.UsedRange.columns(F.columns.service_state.index).value = arr2
        End If
    End If
    
    progress.complete
    
    If EDC.ruleset_name = "FE" Then GoTo FE_replace_mail
    
    progress.start ("Cleaning Mail Addresses")
    If EDC.mail_city = EDC.mail_state Then
        If EDC.mail_state = EDC.mail_zip Then
            arr1 = filter_col(F.columns.mail_city)
            arr2 = filter_col(F.columns.mail_state)
            arr3 = filter_col(F.columns.mail_zip)
            num_rows = UBound(arr1, 1)
            For j = 2 To num_rows
                If mismatch_arr(j, 1) <> "Y" Then
                    If arr1(j, 1) <> "" Then
                        k = split_city_state_zip(arr1(j, 1))
                        arr1(j, 1) = k(1)
                        arr2(j, 1) = k(2)
                        arr3(j, 1) = k(3)
                    End If
                Else
                    arr1(j, 1) = arr1(j, 1)
                    arr2(j, 1) = arr2(j, 1)
                    arr3(j, 1) = arr3(j, 1)
                End If
                progress.activity (j)
            Next
            ff.UsedRange.columns(F.columns.mail_city.index).value = arr1
            ff.UsedRange.columns(F.columns.mail_state.index).value = arr2
            ff.UsedRange.columns(F.columns.mail_zip.index).value = arr3
        Else
            arr1 = filter_col(F.columns.mail_city)
            arr2 = filter_col(F.columns.mail_state)
            num_rows = UBound(arr1, 1)
            For j = 2 To num_rows
                If mismatch_arr(j, 1) <> "Y" Then
                    k = split_city_state(arr1(j, 1))
                    If UBound(k) = -1 Then GoTo next_row
                    arr1(j, 1) = k(1)
                    arr2(j, 1) = k(2)
                Else
                    arr1(j, 1) = arr1(j, 1)
                    arr2(j, 1) = arr2(j, 1)
                End If
                progress.activity (j)
next_row:
            Next
            ff.UsedRange.columns(F.columns.mail_city.index).value = arr1
            ff.UsedRange.columns(F.columns.mail_state.index).value = arr2
        End If
    End If
    
    progress.complete
    
    ff.columns.AutoFit
    
    Exit Sub
    
FE_replace_mail:
    FE_address_replace
    
    ff.columns.AutoFit
    
End Sub

Sub clean_mail_addresses()
    progress.start "Cleaning Mail Addresses"
    arr1 = filter_col(F.columns.mail_address)
    arr4 = filter_col(F.columns.mail_zip)
    num_rows = UBound(arr1, 1)
    For j = 2 To num_rows
        arr1(j, 1) = clean_mail_address(arr1(j, 1))
        arr4(j, 1) = format(Left$(arr4(j, 1), 5), "00000")
    Next
    filter_tab().UsedRange.columns(F.columns.mail_address.index).value = arr1
    filter_tab().UsedRange.columns(F.columns.mail_zip.index).value = arr4
    filter_tab().columns.AutoFit
    progress.complete
End Sub

Function abbreviate_states(str)
    Set dict = state_dict()
    For Each key In dict.keys()
        'k = InStrRev(str, key)
        If str Like "* " & key Then
            str = Left$(str, Len(str) - Len(key)) & dict(key)
            Exit For
        End If
    Next
    abbreviate_states = str
End Function

Function state_dict() As Object
    
    Dim dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Z–A order so WV is before VA
    dict.Add "WYOMING", "WY"
    dict.Add "WISCONSIN", "WI"
    dict.Add "WEST VIRGINIA", "WV"
    dict.Add "WASHINGTON", "WA"
    dict.Add "VIRGINIA", "VA"
    dict.Add "VERMONT", "VT"
    dict.Add "UTAH", "UT"
    dict.Add "TEXAS", "TX"
    dict.Add "TENNESSEE", "TN"
    dict.Add "SOUTH DAKOTA", "SD"
    dict.Add "SOUTH CAROLINA", "SC"
    dict.Add "RHODE ISLAND", "RI"
    dict.Add "PENNSYLVANIA", "PA"
    dict.Add "OREGON", "OR"
    dict.Add "OKLAHOMA", "OK"
    dict.Add "OHIO", "OH"
    dict.Add "NORTH DAKOTA", "ND"
    dict.Add "NORTH CAROLINA", "NC"
    dict.Add "NEW YORK", "NY"
    dict.Add "NEW MEXICO", "NM"
    dict.Add "NEW JERSEY", "NJ"
    dict.Add "NEW HAMPSHIRE", "NH"
    dict.Add "NEVADA", "NV"
    dict.Add "NEBRASKA", "NE"
    dict.Add "MONTANA", "MT"
    dict.Add "MISSOURI", "MO"
    dict.Add "MISSISSIPPI", "MS"
    dict.Add "MINNESOTA", "MN"
    dict.Add "MICHIGAN", "MI"
    dict.Add "MASSACHUSETTS", "MA"
    dict.Add "MARYLAND", "MD"
    dict.Add "MAINE", "ME"
    dict.Add "LOUISIANA", "LA"
    dict.Add "KENTUCKY", "KY"
    dict.Add "KANSAS", "KS"
    dict.Add "IOWA", "IA"
    dict.Add "INDIANA", "IN"
    dict.Add "ILLINOIS", "IL"
    dict.Add "IDAHO", "ID"
    dict.Add "HAWAII", "HI"
    dict.Add "GEORGIA", "GA"
    dict.Add "FLORIDA", "FL"
    dict.Add "DELAWARE", "DE"
    dict.Add "CONNECTICUT", "CT"
    dict.Add "COLORADO", "CO"
    dict.Add "CALIFORNIA", "CA"
    dict.Add "ARKANSAS", "AR"
    dict.Add "ARIZONA", "AZ"
    dict.Add "ALASKA", "AK"
    dict.Add "ALABAMA", "AL"
    
    Set state_dict = dict
    
End Function

Function state_abbrev_dict() As Object
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Z–A order so WV is before VA
    dict.Add "WY", "WYOMING"
    dict.Add "WI", "WISCONSIN"
    dict.Add "WV", "WEST VIRGINIA"
    dict.Add "WA", "WASHINGTON"
    dict.Add "VA", "VIRGINIA"
    dict.Add "VT", "VERMONT"
    dict.Add "UT", "UTAH"
    dict.Add "TX", "TEXAS"
    dict.Add "TN", "TENNESSEE"
    dict.Add "SD", "SOUTH DAKOTA"
    dict.Add "SC", "SOUTH CAROLINA"
    dict.Add "RI", "RHODE ISLAND"
    dict.Add "PA", "PENNSYLVANIA"
    dict.Add "OR", "OREGON"
    dict.Add "OK", "OKLAHOMA"
    dict.Add "OH", "OHIO"
    dict.Add "ND", "NORTH DAKOTA"
    dict.Add "NC", "NORTH CAROLINA"
    dict.Add "NY", "NEW YORK"
    dict.Add "NM", "NEW MEXICO"
    dict.Add "NJ", "NEW JERSEY"
    dict.Add "NH", "NEW HAMPSHIRE"
    dict.Add "NV", "NEVADA"
    dict.Add "NE", "NEBRASKA"
    dict.Add "MT", "MONTANA"
    dict.Add "MO", "MISSOURI"
    dict.Add "MS", "MISSISSIPPI"
    dict.Add "MN", "MINNESOTA"
    dict.Add "MI", "MICHIGAN"
    dict.Add "MA", "MASSACHUSETTS"
    dict.Add "MD", "MARYLAND"
    dict.Add "ME", "MAINE"
    dict.Add "LA", "LOUISIANA"
    dict.Add "KY", "KENTUCKY"
    dict.Add "KS", "KANSAS"
    dict.Add "IA", "IOWA"
    dict.Add "IN", "INDIANA"
    dict.Add "IL", "ILLINOIS"
    dict.Add "ID", "IDAHO"
    dict.Add "HI", "HAWAII"
    dict.Add "GA", "GEORGIA"
    dict.Add "FL", "FLORIDA"
    dict.Add "DE", "DELAWARE"
    dict.Add "CT", "CONNECTICUT"
    dict.Add "CO", "COLORADO"
    dict.Add "CA", "CALIFORNIA"
    dict.Add "AR", "ARKANSAS"
    dict.Add "AZ", "ARIZONA"
    dict.Add "AK", "ALASKA"
    dict.Add "AL", "ALABAMA"
    
    Set state_abbrev_dict = dict
    
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function AES_service_city(str)
    s1 = Application.Trim(str)
    s2 = InStrRev(s1, "OH")
    If s2 <= 0 Then
        AES_service_city = "-"
    Else
        AES_service_city = Left$(s1, s2 - 2)
    End If
End Function

Function AES_service_state(str)
    If InStrRev(1, str, "OH") > 0 Then
        AES_service_state = "OH"
    Else
        AES_service_state = "-"
    End If
End Function

Function AEP_mail_address(m1, m2)
    m3 = Application.Trim(m1 & " " & m2)
    AEP_mail_address = m3
End Function

Function AEP_service_address(s1, s2)
    s3 = Application.Trim(s1 & " " & s2)
    For Each str In Array("TEMP", "LIGHTS", "POND", "TS")
        If s3 Like "* UNIT " & str Then
            s3 = Left$(s3, Len(s3) - 6 - Len(str))
        End If
    Next
    If s3 Like "* HSE" Then
        s3 = Left$(s3, Len(s3) - 4)
    End If
    AEP_service_address = Application.Trim(s3)
End Function

Function AES_mail_address(m1, m2)
    AES_mail_address = Application.Trim(m1 & " " & m2)
End Function

Function AES_service_address(s1, s2)
    AES_service_address (s1 & " " & s2)
End Function

Function AES_mail_city(str)
    m1 = InStrRev(str, ",")
    AES_mail_city = Left$(str, m1 - 1 - 2 - 1)
End Function

Function AES_mail_state(str)
    m1 = InStrRev(str, ",")
    AES_mail_state = Mid$(str, m1 - 1 - 1, 2)
End Function

Function AES_mail_zip(m1)
    AES_mail_zip = Left$(m1, 5)
End Function

''''''''''''''''''''''''''''''''
'''rewrite code below'''''''''''
''''''''''''''''''''''''''''''''

Function AM_mail_address(m1, m2, m3)
    
    m1 = Application.Trim(m1)
    m2 = Application.Trim(m2)
    m3 = Application.Trim(m3)
    
    If m1 = m2 Then m2 = ""
    
    If m2 = m3 Then m = m1
    If m1 = m2 Then
        a3 = InStr(m3, ",")
        If a3 <> 0 Then
            m3 = Left(m3, a3 - 1)
        End If
        m = m3
    End If
    
    If m = "" Then
        If m1 Like "PO BOX*" Then m = m1
        If m2 Like "PO BOX*" Then m = m2
        If m3 Like "PO BOX*" Then m = m3
    End If
    
    If m = "" Then
        If m3 Like "#* [A-Z]*" Then
            m = m3
        ElseIf m2 Like "#* [A-Z]*" Then
            m = m2
        Else
            m = m1
        End If
        
    End If
    
    If m = "" Then
        If m1 Like "[A-Z]*" Then m1 = ""
        If m2 Like "[A-Z]*" Then m2 = ""
        If m3 Like "[A-Z]*" Then m3 = ""
        a1 = InStr(m3, ",")
        If A <> 0 Then
            m3 = Left(m3, A - 1)
        End If
    End If
    
    If m = "" Then
        A = InStr(1, m1, ",")
        If A <> 0 And Not m1 Like "[A-Z]*" Then
            m1 = Left(m1, A - 1)
            m2 = ""
            m3 = ""
        End If
    End If
    
    If m = "" Then
        For Each A In Array("APT", "UNIT", "STE")
            If m1 Like A & " *" Then
                suffix = m1
                m1 = ""
                Exit For
            End If
        Next
        m = m1 & " " & m2 & " " & m3 & " " & suffix
    End If
    
    AM_mail_address = Application.Trim(m)
    
End Function

Function AM_service_address(s1, s2, s3)

    s1 = Application.Trim(s1)
    s2 = Application.Trim(s2)
    s3 = Application.Trim(s3)
    
    A = InStr(s1, ",")
    If A <> 0 And Not s1 Like "[A-Z]*" Then
        s1 = Left(s1, A - 1)
        s2 = ""
        s3 = ""
    End If
    
    If s1 Like "* UNIT [A-Z]*" Then
        b = InStrRev(s1, " UNIT ")
        s1 = Left(s1, b - 1)
    End If
    
    If s2 Like "*POLE #*" Or s2 Like "*CATV*" Then
        s2 = ""
        s3 = ""
    End If
    
    A = InStr(1, s1, "CATV") - 2
    If A > 0 Then
        s1 = Left(s1, A)
        s2 = ""
        s3 = ""
    End If
    
    S = s1 & " " & s2
    
    AM_service_address = Application.Trim(S)
    
End Function

Function COMED_service_address(s1)
    A = InStrRev(s1, "  ")
    's1 = Application.Trim(s1)
    If A <> 0 Then
        S = Application.Trim(Left(s1, A))
    Else
        S = s1
    End If
    b = InStrRev(S, ",")
    If b <> 0 And b > 10 Then
        S = Left(S, b - 1)
    End If
    S = Replace(S, " BD", "")
    If S Like "*# # *? ?*" Then
        S = Mid(S, InStr(1, S, " ") + 1)
    End If
    If S Like "HSE *" Or S Like "BLDG *" Or S Like "LIGHTING *" Then
        S = Mid(S, InStr(1, S, " ") + 1)
    End If
    If S Like "UNIT *" Or S Like "APT *" Then
        s2 = ""
        k = Split(S, " ")
        For j = 0 To 1
            s2 = Application.Trim(s2 & " " & k(j))
        Next
        S = Application.Trim(Replace(S, s2, "") & " " & s2)
    End If
    If S Like "# #*" Then
        S = Mid(S, 2)
    End If
    While Not IsNumeric(Left(m, 1)) And Len(m) > 0
        m = Mid(m, 2)
    Wend
    COMED_service_address = Replace(S, ",", "")
End Function

Function COMED_mail_address(m1)
    m1 = Application.Trim(m1)
    A = InStrRev(m1, "  ")
    's1 = Application.Trim(m1)
    If A <> 0 Then
        m = Application.Trim(Left(m1, A))
    Else
        m = m1
    End If
    b = InStrRev(m, ",")
    If b <> 0 Then
        m = Left(m, b - 1)
    End If
    m = Replace(m, " BD", "")
    If m Like "UNIT *" Or m Like "APT *" Or m Like "SUITE *" Or m Like "STE *" Then
        m2 = ""
        For k = 0 To 1
            m2 = Application.Trim(m2 & " " & Split(m, " ")(k))
        Next
        m = Application.Trim(Replace(m, m2, "") & " " & m2)
        If m Like "UNIT *" Or m Like "APT *" Or m Like "SUITE *" Or m Like "STE *" Then
            COMED_mail_address = m
            Exit Function
        End If
    End If
    If m Like "*UNIT [A-Z][A-Z]*" Then
        C = InStrRev(m, " UNIT")
        m = Left(m, C - 1)
    End If
    If m Like "# #*" Then m = Mid(m, 2)
    If m Like "[A-Z]* [A-Z]* #* *" Then
        While Not m Like " #*"
            m = Mid(m, 2)
        Wend
        m = Mid(m, 2)
    End If
    m = Replace(m, "*", "")
    COMED_mail_address = m
End Function

Function COMED_mail_city(m1)
    If m1 Like "*  I*" Then
        m = Left(m1, Len(m1) - 1)
    Else
        m = m1
    End If
    COMED_mail_city = Application.Trim(m)
End Function

Function COMED_service_city(S)
    S = Trim(S)
    If S = "" Then
        COMED_service_city = ""
        Exit Function
    End If
    k1 = InStrRev(S, "  ")
    k2 = InStrRev(S, ",")
    If k1 = 0 Then
        k1 = InStrRev(Left(S, k2 - 1), ",")
        k = Replace(Mid(S, k1, k2 - k1), ", ", "")
        COMED_service_city = Mid(k, InStr(1, k, " ") + 1)
    Else
        COMED_service_city = Mid(S, k1, k2 - k1)
    End If
End Function

Function COMED_mail_state(m1)
    A = InStrRev(m1, ",")
    If A <> 0 Then
        COMED_mail_state = Mid(m1, A + 1, 2)
    Else
        COMED_mail_state = Application.Trim(m1)
    End If
End Function

Function COMED_service_state(S)
    COMED_service_state = "IL"
End Function

Function COMED_service_zip(S)
    S = Application.Trim(S)
    k = InStrRev(S, " ")
    COMED_service_zip = Mid(S, k + 1)
End Function

Function DUKE_mail_address(m1, m2)
    
    'remove BLDG,FL,LOT,MISC,STORE
    
    m1 = Application.Trim(UCase(m1))
    m2 = Application.Trim(UCase(m2))
    
    m = m1
    
    If m1 = m2 Then m2 = ""
    
    If m1 Like "[#]*" Then m1 = ""
    
    If m1 Like "APT*" Then m1 = ""
    
    If m1 Like "MISC:*" Then m1 = ""
    
    'For Each s In Array("MISC:", "BLDG:", "STORE:")
    '    If m1 Like "*" & s & "*" Then
    '        m1 = Replace(m1, s, "")
    '        m1 = Application.Trim(Replace(m1, ",", " "))
    '        m3 = m3 & " " & m1
    '    End If
    'Next
    
    For Each S In Array("MISC:", "BLDG:", "STORE:")
        If m2 Like "*" & S & "*" Then
            m2 = Replace(m1, S, "")
            m2 = Application.Trim(Replace(m1, ",", " "))
            m3 = m3 & " " & m1
        End If
    Next
    
    If m1 Like "# #*" Then m1 = ""
    
    If m2 Like "# #*" Then m2 = Mid(m2, InStr(1, m2, " ") + 1)
    
    If Not m1 Like "#* *" Then m1 = ""
    
    If m1 Like "#* *" And m2 Like "#* *" Then m1 = ""
    
    If Not m1 Like "*#*" And Not m2 Like "*#*" Then
        m1 = ""
        m2 = ""
    End If
    
    If m1 <> m2 Then m = m1 & " " & m2
    
    If m1 Like "PO BOX *" Then
        m = m1
    ElseIf m2 Like "PO BOX *" Then
        m = m2
    End If
    
    'm = Replace(m, ":", "")
    'm = Replace(m, "# ", "#")
    'm = Replace(m, ",", "")
    'm = Replace(m, "-", " ")
    'm = Replace(m, ".", "")
    m = Trim(m)
    If m = "X, X, X" Or m = "X X, X" Then m = ""
    DUKE_mail_address = m
    
End Function

Function DUKE_mail_city(S)
    
    S = Replace(S, ".", ", ")
    S = Application.Trim(UCase(S))
    S = trimjunk(S)
    
    If S = "" Then
        DUKE_mail_city = ""
        Exit Function
    End If
    
    S = Replace(S, "WYOMING", "WY")
    S = Replace(S, "WISCONSIN", "WI")
    S = Replace(S, "WEST VIRGINIA", "WV")
    S = Replace(S, "WASHINGTON", "WA")
    S = Replace(S, "VIRGINIA", "VA")
    S = Replace(S, "VERMONT", "VT")
    S = Replace(S, "UTAH", "UT")
    S = Replace(S, "TEXAS", "TX")
    S = Replace(S, "TENNESSEE", "TN")
    S = Replace(S, "SOUTH DAKOTA", "SD")
    S = Replace(S, "SOUTH CAROLINA", "SC")
    S = Replace(S, "RHODE ISLAND", "RI")
    S = Replace(S, "PENNSYLVANIA", "PA")
    S = Replace(S, "OREGON", "OR")
    S = Replace(S, "OKLAHOMA", "OK")
    S = Replace(S, "OHIO", "OH")
    S = Replace(S, "NORTH DAKOTA", "ND")
    S = Replace(S, "NORTH CAROLINA", "NC")
    S = Replace(S, "NEW YORK", "NY")
    S = Replace(S, "NEW MEXICO", "NM")
    S = Replace(S, "NEW JERSEY", "NJ")
    S = Replace(S, "NEW HAMPSHIRE", "NH")
    S = Replace(S, "NEVADA", "NV")
    S = Replace(S, "NEBRASKA", "NE")
    S = Replace(S, "MONTANA", "MT")
    S = Replace(S, "MISSOURI", "MO")
    S = Replace(S, "MISSISSIPPI", "MS")
    S = Replace(S, "MINNESOTA", "MN")
    S = Replace(S, "MICHIGAN", "MI")
    S = Replace(S, "MASSACHUSETTS", "MA")
    S = Replace(S, "MARYLAND", "MD")
    S = Replace(S, "MAINE", "ME")
    S = Replace(S, "LOUISIANA", "LA")
    S = Replace(S, "KENTUCKY", "KY")
    S = Replace(S, "KANSAS", "KS")
    S = Replace(S, "IOWA", "IA")
    S = Replace(S, "INDIANA", "IN")
    S = Replace(S, "ILLINOIS", "IL")
    S = Replace(S, "IDAHO", "ID")
    S = Replace(S, "HAWAII", "HI")
    S = Replace(S, "GEORGIA", "GA")
    S = Replace(S, "FLORIDA", "FL")
    S = Replace(S, "DISTRICT OF COLUMBIA", "DC")
    S = Replace(S, "DELAWARE", "DE")
    S = Replace(S, "CONNECTICUT", "CT")
    S = Replace(S, "COLORADO", "CO")
    S = Replace(S, "CALIFORNIA", "CA")
    S = Replace(S, "ARKANSAS", "AR")
    S = Replace(S, "ARIZONA", "AZ")
    S = Replace(S, "ALASKA", "AK")
    S = Replace(S, "ALABAMA", "AL")
    
    b = InStrRev(S, " ")
    C = Left(S, b - 1)
    
    'c = Replace(c, ",", "")
    'DUKE_mail_city = Replace(c, ".", "")
    DUKE_mail_city = Application.Trim(C)
End Function

Function DUKE_mail_state(S)
    
    S = UCase(S)
    
    A = Trim(Replace(S, ",", " "))
    b = InStrRev(A, " ")
    
    If A = "" Then
        DUKE_mail_state = ""
        Exit Function
    End If
    
    S = Replace(S, "WYOMING", "WY")
    S = Replace(S, "WISCONSIN", "WI")
    S = Replace(S, "WEST VIRGINIA", "WV")
    S = Replace(S, "WASHINGTON", "WA")
    S = Replace(S, "VIRGINIA", "VA")
    S = Replace(S, "VERMONT", "VT")
    S = Replace(S, "UTAH", "UT")
    S = Replace(S, "TEXAS", "TX")
    S = Replace(S, "TENNESSEE", "TN")
    S = Replace(S, "SOUTH DAKOTA", "SD")
    S = Replace(S, "SOUTH CAROLINA", "SC")
    S = Replace(S, "RHODE ISLAND", "RI")
    S = Replace(S, "PENNSYLVANIA", "PA")
    S = Replace(S, "OREGON", "OR")
    S = Replace(S, "OKLAHOMA", "OK")
    S = Replace(S, "OHIO", "OH")
    S = Replace(S, "NORTH DAKOTA", "ND")
    S = Replace(S, "NORTH CAROLINA", "NC")
    S = Replace(S, "NEW YORK", "NY")
    S = Replace(S, "NEW MEXICO", "NM")
    S = Replace(S, "NEW JERSEY", "NJ")
    S = Replace(S, "NEW HAMPSHIRE", "NH")
    S = Replace(S, "NEVADA", "NV")
    S = Replace(S, "NEBRASKA", "NE")
    S = Replace(S, "MONTANA", "MT")
    S = Replace(S, "MISSOURI", "MO")
    S = Replace(S, "MISSISSIPPI", "MS")
    S = Replace(S, "MINNESOTA", "MN")
    S = Replace(S, "MICHIGAN", "MI")
    S = Replace(S, "MASSACHUSETTS", "MA")
    S = Replace(S, "MARYLAND", "MD")
    S = Replace(S, "MAINE", "ME")
    S = Replace(S, "LOUISIANA", "LA")
    S = Replace(S, "KENTUCKY", "KY")
    S = Replace(S, "KANSAS", "KS")
    S = Replace(S, "IOWA", "IA")
    S = Replace(S, "INDIANA", "IN")
    S = Replace(S, "ILLINOIS", "IL")
    S = Replace(S, "IDAHO", "ID")
    S = Replace(S, "HAWAII", "HI")
    S = Replace(S, "GEORGIA", "GA")
    S = Replace(S, "FLORIDA", "FL")
    S = Replace(S, "DISTRICT OF COLUMBIA", "DC")
    S = Replace(S, "DELAWARE", "DE")
    S = Replace(S, "CONNECTICUT", "CT")
    S = Replace(S, "COLORADO", "CO")
    S = Replace(S, "CALIFORNIA", "CA")
    S = Replace(S, "ARKANSAS", "AR")
    S = Replace(S, "ARIZONA", "AZ")
    S = Replace(S, "ALASKA", "AK")
    S = Replace(S, "ALABAMA", "AL")
    
    A = Split(S, " ")
    b = A(UBound(A))
    If Len(b) = 2 Then
        DUKE_mail_state = b
    Else
        DUKE_mail_state = ""
    End If
    
End Function

Function DUKE_service_address(s1, s2)
    S = Application.Trim(s1 & " " & s2)
    If S Like "* MISC*" Then
        S = Trim(Left(S, InStr(1, S, "MISC") - 2))
    End If
    S = Replace(S, ":", "")
    DUKE_service_address = S
End Function

Function JCPL_mail_address(m1, m2, Optional m3 = "", Optional m4 = "", Optional m5 = "", Optional pobox = "")
    m1 = Application.Trim(m1)
    m2 = Application.Trim(m2)
    If m1 = m2 Then m2 = ""
    pobox = Application.Trim(pobox)
    If pobox <> "" Then
        JCPL_mail_address = pobox
    Else
        JCPL_mail_address = Application.Trim(m1 & " " & m2 & " " & " " & m5)
    End If
    JCPL_mail_address = Replace(JCPL_mail_address, "P O BOX", "PO BOX")
End Function

Function FE_mail_address(m1, m2)
    m1 = Application.Trim(m1)
    m2 = Application.Trim(m2)
    If m1 = m2 Then m2 = ""
    If m1 = "-" And m2 <> "-" Then
        m = m2
    ElseIf m1 <> "-" Then
        m = m1
    Else
        FE_mail_address = ""
        Exit Function
    End If
    m = Replace(m, "P O ", "PO ")
    m = Replace(m, " AVENUE", " AVE")
    m = Replace(m, " STREET", " ST")
    m = Replace(m, ".5", " 1/2")
    m = UCase(m)
    
    FE_mail_address = m
End Function

Function FE_service_address(s1)
    s1 = Application.Trim(s1)
    s1 = Replace(s1, " BLK", "")
    s1 = Replace(s1, " HSM", "")
    s1 = Replace(s1, " GARAGE", "")
    s1 = Replace(s1, " BASE", "")
    If s1 Like "* SIGN" Then
        s1 = Replace(s1, " SIGN", "")
    End If
    If s1 Like "* BILL BOX #*" Then
        s1 = Left(s1, InStrRev(s1, " BILL BOX") - 1)
    End If
    If s1 Like "* 0*" Then
        While s1 Like "* 0*"
            s1 = Replace(s1, " 0", " ")
        Wend
    End If
    FE_service_address = Application.Trim(s1)
End Function

Function FE_service_city(S)
    A = InStr(1, S, ", ", vbTextCompare)
    FE_service_city = Trim(Left(S, A - 1))
End Function

Function FE_service_state(S)
    A = InStr(1, S, ", ", vbTextCompare)
    FE_service_state = Trim(Mid(S, A + 2, 2))
End Function

Function FE_service_zip(S)
    If Right(S, 10) = format(Right(S, 10), "!#####-####") Then
        FE_service_zip = Trim(Right(S, 10))
    ElseIf Right(S, 5) = format(Right(S, 5), "#####") Then
        FE_service_zip = Right(S, 5)
    Else
        FE_service_zip = ""
    End If
End Function

Function JCPL_service_city(S)
    'city, st zip
    A = Application.Trim(S)
    b = InStrRev(A, ",")
    If b <> 0 Then
        C = Left(A, b - 1)
    Else
        C = S
    End If
    JCPL_service_city = C
End Function

Function JCPL_service_state(S)
    'city, st zip
    A = Application.Trim(S)
    b = InStrRev(A, ",")
    C = Mid(A, b + 1, 3)
    JCPL_service_state = C
End Function

Function JCPL_service_zip(S)
    'city, st zip
    A = Application.Trim(S)
    JCPL_service_zip = Right(A, 5)
End Function

Function service_address_old(Optional no_apt = False)
    If EDC_name Like "FE*" Then
        s1 = gagg.Cells(i, s_1)
        S = FE_service_address(s1)
    ElseIf EDC_name Like "AEP*" Then
        s1 = gagg.Cells(i, s_1)
        s2 = gagg.Cells(i, s_2)
        S = AEP_service_address(s1, s2)
    ElseIf EDC_name = "AES" Then
        s1 = gagg.Cells(i, s_1)
        s2 = gagg.Cells(i, s_2)
        S = AES_service_address(s1, s2)
    ElseIf EDC_name = "DUKE" Then
        s1 = gagg.Cells(i, s_1)
        If no_apt Then
            s2 = gagg.Cells(i, s_2)
        Else
            s2 = gagg.Cells(i, s_2) & " " & gagg.Cells(i, duke_apt) & " " & gagg.Cells(i, duke_floor)
        End If
        s2 = Application.Trim(s2)
        S = DUKE_service_address(s1, s2)
    ElseIf EDC_name = "COM" Then
        s1 = gagg.Cells(i, s_1)
        S = COMED_service_address(s1)
    ElseIf EDC_name = "AM" Then
        s1 = gagg.Cells(i, s_1)
        s2 = gagg.Cells(i, s_2)
        If no_apt Then
            s3 = ""
        Else
            s3 = gagg.Cells(i, s_2 + 1)
        End If
        S = AM_service_address(s1, s2, s3)
    ElseIf EDC_name = "JCPL" Then
        S = gagg.Cells(i, s_1)
    End If
    While Left(S, 1) = "0"
        S = Mid(S, 2)
    Wend
    S = Replace(S, ".", "")
    S = Replace(S, ",", "")
    S = Replace(S, ";", "")
    S = Replace(S, ":", "")
    S = Replace(S, "&", " AND ")
    's = Replace(s, " 1/2", ".5")
    's = Replace(s, "-", " ")
    service_address = Application.Trim(UCase(S))
End Function

Function mail_address_old()
    If EDC_name Like "FE*" Then
        m1 = gagg.Cells(i, m_1)
        m2 = gagg.Cells(i, m_1 + 1)
        m3 = gagg.Cells(i, m_2)
        m = FE_mail_address(m1 & " " & m2, m3)
    ElseIf EDC_name Like "AEP*" Then
        m1 = gagg.Cells(i, m_1)
        m2 = gagg.Cells(i, m_2)
        m = AEP_mail_address(m1, m2)
    ElseIf EDC_name = "AES" Then
        m1 = gagg.Cells(i, m_1)
        m2 = gagg.Cells(i, m_2)
        m = AES_mail_address(m1, m2)
    ElseIf EDC_name = "DUKE" Then
        m1 = gagg.Cells(i, m_1)
        m2 = gagg.Cells(i, m_2)
        'm3 = gagg.Cells(i, duke_apt) & " " & gagg.Cells(i, duke_floor)
        m = DUKE_mail_address(m1, m2)
    ElseIf EDC_name = "COM" Then
        m = COMED_mail_address(gagg.Cells(i, m_1))
    ElseIf EDC_name = "AM" Then
        m1 = gagg.Cells(i, m_1)
        m2 = gagg.Cells(i, m_2)
        m3 = gagg.Cells(i, m_2 + 1)
        m = AM_mail_address(m1, m2, m3)
    ElseIf EDC_name = "JCPL" Then
        m1 = gagg.Cells(i, m_1)
        m2 = gagg.Cells(i, m_1 + 1)
        m3 = gagg.Cells(i, m_1 + 2)
        m4 = gagg.Cells(i, m_1 + 3)
        m5 = gagg.Cells(i, m_1 + 4)
        pobox = gagg.Cells(i, m_2)
        m = JCPL_mail_address(m1, m2, m3, m4, m5, pobox)
    End If
    m = Replace(m, ".", "")
    m = Replace(m, ",", "")
    m = Replace(m, ";", "")
    m = Replace(m, ":", "")
    m = Replace(m, "&", " AND ")
    'm = Replace(m, " 1/2", ".5")
    'm = Replace(m, "-", " ")
    While m Like "0*" And Not m Like "0 *"
        m = Mid(m, 2)
    Wend
    If m Like "* # @*" Then
        m = ""
    End If
    If m Like "POB #*" Then
        m = Replace(m, "POB ", "PO BOX")
    End If
    If m Like "P O *" Then
        m = Replace(m, "P O", "PO")
    End If
    m = UCase(m)
    mail_address = Application.Trim(UCase(m))
End Function

Function service_state_old()
    If EDC_name Like "FE*" Then
        S = FE_service_state(gagg.Cells(i, s_state))
    ElseIf EDC_name Like "AEP*" Then
        S = gagg.Cells(i, s_state)
    ElseIf EDC_name = "AES" Then
        S = AES_service_state(gagg.Cells(i, s_state))
    ElseIf EDC_name = "DUKE" Then
        S = gagg.Cells(i, s_state)
    ElseIf EDC_name = "COM" Then
        S = COMED_service_state(gagg.Cells(i, s_state))
    ElseIf EDC_name = "AM" Then
        S = Application.Trim(gagg.Cells(i, s_state))
    ElseIf EDC_name = "JCPL" Then
        S = JCPL_service_state(gagg.Cells(i, s_state))
    End If
    service_state = UCase(Application.Trim(S))
End Function

Function mail_state_old()
    If EDC_name Like "FE*" Then
        m = gagg.Cells(i, m_state)
    ElseIf EDC_name Like "AEP*" Then
        m = gagg.Cells(i, m_state)
    ElseIf EDC_name = "AES" Then
        m = AES_mail_state(gagg.Cells(i, m_state))
    ElseIf EDC_name = "DUKE" Then
        m = DUKE_mail_state(gagg.Cells(i, m_state))
    ElseIf EDC_name = "COM" Then
        m = COMED_mail_state(gagg.Cells(i, m_state))
    ElseIf EDC_name = "AM" Then
        m = gagg.Cells(i, m_state)
    ElseIf EDC_name = "JCPL" Then
        m = gagg.Cells(i, m_state)
    End If
    mail_state = UCase(Application.Trim(m))
End Function

Function service_city_old() As String
    If EDC_name Like "FE*" Then
        S = FE_service_city(gagg.Cells(i, s_city))
    ElseIf EDC_name Like "AEP*" Then
        S = gagg.Cells(i, s_city)
    ElseIf EDC_name = "AES" Then
        S = AES_service_city(gagg.Cells(i, s_city))
    ElseIf EDC_name = "DUKE" Then
        S = gagg.Cells(i, s_city)
    ElseIf EDC_name = "COM" Then
        S = COMED_service_city(gagg.Cells(i, s_city))
    ElseIf EDC_name = "AM" Then
        S = Application.Trim(gagg.Cells(i, s_city))
    ElseIf EDC_name = "JCPL" Then
        S = JCPL_service_city(gagg.Cells(i, s_city))
    End If
    service_city = spellcheck(UCase(Application.Trim(S)))
End Function

Function mail_city_old()
    If EDC_name Like "FE*" Then
        m = gagg.Cells(i, m_city)
    ElseIf EDC_name Like "AEP*" Then
        m = gagg.Cells(i, m_city)
    ElseIf EDC_name = "AES" Then
        m = AES_mail_city(gagg.Cells(i, m_city))
    ElseIf EDC_name = "DUKE" Then
        m = DUKE_mail_city(gagg.Cells(i, m_city))
    ElseIf EDC_name = "COM" Then
        m = COMED_mail_city(gagg.Cells(i, m_city))
    ElseIf EDC_name = "AM" Then
        m = gagg.Cells(i, m_city)
    ElseIf EDC_name = "JCPL" Then
        m = gagg.Cells(i, m_city)
    End If
    mail_city = spellcheck(UCase(Application.Trim(m)))
End Function

Function trimjunk(S)
    S = Application.Trim(S)
    A = Split(S, " ")
    If UBound(A) < 2 Then
        trimjunk = S
        Exit Function
    End If
    If S Like "APT *" Then
        trimjunk = Mid(S, InStr(1, S, A(1)) + Len(A(1)) + 1)
        Exit Function
    End If
    For k = 1 To 0 Step -1
        If A(k) Like "*[0-9]*" Then
            trimjunk = Mid(S, InStr(1, S, A(k)) + Len(A(k)) + 1)
            Exit Function
        End If
    Next
    
    trimjunk = S
    
End Function

Function trim_junk(S)
    'remove unncessary sffixes from mail address if using service address
    trim_junk = S
    
    A = InStrRev(trim_junk, " LT ")
    If A <> 0 Then
        trim_junk = Left(trim_junk, A - 1)
    End If
    
    'b = InStrRev(trim_junk, " LOT #")
    'If b <> 0 Then
    '    trim_junk = Left(trim_junk, b - 1)
    'End If
    
    If trim_junk Like "* LOT #*" Then
        k = Split(trim_junk, " ")
        b = ""
        For j = 0 To UBound(k) - 2
            b = b & " " & k(j)
        Next
        trim_junk = Trim(b)
    End If
    
    If trim_junk Like "* [#] [A-Z][A-Z]*" Then
        trim_junk = Left(S, InStrRev(trim_junk, "#") - 2)
    End If

End Function

Function AEP_service_zip(S)
    AEP_service_zip = format_zip(S)
End Function

Function AEP_mail_zip(m)
    AEP_mail_zip = format_zip(m)
End Function

Function service_zip_old()
    If EDC_name Like "FE*" Then
        S = FE_service_zip(gagg.Cells(i, s_city))
    ElseIf EDC_name Like "AEP*" Then
        S = AEP_service_zip(gagg.Cells(i, s_zip))
    ElseIf EDC_name = "AES" Then
        S = gagg.Cells(i, s_zip)
    ElseIf EDC_name = "DUKE" Then
        S = gagg.Cells(i, s_zip)
    ElseIf EDC_name = "COM" Then
        S = COMED_service_zip(gagg.Cells(i, s_zip))
    ElseIf EDC_name = "AM" Then
        S = Application.Trim(gagg.Cells(i, s_zip))
    ElseIf EDC_name = "JCPL" Then
        S = JCPL_service_zip(gagg.Cells(i, s_zip))
    End If
    service_zip = format_zip(S)
End Function

Function mail_zip_old()
    If EDC_name Like "FE*" Then
        m = gagg.Cells(i, m_zip)
    ElseIf EDC_name Like "AEP*" Then
        m = AEP_mail_zip(gagg.Cells(i, m_zip))
    ElseIf EDC_name = "AES" Then
        m = AES_mail_zip(gagg.Cells(i, m_zip - 1), gagg.Cells(i, m_zip))
    ElseIf EDC_name = "DUKE" Then
        m = gagg.Cells(i, m_zip)
    ElseIf EDC_name = "COM" Then
        m = gagg.Cells(i, m_zip)
    ElseIf EDC_name = "AM" Then
        m = gagg.Cells(i, m_zip)
    ElseIf EDC_name = "JCPL" Then
        m = gagg.Cells(i, m_zip)
    End If
    mail_zip = format_zip(m)
End Function

Sub FE_address_replace()

    On Error Resume Next
    
    Set ff = filter_tab()
    Set s1 = Sheets("FE Mail")
    
    If s1 Is Nothing Then Exit Sub
    
    Call sort_sheet_col(ff, 1, "A")
    Call sort_sheet_col(s1, 1, "A")
    
    s1.UsedRange.Sort key1:=s1.columns(1), Order1:=xlAscending, header:=xlYes
    
    arr_1 = s1.UsedRange.columns(1).value
    arr_6 = s1.UsedRange.columns(6).value
    arr_7 = s1.UsedRange.columns(7).value
    arr_8 = s1.UsedRange.columns(8).value
    arr_9 = s1.UsedRange.columns(9).value
    arr_10 = s1.UsedRange.columns(10).value
    arr_11 = s1.UsedRange.columns(11).value
    arr_12 = s1.UsedRange.columns(12).value
    
    arr_1 = flatten_array(arr_1)
    
    n = UBound(arr_1)
    
    With F.columns
        account_arr = filter_col(.account_number)
        mail_address_arr = filter_col(.mail_address)
        mail_city_arr = filter_col(.mail_city)
        mail_state_arr = filter_col(.mail_state)
        mail_zip_arr = filter_col(.mail_zip)
    End With
    
    search_start = 2
    
    progress.start ("Replacing FE Addresses")
    
    For i = 2 To UBound(account_arr)
        j = array_binary_search(account_arr(i, 1), arr_1, search_start, n)
        If j > 0 Then
            search_start = j
            x1 = arr_12(j, 1)
            If x1 = "" Then
                x1 = arr_7(j, 1) & " " & arr_8(j, 1) & " " & arr_6(j, 1)
                x1 = Application.Trim(UCase(x1))
            End If
            x2 = UCase(arr_9(j, 1))
            x3 = UCase(arr_10(j, 1))
            x4 = format(arr_11(j, 1), "00000")
            mail_address_arr(i, 1) = x1
            mail_city_arr(i, 1) = x2
            mail_state_arr(i, 1) = x3
            mail_zip_arr(i, 1) = x4
        End If
    Next
    
    delete_sheet ("FE Mail")
    
    ff.UsedRange.columns(F.columns.mail_address.index).value = mail_address_arr
    ff.UsedRange.columns(F.columns.mail_city.index).value = mail_city_arr
    ff.UsedRange.columns(F.columns.mail_state.index).value = mail_state_arr
    ff.UsedRange.columns(F.columns.mail_zip.index).value = mail_zip_arr

    progress.complete
    
End Sub
