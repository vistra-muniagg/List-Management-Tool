Function find_dna_list(Optional try_day = "")

    find_dna_list = ""
    
    dna_folder = onedrive_dna_folder()
    
    'try to open file
    mondays = 0
    If try_day = "" Then try_day = format(Date, "m-d-yy")
    While mondays < 100
        try_path = onedrive_dna_file_path(dna_folder, try_day)
        If try_path = False Then
            try_day = previous_day(try_day)
        Else
            GoTo found_list
        End If
        If format(try_day, "dddd") = "Monday" Then mondays = mondays + 1
    Wend
    
found_list:
    find_dna_list = try_path
    
End Function

Function onedrive_dna_file_path(folder, try_day)

    onedrive_dna_file_path = False
    
    p1 = folder
    
    p2 = "PUCO - Do Not Aggregate List (MM-DD-YY).xlsx"
    
    try_path = Replace(p1 & p2, "MM-DD-YY", try_day)
    
    x = Dir(try_path, vbNormal)
    
    If x = "" Then
        onedrive_dna_file_path = False
    Else
        onedrive_dna_file_path = try_path
    End If
    
End Function

Function onedrive_dna_folder_old()

    onedrive_dna_folder_old = ""
    
    x = Environ("USERPROFILE")
    
    folder_pattern = "*PUCO Do Not Aggregate (DNA) List"
    
    b1 = "\OneDrive - Vistra Corp\(1) Operations\(6) List Management\" & folder_pattern
    b2 = "\OneDrive - Vistra Corp\(6) List Management\" & folder_pattern
    b3 = "\OneDrive - Vistra Corp\MUNI AGG\(1) Operations\(6) List Management\" & folder_pattern
    b4 = "\OneDrive - Vistra Corp\Shared Documents - Muni-Agg\(1) Operations\(6) List Management\" & folder_pattern
    
    d = ""
    
    If d = "" Then
        b = b1
        d = Dir(x & b1, vbDirectory)
        If d <> "" Then onedrive_dna_folder_old = x & Replace(b1, folder_pattern, d)
    End If
    
    If d = "" Then
        b = b2
        d = Dir(x & b2, vbDirectory)
        If d <> "" Then onedrive_dna_folder_old = x & Replace(b2, folder_pattern, d)
    End If
    
    If d = "" Then
        b = b3
        d = Dir(x & b3, vbDirectory)
        If d <> "" Then onedrive_dna_folder_old = x & Replace(b3, folder_pattern, d)
    End If
    
    If d = "" Then
        b = b4
        d = Dir(x & b4, vbDirectory)
        If d <> "" Then onedrive_dna_folder_old = x & Replace(b4, folder_pattern, d)
    End If
    
    If Not onedrive_dna_folder_old Like "*\" Then onedrive_dna_folder_old = onedrive_dna_folder_old & "\"
    
End Function

