Public Type CheckListItem
    name As String
    label As String
    index As Long
End Type

Public Type CheckList
    location As String
    title As String
    data_range As String
    items() As CheckListItem
End Type

Sub test_checklist()
    init
    Call update_checklist(S.QC.audit_checklist, "audit2", 1)
    Call update_checklist(S.QC.audit_checklist, "audit3", 0)
    Call update_checklist(S.QC.audit_checklist, "audit4", -1)
    Call update_checklist(S.QC.audit_checklist, "audit5")
    Call update_checklist(S.QC.qc_checklist, "item2", 1)
    Call update_checklist(S.QC.qc_checklist, "item3", 0)
    Call update_checklist(S.QC.qc_checklist, "item4", -1)
    Call update_checklist(S.QC.qc_checklist, "item5")
End Sub

Sub define_checklists()
    With S.QC.audit_checklist
        .location = S.HOME.audit_checklist_location
        .title = "Audit Checklist"
        .items = audit_checklist_items()
    End With
    With S.QC.qc_checklist
        .location = S.HOME.qc_checklist_location
        .title = "QC Checklist"
        .items = qc_checklist_items()
    End With
End Sub

Function checklist_item(index As Long, name As String, label As String) As CheckListItem
    Dim item As CheckListItem
    item.index = index
    item.name = name
    item.label = label
    checklist_item = item
End Function

Sub checklist_append(ByRef arr() As CheckListItem, item As CheckListItem)
    
    If arr(1).name = "" Then
        arr(1) = item
    Else
        n = UBound(arr)
        ReDim Preserve arr(LBound(arr) To n + 1)
        arr(n + 1) = item
    End If
    
End Sub

Function audit_checklist_items() As CheckListItem()
    
    Dim arr() As CheckListItem
    
    ReDim arr(1 To 1)
    
    Call checklist_append(arr, checklist_item(1, "audit_pipp", "audit1"))
    Call checklist_append(arr, checklist_item(2, "audit_usage", "audit2"))
    Call checklist_append(arr, checklist_item(3, "audit_shopping", "audit3"))
    Call checklist_append(arr, checklist_item(4, "audit_arrears", "audit4"))
    Call checklist_append(arr, checklist_item(5, "audit_mercantile_national", "audit5"))
    Call checklist_append(arr, checklist_item(6, "audit_dna", "audit6"))
    Call checklist_append(arr, checklist_item(7, "audit_hourly_pricing", "audit7"))
    Call checklist_append(arr, checklist_item(8, "audit_solar", "audit8"))
    Call checklist_append(arr, checklist_item(9, "audit_free_service", "audit9"))
    Call checklist_append(arr, checklist_item(10, "audit_bgs_hold", "audit10"))
    Call checklist_append(arr, checklist_item(11, "audit_mapping", "audit11"))
    
    audit_checklist_items = arr
    
End Function

Function qc_checklist_items() As CheckListItem()
    
    Dim arr() As CheckListItem
    
    ReDim arr(1 To 1)
    
    Call checklist_append(arr, checklist_item(1, "account_number_format", "item1"))
    Call checklist_append(arr, checklist_item(2, "all_files_present", "item2"))
    Call checklist_append(arr, checklist_item(3, "correct_mapping", "item3"))
    Call checklist_append(arr, checklist_item(4, "apt_numbers", "item4"))
    Call checklist_append(arr, checklist_item(5, "valid_states", "item5"))
    Call checklist_append(arr, checklist_item(6, "valid_zips", "item6"))
    'Call checklist_append(arr, checklist_item(7, "item7", "item7"))
    'Call checklist_append(arr, checklist_item(8, "item8", "item8"))
    'Call checklist_append(arr, checklist_item(9, "item9", "item9"))
    'Call checklist_append(arr, checklist_item(10, "item10", "item10"))
    'Call checklist_append(arr, checklist_item(11, "item11", "item11"))
    'Call checklist_append(arr, checklist_item(12, "item12", "item12"))
    
    qc_checklist_items = arr
    
End Function

Function new_checklist(list_type) As CheckList
    Dim check As CheckList
    With check
        If list_type = "AUDIT" Then
            .location = S.HOME.audit_checklist_location
            .title = S.QC.audit_checklist_title
            .items = audit_checklist_items()
        ElseIf list_type = "FILTER" Then
            .location = S.HOME.qc_checklist_location
            .title = S.QC.qc_checklist_title
            .items = qc_checklist_items()
        Else
            'idk
        End If
    End With
End Function

Sub update_checklist(list As CheckList, name As String, Optional value = "")
    index = checklist_index(list, name)
    If index = -1 Then Exit Sub
    Set cell = home_tab().Range(list.location).Offset(index + 1, 1)
    Select Case value
        Case 1:
            'green check
            cell.value = ChrW(&H2714)
            cell.Font.color = RGB(0, 175, 0)
        Case -1:
            'red x
            cell.value = ChrW(&H2718)
            cell.Font.color = RGB(255, 0, 0)
        Case 0:
            'yellow circle
            cell.value = ChrW(&H25C9)
            cell.Font.color = RGB(225, 200, 0)
        Case Else:
            'blue square
            cell.value = ChrW(&H25A0)
        cell.Font.color = RGB(0, 0, 255)
    End Select
End Sub

Function checklist_index(list As CheckList, name As String)
    checklist_index = -1
    For j = 1 To UBound(list.items)
        'dp list.items(j).name
        If list.items(j).name = name Then
            checklist_index = j
            Exit For
        End If
    Next
End Function

Function compare_checksums(dict As Object) As Boolean
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Set vbProj = ThisWorkbook.VBProject
    compare_checksums = True
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type < 100 Then
            Set codeMod = vbComp.CodeModule
            checksum = module_checksum(codeMod)(0)
            validated_checksum = dict(vbComp.name)
            If checksum <> validated_checksum Then
                compare_checksums = False
                Exit Function
            End If
        End If
    Next
    If dict("checksum") <> GetVBAChecksum() Then compare_checksums = False
End Function

Function GetVBAChecksum()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim fullCode As String
    Set vbProj = ThisWorkbook.VBProject
    fullCode = ""
    For Each vbComp In vbProj.VBComponents
        Set codeMod = vbComp.CodeModule
        lineCount = codeMod.CountOfLines
        If lineCount > 0 Then
            codeText = codeMod.Lines(1, lineCount)
            fullCode = fullCode & codeText & vbLf
        End If
    Next
    'build = format(Now, "yyyymmddhhnn")
    GetVBAChecksum = Hex(CRC32(fullCode))
End Function

Sub WriteDetailedVersionJSON()
    Dim version As String
    Dim buildDate As String
    Dim buildNotes As String
    Dim overallChecksum As String
    Dim filePath As String
    Dim fileNum As Integer
    
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim codeText As String
    Dim lineCount As Long
    Dim moduleChecksum As Long
    Dim modulesJson As String
    
    Dim allCode As String
    
    ' Set version info
    'version = home_tab().Range(S.HOME.version_location)
    
    version = "v0.0.1"
    
    buildDate = format(Date, "yyyymmdd")
    buildNotes = "build log"
    
    Set vbProj = ThisWorkbook.VBProject
    allCode = ""
    modulesJson = ""
    
    ' Compute module checksums & build JSON entries
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type < 100 Then
            Set codeMod = vbComp.CodeModule
            arr = module_checksum(codeMod)
            Set codeMod = vbComp.CodeModule
            modulesJson = modulesJson & "    """ & vbComp.name & """: """ & arr(0) & """, " & vbLf
            allCode = allCode & arr(1)
        End If
    Next
    
    ' Remove trailing comma and newline from modulesJson
    If Len(modulesJson) > 2 Then
        modulesJson = Left(modulesJson, Len(modulesJson) - 3) & vbLf
    End If
    
    ' Compute overall checksum of all code combined
    overallChecksum = Hex(CRC32(allCode))
    
    ' Compose JSON string
    Dim jsonText As String
    jsonText = "{" & vbLf & _
        "  ""version"": """ & version & """," & vbLf & _
        "  ""build_date"": """ & buildDate & """," & vbLf & _
        "  ""checksum"": """ & overallChecksum & """," & vbLf & _
        "  ""build_notes"": """ & Replace(buildNotes, """", "\""") & """," & vbLf & _
        "  ""modules"": {" & vbLf & _
        modulesJson & _
        "  }" & vbLf & _
        "}"
        
    file_name = version & "." & buildDate
    
    ' Save JSON file next to workbook
    filePath = onedrive_list_management_folder & "\" & file_name & ".json"
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, jsonText
    Close #fileNum
    
End Sub

Function module_checksum(codeMod As VBIDE.CodeModule) As String()
    Dim arr() As String
    Dim codeText As String
    ReDim arr(0 To 1)
    lineCount = codeMod.CountOfLines
    If lineCount > 0 Then
        codeText = codeMod.Lines(1, lineCount)
        allCode = allCode & codeText & vbLf
        moduleChecksum = CRC32(codeText)
        arr(0) = Hex(moduleChecksum)
    Else
        allCode = ""
        arr(0) = "(empty)"
    End If
    arr(1) = allCode
    module_checksum = arr
End Function

Function Base64Decode(ByVal base64String As String) As Byte()
    Dim xmlObj As Object
    Set xmlObj = CreateObject("MSXML2.DOMDocument")
    
    Dim node As Object
    Set node = xmlObj.createElement("b64")
    
    node.DataType = "bin.base64"
    node.text = base64String
    
    Base64Decode = node.nodeTypedValue
End Function

Private Function CRC32(strData As String) As Long
    Dim crcTable(0 To 255) As Long
    Dim crc As Long
    Dim i As Long, j As Long
    Dim byteVal As Integer
    
    ' Build the CRC table
    For i = 0 To 255
        crc = i
        For j = 1 To 8
            If (crc And 1) Then
                crc = &HEDB88320 Xor (crc \ 2)
            Else
                crc = crc \ 2
            End If
        Next j
        crcTable(i) = crc
    Next i

    ' Compute the CRC
    crc = &HFFFFFFFF
    For j = 1 To Len(strData)
        byteVal = Asc(Mid$(strData, j, 1))
        crc = crcTable((crc Xor byteVal) And &HFF) Xor (crc \ 256)
    Next
    CRC32 = Not crc
End Function

Sub ExportAllModules()

    Dim vbComp As Object
    
    export_path = Environ("UserProfile") & "\Documents\macros\code\" & format(Date, "MM-DD-YY")
    export_path = export_path & "\"
    
    make_dir (export_path)
    
        ' Loop through all components in the project
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Set codeMod = vbComp.CodeModule
        
        ' Build file name depending on type
        Select Case vbComp.Type
            Case 1 ' Standard module
                fileName = export_path & vbComp.name & ".bas"
            Case 2 ' Class module
                fileName = export_path & vbComp.name & ".cls"
            Case 3 ' UserForm (export code only)
                fileName = export_path & vbComp.name & "_Code.bas"
            Case Else
                GoTo NextComp
        End Select
        
        ' Write the code from the module to file
        fnum = FreeFile
        Open fileName For Output As #fnum
        For i = 1 To codeMod.CountOfLines
            Print #fnum, codeMod.Lines(i, 1)
        Next i
        Close #fnum
        
        dp "Exported code: " & fileName
        
NextComp:
    Next
End Sub
