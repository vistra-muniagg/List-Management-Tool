Public Type SheetNames
    HOME As String
    README As String
    Stats As String
    Filter As String
    mapping As String
    Utility As String
    Active As String
    Supplier As String
    DNA As String
    Snowflake As String
    contracts As String
    LP As String
    ren_drops As String
    opt_in As String
    mail_list As String
    duke_siblings As String
    premise_mismatch As String
End Type

Public Type CellColors
    FontColor As Long
    InteriorColor As Long
End Type

Public Type MacroColors
    NONE As CellColors
    GRAY_1 As CellColors
    GRAY_2 As CellColors
    GRAY_3 As CellColors
    LIGHT_ORANGE As CellColors
    ORANGE As CellColors
    DARK_ORANGE As CellColors
    YELLOW As CellColors
    GREEN As CellColors
    GREEN_0 As CellColors
    GREEN_1 As CellColors
    GREEN_2 As CellColors
    GREEN_3 As CellColors
    LIGHT_PURPLE As CellColors
    PINK As CellColors
    GOLD As CellColors
    MAGENTA As CellColors
    DARK_GRAY As CellColors
    RED As CellColors
    DARK_RED As CellColors
    BLUE As CellColors
    BLUE_1 As CellColors
    BLUE_2 As CellColors
    BLUE_3 As CellColors
    PURPLE As CellColors
    DARK_PURPLE As CellColors
End Type

Public Type CellFormatting
    add_utility_list_dupe_comments As Boolean
    add_utility_list_dupe_colors As CellColors
    utility_list_dupes_color As CellColors
    utility_list_dupe_comment As String
    add_overwrite_comments As Boolean
    add_overwrite_colors As Boolean
    add_FE_overwrite_comments As Boolean
    add_FE_overwrite_colors As Boolean
End Type

Sub define_colors()
    With C.NONE
        .InteriorColor = xlColorIndexNone
        .FontColor = vbBlack
    End With
    With C.GRAY_1
        .InteriorColor = RGB(217, 217, 217)
        .FontColor = vbBlack
    End With
    With C.GRAY_2
        .InteriorColor = RGB(173, 173, 173)
        .FontColor = vbWhite
    End With
    With C.GRAY_3
        .InteriorColor = RGB(116, 116, 116)
        .FontColor = vbWhite
    End With
    With C.LIGHT_ORANGE
        .InteriorColor = RGB(255, 204, 153)
        .FontColor = vbBlack
    End With
    With C.ORANGE
        .InteriorColor = RGB(255, 153, 0)
        .FontColor = vbBlack
    End With
    With C.DARK_ORANGE
        .InteriorColor = RGB(255, 102, 0)
        .FontColor = vbBlack
    End With
    With C.YELLOW
        .InteriorColor = RGB(255, 255, 0)
        .FontColor = vbBlack
    End With
    With C.GREEN
        .InteriorColor = RGB(0, 255, 0)
        .FontColor = vbBlack
    End With
    With C.GREEN_0
        .InteriorColor = RGB(181, 230, 162)
        .FontColor = vbBlack
    End With
    With C.GREEN_1
        .InteriorColor = RGB(131, 226, 142)
        .FontColor = vbBlack
    End With
    With C.GREEN_2
        .InteriorColor = RGB(146, 208, 80)
        .FontColor = vbBlack
    End With
    With C.GREEN_3
        .InteriorColor = RGB(60, 125, 34)
        .FontColor = vbWhite
    End With
    With C.PINK
        .InteriorColor = RGB(255, 192, 203)
        .FontColor = vbBlack
    End With
    With C.GOLD
        .InteriorColor = RGB(255, 215, 0)
        .FontColor = vbBlack
    End With
    With C.MAGENTA
        .InteriorColor = RGB(255, 0, 255)
        .FontColor = vbBlack
    End With
    With C.DARK_GRAY
        .InteriorColor = RGB(128, 128, 128)
        .FontColor = vbWhite
    End With
    With C.RED
        .InteriorColor = RGB(255, 0, 0)
        .FontColor = vbWhite
    End With
    With C.DARK_RED
        .InteriorColor = RGB(128, 0, 0)
        .FontColor = vbWhite
    End With
    With C.BLUE
        .InteriorColor = RGB(0, 0, 255)
        .FontColor = vbWhite
    End With
    With C.BLUE_1
        .InteriorColor = RGB(0, 176, 240)
        .FontColor = vbWhite
    End With
    With C.BLUE_2
        .InteriorColor = RGB(0, 112, 192)
        .FontColor = vbWhite
    End With
    With C.BLUE_3
        .InteriorColor = RGB(0, 32, 96)
        .FontColor = vbWhite
    End With
    With C.PURPLE
        .InteriorColor = RGB(128, 0, 128)
        .FontColor = vbWhite
    End With
End Sub

Sub define_sheet_names()
    With SN
        .HOME = "HOME"
        .Filter = "FILTER"
        .mapping = "Geocoding"
        .README = "README"
        .Stats = "STATS"
        .Utility = "UTILITY"
        .Active = "ACTIVE"
        .Supplier = "SUPPLIER"
        .DNA = "DNA"
        .Snowflake = "Snowflake Query"
        .contracts = "Contracts Query"
        .LP = "LP"
        .ren_drops = "Drop At Renewal"
        .opt_in = "Opt In Eligible"
        .mail_list = "Mail List"
        .duke_siblings = "DUKE Sibling Accounts"
        .premise_mismatch = "Premise Mismatch"
    End With
End Sub

Function find_sheet(sheet_name)
    If sheet_name = "" Then
        Set find_sheet = Nothing
        Exit Function
    End If
    For Each ws In ThisWorkbook.Sheets
        If ws.name Like sheet_name & "*" Then
            Set find_sheet = ws
            Exit Function
        End If
    Next
End Function

Function home_tab()
    On Error Resume Next
    Set home_tab = Nothing
    Set home_tab = find_sheet(SN.HOME)
End Function

Function readme_tab()
    On Error Resume Next
    Set readme_tab = Nothing
    Set readme_tab = find_sheet(SN.README)
End Function

Function filter_tab()
    On Error Resume Next
    Set filter_tab = Nothing
    Set filter_tab = find_sheet(SN.Filter)
End Function

Function utility_tab()
    On Error Resume Next
    Set utility_tab = Nothing
    Set utility_tab = find_sheet(SN.Utility)
End Function

Function active_tab()
    On Error Resume Next
    Set active_tab = Nothing
    Set active_tab = find_sheet(SN.Active)
End Function

Function supplier_tab()
    On Error Resume Next
    Set supplier_tab = Nothing
    Set supplier_tab = find_sheet(SN.Supplier)
End Function

Function stats_tab()
    On Error Resume Next
    Set stats_tab = Nothing
    Set stats_tab = find_sheet(SN.Stats)
End Function

Function dna_tab()
    On Error Resume Next
    Set dna_tab = Nothing
    Set dna_tab = find_sheet(SN.DNA)
End Function

Function contracts_tab()
    On Error Resume Next
    Set contracts_tab = Nothing
    Set contracts_tab = find_sheet(SN.contracts)
End Function

Function mapping_tab()
    On Error Resume Next
    Set mapping_tab = Nothing
    Set mapping_tab = find_sheet(SN.mapping)
End Function

Function LP_tab()
    On Error Resume Next
    Set LP_tab = Nothing
    Set LP_tab = find_sheet(SN.LP)
End Function

Function mail_tab()
    On Error Resume Next
    Set mail_tab = Nothing
    Set mail_tab = find_sheet(SN.mail_list)
End Function

Function opt_in_tab()
    On Error Resume Next
    Set opt_in_tab = Nothing
    Set opt_in_tab = find_sheet(SN.opt_in)
End Function

Function drop_tab()
    On Error Resume Next
    Set drop_tab = Nothing
    Set drop_tab = find_sheet(SN.ren_drops)
End Function

Function sibling_tab()
    On Error Resume Next
    Set sibling_tab = Nothing
    Set sibling_tab = find_sheet(SN.duke_siblings)
End Function
