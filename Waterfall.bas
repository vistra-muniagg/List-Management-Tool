Sub make_filter_waterfall()

    progress.start ("Making Waterfall")
    
    Set ff = filter_tab()
    
    Set h = home_tab()
    
    waterfall_name = MT.waterfall_title
    
    Set table_location = h.Range(S.HOME.filter_waterfall_location)
    
    On Error Resume Next
    h.PivotTables(waterfall_name).TableRange2.Clear
    
    'create waterfall pivot table
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ff.UsedRange)

    Set pvt = h.PivotTables.Add(PivotCache:=pvtCache, TableDestination:=table_location, tableName:=waterfall_name)
    
    With h.PivotTables(MT.waterfall_title)
        .PivotFields(F.columns.status.header).Orientation = xlRowField
        .PivotFields(F.columns.status.header).Position = 1
        .PivotFields(F.columns.mail_category.header).Orientation = xlRowField
        .PivotFields(F.columns.mail_category.header).Position = 1
        .PivotFields(F.columns.mail_category.header).AutoSort xlAscending, F.columns.mail_category.header
        .HasAutoFormat = False
    End With
    
    pvt.AddDataField pvt.PivotFields(F.columns.account_number.header), "Count", xlCount
    
    h.PivotTables(1).PivotSelect "", xlDataAndLabel
    
    'sort pivot table filters A-Z
    h.PivotTables(waterfall_name).PivotFields(F.columns.status.header).AutoSort xlAscending, F.columns.status.header
    
    table_location.value = waterfall_name
    
    progress.finish
    
    Range("A1").Select
    
End Sub

Sub make_geocode_waterfall()
    
    progress.start ("Making Geocoding Waterfall")
    
    Set h = home_tab()
    Set ff = filter_tab()
    
    On Error Resume Next
    h.PivotTables(S.HOME.mapping_waterfall_name).TableRange2.Clear
    
    Set table_location = h.Range(S.HOME.mapping_waterfall_location)
    table_name = S.HOME.mapping_waterfall_name
    
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ff.UsedRange)
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=table_location, tableName:=table_name)
    
    account_header = F.columns.account_number.header
    mapping_header = F.columns.mapping_result.header
    community_header = F.columns.community_mapped_into.header
    
    map_in_label = FS.mapping.maps_in_label
    map_out_label = FS.mapping.mapped_out_label
    no_results_label = FS.mapping.no_results_label
    retained_label = FS.mapping.mapped_out_retained_label
    
    With h.PivotTables(table_name)
            
        .AddDataField .PivotFields(account_header), "Accounts", xlCount
        .AddDataField .PivotFields(account_header), "% of Total", xlCount
        
        .PivotFields("Accounts").Position = 1
        
        .PivotFields(mapping_header).Orientation = xlRowField
        .PivotFields(mapping_header).Position = 1
        .PivotFields(community_header).Orientation = xlRowField
        .PivotFields(community_header).Position = 2
        .PivotFields(mapping_header).AutoSort xlAscending, mapping_header
        
        .PivotFields(mapping_header).PivotItems(map_in_label).ShowDetail = True
        .PivotFields(mapping_header).PivotItems(map_out_label).ShowDetail = False
        .PivotFields(mapping_header).PivotItems(no_results_label).ShowDetail = False
        .PivotFields(mapping_header).PivotItems(retained_label).ShowDetail = False
        
        .ShowValuesRow = False
        .HasAutoFormat = False
        
        .PivotFields("% of Total").Calculation = xlPercentOfTotal
        .PivotFields("% of Total").NumberFormat = "0.00%"
        
        .PivotSelect "", xlDataAndLabel, True
        
        'add_borders xlMedium
        
        table_location.value = S.HOME.mapping_waterfall_caption
    
    End With
    
    progress.finish
    
    Range("A1").Select
    
    ThisWorkbook.RefreshAll
    
End Sub

Sub make_cycle_waterfall()

    Set h = home_tab()
    Set ff = filter_tab()
    
    row_count = Application.CountA(ff.columns(1)) - 1
    
    Set table_location = h.Range(S.HOME.cycle_pivot_location)
    table_name = S.HOME.cycle_pivot_name
    
    eligible_col_label = F.columns.eligible_opt_out.header
    cycle_col_label = F.columns.read_cycle.header
    category_label = F.columns.mail_category.header
    
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    h.Activate
    
    On Error Resume Next
    h.PivotTables(table_name).TableRange2.Clear
    
    'create waterfall pivot table
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ff.UsedRange)
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=table_location, tableName:=table_name)
    
    With h.PivotTables(table_name)
        .AddDataField .PivotFields(cycle_col_label), "Count", xlCount
        .PivotFields("Count").Position = 1
        .PivotFields(cycle_col_label).Orientation = xlRowField
        .PivotFields(cycle_col_label).Position = 1
        .ShowValuesRow = S.HOME.add_cycle_pivots
        .PivotSelect "", xlDataAndLabel, True
        table_location.value = S.HOME.cycle_pivot_caption
        .PivotFields(eligible_col_label).Orientation = xlPageField
        .PivotFields(eligible_col_label).Position = 1
        .PivotFields(eligible_col_label).CurrentPage = "Y"
        If MT.cycle_pivot_colors And row_count >= S.HOME.large_community_limit Then
            Set r = .DataBodyRange
            r.FormatConditions.Delete
            Set r = r.Resize(r.Cells.count - 1)
            Set fc = r.FormatConditions.AddColorScale(ColorScaleType:=2)
            With fc.ColorScaleCriteria(1)
                .Type = xlConditionValueLowestValue
                .FormatColor.color = C.NONE.InteriorColor
            End With
            With fc.ColorScaleCriteria(2)
                .Type = xlConditionValueHighestValue
                .FormatColor.color = C.BLUE_1.InteriorColor
            End With
        End If
        .PivotFields(category_label).Orientation = xlRowField
        .PivotFields(category_label).Position = 1
        .PivotFields(category_label).ShowDetail = False
    End With
    
    Range("A1").Select
    
End Sub
