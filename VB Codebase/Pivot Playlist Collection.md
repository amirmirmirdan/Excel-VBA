### Format Data Range as Table (Create List Object)
**Input:** Table Name (String)
```vb
Sub CreateListObject(strTableName As String)
	Dim rng As Range
    Set rng = ActiveSheet.Range("A1")

	Application.CutCopyMode = False
	ActiveSheet.ListObjects.Add(xlSrcRange, rng.CurrentRegion, , xlYes).Name = strTableName
        ActiveSheet.ListObjects(strTableName).TableStyle = ""
        
    Set rng = Nothing
End Sub
```
---
### Create Pivot Cache
```vb
Sub CreatePvtCache()
    Application.CutCopyMode = False
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:="Report21DataTable", _
        Version:=8).CreatePivotTable TableDestination:="Sheet1!R4C1", _
                TableName:="PvtClearing", DefaultVersion:=8
                
    ' Possible to add a select case stmt for each playlist.
    ' Create Pivot Cache Subroutine could also include arguement "TableName" as string
    ' Same as TableDestination & Source Data as well.
End Sub
```
---
### Format Pivot Table
```vb
Sub FormatPvtTable()
    ActiveWorkbook.Worksheets("Sheet1").Activate
    With ActiveSheet.PivotTables("PvtClearing")
        .ColumnGrand = False
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = False
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = False
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
    End With
    With ActiveSheet.PivotTables("PvtClearing").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
End Sub
```
### Clear Pivot Cache
```vb
Sub ClearPvtCache()
    ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlDataAndLabel, True
    ActiveSheet.PivotTables("PivotTable1").ClearTable
End Sub
```
---
### Refresh All Pivot Tables
```vb
Sub RefereshAllPvtTable()
    Dim pt As PivotTable
        For Each pt In ActiveWorkbook.PivotTables
            pt.RefreshTable
        Next pt
End Sub
```
---
### Refresh All
```vb
Sub RefreshAll()
	ActiveWorkbook.RefreshAll
End Sub
```
---


### Pivot Field Playlist
- [ ] [[PVT Clearing System Transactions]]
- [ ] 