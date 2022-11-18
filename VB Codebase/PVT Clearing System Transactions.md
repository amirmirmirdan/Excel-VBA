### Pivot Fields
```vb
Sub PvtFieldPlaylist1()
' Designning The First Pivot Playlist for Clearing External Transactions items (Items that is reconciled)
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("SEGMENT1")
        .Orientation = xlPageField
        .Position = 1
        .CurrentPage = "(All)"
        .EnableMultiplePageItems = True
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("JV_HEADER_DESCRIPTION")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("JV_NAME")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("JE_CATEGORY")
        .Orientation = xlRowField
        .Position = 3
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("JE_SOURCE")
        .Orientation = xlRowField
        .Position = 4
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("EVENT_CLASS_TYPE")
        .Orientation = xlRowField
        .Position = 5
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("SEGMENT3")
        .Orientation = xlRowField
        .Position = 6
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("SEGMENT3_DESC")
        .Orientation = xlRowField
        .Position = 7
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("REFERENCE_NUMBER")
        .Orientation = xlRowField
        .Position = 8
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("ACCOUNTING_SEQUENCE_NUMBER")
        .Orientation = xlRowField
        .Position = 9
    End With
    
    With ActiveSheet.PivotTables("PvtClearing").PivotFields("CREATED_BY")
        .Orientation = xlRowField
        .Position = 10
    End With
    ActiveWorkbook.ShowPivotTableFieldList = False
End Sub
```

### Filtering Pivot Field
```vb
Sub FilterPvtTablePlaylist1()
    With ActiveSheet.PivotTables("PvtClearing")
        .PivotFields("JE_CATEGORY").PivotFilters.Add2 _
            Type:=xlCaptionEquals, Value1:="Miscellaneous"
            
        .PivotFields("SEGMENT3").PivotFilters. _
        Add2 Type:=xlCaptionBeginsWith, Value1:="123"
    End With
    ' Write Formulas to Count the items Filtered & Pivotted    
    Range("E2").Formula2R1C1 = "=COUNTA(C[3])-1"
End Sub
```
