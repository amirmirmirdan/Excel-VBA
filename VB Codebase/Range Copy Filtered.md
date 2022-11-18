[[20220918 (SUN)|Back]]
### Copy Filtered Range
Reference[^RangeCodebase]
```vb
Sub CopyFilteredRange(rgData As Range, rgCriteria As Range, rgCopyTo As Range)
    rgData.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=rgCriteria, _
        CopyToRange:=rgCopyTo, _
        Unique:=False
End Sub
```
1. `rgData` = DataSource Range. Preferably refers to a ListObject
Exp: *"Table1['# All]"*
2. `rgCriteria` = The Criteria Range (Must Include the Column Header as part of the criteria range).
3. `rgCopyTo` = The Destination Range to copy the filtered data to. Destination Range, only include the Header Column of Data we wanted.

[^RangeCodebase]: References [[Base]] & [AdvancedFilter](AdvancedFilter.bas) for the BAS File.
