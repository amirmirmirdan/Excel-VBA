Sub CopyFilteredRange(rgData As Range, rgCriteria As Range, rgCopyTo As Range)
    ' rgData = DataSource Range. Preferably refers to a ListObject
    ' Exp: "Table1[#All]"
    ' rgCriteria = The Criteria Range (Must Include the Column Header as part of the criteria range).
    ' rgCopyTo = The Destination Range to copy the filtered data to. Destination Range, only include the Header Column of Data we wanted.
    rgData.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=rgCriteria, _
        CopyToRange:=rgCopyTo, _
        Unique:=False
End Sub