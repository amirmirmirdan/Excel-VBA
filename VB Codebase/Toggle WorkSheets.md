### Get Sheet Count (Function)
```vb
Function GetSheetsCount() As Integer
        GetSheetsCount = ActiveWorkbook.Worksheets.Count
End Function
```
### ActiveSheet Index (Function)
```vb
Function CurrentSheetIndex() As Integer
    Dim intCurrentSheet As Integer, intCounter As Integer
    Dim sh As Worksheet
    Dim shName As String
        
    shName = ActiveSheet.Name
    intCounter = 0
    If GetSheetsCount > 0 Then
        
        For Each sh In Worksheets
            If sh.Name = shName Then
                CurrentSheetIndex = intCounter + 1
            Else
                intCounter = intCounter + 1
            End If
        Next sh
    
    Else
    End If
End Function
```
### Activate Next Sheet
```vb
Sub NextSheet()
    If CurrentSheetIndex = GetSheetsCount Then
        ActiveWorkbook.Worksheets(1).Activate
    Else
        ActiveWorkbook.Worksheets(CurrentSheetIndex + 1).Activate
    End If
End Sub
```
### Activate Previous Sheet
```vb
Sub PreviousSheet()
    If CurrentSheetIndex = 1 Then
        ActiveWorkbook.Worksheets(GetSheetsCount).Activate
    Else
        ActiveWorkbook.Worksheets(CurrentSheetIndex - 1).Activate
    End If
End Sub
```