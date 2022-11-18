[[Base|Home]]

---
#### Unhide All Rows & Columns
```vb
Sub UnhideRowsColumns()
    Columns.EntireColumn.Hidden = False
    Rows.EntireRow.Hidden = False
End Sub
```
#### Find Last Row (Function)
```vb
Function fxLastRow() As Long
    With ActiveSheet.UsedRange
        fxLastRow = .Rows.Count
    End With
End Function
```
#### Delete Empty Rows
```vb
Sub DeleteEmptyRows()
Dim l As Long
    For l = fxLastRow To 1 Step -1 ' This will loop from the lastrow, and up.
        If Application.WorksheetFunction.CountA(Rows(l)) = 0 Then
            Rows(l).Delete
        End If
    Next l
End Sub
```
