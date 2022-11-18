[[Base|Home]]

---
### Copy ActiveSheet to New Workbook
```vb
Sub CopyWorksheetToNewWorkbook()
    ThisWorkbook.ActiveSheet.Copy _
        Before:=Workbooks.Add.Worksheets(1)
End Sub
```

### Create a Table of Content Sheet
```vb
Sub TableofContent()
    Dim i As Long
        On Error Resume Next
            Application.DisplayAlerts = False
            Worksheets("Table of Content").Delete
            Application.DisplayAlerts = True
        On Error GoTo 0
            ThisWorkbook.Sheets.Add Before:=ThisWorkbook.Worksheets(1)
            ActiveSheet.Name = "Table of Content"
        For i = 1 To Sheets.Count
            With ActiveSheet
                .Hyperlinks.Add _
                Anchor:=ActiveSheet.Cells(i, 1), _
                Address:="", _
                SubAddress:="'" & Sheets(i).Name & "'!A1", _
                ScreenTip:=Sheets(i).Name, _
                TextToDisplay:=Sheets(i).Name
            End With
        Next i
End Sub
```