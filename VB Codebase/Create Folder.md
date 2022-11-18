### Create Folders
```vb
Sub RangeFolderName()
	Dim LRow As Long
	LRow = ActiveSheet.UsedRange.Rows.Count
	Dim i As Long
	i = 0
	Dim strFolder As String
	i = 1    
		For i = 1 To LRow
			strFolder = Cells(i, 1).Value
			MkDir strFolder
		Next i
	MsgBox "Done! " + LRow + " was created successfully"
End Sub
```
---
### Import a CSV File
```vb
Sub CopySheetToCurrentWorkbook()
    Dim ws As Worksheet
        Set ws = ActiveSheet
        
    Dim strFile As String
        strFile = Application.GetOpenFilename()
        ' this enables user to determine which file they would like to copy to current sheet.
    
    Dim closedbook As Workbook
        Set closedbook = Workbooks.Open(strFile)
        closedbook.Sheets(1).Copy After:=ws
        closedbook.Close SaveChanges:=False
        Set closedbook = Nothing
End Sub
```
---

---
Related Codes:
- [[Application]]
- [[Excel WorkSheet Code|WorkSheets]]
- 