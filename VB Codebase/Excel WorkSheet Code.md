Return [Base](Base.md) | 

## Properties 
### Name
### Count
Example: Sort WorkSheets in ActiveWorkbook
```vb
Sub SortWorksheets()
    Dim i As Integer
    Dim j As Integer
    Dim iAnswer As VbMsgBoxResult
        iAnswer = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
        & "Clicking No will sort in Descending Order", _
        vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
        For i = 1 To Sheets.Count
            For j = 1 To Sheets.Count - 1
                If iAnswer = vbYes Then
                    If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
                        Sheets(j).Move After:=Sheets(j + 1)
                    End If
                ElseIf iAnswer = vbNo Then
                    If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then Sheets(j).Move After:=Sheets(j + 1)
                    End If
                End If     
            Next j
        Next i
End Sub
```






### Visible
Return either:
1. xlSheetVisible
2. xlSheetHidden
3. xlSheetVeryHidden

Example: Hide All WorkSheet
```vb
Sub HideWorksheet()
    Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
                ws.Visible = xlSheetHidden
            End If
        Next ws
End Sub
```
Example: Unhide All WorkSheets
```vb
Sub UnhideAllWorksheet()
    Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws
End Sub
```


## Methods
### Copy
Example: Copy ActiveSheet.
```vb
Sub DuplicateSheet()
    With ActiveSheet
        .Copy After:=ActiveSheet
    End With
End Sub
```

### Delete 
Example: Delete all WorkSheets (Except ActiveSheet)
```vb
Sub DeleteWorksheets()
    Dim ws As Worksheet
    Dim iAnswer As VbMsgBoxResult
        iAnswer = MsgBox("Delete All Sheets?" & Chr(10) _
        & "Clicking No will stop the subroutine", _
        vbYesNo + vbQuestion + vbDefaultButton1, "Delete WorkSheets")
        
        If iAnswer = vbYes Then
            For Each ws In ActiveWorkbook.Worksheets
                If ws.Name <> ActiveWorkbook.ActiveSheet.Name Then
                    Application.DisplayAlerts = False
                        ws.Delete
                    Application.DisplayAlerts = True
                End If
            Next ws
        Else
            MsgBox "No WorkSheet was deleted"
            Exit Sub
        End If
End Sub
```





---
## Others
### Clear ActiveSheet Contents
```vb
Sub ClearSheetContents()
	ActiveSheet.UsedRange.ClearContents
End Sub
```
### Clear ActiveSheet Formatting
```vb
Sub ClearSheetFormats()
	ActiveSheet.UsedRange.ClearFormats
End Sub
```
### Delete Blank WorkSheets
```vb
Sub deleteBlankWorksheets()
' Check all the worksheets in the active workbook and delete if a worksheet is blank.
    Dim ws As Worksheet
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        For Each ws In Application.Worksheets
            If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
                ws.Delete
            End If
        Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
```
