## Class: clsFilterCopy
### Process
1. import CSV
2. Massage Data
3. Copy Criteria Data based on selected budget code
4. Copy Bank Stmt Data based on Criteria Selected.
5. 

### Class Event
#### Initialise
```vb
Private Sub Class_Initialise()    
	ExcelIsBusy True
End Sub
```
#### Terminate
```vb
Private Sub Class_Terminate()    
	ExcelIsBusy False
End Sub
```
### Public Sub
#### ExcelIsBusy
```vb
Public Sub ExcelIsBusy(boolUpdate As Boolean)
    If boolUpdate = True Then
        With Application
            .Calculation = xlCalculationManual
            .DisplayAlerts = False
            .DisplayStatusBar = False
            .DisplayFormulaBar = False
            .EnableEvents = False
            .ScreenUpdating = False
        End With
        ActiveSheet.DisplayPageBreaks = False
    Else
        With Application
            .Calculation = xlCalculationAutomatic
            .DisplayAlerts = True
            .DisplayStatusBar = True
            .DisplayFormulaBar = True
            .EnableEvents = True
            .ScreenUpdating = True
        End With
        ActiveSheet.DisplayPageBreaks = False
    End If
End Sub
```

#### Clear Data Only, Remain Hearder
```vb

```

#### Import Bank Stmt CSV
```vb
Sub Import_AccountStmtCSV()
    Dim wb As Workbook
        Set wb = ActiveWorkbook
    
    Dim ws As Worksheet
        Set ws = wb.ActiveSheet
        
Dim strFilePath As String
strFilePath = Application.GetOpenFilename()
    If strFilePath = "" Then
        MsgBox "Let me know when you need me!"
    Else
        Dim TargetBook As Workbook
            Set TargetBook = Workbooks.Open(strFilePath)
            closedbook.Sheets(1).Copy Before:=ws
            closedbook.Close SaveChanges:=False
    End If
Set wb = Nothing
Set ws = Nothing
strFilePath = vbNull
Set TargetBook = Nothing
End Sub
```
#### Filter Copy
```vb
Sub FilterCriteria(rgData As Range, rgCriteria As Range, rgCopyTo As Range)
    With rgData
        .AdvancedFilter xlFilterCopy, rgCriteria, rgCopyTo, False
    End With
End Sub
```

#### Delete Empty Rows
```vb
Public Sub DeleteEmptyRows()
    For r = LastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(Rows(r)) = 0 Then
            Rows(r).Delete
        End If
    Next r
End Sub
```
#### Get Last Row
```vb
Public Property Get LastRow() As Long
    LastRow = ActiveSheet.UsedRange.Rows.Count + _
                   ActiveSheet.UsedRange.Rows(1).Row - 1
End Property
```





# Archive
## Ref
### Clear Data while remain Header Rows
```vb 
Sub ClearData()
Dim rg As Range
	Set rg = Activesheet.Range("A1").CurrentRegion.Offset(1, 0)
	rg.ClearContents
	Set rg = Nothing
End Sub
```

### Create Empty Zip File
```vb
Sub CreateZipFile()
'Select Folder to ZIP the CSV File
Dim FolderName As Variant
Dim InitialFIleName As String
Dim FileFilter As String
Dim Title As String
    InitialFIleName = "TestZip"
    FileFilter = "Zip Files (*.zip), *.zip"
    Title = "Please select a location and file name for ZIP File"
FolderName = Application.GetSaveAsFilename(InitialFIleName, FileFilter, , Title)
    'Open a Empty ZIP
    Open FolderName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
Set FolderName = Nothing
InitialFIleName = ""
FileFilter = ""
Title = ""
End Sub
```

### Refresh All Pivot Table in the Workbook
```vb
Sub vba_referesh_all_pivots()
Dim pt As PivotTable
	For Each pt In ActiveWorkbook.PivotTables
		pt.RefreshTable
	Next pt
End Sub
```
### Unhide All Rows & Columns
*Use for the GPAY 7-Eleven Settlement Report*
```vb
Sub UnhideRowsColumns()
Columns.EntireColumn.Hidden = False
Rows.EntireRow.Hidden = False
End Sub
```
### Convert all to Values
```vb
Sub convertToValues()
Dim MyRange As Range
Dim MyCell As Range
	Select Case _
		MsgBox("You Can't Undo This Action. " _
		& "Save Workbook First?", vbYesNoCancel, _
		"Alert")
	Case Is = vbYes
		ThisWorkbook.Save
	Case Is = vbCancel
		Exit Sub
	End Select
        
	Set MyRange = ActiveSheet.UsedRange
	For Each MyCell In MyRange
		If MyCell.HasFormula Then
			MyCell.Formula = MyCell.Value
		End If
	Next MyCell
End Sub
```

### Clear WorkSheet
```vb
Sub Worksheet_Clear(wksWorksheet As Worksheet)
With wksWorksheet.Cells
    .ClearContents
    .EntireColumn.Hidden = False
    .EntireRow.Hidden = False
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .IndentLevel = 0
    .ShrinkToFit = False
    .MergeCells = False
    .Font.Bold = False
    .Font.Italic = False
    .Font.Name = "Arial"
    .Font.Size = 12
    .Font.Color = RGB(0, 0, 0)
    .Font.Strikethrough = False
    .Font.Subscript = False
    .Font.Superscript = False
    .Font.Underline = False
    .Interior.ColorIndex = xlColorIndexNone
    .NumberFormat = xlGeneral
End With
End Sub
```

---

## Notes
For GPAY 6421 Settlement Bank...
Add the following:
1. Convert ListObject To Range
2. Duplicate Sheet
3. Delete Specific Range Contents
4. Move Sheet to new Workbook
5. Save New Workbook as CSV



### Workbook Path & Name
```vb
Function Path()As String
Dim vStr As String
vStr = Activeworkbook.Path
Path = vStr
End Function
```
.
```vb
Function Name()As String
Dim vStr As String
vStr = Activeworkbook.Name
Name = vStr
End Function
```
# Custom
## Module

### mod1
```vb
Option Explicit

Sub Massage_UOB()
    Dim oBnkStmt As New BankBranchClass
        oBnkStmt.MassageBnkStmt "UOB"
    Set oBnkStmt = Nothing
End Sub
Sub Massage_BDO()
    Dim oBnkStmt As New BankBranchClass
        oBnkStmt.MassageBnkStmt "BDO"
    Set oBnkStmt = Nothing
End Sub

Sub TEst()

        Sheets("BNK STMT").Range("GTH[#All]").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Range("Criteria"), _
        CopyToRange:=Range("CopyToDestination"), _
        Unique:=False
        
        ActiveSheet.UsedRange.EntireColumn.AutoFit
        
End Sub

Sub CreateZipFile()
'Select Folder to ZIP the CSV File
Dim FolderName As Variant
Dim InitialFIleName As String
Dim FileFilter As String
Dim Title As String
    InitialFIleName = "TestZip"
    FileFilter = "Zip Files (*.zip), *.zip"
    Title = "Please select a location and file name for ZIP File"
FolderName = Application.GetSaveAsFilename(InitialFIleName, FileFilter, , Title)
    'Open a Empty ZIP
    Open FolderName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
Set FolderName = Nothing
InitialFIleName = ""
FileFilter = ""
Title = ""
End Sub

```




## Class
### Bnk Stmt
```vb
Option Explicit
Sub MassageBnkStmt(strBnkBranch As String)
    
    Select Case strBnkBranch
    Case "BDO"
        CleanTrans_History
        
    Case "UOB"
        CleanAccount_Stmt
        
    Case Else
        MsgBox "SORRY, DONT UNDERSTAND HUMAN LANGUAGE :["
    
    End Select

End Sub
```

```vb
Private Sub CleanTrans_History()
    Dim i As Long, j As Long
        i = vbNull
        j = vbNull
    Dim arrHeader(1 To 7) As String
        arrHeader(1) = "BUDGET CODE | DESCRIPTION"
        arrHeader(2) = "TRANSACTION REFERENCE"
        arrHeader(3) = "BLANK"
        arrHeader(4) = "TRANSACTION SEQ"
        arrHeader(5) = "RECEIPT NUMBER"
        arrHeader(6) = "ORACLE DOC NUMBER"
        arrHeader(7) = "ADDITIONAL COMMENT"

        ActiveSheet.Rows("1:3").Delete
        For i = 1 To 7
            j = 7
            ActiveSheet.Cells(1, j + i).Value = arrHeader(i)
        Next i
        
        i = vbNull
        j = vbNull

        With ActiveSheet
            .Columns("A:A").ColumnWidth = 12
            .Columns("D:F").Style = "Comma"
            .Columns("B:N").EntireColumn.AutoFit
            .Columns("C:C").ColumnWidth = 50
        End With
End Sub
```

```vb
Private Sub CleanAccount_Stmt()
    
    With ActiveSheet.UsedRange
        .ClearFormats
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .MergeCells = False
        .Orientation = 0
        .AddIndent = False
        .ReadingOrder = xlContext
    End With
    
    With ActiveSheet
        .Columns("C:E").ColumnWidth = 12
        .Rows("1:3").Delete shift:=xlUp
        .Columns(7).NumberFormat = "0"
        .Columns(11).NumberFormat = "0"
    End With
    
    Dim iClmn As Integer
        iClmn = 18
    
    With ActiveWorkbook
        For iClmn = 18 To 20
            ReplaceZeros_AmountColumn (iClmn)
        Next iClmn
        .ActiveSheet.Columns("L:Q").Columns.Group
        .ActiveSheet.Columns("I:J").Columns.Group
    End With

End Sub
```

```vb
Private Sub ReplaceZeros_AmountColumn(iColumn As Integer)
    With ActiveSheet.Columns(iColumn)
		.Replace _
			What:="0", _
			Replacement:="", _
			LookAt:=xlWhole, _
			SearchOrder:=xlByRows, _
			MatchCase:=False, _
			SearchFormat:=False, _
			ReplaceFormat:=False, _
			FormulaVersion:=xlReplaceFormula2
		.Style = "Comma"
    End With
End Sub
```

```vb
Private Sub RenameSheets(strNew_SheetName As String)
	Dim ws As Worksheet
	Set ws = ActiveSheet
	With ActiveWorkbook
		.Sheets(ws.Index + 1).Name = strNew_SheetName
	End With
End Sub
```

```vb

Private Sub CreateListObj(strTblName As String)    
    With ActiveSheet
        .ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = _
        strTblName
        .ListObjects(strTblName).TableStyle = ""
    End With
End Sub
```

```vb
Sub Import_AccountStmtCSV()
    Dim wb As Workbook
        Set wb = ActiveWorkbook
    
    Dim ws As Worksheet
        Set ws = wb.ActiveSheet     
	Dim strFilePath As String
	strFilePath = Application.GetOpenFilename()
    If strFilePath = "" Then
        MsgBox "Let me know when you need me!"
    Else
        Dim TargetBook As Workbook
            Set TargetBook = Workbooks.Open(strFilePath)
            closedbook.Sheets(1).Copy Before:=ws
            closedbook.Close SaveChanges:=False
    End If
	Set wb = Nothing
	Set ws = Nothing
	strFilePath = vbNull
	Set TargetBook = Nothing
End Sub
```



### Import Sheet
```vb
Public Sub ImportExternalSheet(IndexSheet As Long) ' Index 1 for SME, Index 2 for IPAY
    Set mWB = ActiveWorkbook
    Set mWS = mWB.ActiveSheet
        
        mSTR = Application.GetOpenFilename()
            Dim closedbook As Workbook
                Set closedbook = Workbooks.Open(mSTR)
                closedbook.Sheets(IndexSheet).Copy After:=mWS
                closedbook.Close SaveChanges:=False
    
    Set closedbook = Nothing
    Set mWS = Nothing
    Set mWB = Nothing
End Sub
```






# Links
- [ ] [[UserForm]]
- [ ] Modules
- [ ] Class Module
- [ ] Pivot Playlists 