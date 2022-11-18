# Best Practice
## Naming Convention
Follow these syntax/structure to name variables, functions and/or subroutines.

> *scope* `DataType`*_VariableName*`(Qualifier)`

![[NamingVariable.png]][^log20220907]

[^log20220907]:Log [20220907](20220907%20(WED).md)

---
## Turn off Screen updating
Use [[IsExcelBusy Function|IsExcelBusy]] Function at the start & end of a subroutine. For example, put the below code, at the beginning of the subroutine;
```vb
	IsExcelBusy True
```
Above code will turn the updating off. It will turn back on when 

```vb
Public Sub IsExcelBusy(boolUpdate As Boolean)
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








# Current
## Note: Custom Add-ins

### Currently Used: Quick Access Tool Bar.
1. [[BDO Bank Stmt Download (PH)]]
2. UOB **(SG)**
3. GPAY 6421
4. GPAY 0699 (INSTAPAY)
5. Duplicate Sheet
6. Import csv File
7. Set Table
8. Toggle Style Reference

### Using: Ribbon
---

#### Group: ActiveWorkbook
1. Import csv file
2. 


#### Group: Sheets
1. Duplicate Sheets
2. 

#### Group: FileSystemObject
1. Create Zip Folder
2. Copy Selected Files to Zip Folder
3. List Files in Selected Folder

#### Group: BDO Bank (PH)
1. BDO bank stmt (massaging)
2. GPAY 6421 Tagging
3. GPAY 0699 Tagging

#### Group: UOB Bank (SG)
1. UOB bank stmt (massaging)

#### Group: Modes & Playlists
1. Toggle Reference Style (R1C1/A1)
2. Return to normal

#### Group: Oracle


















# Frequently Used
## Scope
1. Excel Objects
2. Modules
3. **[[VBA Codebase#Class Modules|Class Modules]]**
	1. Event
	2. Property
	3. Metthods
4. Forms


## Modules
### Subroutine


### Function (f)
#### 1: SheetIsBlank
```vb
Function SheetIsBlank() As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error Resume Next
    Set wb = ActiveWorkbook
        wb.Activate
    Set ws = wb.ActiveSheet
    With ws
        If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
            SheetIsBlank = True
        Else
            SheetIsBlank = False
        End If
    End With
End Function
```
#### 2: List Files in a Folder
```vb
Public Function Files_GetNames(ByVal StrPath As String) As Collection
' Declare Variables
Dim fso As New FileSystemObject
Dim fsoFolder As Folder
Dim fsoFile As File
Dim collFiles As Collection
    
' Set & Initialize Collection Object
    Set collFiles = New Collection
    Set fsoFolder = fso.GetFolder(StrPath & "\")
    
' Start
' Check <stringPath> to ensure it is not a zero length string before you call this function.
    If fsoFolder.Files.Count > 0 Then        
        For Each fsoFile In fsoFolder.Files
            collFiles.Add Item:=fsoFile.Path
        Next fsoFile
    Else
    ' do nothing
    End If
    
    ' Return Value (Collection)
    Set Files_GetNames = collFiles
End Function
```


#### Array (arr)
##### Return Array from a function
```vb
Public Function MyFunction() As String()


```


##### Array as Arguement
```vb
Public Sub ArrayHeaderName(ByRef arValues() As String)








```






## Class Modules (cls)
### Event Sub
```vb
Private Sub Class_Initialize()    
	ExcelIsBusy True
End Sub
```

```vb
Private Sub Class_Terminate()    
	ExcelIsBusy False
End Sub
```

^638959

### Methods
#### 1: ExcelIsBusy


#### 2: ClearSheetContents
```vb
Public Sub ClearSheetContents()
	ActiveSheet.UsedRange.ClearContents
End Sub
```
To Clear only the formatting for the ActiveSheet.
```vb
Public Sub ClearSheetFormatting()
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
End Sub
```
#### 3: fxFiles_GetNames() As Collection
```vb
Public Function Files_GetNames(ByVal StrPath As String) As Collection
' Declare Variables
Dim fso As New FileSystemObject
Dim fsoFolder As Folder
Dim fsoFile As File
Dim collFiles As Collection
    
' Set & Initialize Collection Object
    Set collFiles = New Collection
    Set fsoFolder = fso.GetFolder(StrPath & "\")
    
' Start
' Check <stringPath> to ensure it is not a zero length string before you call this function.
    If fsoFolder.Files.Count > 0 Then        
        For Each fsoFile In fsoFolder.Files
            collFiles.Add Item:=fsoFile.Path
        Next fsoFile
    Else
    ' do nothing
    End If
    
    ' Return Value (Collection)
    Set Files_GetNames = collFiles
End Function

```

#### 4: ImportFile
```vb
Public Sub ImportFile(vStrFilePath As String)
Dim OpenBook As Workbook
    Set OpenBook = Application.Workbooks.Open(vStrFilePath)
    
    ' Copy Report 21GL
    With OpenBook.Sheets(1).Range("A1")
        .Activate
        .CurrentRegion.Copy
    End With
    
    ' Paste Report 21GL to Data Sheet in Pvt Count Workbook
    ThisWorkbook.Worksheets("RPT21").Range("A1").PasteSpecial xlPasteValues
    
    ' AdHoc: Logging File Properties to Class
    With OpenBook
        mFullName = .FullName
        mName = .Name
        mPath = .Path
    ' Close Report 21 without saving changes
        .Application.CutCopyMode = False
        .Close False
    End With
    Set OpenBook = Nothing
End Sub
```

#### 5: CreateTableObject
```vb
Public Sub CreateTableObject(vRange As Range, vStrTableName As String)
    Dim Rng As Range
        Set Rng = vRange.CurrentRegion
        With ActiveSheet
            .ListObjects.Add(xlSrcRange, Rng, , xlYes).Name = vStrTableName
            .ListObjects(vStrTableName).TableStyle = ""
        End With
End Sub
```


## Advanced Filter

- [ ] [[Advanced Filtering]]
