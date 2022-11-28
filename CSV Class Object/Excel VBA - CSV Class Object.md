[README](https://github.com/amirmirmirdan/Excel-VBA#project-list) | [Project List](https://github.com/amirmirmirdan/Excel-VBA#project-list) | [FileDialog](https://github.com/amirmirmirdan/Excel-VBA/blob/main/CSV%20Class%20Object/FileDialog.md#file-dialog)

---

# CSV Class Object
> Focus on the object methods first & along the way, you'll figure out the required object properties.
> The CSV Class Object documents the methods & functions used in order to assist in Importing & Exporting Data.  

## Contents
- [Installation](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#installation)
- [Usage](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Usage)
- [Documentation](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Documentation)
	- [Example Code](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#example-code) 	
- [Configuration](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Configuration)
- [Roadmap](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Roadmap)
- [Reference](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Reference)

---

## Installation
### Install as an Excel File Add-Ins


## Usage
### Basic Module
Bas File Link.

```vb
Option Explicit
Sub ImportDataToWorkbook()
    Dim CsvHelper As New CSV_class
        CsvHelper.ImportSheet
    Set CsvHelper = Nothing
End Sub
```

## Documentation
| Legend | Remarks |
|---|---|
| 🟢 | Public |
| 📵 | Private |

### Properties
1. 📵 **[MainSheet](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#property-mainsheet)** As _WorkSheet_

### Event
1. 📵 **[Class_Initialize](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#class-initialize)**
2. 📵 **[Class_Terminate](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#class-terminate)**

### Methods
1. 🟢 **Sub**: [ImportSheet](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#import-sheet)
2. 📵 **Function**: [SelectCsvFile_Path](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#selectcsvfile_path-function) () **Return** _String_
3. 📵 **Sub**: [CopyFileSheet](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#copy-file-sheet)
	- SourceFile:= _[SelectCsvFile_Path](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#selectcsvfile_path-function)_
	- Destination:= _[MainSheet](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#property-mainsheet)_
4. 📵 **Sub**: [ExcelBusy](https://github.com/amirmirmirdan/Excel-VBA)

#### Example Code
##### Property MainSheet 
```vb
Private mvSheet As Worksheet
Option Explicit

Private Property Get MainSheet() As Worksheet
    Set MainSheet = mvSheet
End Property

Private Property Set MainSheet(vSheet As Worksheet)
    Set mvSheet = vSheet
End Property
```

##### Class Initialize
```vb
Private Sub Class_Initialize()
    Set MainSheet = ActiveSheet
End Sub
```
##### Class Terminate
```vb
Private Sub Class_Terminate()
        Set MainSheet = Nothing
End Sub
```

#### Import Sheet
```vb
Public Sub ImportSheet()
    ' Declaring variables
    Dim ws As Worksheet
    Dim TargetBook As Workbook
    Dim StrFile As String

    ' Define Variables
    Set ws = MainSheet

    ' this enables user to determine which file they would like to copy to current sheet.
    StrFile = SelectCsvFile_Path()

    ' Avoid cause error If user close the File Picker Dialog _
    or click on cancel
    If WorksheetFunction.IsText(StrFile) <> True Then Exit Sub
    Else: Set TargetBook = Workbooks.Open(StrFile)

    ' Procedure starts here
    CopyFileSheet _
	SourceFile:=TargetBook, _
	DestinationSheet:=ws

    ' Clearing VBA Memory
    Set ws = Nothing
    Set TargetBook = Nothing
    StrFile = vbNull
End Sub
```
#### SelectCsvFile_Path Function
```vb
Private Function SelectCsvFile_Path() As String
    Dim Fd As FileDialog
    Dim Filter As FileDialogFilters
    
    Set Fd = Application.FileDialog(3)  ' An Enum Integer representation of the FileDialog Type Object
    With Fd
        ' Define a Filters object
        Set Filter = .Filters
        
        With Filter
            ' Clear the default filters
            .Clear
            ' Add new filter
            .Add "CSV", "*.csv"
            .Add "All Files", "*.*"
        End With
        
            ' Either to allow or disable the multiselect file.
            .AllowMultiSelect = False
            
            If .Show = False Then
                Exit Function
            End If
        SelectCsvFile_Path = .SelectedItems(1)
    End With
    
    ' Clearing VBA Memory
    Set Fd = Nothing
    Set Filter = Nothing
End Function
```

#### Copy File Sheet
```vb
Private Sub CopyFileSheet(ByRef SourceFile As Workbook, ByRef DestinationSheet As Worksheet)
    ' For reuse in another class method where
    ' - The FileDialogFilePicker.AllowMultiSelect = True, returning an array of file path.
    ' - The idea is to loop openning the workbook &
    ' - call this subroutine and providing the arguements required for the loop.
            SourceFile.Sheets(1).Copy After:=DestinationSheet
            SourceFile.Close SaveChanges:=False
End Sub
```

#### Import to Range
> Note: This has not been tested yet.

```vb
Public Sub ImportRange(Optional ByRef vImportToRange As Range)	
	Dim rng As Range
	Dim TargetBook As Workbook, TargetRange As Range
	Dim strFile As String, strFilter As String, strCaption As String

	If vImportToRange = Nothing Then
		With ActiveWorkbook.Activesheet
			Set rng = .Activecell
		End With
	Else
		Set rng = vImportToRange
	End If
		strFilter = "Text Files (*.prn ; *.txt ; *.csv)"
		strCaption = "Please Select the CSV file to import."
	' this enables user to determine which file they would like to copy to current sheet.
		strFile = Application.GetOpenFilename(strFilter, , strCaption)
			If strFile = "" Then Exit Sub
		Set TargetBook = Workbooks.Open(strFile)
		Set TargetRange = TargetBook.Sheets(1).UsedRange
	' Copy the value from CSV file to the target range
			rng.value = TargetRange.Value
			TargetBook.Close SaveChanges:=False	
	'Clearing VBA Memory
		Set rng = Nothing
		Set TargetRange = Nothing
		Set TargetBook = Nothing
		strFile = vbNull
		strFilter = vbNull
		strCaption = vbNull
End Sub
```

## Configuration

## Roadmap


## Reference
- [x] [Application.FileDialog Issue](https://github.com/amirmirmirdan/Excel-VBA/issues/10#issuecomment-1328392572)
- [x] [FileDialog Object](https://github.com/amirmirmirdan/Excel-VBA/blob/main/CSV%20Class%20Object/FileDialog.md#file-dialog)
- [x] [Better Solutions: FileDialog](https://bettersolutions.com/vba/files-directories/filedialog.htm)
- [x] [Better Solutions: FileDialog Type](https://bettersolutions.com/vba/enumerations/msofiledialogtype.htm)
- [x] [Learn Microsoft: Application FileDialog](https://learn.microsoft.com/en-us/office/vba/api/office.filedialog)
- [x] [Excel How: How to Import Data from Another Workbook in Excel](https://www.excelhow.net/how-to-import-data-from-another-workbook-in-excel.html)

---
