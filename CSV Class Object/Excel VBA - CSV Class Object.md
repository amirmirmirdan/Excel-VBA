[README](https://github.com/amirmirmirdan/Excel-VBA#project-list) | [Project List](https://github.com/amirmirmirdan/Excel-VBA#project-list) 

---

# CSV Class Object
> Focus on the object methods first & along the way, you'll figure out the required object properties.
> The CSV Class Object documents the methods & functions used in order to assist in Importing & Exporting Data.  

## Topic
- [Installation](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#installation)
- [Usage](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Usage)
- [Documentation](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Documentation)
	- [Example Code](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#example-code) 	
- [Configuration](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Configuration)
- [Roadmap](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Roadmap)
- [Reference](https://github.com/amirmirmirdan/Excel-VBA/edit/main/CSV%20Class%20Object/Excel%20VBA%20-%20CSV%20Class%20Object.md#Reference)

---

## Installation

## Usage

## Documentation
### Properties
1. **MainSheet** As _WorkSheet_
2. f

### Event
1. **Class_Initialize**
2. **Class_Terminate**

### Methods
1. m
2. m
3. m
4. m

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

##### Event: Class Initialize
```vb
Private Sub Class_Initialize()
    Set MainSheet = ActiveSheet
End Sub
```
##### Event: Class Terminate
```vb
Private Sub Class_Terminate()
        Set MainSheet = Nothing
End Sub
```

#### Methods: Import Sheet
```vb
Public Sub ImportSheet()
    ' Declaring variables
        Dim ws As Worksheet
        Dim TargetBook As Workbook
        Dim StrFile As String
    
    ' Define Variables
            Set ws = MainSheet
            
    ' this enables user to determine which file they would like to copy to current sheet.
            StrFile = Application.GetOpenFilename()
            
            ' Avoid cause error If user _
                Close the File Picker Dialog or _
                Click on cancel
            If WorksheetFunction.IsText(StrFile) = True Then
                Set TargetBook = Workbooks.Open(StrFile)
            Else
                Exit Sub
            End If
        
    ' Procedure starts here
            TargetBook.Sheets(1).Copy After:=ws
            TargetBook.Close SaveChanges:=False
    
    ' Clearing VBA Memory
            Set ws = Nothing
            Set TargetBook = Nothing
            StrFile = vbNull

End Sub
```

#### Method: Import to Range
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
