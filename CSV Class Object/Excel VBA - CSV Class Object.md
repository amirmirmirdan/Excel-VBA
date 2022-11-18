### Basic Module
```vb
Option Explicit
Sub ExportARange()
	Dim RangeToExport As Range
	Set RangeToExport = ActiveWindow.RangeSelection
	
	If Application.CountA(RangeToExport) = 0 Then
		MsgBox "The selection is empty."
		Exit Sub
	End If
	
	Dim CSVFile As New CSVFileClass
	On Error Resume Next
	With CSVFile
		.ExportRange = RangeToExport
		.Export CSVFileName:=Application.DefaultFilePath & "\temp.csv"
		If Err <> 0 Then MsgBox "Cannot export" & Application.DefaultFilePath & "\temp.csv"
	End With
End Sub

Sub ImportAFile()
	Dim CSVFile As New CSVFileClass
	
	On Error Resume Next
	With CSVFile
		.ImportRange = ActiveCell
		.Import CSVFileName:=Application.DefaultFilePath & "\temp.csv"
		If Err <> 0 Then MsgBox "Cannot import " & Application.DefaultFilePath & "\temp.csv"
	End With
	
End Sub
```
created on 2022-11-18 at 20:39 #VBA 

---
### Class Module

```vb

Option Explicit
'CSVFileClass
'''''''''''''
'PROPERTIES
'  ExportRange
'  ImportRange

'METHOD
'  Import
'  Export

Private RangeToExport As Range
Private ImportToCell As Range

Property Get ExportRange() As Range
    Set ExportRange = RangeToExport
End Property

Property Let ExportRange(rng As Range)
    Set RangeToExport = rng
End Property

Property Get ImportRange() As Range
    Set ImportRange = ImportToCell
End Property

Property Let ImportRange(rng As Range)
    Set ImportToCell = rng
End Property

Sub Export(CSVFileName)
'   Exports a range to CSV file
    Dim ExpBook As Workbook

    If RangeToExport Is Nothing Then
        MsgBox "ExportRange not specified"
        Exit Sub
    End If

    On Error GoTo ErrHandle
    Application.ScreenUpdating = False
    Set ExpBook = Workbooks.Add(xlWorksheet)
    RangeToExport.Copy
    Application.DisplayAlerts = False
    With ExpBook
        .Sheets(1).Paste
        .SaveAs FileName:=CSVFileName, FileFormat:=xlCSV
        .Close SaveChanges:=False
    End With
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrHandle:
    ExpBook.Close SaveChanges:=False
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err & vbCrLf & vbCrLf & Error(Err), _
      vbCritical, "Export Method Error"
End Sub

Sub Import(CSVFileName)
'   Imports a CSV file to a range
    Dim CSVFile As Workbook
    
    If ImportToCell Is Nothing Then
        MsgBox "ImportRange not specified"
        Exit Sub
    End If
    
    If CSVFileName = "" Then
        MsgBox "Import FileName not specified"
        Exit Sub
    End If
    
	On Error GoTo ErrHandle
	    Application.ScreenUpdating = False
	    Application.DisplayAlerts = False
	    Workbooks.Open CSVFileName
	    
		Set CSVFile = ActiveWorkbook
	    ActiveSheet.UsedRange.Copy Destination:=ImportToCell
	    CSVFile.Close SaveChanges:=False
	    Application.ScreenUpdating = True
	    Application.DisplayAlerts = True
	    Exit Sub
	ErrHandle:
	    CSVFile.Close SaveChanges:=False
	    Application.ScreenUpdating = True
	    Application.DisplayAlerts = True
	    MsgBox "Error " & Err & vbCrLf & vbCrLf & Error(Err), _
	      vbCritical, "Import Method Error"
End Sub

```
created on 2022-11-18 at 20:42 #VBA 

---

### Class Object
Focus on the object methods first & along the way, you'll figure out the required object properties.

#### Method - Import Sheet
```vb
Public Sub ImportSheet()
	' Declaring variables
		Dim ws As Worksheet
	    Dim strFile As String, strFilter As String, strCaption As String
		Dim TargetBook As Workbook
	' Define Variables
		Set ws = ActiveSheet
		strFilter = "Text Files (*.prn ; *.txt ; *.csv)"
		strCaption = "Please Select a file to import."
	' this enables user to determine which file they would like to copy to current sheet.
		strFile = Application.GetOpenFilename(strFilter, , StrCaption)
			If strFile = "" Then Exit Sub
		Set TargetBook = Workbooks.Open(strFile)
	' Procedure starts here
		TargetBook.Sheets(1).Copy After:=ws
		TargetBook.Close SaveChanges:=False
	'Clearing VBA Memory
		Set ws = Nothing
		Set TargetBook = Nothing
		strFile = vbNull
		strFilter = vbNull
		strCaption = vbNull
End Sub
```
created on 2022-11-18 at 20:46 #VBA 

#### Method - Import to Range
```vb
Public Sub ImportRange(Optional By Ref vImportToRange As Range)	
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
created on 2022-11-18 at 21:25 #VBA 



### Basic Module
```vb
'*****************************************************
' Purpose: 
' Inputs: 
' Returns:  
Public Sub SomeSubroutineName()
	
	' Write code here'
	
End Sub
```
created on 2022-11-18 at 20:51 #VBA 
