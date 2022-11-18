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
