> Description here.

---

## Sheets Object
### Loop Sheets in ActiveWorkbook 
#### Create Table of Content
```vb
Sub CreateTableOfContents()
    Dim i As Long
    Dim TOCName As String
	'Name of the Table of contents
	TOCName = "TOC"
	'Delete the existing Table of Contents sheet if it exists
	Application.DisplayAlerts = False
		' Function to check if worksheet already exists.
	    If SheetExist(TOCName) Then 
		    ActiveWorkbook.Worksheets(TOCName).Delete
		End If
	Application.DisplayAlerts = True
	
	'Create a new worksheet
	ActiveWorkbook.Sheets.Add before:=ActiveWorkbook.Worksheets(1)
	ActiveSheet.Name = TOCName
	
	'Loop through the worksheets
	For i = 1 To Sheets.Count
	    'Create the table of contents
	    ActiveSheet.Hyperlinks.Add _
	        Anchor:=ActiveSheet.Cells(i, 1), _
	        Address:="", _
	        SubAddress:="'" & Sheets(i).Name & "'!A1", _
	        ScreenTip:=Sheets(i).Name, _
	        TextToDisplay:=Sheets(i).Name
	Next i
End Sub
```

#### Check if Worksheet Exist
```vb
Function SheetExist(shtName As String) As Boolean
    Dim sht As Worksheet
        
    For Each sht In ActiveWorkbook.Worksheets
    
        If sht.Name = shtName Then
            SheetExist = True
            Exit Function
        End If
    Next sht
    
    SheetExist = False
    
End Function
```

## Properties of a Cell Object

### Get Color Code From Cell Fill
```vb
Sub GetColorCodeFromCellFill()

'Create variables hold the color data
Dim fillColor As Long
Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim Hex As String

'Get the fill color
fillColor = ActiveCell.Interior.Color

'Convert fill color to RGB
R = (fillColor Mod 256)
G = (fillColor \ 256) Mod 256
B = (fillColor \ 65536) Mod 256

'Convert fill color to Hex
Hex = "#" & Application.WorksheetFunction.Dec2Hex(fillColor)

'Display fill color codes
MsgBox "Color codes for active cell" & vbNewLine & _
    "R:" & R & ", G:" & G & ", B:" & B & vbNewLine & _
    "Hex: " & Hex, Title:="Color Codes"

End Sub
```

HEX: # 40B100
R 0
G 177
B 64


