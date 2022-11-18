# Description
The workbook object is a member of the [[Workbooks]] collection. The workbooks collection contains all the workbook object currently open [^2]

## Contents
### Event
- [[VBA Workbook Object#Workbook Save|Workbook_BeforeSave Event]]
- Workbook_Open
- Workbook_Activate
- Workbook_Close
- Etc

### Methods
- [[VBA Workbook Object#Workbook Save|Save]]
- [[VBA Workbook Object#Workbook SaveAs|SaveAs]]
- [[VBA Workbook Object#Open a Workbook|Open]]
- [[Add new WorkbookQuery Object| Queries Add]]
- Close
- Close and Save
- SaveCopy
- etc

### Property
1. [[VBA Workbook Object#Workbook Name|Name]]
2. [[VBA Workbook Object#Workbook Index|Index]]
3. Count

---
### Workbook Save
To save a workbook and mark it as *saved*, use the following code.
```vb
ActiveWorkbook.Save
```
Below code will save all open workbooks and then closes Microsoft Excel.
```vb
For Each w In Application.Workbooks 
    w.Save 
Next w 
Application.Quit
```
However, if it's the first time you save the workbook, use the [[SaveAs]] method below to specify the name for the Excel file. [^3]

---
### Workbook SaveAs
Save changes to the workbook in a different file.

**Syntax**
*Expression*.**SaveAs** (*FileName, FileFormat, ~~Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConfilctResolution, AddToMru, TextCodepage, TextVisualLayout, Local*)~~


**Parameters**
To simplify this note, it will only elaborate the *FileName* & *FileFormat* Parameters for the **SaveAs** method as all of the arguements are optional. For more extensive explaination on the other Parameters, please refer to the [Microsoft Documentation: Workbook.SaveAs method (Excel).](https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.saveas)

**FileName:** A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.

**FileFormat:** The file format to use when you save the file. For a list of valid choices, see the [[XlFileFormat Enumeration|XlFileFormat]] enumeration. For an existing file, the default format is the last file format specified; for a new file, the default is the format of the version of Excel being used.

---
**Example:** This example creates a new workbook, prompts the user for a file name, and then saves the workbook.
```vb
Set NewBook = Workbooks.Add 
Do 
    fName = Application.GetSaveAsFilename 
Loop Until fName <> False 
NewBook.SaveAs Filename:=fName
```
---
### Open a Workbook
We can access any open workbook using the code below:
```vb
Workbooks("C:\Example.xlsm").Open 
```
*You may change the quoted part inside the bracket to reference the workbook file that suits your needs.*

**Example:** This example will copy a csv file to the current Active Workbook. 
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
It will prompt the user to select the csv file using the [[GetOpenFileName]] Dialog, to get the file name and pass it to the Workbook.Open method to open the workbook.

Once it is open, it will copy a WorkSheet to the current Workbook & then closes the CSV file without saving any changes.

---
### Workbook Name
The Name property will return the workbook name.
```vb
Function WorkbookName() As String
	WorkbookName = ActiveWorkbook.Name
End Function
```
You **cannot** set the name by using this property; if you need to change the name, use the [[VBA Workbook Object#Workbook SaveAs|SaveAs]] method to save the workbook under a different name.

---
### Workbook Index
You can also use an Index number with **Workbooks()**. The index refers to the order the Workbook was open or created. **Workbooks(1)** refers to the workbook that was opened first. **Workbooks(2)** refers to the workbook that was opened second and so on.
```vb
Workbooks(1).Activate
```
---



---
### Footnote

[^2]: Reference [Workbook Object(Excel).](https://docs.microsoft.com/en-us/office/vba/api/excel.workbook)
[^3]: Reference: [Workbook.Save method.](https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.save)

[[Base|Home]]