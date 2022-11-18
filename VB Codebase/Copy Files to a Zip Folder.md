### TOC
1. [[Copy Files to a Zip Folder#Get Folder Path Function|Get Folder Path (Function)]] ^d4f85a
2. [[Copy Files to a Zip Folder#List Files in a Folder Function|List Files in a Folder (Function)]] ^6e0e68
3. [[Copy Files to a Zip Folder#Create an Empty Zip Folder|Create an Empty Zip Folder]]
4. [[Copy Files to a Zip Folder#Copy Files into a Zip Folder|Copy Files into a Zip Folder]]

---
### Get Folder Path (Function)
---
#### Description
**Purpose:** Searches a directory/folder provided and adds the file names to a collection
**Inputs:** StrPath
**Type:** String
**Desc:** The pathway to the folder / directory
**Return:** Object (Collection)

#### Example:
```vb
Function FolderPath() As String
' Declare Variables
	Dim oFolder As FileDialog
    Set oFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With oFolder
        .AllowMultiSelect = False
        .Title = "Select a Folder"
        .Show
        FolderPath = .SelectedItems(1)
    End With 
End Function
```

^3556a9

### List Files in a Folder (Function)
---
#### Description
**Purpose:** Searches a directory/folder provided and adds the file names to a collection
**Inputs:** StrPath
**Type:** String
**Desc:** The pathway to the folder / directory
**Return:** Object (Collection)

#### Example
```vb
Function Files_GetNames(ByVal StrPath As String) As Collection
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

^4465eb

---

### Create an Empty Zip Folder
---
```vb
Private Sub CreateZipFile(ByVal strZipFolderName As String)
    'Select Folder to ZIP the CSV File
    Dim FolderName As Variant
    Dim FileFilter As String
    Dim Title As String
        
    FileFilter = "Zip Files (*.zip), *.zip"
    Title = "Please select a location and file name for ZIP File"
    FolderName = Application.GetSaveAsFilename(strZipFolderName, FileFilter, , Title)
        
    'Open a Empty ZIP
    Open FolderName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
        
    Set FolderName = Nothing
    FileFilter = ""
    Title = ""
End Sub
```
---
### Copy Files into a Zip Folder
```vb
'Copy the files & folders into the zip file
Set ShellApp = CreateObject("Shell.Application")
ShellApp.Namespace(zippedFileFullName).CopyHere ShellApp.Namespace(folderToZipPath).items

'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
On Error Resume Next
Do Until ShellApp.Namespace(zippedFileFullName).items.Count = ShellApp.Namespace(folderToZipPath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0
```
