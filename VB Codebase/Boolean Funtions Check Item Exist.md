### TOC
- [ ] [[Boolean Funtions Check Item Exist#Folder Exist|Folder Exist]]
- [ ] File
- [ ] Workbook
- [ ] Worksheet
- [ ] Pivot Table
- [ ] 

### Drive Exist
```vb
Function DriveExist(StrDrive As String) As Boolean
    If Dir(StrDrive, vbDirectory) <> "" Then DriveExist = True
    Else: DriveExist = False
End Function
```

### Folder Exist
```vb
Function FolderExist(strFolderPath As String) Boolean
	Dim strDir As String
	strDir = Dir(strFolderPath, vbDirectory)
	
	If strDir <> "" Then FolderExist = True
	Else: FolderExist = False
End Function
```

### File Exist
```vb
Function FileExist(strFilePath As String) Boolean
	Dim strDir As String
	strDir = Dir(strFilePath)
	
	If strDir <> "" Then FileExist = True
	Else: FileExist = False
End Function
```

