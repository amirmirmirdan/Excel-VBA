Provide Access to a computer's file system.

**Syntax**
_Scripting_.**FileSystemObject**

### Methods
| Method | Description |
| --- | --- |
| [BuildPath](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/buildpath-method) | Appends a name to an existing path. |
| [CopyFile](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/copyfile-method) | Copies one or more files from one location to another. |
| [CopyFolder](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/copyfolder-method) | Copies one or more folders from one location to another. |
| [CreateFolder](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/createfolder-method) | Creates a new folder. |
| [CreateTextFile](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/createtextfile-method) | Creates a text file and returns a TextStream object that can be used to read from, or write to the file. |
| [DeleteFile](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/deletefile-method) | Deletes one or more specified files. |
| [DeleteFolder](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/deletefolder-method) | Deletes one or more specified folders. |
| [DriveExists](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/driveexists-method) | Checks if a specified drive exists. |
| [FileExists](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/fileexists-method) | Checks if a specified file exists. |
| [FolderExists](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/folderexists-method) | Checks if a specified folder exists. |
| [GetAbsolutePathName](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getabsolutepathname-method) | Returns the complete path from the root of the drive for the specified path. |
| [GetBaseName](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getbasename-method) | Returns the base name of a specified file or folder. |
| [GetDrive](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getdrive-method) | Returns a Drive object corresponding to the drive in a specified path. |
| [GetDriveName](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getdrivename-method) | Returns the drive name of a specified path. |
| [GetExtensionName](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getextensionname-method) | Returns the file extension name for the last component in a specified path. |
| [GetFile](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getfile-method) | Returns a File object for a specified path. |
| [GetFileName](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getfilename-method-visual-basic-for-applications) | Returns the file name or folder name for the last component in a specified path. |
| [GetFolder](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getfolder-method) | Returns a Folder object for a specified path. |
| [GetParentFolderName](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getparentfoldername-method) | Returns the name of the parent folder of the last component in a specified path. |
| [GetSpecialFolder](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/getspecialfolder-method) | Returns the path to some of Windows' special folders. |
| [GetTempName](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/gettempname-method) | Returns a randomly generated temporary file or folder. |
| [Move](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/move-method-filesystemobject-object) | Moves a specified file or folder from one location to another. |
| [MoveFile](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/movefile-method) | Moves one or more files from one location to another. |
| [MoveFolder](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/movefolder-method) | Moves one or more folders from one location to another. |
| [OpenAsTextStream](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/openastextstream-method) | Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file. |
| [OpenTextFile](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/opentextfile-method) | Opens a file and returns a TextStream object that can be used to access the file. |
| [WriteLine](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/writeline-method) | Writes a specified string and new-line character to a TextStream file. |

### Properties

| Property | Description |
| --- | --- |
| [Drives](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/drives-property) | Returns a collection of all Drive objects on the computer. |
| [Name](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/name-property-filesystemobject-object) | Sets or returns the name of a specified file or folder. |
| [Path](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/path-property-filesystemobject-object) | Returns the path for a specified file, folder, or drive. |
| [Size](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/size-property-filesystemobject-object) | For files, returns the size, in bytes, of the specified file; for folders, returns the size, in bytes, of all files and subfolders contained in the folder. |
| [Type](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/type-property-filesystemobject-object) | Returns information about the type of a file or folder (for example, for files ending in .TXT, "Text Document" is returned). |


## Example:
Example below is a function that returns a **collection** of files in the `StrFolderPath` folder that was provided in the arguement.
```vb
Function ListFileNames(ByVal StrFolderPath As String) As Collection
' Declare Variables
	Dim fso As Object
	Dim fsoFolder As Object
	Dim fsoFiles As Object
	Dim colFiles As Collection
' Set & Initialize Collection Object
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fsoFolder = fso.GetFolder(StrFolderPath)
	Set colFiles = New Collection
' Start
' Check <stringPath> to ensure it is not a zero length string before you call this function.
    If fsoFolder.Files.Count > 0 Then
        For Each fsoFiles In fsoFolder.Files
            colFiles.Add Item:=fsoFiles.Path
        Next fsoFiles
    Else
	    MsgBox "There's No File in the Folder", vbOk, "Get Files" 
    End If
	' Return Value (Collection)
    Set ListFileNames = colFiles
End Function
```
