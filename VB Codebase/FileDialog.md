## Application.FileDialog property (Excel)
Returns a [[FileDialog]] object representing an instance of the file dialog.[^FileDialog]

**Syntax**

_expression_.**FileDialog** (_fileDialogType_)

_expression_ A variable that represents an [[Base#Application Object|Application]] object.

**Parameters**

| Name | Required/Optional | Data type | Description |
| --- | --- | --- | --- |
| _fileDialogType_ | Required | **[MsoFileDialogType](chrome-extension://pcmpcfapbekmbjjkdalcgopdkipoggdi/office.msofiledialogtype)** | The type of file dialog. |

## Remarks

**MsoFileDialogType** can be one of these constants:

-   **msoFileDialogFilePicker**. Allows user to select a file.
-   **msoFileDialogFolderPicker**. Allows user to select a folder.
-   **msoFileDialogOpen**. Allows user to open a file.
-   **msoFileDialogSaveAs**. Allows user to save a file.

## Example

In this example, Microsoft Excel opens the file dialog allowing the user to select one or more files. After these files are selected, Excel displays the path for each file in a separate message.

```vb
Sub UseFileDialogOpen()
    Dim lngCount As Long 
    ' Open the file dialog 
    With Application.FileDialog(msoFileDialogOpen) 
        .AllowMultiSelect = True 
        .Show 
        ' Display paths of each file selected 
        For lngCount = 1 To .SelectedItems.Count 
            MsgBox .SelectedItems(lngCount) 
        Next lngCount 
    End With  
End Sub
```

Another example, Microdoft Excel opens the file dialog allowing the user to select one or more files. After these files are selected, the function will return a collection of path for each file.[^project]
```vb
Function SelectFilesPath() As Collection
    Dim lngCount As Long 
    Dim col As Collection
    Set col = New Collection
    ' Open the file dialog 
    With Application.FileDialog(msoFileDialogOpen) 
        .AllowMultiSelect = True 
        .Show 
        ' adding the file path to the collection object 
        For lngCount = 1 To .SelectedItems.Count 
            col.AddItem(lngCount) = .SelectedItems(lngCount) 
        Next lngCount 
    End With
    ' returning the function as a collection object.
	SelectFilesPath = col
End Function
```

[[Base|Return]]

[^FileDialog]: [Microsoft Documentation:FileDialog Property](https://docs.microsoft.com/en-us/office/vba/api/excel.application.filedialog)
[^project]: Used in the [[Batch Run FTE Count (Excel VBA)]]
