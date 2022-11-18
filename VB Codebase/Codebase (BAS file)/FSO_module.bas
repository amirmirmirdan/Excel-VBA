Attribute VB_Name = "FSO_module"
Option Explicit
Public Sub CreateFolders(Optional StrFolderName As String, Optional StrFolderPath As String)
    If StrFolderName = "" Then
        StrFolderName = InputBox("New Folder Name", "Create a Folder")
    Else
    End If

    If StrFolderPath = "" Then StrFolderPath = GetPath

    If PathExist(StrFolderPath & "\" & StrFolderName) Then
        MsgBox "Sorry, that folder already exists"
        Exit Sub
    Else
        MkDir (StrFolderPath & "\" & StrFolderName)
    End If
End Sub
Private Function GetPath() As String
    Dim NameHolder As String
        With Application.FileDialog(msoFileFolderPicker)
            .AllowMultiSelect = False
            .Show
            NameHolder = .SelectedItems(1)
        End With
        GetPath = NameHolder
End Function
Private Function PathExist(StrFolderName As String) As Boolean

    Dim StrFolderPath As String
        StrFolderPath = Dir(StrFolderName, vbDirectory)

    If StrFolderPath <> "" Then PathExist = True
    Else: PathExist = False

End Function
Private Sub CreateFolder(StrFolderName As String)
    MkDir StrFolderName
End Sub
