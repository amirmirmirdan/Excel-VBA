Public Sub MTPI_GCashMain()
    Dim oHandler As clsErrorHandler
        Set oHandler = New clsErrorHandler
    Dim colFiles As New Collection, i As Integer
        Set colFiles = ListFileNames(FolderPath)
        i = 1

        For i = 1 To colFiles.Count
            CopyTransactionSheet(colFiles.Item(i))
        Next i
    ' Done Import & Compile. Next is processing the data
End Sub
' GCash

Private Sub ExcelIsBusy(boolUpdate As Boolean)
    If boolUpdate = True Then
        With Application
            .Calculation = xlCalculationManual
            .DisplayAlerts = False
            .DisplayStatusBar = False
            .DisplayFormulaBar = False
            .EnableEvents = False
            .ScreenUpdating = False
        End With
        ActiveSheet.DisplayPageBreaks = False
    Else
        With Application
            .Calculation = xlCalculationAutomatic
            .DisplayAlerts = True
            .DisplayStatusBar = True
            .DisplayFormulaBar = True
            .EnableEvents = True
            .ScreenUpdating = True
        End With
        ActiveSheet.DisplayPageBreaks = False
    End If
End Sub
Sub ClearRangeContent(ws As WorkSheet)
    Dim Rng As Worksheet
    With activesheet.Range("A1").CurrentRegion
        Set Rng = .Offset(1, 0)
        Rng.ClearContent
    End With
End Sub
Function ListFileNames(ByVal strFolderPath As String) As Collection
    ' Declare Variables
    Dim fso As Object
    Dim fsoFolder As Object
    Dim fsoFiles As Object
    Dim colFiles As Collection
        ' Set & Initialize Collection Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fsoFolder = fso.GetFolder(strFolderPath)
        Set colFiles = New Collection
    ' Start
    ' Check <stringPath> to ensure it is not a zero length string before you call this function.
        If fsoFolder.Files.Count > 0 Then
            For Each fsoFiles In fsoFolder.Files
                colFiles.Add Item:=fsoFiles.Path
            Next fsoFiles
        Else
            MsgBox "There's No File in the Folder", vbOK, "Get Files"
        End If
    ' Return Value (Collection)
    Set Files_GetNames = colFiles
End Function
Function FolderPath() As String
    ' Declare Variables
    Dim oFolder As FileDialog
    Set oFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With oFolder
        .AllowMultiSelect = False
        .Title = "Select the Folder with the files"
        .Show
        FolderPath = .SelectedItems(1)
    End With
End Function
Sub DeleteTopRow()
    activesheet.Rows("1:4").Delete
End Sub
Sub CopyWorksheetToNewWorkbook()
    ThisWorkbook.ActiveSheet.Copy _
        Before:=Workbooks.Add.Worksheets(1)
End Sub
Function ArrTransactions(rng As Range, RowsCount As Integer, ColumnsCount As Integer) As Variant()
    Dim TempArray As Variant()
    Dim i As Integer, j As Integer
    ReDim TempArray(0 To RowsCount, 0 To ColumnsCount)
    For i = 0 To UBound(TempArray, 1)
        For j = 0 To UBound(TempArray, 2)
            TempArray(i, j) = rng.Offset(i, j).Value
        Next j
    Next i
    ArrTransactions = TempArray
End Function
Sub AppendSheetData(MyArray() As Variant)
    Dim iRow As Integer
    iRow = MainShRowCount + 1
    With ThisWorkbook.WorkSheet(2)
        For i = 0 To UBound(MyArray, 1)
            For j = 0 To UBound(MyArray, 2)
                .Range("A1").Offset(i+iRow, 0).Value = MyArray(i, j) ' Append the Compile Sheet with Array Values.
            Next j
        Next i
    End With
    iRow = vbNull
End Sub