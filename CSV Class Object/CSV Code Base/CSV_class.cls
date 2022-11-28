VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSV_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvSheet As Worksheet

Option Explicit
Private Property Get MainSheet() As Worksheet
    Set MainSheet = mvSheet
End Property
Private Property Set MainSheet(vSheet As Worksheet)
    Set mvSheet = vSheet
End Property

Private Sub Class_Initialize()
    Set MainSheet = ActiveSheet
    ExcelBusy True
End Sub

Public Sub ImportSheet()
    ' Declaring variables
    Dim ws As Worksheet
    Dim TargetBook As Workbook
    Dim StrFile As String
    
    ' Define Variables
    Set ws = MainSheet
            
    ' this enables user to determine which file they would like to copy to current sheet.
    StrFile = SelectCsvFile_Path()
            
    ' Avoid cause error If user close the File Picker Dialog _
    or click on cancel
    If WorksheetFunction.IsText(StrFile) <> True Then
        Exit Sub
    Else
        Set TargetBook = Workbooks.Open(StrFile)
    End If
    
    
    ' Procedure starts here
    CopyFileSheet _
        SourceFile:=TargetBook, _
        DestinationSheet:=ws
    
    ' Clearing VBA Memory
    Set ws = Nothing
    Set TargetBook = Nothing
    StrFile = vbNull
End Sub

Private Function SelectCsvFile_Path() As String
    Dim Fd As FileDialog
    Dim Filter As FileDialogFilters
    
    Set Fd = Application.FileDialog(3)  ' An Enum Integer representation of the FileDialog Type Object
    With Fd
        ' Define a Filters object
        Set Filter = .Filters
        
        With Filter
            ' Clear the default filters
            .Clear
            ' Add new filter
            .Add "CSV", "*.csv"
            .Add "All Files", "*.*"
        End With
        
            ' Either to allow or disable the multiselect file.
            .AllowMultiSelect = False
            
            If .Show = False Then
                Exit Function
            End If
        SelectCsvFile_Path = .SelectedItems(1)
    End With
    
    ' Clearing VBA Memory
    Set Fd = Nothing
    Set Filter = Nothing
End Function

Private Sub CopyFileSheet(ByRef SourceFile As Workbook, ByRef DestinationSheet As Worksheet)
    ' For reuse in another class method where
    ' - The FileDialogFilePicker.AllowMultiSelect = True, returning an array of file path.
    ' - The idea is to loop openning the workbook &
    ' - call this subroutine and providing the arguements required for the loop.
            SourceFile.Sheets(1).Copy After:=DestinationSheet
            SourceFile.Close SaveChanges:=False
End Sub

Private Sub ExcelBusy(bUpdate As Boolean)
    If bUpdate = True Then
        StartRoutine
    Else
        EndRoutine
    End If
End Sub
Private Sub StartRoutine()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationManual
End Sub
Private Sub EndRoutine()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub Class_Terminate()
        Set MainSheet = Nothing
        ExcelBusy False
End Sub
