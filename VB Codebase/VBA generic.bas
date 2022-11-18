Attribute VB_Name = "FolderFileSub"
Option Explicit
' ---------------------------------------------
' Written by Amir Danial Halim, 22 Jan 2022.
' ---------------------------------------------
Sub CopySheetToCurrentWorkbook()
    ' Purpose: Copy a close workbook's sheet,
    '          into the active workbook.
    ' Inputs: sngSquareMe, the value to be squared
    ' Returns: The input value squared
    '**********************************************
    Dim oHelper As HelperCls
    Dim wb As Workbook, ws As Worksheet
        Set oPretify = New Pretify
        oHelper.StartRoutine
        Set ws = ActiveSheet
    Dim strFile As String
        ' GetOpenFilename will use the Select File Dialog & let user to navigate their Files and select
        ' for files they.
        strFile = Application.GetOpenFilename()
    Dim closedbook As Workbook
        Set closedbook = Workbooks.Open(strFile)
        closedbook.Sheets(1).Copy After:=ws
        closedbook.Close SaveChanges:=False
        oHelper.EndRoutine

        Set oHelper = Nothing
        Set closedbook = Nothing
End Sub

Sub CreateFoldersfromRAnge()
    ' Count ActiveSheets Row Data to be used in the loop statement.
    Dim LRow As Long
        LRow = ActiveSheet.UsedRange.Rows.Count
    Dim i As Long
    i = 0
    ' Loop through the worksheet range. to get it's value
    Dim strFolder As String
    i = 1
        For i = 1 To LRow
            strFolder = Cells(i, 1).Value
            ' Create a folder
            MkDir strFolder
        Next i
    MsgBox "Done! " + LRow + " was created successfully"
End Sub
Sub FormatDataTable()
    ' Purpose:
    ' 1. Quick Format, Current Active Cell and all filled cell adjacent to it, in a table format.
    ' 2. Force to always name the DataTable (Input Box)
    ' Inputs: DataTable Name As String
    ' Returns: Produced a Table formatted data, with style = ""
    '*****************************************************
        ' Declaring variables
        Dim strTableName As String

        ' Input Box
        strTableName = fTableNameInput

        ' Select Current Region and With Selection Code Block to add ListObject
        Dim rg As Range
        With ActiveWorkbook
            Set rg = ActiveSheet.Range("A1").CurrentRegion
        ' Rename DataTable
            ActiveSheet.ListObjects.Add(xlSrcRange, rg, , xlYes).Name = strTableName

        ' Removing Object variable from memory
            Set rg = Nothing

        ' Change Table Style
            ActiveSheet.ListObjects(strTableName).TableStyle = ""
        End With

        ' Removing String variable from memory
        strTableName = Empty

        ' Autofit Column and Rows
        With ActiveSheet
            .Cells.Select

            With Selection
                .EntireRow.AutoFit
                .EntireColumn.AutoFit
                .ColumnWidth = 25
            End With

        End With
End Sub
Function fTableNameInput() As String
    ' Declaring variables
    Dim GiveName As String, BoxPrompt As String
    Dim BoxTitle As String, DefaultName As String
        BoxPrompt = "Hi, Boss! What is my Name?"
        BoxTitle = "Create Data Table Name"
        DefaultName = "DataTable"
            ' Open User Input Dialog Box
            GiveName = InputBox(BoxPrompt, BoxTitle, DefaultName)
            fTableNameInput = GiveName
End Function
'-----------Below this line, not really reviewed yet
Sub Copy_ActiveSheet()
    ' Simple procedure of duplicating the current active worksheet.
    Dim SampleSheet As Worksheet
    Set SampleSheet = ActiveSheet
    SampleSheet.Copy After:=SampleSheet
End Sub

Sub BDO_Bank_Massaging()
    Dim thisSheet As Worksheet
        Cells.Select
        Dim i As Long
    ' Change this block to a function to return array string. Encapsulate.
        Dim Column_Header(1 To 7) As String
                Column_Header(1) = "BUDGET CODE | DESCRIPTION"
                Column_Header(2) = "TRANSACTION REFERENCE"
                Column_Header(3) = "x"
                Column_Header(4) = "TRANSACTION SEQ"
                Column_Header(5) = "RECEIPT NUMBER"
                Column_Header(6) = "ORACLE DOC NUMBER"
                Column_Header(7) = "ADDITIONAL COMMENT"
    ' Until here
        Columns("A:A").ColumnWidth = 12
        Columns("D:F").Style = "Comma"
        Columns("B:N").EntireColumn.AutoFit
        Cells.Select
        For i = 1 To 7
            Cells(4, 7 + i).Value = Column_Header(i)
        Next i
        Columns("B:N").EntireColumn.AutoFit
        Columns("C:C").ColumnWidth = 50
End Sub

Sub UOB_Bank_Massaging()
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("C:E").Select
    Selection.ColumnWidth = 12
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    Columns("G:G").Select
    Selection.NumberFormat = "0"
    Columns("G:H").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("K:K").Select
    Selection.NumberFormat = "0"
    Columns("K:K").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("R:S").Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Style = "Comma"
    Columns("L:Q").Select
    Selection.Columns.Group
    Columns("I:J").Select
    Selection.Columns.Group
End Sub

Sub Refresh_All_Pivot()
    '  Pivot file based on FTE count -Jan21 onwards file
    '  Apply macro when click Refresh icon/button
    ActiveWorkbook.RefreshAll
End Sub

Sub SelectAll_Valued_Range()
    '
    Range(Selection, Selection.End(xlToRight)).Select '  From selection, select till end to the right (Ctrl + Right arrow)
    Range(Selection, Selection.End(xlDown)).Select  ' From current selection, select till end downwards (Ctrl + Down arrow)
    '
End Sub

Sub Clear_ws()
    Cells.Select
    Selection.ClearContents
End Sub

Sub ListSheets()
    Dim ws As Worksheet, wsList As Worksheet
    Dim i As Integer

    Set wsList = Sheets(1)
    Const X = 5
    i = 0
        wsList.Range("B:B").ClearContents
    For Each ws In Worksheets
        wsList.Cells(X + i, 1) = FormatDateTime(Now, 2)
        wsList.Cells(X + i, 2) = ws.Name
        i = i + 1
    Next ws
End Sub

Sub hideAll_Ws_ExceptActive()
    ' --------
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> ActiveSheet.Name Then ws.Visible = xlSheetHidden
    Next ws
    '
End Sub

Sub unhide_All_ws()
    ' ----
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    ActiveWorkbook.Sheets(1).Activate
    '
End Sub

Sub RunFast()
    ' ----
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationManual

    'Your Code Here

    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.Calculation = xlCalculationAutomatic
    '
End Sub

Sub Import_AA()
    '          NOT ADJUSTED YET...
    ' SOME ARE HARD CODED.
    Dim FileToOpen As Variant
    Dim OpenBook As Workbook
        Application.ScreenUpdating = False
        Sheets("Exported").Visible = False
        Sheets("DATA").Select
            Call Clear_ws
        FileToOpen = Application.GetOpenFilename(Title:="Select the Account Analysis by Legal Entity csv file")
            If FileToOpen <> False Then
                Set OpenBook = Application.Workbooks.Open(FileToOpen)
                OpenBook.Sheets(1).Range("A1").Select
                    Call SelectAll_Valued_Range
                Selection.Copy
            ThisWorkbook.Worksheets("DATA").Range("A1").PasteSpecial xlPasteValues
                    OpenBook.Application.CutCopyMode = False
                OpenBook.Close False
            End If
        Application.ScreenUpdating = True
        Range("A1").Select
            Call SelectAll_Valued_Range
        Selection.Name = "SelData"
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("SelData"), , xlYes).Name = "Table1"
        Range("Table1[#All]").Select
        ActiveSheet.ListObjects("Table1").Name = "DATA"
            Call RefreshPvtAll
        Sheets("Count").Select
        MsgBox ("Done Counting the pivots. Good Jobs!")
End Sub

Sub Get_Data_from_File()
    Dim FileToOpen As Variant
    Dim OpenBook As Workbook
    Application.ScreenUpdating = False
    FileToOpen = Application.GetOpenFilename(Title:="Select the Account Analysis by Legal Entity csv file")
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
            OpenBook.Sheets(1).Range("B1").Select
                Call SelectAll_Valued_Range
                Selection.Copy
            ThisWorkbook.Worksheets("Exported").Range("D1").PasteSpecial xlPasteValues
        OpenBook.Close False
    End If
    Application.ScreenUpdating = True
End Sub

Sub CMFindReplaceReconciled()
    ActiveSheet.Select
    Columns("C:XFD").Select
    Columns("C:XFD").AutoFit
    Cells.Replace What:="TRUE", Replacement:="Reconciled", _
        LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Cells.Replace What:="FALSE", Replacement:="", _
        LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub