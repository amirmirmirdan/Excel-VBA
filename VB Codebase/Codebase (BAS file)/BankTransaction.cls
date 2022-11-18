Option Explicit
' Event Sub Procedures
Private Sub Class_Initialize()
    ExcelIsBusy True
    ResetFormatting
End Sub
Private Sub Class_Terminate()
    ExcelIsBusy False
End Sub
' Public Sub Procedures
Public Sub ExcelIsBusy(bUpdate As Boolean)
    If bUpdate = True Then
        StartRoutine
    Else
        EndRoutine
    End If
End Sub
Public Sub StartRoutine()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationManual
End Sub
Public Sub EndRoutine()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationAutomatic
End Sub
Public Sub MassageBnkStmt(lBnkBranchEnum As Long)
    Select Case lBnkBranchEnum
        Case 0      ' for SG UOB Bank Statements
            CleanAccount_Stmt
        Case 1      ' for PH BDO Bank Statements
            CleanTrans_History
        Case Else
            MsgBox "Other Bank Branch to be added soon..."
    End Select
End Sub
' Massage UOB Bank Stmt for SG entities
Private Sub CleanAccount_Stmt()
    ' Delete 3 Top Rows
    DeleteRows 3, 1
    With ActiveSheet
        .Columns("C:E").Select
        FormatSelection DataTypeColumn.ColumnDate
        '
        .Columns("G:K").Select
        FormatSelection DataTypeColumn.ColumnDescription
        '
        ReplaceZeros_AmountColumn "R:T"
        FormatSelection DataTypeColumn.ColumnAmount
        '
        .Columns("L:Q").Columns.Group
        .Columns("I:J").Columns.Group
        .Columns("D:E").Columns.Group
        .Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    End With
End Sub
' Massage BDO Bank Stmt for PH entities
Private Sub CleanTrans_History()
    ' Delete 3 Top Rows
    DeleteRows 3, 1
    ExtraColumnHeader_BDO
    With ActiveSheet
        .Columns("A:A").ColumnWidth = 12
        .Columns("D:F").Style = "Comma"
        .Columns("B:N").EntireColumn.AutoFit
        .Columns("C:C").ColumnWidth = 50
    End With
End Sub
Private Sub ExtraColumnHeader_BDO()
    Dim arrColumnHeader(1 To 7) As String
    Dim i As Integer
        arrColumnHeader(1) = "BUDGET CODE | DESCRIPTION"
        arrColumnHeader(2) = "TRANSACTION REFERENCE"
        arrColumnHeader(3) = "BLANK"
        arrColumnHeader(4) = "TRANSACTION SEQ"
        arrColumnHeader(5) = "RECEIPT NUMBER"
        arrColumnHeader(6) = "ORACLE DOC NUMBER"
        arrColumnHeader(7) = "ADDITIONAL COMMENT"
    For i = 1 To UBound(arrColumnHeader)
        ActiveSheet.Cells(1, 7 + i).Value = arrColumnHeader(i)
    Next i
    Erase arrColumnHeader
    i = vbNull
End Sub
Private Sub FormatSelection(lDataTypeEnum As Long)
    Select Case lDataTypeEnum
        Case 0 ' Description Column
            With ActiveSheet
                Selection.NumberFormat = "0"
                Selection.EntireColumn.AutoFit
            End With
        Case 1 ' Date Column
            With ActiveSheet
                Selection.NumberFormat = "DD/MM/YYYY"
                Selection.EntireColumn.ColumnWidth = 12
            End With
        Case 2 ' Amount
            With ActiveSheet
                Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
                Selection.EntireColumn.ColumnWidth = 14
                Selection.HorizontalAlignment = xlRight
            End With
        Case Else
            With ActiveSheet
                Selection.EntireColumn.ColumnWidth = 50
            End With
    End Select
End Sub
Private Sub ResetFormatting()
    With ActiveSheet.UsedRange
        .ClearFormats
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .MergeCells = False
        .Orientation = 0
        .AddIndent = False
        .ReadingOrder = xlContext
    End With
End Sub
Private Sub DeleteRows(lLastRow As Long, Optional lStartRow As Long)
    Dim lCount As Long
    lCount = 1
        If lStartRow = 0 Then
            For lCount = 1 To lLastRow
                ActiveSheet.Rows(1).Delete
            Next lCount
        Else

            If lStartRow <= lLastRow Then
                For lCount = lStartRow To lLastRow
                    ActiveSheet.Rows(lStartRow).Delete
                Next lCount

            ElseIf lStartRow > lLastRow Then
                MsgBox "The Start Row must be smaller than the last row."
            Else
            End If
        End If
    lCount = vbNull
End Sub
Private Sub ReplaceZeros_AmountColumn(strColumn As String)
    With ActiveSheet.Columns(strColumn)
        .Select
        .Replace _
            What:="0", _
            Replacement:="", _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            MatchCase:=False, _
            SearchFormat:=False, _
            ReplaceFormat:=False, _
            FormulaVersion:=xlReplaceFormula2
        .Style = "Comma"
    End With
End Sub