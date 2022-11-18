' Assign Array
Option Explicit
Private Sub arrCriteria()
    Dim arrBudgetCodeTag() As String
    Dim arrBranch() As String
    Dim arrDesc() As String
    Dim rg As Range
    Dim lr As Long
    Dim i As Long, j As Long
    Dim r As Long, c As Long
    With ThisWorkbook.Worksheets("GPAY6421").Range("A1")
        r = .CurrentRegion.Offset(1, 0).Rows.Count
        c = 4
        i = 1
        j = 1
        ReDim arrBudgetCodeTag(i To r, i To c) As String
        For i = 1 To r
            For j = 1 To c
                arrBudgetCodeTag(i, j) = .Offset(i, j - 1).Value
            Next j
        Next i
    End With
    With ActiveSheet
        Set rg = .Range("A1")
        r = vbNull
        c = vbNull
        i = 1
        j = 1
        lr = LastRow - 1
    ReDim arrBranch(i To lr) As String
    ReDim arrDesc(i To lr) As String
        For i = 1 To lr
            arrBranch(i) = rg.Offset(i, 1).Value
            arrDesc(i) = rg.Offset(i, 2).Value
        Next i
        For i = 1 To lr
            j = 1
            Select Case arrBranch(i)
            Case arrBudgetCodeTag(1, 1)
                For j = 1 To 8
                    If InStr(1, arrDesc(i), arrBudgetCodeTag(j, 2), vbTextCompare) <> 0 Then
                        rg.Offset(i, 7).Value = arrBudgetCodeTag(j, 3)
                        rg.Offset(i, 8).Value = arrBudgetCodeTag(j, 4)
                        j = 9
                    Else
                        rg.Offset(i, 7).Value = "Check"
                        rg.Offset(i, 8).Value = "Check"
                    End If
                    rg.Offset(i, 9).Value = i
                Next j
            Case arrBudgetCodeTag(9, 1)
                For j = 9 To 14
                    If InStr(1, arrDesc(i), arrBudgetCodeTag(j, 2), vbTextCompare) <> 0 Then
                        rg.Offset(i, 7).Value = arrBudgetCodeTag(j, 3)
                        rg.Offset(i, 8).Value = arrBudgetCodeTag(j, 4)
                        j = 14
                    Else
                        rg.Offset(i, 7).Value = "Check"
                        rg.Offset(i, 8).Value = "Check"
                    End If
                    rg.Offset(i, 9).Value = i
                Next j
            Case arrBudgetCodeTag(14, 1)
                If InStr(1, arrDesc(i), arrBudgetCodeTag(14, 2), vbTextCompare) <> 0 Then
                    rg.Offset(i, 7).Value = arrBudgetCodeTag(14, 3)
                    rg.Offset(i, 8).Value = arrBudgetCodeTag(14, 4)
                    j = 14
                Else
                    rg.Offset(i, 7).Value = "Check"
                    rg.Offset(i, 8).Value = "Check"
                End If
                rg.Offset(i, 9).Value = i

            Case arrBudgetCodeTag(15, 1)
                    rg.Offset(i, 7).Value = arrBudgetCodeTag(15, 3)
                    rg.Offset(i, 8).Value = arrBudgetCodeTag(15, 4)
                    rg.Offset(i, 9).Value = i

            Case arrBudgetCodeTag(16, 1)
                    rg.Offset(i, 7).Value = arrBudgetCodeTag(16, 3)
                    rg.Offset(i, 8).Value = arrBudgetCodeTag(16, 4)
                    rg.Offset(i, 9).Value = i

            Case Else
                    rg.Offset(i, 7).Value = "Check"
                    rg.Offset(i, 8).Value = "Check"
                    rg.Offset(i, 9).Value = i

            End Select
        Next i
    End With
End Sub
Private Function LastRow() As Long
    With ActiveSheet.Range("A1").CurrentRegion
        LastRow = .Rows.Count
    End With
End Function