VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Csv_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Example As String
' Methods
Public Sub ImportSheet()
    ' Declaring variables
        Dim ws As Worksheet
        Dim strFile As String, strFilter As String, strCaption As String
        Dim TargetBook As Workbook
        
    ' Define Variables
        Set ws = ActiveSheet
        strFilter = "Text Files (*.prn ; *.txt ; *.csv)"
        strCaption = "Please Select a file to import."
    
    ' this enables user to determine which file they would like to copy to current sheet.
        strFile = Application.GetOpenFilename(strFilter, , strCaption)
            If strFile = "" Then Exit Sub
        Set TargetBook = Workbooks.Open(strFile)
    
    ' Procedure starts here
        TargetBook.Sheets(1).Copy After:=ws
        TargetBook.Close SaveChanges:=False
    
    ' Clearing VBA Memory
        Set ws = Nothing
        Set TargetBook = Nothing
        strFile = vbNull
        strFilter = vbNull
        strCaption = vbNull
End Sub

Public Sub ImportRange(Optional ByRef vImportToRange As Range)
    ' Declaring variables
        Dim rng As Range
        Dim TargetBook As Workbook, TargetRange As Range
        Dim strFile As String, strFilter As String, strCaption As String
    
    ' Check if optional arguement was provided.
        If vImportToRange = Nothing Then
            With ActiveWorkbook.ActiveSheet
                Set rng = .ActiveCell
            End With
        Else
            Set rng = vImportToRange
        End If
        
    ' Define Variables
        strFilter = "Text Files (*.prn ; *.txt ; *.csv)"
        strCaption = "Please Select the CSV file to import."
    ' this enables user to determine which file they would like to copy to current sheet.
        strFile = Application.GetOpenFilename(strFilter, , strCaption)
            If strFile = "" Then Exit Sub
        Set TargetBook = Workbooks.Open(strFile)
        Set TargetRange = TargetBook.Sheets(1).UsedRange
    
    ' Copy the value from CSV file to the target range
            rng.Value = TargetRange.Value
            TargetBook.Close SaveChanges:=False
    
    'Clearing VBA Memory
        Set rng = Nothing
        Set TargetRange = Nothing
        Set TargetBook = Nothing
        strFile = vbNull
        strFilter = vbNull
        strCaption = vbNull
End Sub
