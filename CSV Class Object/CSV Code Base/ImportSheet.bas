Attribute VB_Name = "ImportSheet"
Option Explicit

Sub ImportDataToWorkbook()
    Dim CsvHelper As New CSV_class
        CsvHelper.ImportSheet
    Set CsvHelper = Nothing
End Sub
