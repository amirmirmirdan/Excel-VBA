Adds a new **[WorkbookQuery](https://docs.microsoft.com/en-us/office/vba/api/excel.queries.add)** object to the **Queries** collection.

## Syntax

_expression_.**Add** (_Name_, _Formula_, _Description_)

_expression_ A variable that represents a **[Queries](https://docs.microsoft.com/en-us/office/vba/api/excel.queries)** object.

## Parameters

| Name | Required/Optional | Data type | Description |
| --- | --- | --- | --- |
| _Name_ | Required | **String** | The name of the query. |
| _Formula_ | Required | **String** | The Power Query M formula for the new query. |
| _Description_ | Optional | **Variant** | The description of the query. |

---

### Example from record macro done
```vb
    ActiveWorkbook.Queries.Add _
    Name:="Raw files", Formula:="let" & _
        Chr(13) & "" & Chr(10) & "    Source = Folder.Files(""G:\Shared drives\Finance Shared Services Centre  II\2022\RTR\Closing Folder\PH01\04 - Cash Management\0522\GCASH\Raw files"")," & _
        Chr(13) & "" & Chr(10) & "    #""Filtered Rows"" = Table.SelectRows(Source, each [Extension] = "".csv"")," & _
        Chr(13) & "" & Chr(10) & "    #""Filtered Hidden Files1"" = Table.SelectRows(#""Filtered Rows"", each [Attributes]?[Hidden]? <> true)," & _
        Chr(13) & "" & Chr(10) & "    #""Invoke Cu" & "stom Function1"" = Table.AddColumn(#""Filtered Hidden Files1"", ""Transform File"", each #""Transform File""([Content]))," & _
        Chr(13) & "" & Chr(10) & "    #""Renamed Columns1"" = Table.RenameColumns(#""Invoke Custom Function1"", {""Name"", ""Source.Name""})," & _
        Chr(13) & "" & Chr(10) & "    #""Removed Other Columns1"" = Table.SelectColumns(#""Renamed Columns1"", {""Source.Name"", ""Transform File""})," & _
        Chr(13) & "" & Chr(10) & "    #""Expanded Table Column1"" = Table.ExpandTableColumn(#""Removed Other Columns1"", ""Transform File"", Table.ColumnNames(#""Transform File""(#""Sample File"")))," & _
        Chr(13) & "" & Chr(10) & "    #""Removed Top Rows"" = Table.Skip(#""Expanded Table Column1"",3)," & _
        Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(#""Removed Top Rows"", [PromoteAllScalars=true])," & _
        Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Date"", type datetime}})," & _
        Chr(13) & "" & Chr(10) & "    #""Filtered Rows1"" = Table.SelectRows(#""Changed Type"", each true)," & _
        Chr(13) & "" & Chr(10) & "    #""Added Index"" = Table.AddIndexColumn(#""Filtered Rows1"", ""Index"", 1, 1, Int64.Type)," & _
        Chr(13) & "" & Chr(10) & "    #""Changed Type1"" = Table.TransformColumnTypes(#""Added Index"",{{""Date"", type date}, {""Amount"", Currency.Type}, {""Target"", type text}, {""Pre Bal."", Currency.Type}, {""Trans ID"", type text}, {""MSISDN"", type text}, {""Type"", type text}, {""Channel"", type text}, {""State"", type text}, {""Post Bal."", Currency.Type}, {""Details"", type text}})," & _
        Chr(13) & "" & Chr(10) & "    #""Reordered Columns"" = Table.ReorderColumns(#""Changed Type1"",{""Index"", ""20220501_MyTaxi_nobranch_grabtaxi-p3.csv"", ""Date"", ""Trans ID"", ""MSISDN"", ""Type"", ""Channel"", ""State"", ""Amount"", ""Target"", ""Pre Bal."", ""Post Bal."", ""Details""})," & _
        Chr(13) & "" & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Reordered Columns"",{""20220501_MyTaxi_nobranch_grabtaxi-p3.csv""})," & _
        Chr(13) & "" & Chr(10) & "    #""Filtered Rows2"" = Table.SelectRows(#""Removed Columns"", each true)" & _
        Chr(13) & "" & Chr(10) & "in" & _
        Chr(13) & "" & Chr(10) & "    #""Filtered Rows2"""
    
    ActiveWorkbook.Queries.Add _
    Name:="Sample File", Formula:="let" & Chr(13) & "" & Chr(10) & "    Source = Folder.Files(""G:\Shared drives\Finance Shared Services Centre  II\2022\RTR\Closing Folder\PH01\04 - Cash Management\0522\GCASH\Raw files"")," & _
        Chr(13) & "" & Chr(10) & "    #""Filtered Rows"" = Table.SelectRows(Source, each [Extension] = "".csv"")," & _
        Chr(13) & "" & Chr(10) & "    Navigation1 = #""Filtered Rows""{0}[Content]" & _
        Chr(13) & "" & Chr(10) & "in" & _
        Chr(13) & "" & Chr(10) & "    Navigation1"
    
    ActiveWorkbook.Queries.Add _
    Name:="Parameter1", Formula:="#""Sample File"" meta [IsParameterQuery=true, BinaryIdentifier=#""Sample File"", Type=""Binary"", IsParameterQueryRequired=true]"
    
    ActiveWorkbook.Queries.Add _
    Name:="Transform Sample File", Formula:="let" & _
        Chr(13) & "" & Chr(10) & "    Source = Csv.Document(Parameter1,[Delimiter="","", Columns=11, Encoding=1252, QuoteStyle=QuoteStyle.None])" & _
        Chr(13) & "" & Chr(10) & "in" & _
        Chr(13) & "" & Chr(10) & "    Source"
	
	' This one is the M Code function
    ActiveWorkbook.Queries.Add _
    Name:="Transform File", Formula:="let" & _
        Chr(13) & "" & Chr(10) & "    Source = (Parameter1) => let" & _
        Chr(13) & "" & Chr(10) & "        Source = Csv.Document(Parameter1,[Delimiter="","", Columns=11, Encoding=1252, QuoteStyle=QuoteStyle.None])" & _
        Chr(13) & "" & Chr(10) & "    in" & Chr(13) & "" & Chr(10) & "        Source" & _
        Chr(13) & "" & Chr(10) & "in" & _
        Chr(13) & "" & Chr(10) & "    Source"
```

```vb
'  This load the data to the worksheet as a table.
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Raw files"";Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Raw files]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Raw_files"
        .Refresh BackgroundQuery:=False
    End With
    
    Workbooks("Book7").Connections.Add2 "Query - Sample File", _
        "Connection to the 'Sample File' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Sample File"";Extended Properties=""""" _
        , "SELECT * FROM [Sample File]", 2
        
    Workbooks("Book7").Connections.Add2 "Query - Parameter1", _
        "Connection to the 'Parameter1' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Parameter1;Extended Properties=""""" _
        , "SELECT * FROM [Parameter1]", 2
    
    Workbooks("Book7").Connections.Add2 "Query - Transform Sample File", _
        "Connection to the 'Transform Sample File' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Transform Sample File"";Extended Properties=""""" _
        , "SELECT * FROM [Transform Sample File]", 2
    
    Workbooks("Book7").Connections.Add2 "Query - Transform File", _
        "Connection to the 'Transform File' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Transform File"";Extended Properties=""""" _
        , "SELECT * FROM [Transform File]", 2
    
    ActiveSheet.ListObjects("Raw_files").TableStyle = ""
    ActiveSheet.ListObjects("Raw_files").Range.AutoFilter Field:=2, Criteria1:="<>"

```

## Test
Referencing Source File.
```sql
let
// Sample Function created
    fSourceFile = (sFileDir) => let
        fxSourceFile = Csv.Document(sFileDir, [Delimiter=",", Column=11, Encoding=1252, QuoteStyle=QuoteStyle.None])
    in
        fxSourceFile
in
    fSourceFile
```
or
```sql
[ // Sample Function created
	
	fSourceFile = (sFileDir) => [
		fxSourceFile = Csv.Document(sFileDir, [Delimiter=",", Column=11, Encoding=1252, QuoteStyle=QuoteStyle.None])
	][fxSourceFile]
	
][fSourceFile]
```
## Trans
```sql
[
    FilterTable = Table.SelectRows(fSourceFile, each [Column2]<> null and [Column]<> ""),
    HeadName = Table.PromoteHeaders(FilterTable),
    ChgDateColumn = Table.TransformColumnTypes(HeadName, {{"Date", type datetime}}),
    TableOutput = Table.TransformColumnTypes(ChgDateColumn, 
        {{"Date", type date},
        {"Trans ID", type text}, 
        {"MSISDN", type text}, 
        {"Type", type text}, 
        {"Channel", type text}, 
        {"State", type text}, 
        {"Target", type text}, 
        {"Details", type text}, 
        {"Amount", Currency.Type}, 
        {"Pre Bal.", Currency.Type}, 
        {"Post Bal.", Currency.Type}})
][TableOutput]
```

## Append Queries
```SQL
[
	fAppendTbl = (TblMain, TblAdditional) => Table.Combine(TblMain, TblAdditional)
][fAppendTbl]
```

## Loop Stmt in M Code
```m
	each _ + 1
	// above is similar to the below statement
	each [A]
	

```