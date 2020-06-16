Attribute VB_Name = "Module2"
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="Output_Template", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Output_Template""]}[Content]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Source"
    Workbooks("UIP Template.xlsm").Connections.Add2 "Query - Output_Template", _
        "Connection to the 'Output_Template' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Output_Template;Extended Properties=""""" _
        , "SELECT * FROM [Output_Template]", 2
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="Output_0000", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Output_0000""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Added Custom"" = Table.AddColumn(Source, ""Trade Name"", each ""OutputTable_0000"")" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Added Custom"""
    Workbooks("UIP Template.xlsm").Connections.Add2 "Query - Output_0000", _
        "Connection to the 'Output_0000' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Output_0000;Extended Properties=""""" _
        , "SELECT * FROM [Output_0000]", 2
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1"), , xlYes).Name = "Table7"
    Range("Table7[[#All],[Column1]]").Select
    ActiveWorkbook.Queries.Add Name:="BI_ExportTable", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Table7""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Column1"", type any}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=BI_ExportTable;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [BI_ExportTable]")
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
        .ListObject.DisplayName = "BI_ExportTable"
        .Refresh BackgroundQuery:=False
    End With
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
    Range("BI_ExportTable[[#Headers],[Column1]]").Select
    ActiveCell.FormulaR1C1 = "Column1"
    Range("B5").Select
End Sub
Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro8 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="Append1", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Table.Combine({BI_ExportTable, Output_0000, Output_Template})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Source"
    Workbooks("UIP Template.xlsm").Connections.Add2 "Query - Append1", _
        "Connection to the 'Append1' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Append1;Extended Properties=""""" _
        , "SELECT * FROM [Append1]", 2
End Sub
Sub Macro9()
Attribute Macro9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro9 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="Append1", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Table.Combine({BI_TempTable, Output_0000})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Source"
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Append1;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Append1]")
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
        .ListObject.DisplayName = "Append1"
        .Refresh BackgroundQuery:=False
    End With
End Sub
Sub Test45()
Workbooks("New Data.xlsx").Worksheets("Export").Range("A2:D9").Copy _
    Workbooks("Reports.xlsm").Worksheets("Data").Range("A2")
End Sub
