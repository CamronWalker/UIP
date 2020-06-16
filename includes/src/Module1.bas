Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Application.WindowState = xlNormal
    ActiveWorkbook.Queries.Add Name:="Output_Template", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Output_Template""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Date"", type date}, {""Primary Areas"", type any}, {""Weekly Plan"", type number}, {""Weekly Actual"", Int64.Type}, {""Accumulated Plan"", type number}, {""Accumulated Actual"", type any}, {""WP_A1"", type number}, {""WA_A1""" & _
        ", Int64.Type}, {""WP_A2"", type number}, {""WA_A2"", Int64.Type}, {""WP_A3"", type number}, {""WA_A3"", type any}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    Workbooks("UIP Template.xlsm").Connections.Add2 "Query - Output_Template", _
        "Connection to the 'Output_Template' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Output_Template;Extended Properties=""""" _
        , "SELECT * FROM [Output_Template]", 2
    Application.WindowState = xlNormal
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="Output_0000", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Output_0000""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Date"", type date}, {""Primary Areas"", type any}, {""Weekly Plan"", type number}, {""Weekly Actual"", type any}, {""Accumulated Plan"", type number}, {""Accumulated Actual"", type any}, {""WP_ABC"", Int64.Type}, {""WA_ABC"", typ" & _
        "e any}, {""WP_ACS"", type number}, {""WA_ACS"", type any}, {""WP_AL1"", type number}, {""WA_AL1"", type any}, {""WP_HBase"", type number}, {""WA_HBase"", type any}, {""WP_HWST"", type number}, {""WA_HWST"", type any}, {""WP_HWSR"", type number}, {""WA_HWSR"", type any}, {""WP_HL2"", type number}, {""WA_HL2"", type any}, {""WP_HBCeil"", type number}, {""WA_HBCeil"", " & _
        "type any}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    Workbooks("UIP Template.xlsm").Connections.Add2 "Query - Output_0000", _
        "Connection to the 'Output_0000' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Output_0000;Extended Properties=""""" _
        , "SELECT * FROM [Output_0000]", 2
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="Output_0000", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Output_0000""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Date"", type date}, {""Primary Areas"", type any}, {""Weekly Plan"", type number}, {""Weekly Actual"", type any}, {""Accumulated Plan"", type number}, {""Accumulated Actual"", type any}, {""WP_ABC"", Int64.Type}, {""WA_ABC"", typ" & _
        "e any}, {""WP_ACS"", type number}, {""WA_ACS"", type any}, {""WP_AL1"", type number}, {""WA_AL1"", type any}, {""WP_HBase"", type number}, {""WA_HBase"", type any}, {""WP_HWST"", type number}, {""WA_HWST"", type any}, {""WP_HWSR"", type number}, {""WA_HWSR"", type any}, {""WP_HL2"", type number}, {""WA_HL2"", type any}, {""WP_HBCeil"", type number}, {""WA_HBCeil"", " & _
        "type any}})," & Chr(13) & "" & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Changed Type"", ""Trade Name"", each ""Output_0000"")" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Added Custom"""
    Workbooks("UIP Template.xlsm").Connections.Add2 "Query - Output_0000", _
        "Connection to the 'Output_0000' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Output_0000;Extended Properties=""""" _
        , "SELECT * FROM [Output_0000]", 2
End Sub
Sub AddTradeQuery()
Attribute AddTradeQuery.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="Output_Template", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Output_Template""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Date"", type date}, {""Primary Areas"", type any}, {""Weekly Plan"", type number}, {""Weekly Actual"", Int64.Type}, {""Accumulated Plan"", type number}, {""Accumulated Actual"", type any}, {""WP_A1"", type number}, {""WA_" & _
        "A1"", Int64.Type}, {""WP_A2"", type number}, {""WA_A2"", Int64.Type}, {""WP_A3"", type number}, {""WA_A3"", type any}})," & Chr(13) & "" & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Changed Type"", ""Table Name"", each ""OutputTableReference_0000"")" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Added Custom"""
    Workbooks("UIP Template.xlsm").Connections.Add2 "Query - Output_Template", _
        "Connection to the 'Output_Template' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Output_Template;Extended Properties=""""" _
        , "SELECT * FROM [Output_Template]", 2
End Sub
