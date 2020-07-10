Attribute VB_Name = "AutomateAssemble_Module"
Sub Automate_MapQuantities()
    'On Error GoTo ErrorHandler
    'Application.ScreenUpdating = False
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    Dim InputTable As ListObject
    Set InputTable = ActiveSheet.ListObjects("Input_" & CurrentSheetName)
    AddLog ("Starting Automate_MapQuantities Macro on " & CurrentSheetName)
    
    For r1 = 1 To InputTable.DataBodyRange.Rows.Count
        If Not (InputTable.DataBodyRange(r1, 6).Comment Is Nothing) Then InputTable.DataBodyRange(r1, 6).Comment.Delete
        InputTable.DataBodyRange(r1, 6).AddComment "{""QuantityPropertyId"":""Count"",""Filters"":[{""FilterPropertyId"":""ZoneArea_AssembleProperty"",""FilterValues"":[""" & InputTable.DataBodyRange(r1, 2) & """]}]}"
        
        If Not (InputTable.DataBodyRange(r1, 7).Comment Is Nothing) Then InputTable.DataBodyRange(r1, 7).Comment.Delete
        InputTable.DataBodyRange(r1, 7).AddComment "{""QuantityPropertyId"":""Count"",""Filters"":[{""FilterPropertyId"":""ZoneArea_AssembleProperty"",""FilterValues"":[""" & InputTable.DataBodyRange(r1, 2) & """]},{""FilterPropertyId"":""InstallationStatus2_AssembleProperty"",""FilterValues"":[""Completed""]}]}"
    Next r1
    
    AddLog ("Completed Automate_MapQuantities Macro on " & CurrentSheetName)
    ' Error Handler
    Exit Sub
ErrorHandler:
    AddLog ("CreateAssembleBackupFile had an error. Error: " & Err.Number & " -- " & Err.Description)
    e = MsgBox("Error: " & Err.Number & vbNewLine & vbNewLine & Err.Description, vbExclamation, "Error")
    Err.Clear
End Sub
