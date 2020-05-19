Attribute VB_Name = "OutputTable_Module"
Option Explicit
Sub AddColumnsOutputTable()
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    AddLog ("Starting AddColumnsOutputTable Macro on " & CurrentSheetName)
    Dim e1, e2, e3, e4
    Dim r1, r2
    Dim dict As Scripting.Dictionary
    Set dict = New Dictionary
    
    Dim OutputTable As ListObject
    Set OutputTable = ActiveSheet.ListObjects("Output_" & CurrentSheetName)
    Dim InputTable As ListObject
    Set InputTable = ActiveSheet.ListObjects("Input_" & CurrentSheetName)
   
    ' Check if macro has been run before
    If OutputTable.ListColumns.Count <> 6 Then
        e2 = MsgBox("You can only initilize the sheet once. If you need to add an area you will have to create a new trade sheet, copy over the areas and info on the blue table, initilize that worksheet then copy over any actual production to the gray table.", vbExclamation)
        AddLog ("ListColumns.Count = " & OutputTable.ListColumns.Count & ".  It should equal 6. Exiting Sub.")
        Exit Sub
    End If
    
    ' Verify Unique Short Descriptions
    For Each r1 In ActiveSheet.ListObjects("Input_" & CurrentSheetName).ListColumns("Short Description").DataBodyRange
        If r1 = "" Then
            AddLog ("Blank value in the InputTable. Exiting Sub.")
            e4 = MsgBox("There is a blank short description.  Please fill in each short description with a unique value.", vbExclamation)
            Exit Sub
        End If
        If Not dict.Exists(r1.Value) Then
            dict.Add r1.Value, r1.Value
        End If
    Next r1
    
    If InputTable.ListColumns("Short Description").DataBodyRange.Count = dict.Count Then
        AddLog ("Short Descriptions are unique.")
    Else
        AddLog ("Short Descriptions are not unique. Exiting Sub. Number of ununique values: " & InputTable.ListColumns("Short Description").DataBodyRange.Count - dict.Count)
        e3 = MsgBox("Your Short Description values aren't unique. Please rename " & InputTable.ListColumns("Short Description").DataBodyRange.Count - dict.Count & " of them and try again.", vbExclamation)
        Exit Sub
    End If

    ' Warn this macro can only be run 1 time
    e1 = MsgBox("The Initilize Sheet Macro can only be run one time." & vbNewLine & vbNewLine & "Please verify that you have included all the areas you want to include or you will need to create a new Trade sheet and start over.", vbOKCancel)
    If e1 = vbCancel Then
        AddLog ("RunOneTime dialog cancelled. Ending Sub.")
        Exit Sub
    End If
    
    ' Add output table columns
    For Each r2 In InputTable.ListColumns("Short Description").DataBodyRange
        OutputTable.ListColumns.Add.Name = "WP_" & r2
        OutputTable.ListColumns.Add.Name = "WA_" & r2
        
    Next r2
    ' update output table formulas
    
    ' resize number of output rows (should probably be a different maco
        
    
    
End Sub
