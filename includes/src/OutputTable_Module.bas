Attribute VB_Name = "OutputTable_Module"
Option Explicit
Sub AddColumnsOutputTable()
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    AddLog ("Starting AddColumnsOutputTable Macro on " & CurrentSheetName)
    Dim e1, e2, e3, e4
    Dim r1, r2, r3, r4, r5
    Dim dict As Scripting.Dictionary
    Set dict = New Dictionary
    Dim PrimaryAreas_Formula As String
    Dim WeeklyPlan_Formula As String
    Dim WeeklyActual_Formula As String
    Dim WeeklyActual_Alt As String
    
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
    
    '=[@[WP_A1]]+[@[WP_A2]]+[@[WP_A3]]
    ' output formulas preface
    PrimaryAreas_Formula = "=PrimaryAreas(" & InputTable.ListColumns("Short Description").DataBodyRange.Count & ")"
    WeeklyPlan_Formula = "="
    WeeklyActual_Formula = ""
    
    
    ' Add output table columns & continue building output formulas
    For Each r2 In InputTable.ListColumns("Short Description").DataBodyRange
        OutputTable.ListColumns.Add.Name = "WP_" & r2
        OutputTable.ListColumns.Add.Name = "WA_" & r2
        
        WeeklyPlan_Formula = WeeklyPlan_Formula & "[@[WP_" & r2 & "]]+"
        WeeklyActual_Formula = WeeklyActual_Formula & "+[@[WA_" & r2 & "]]"
    Next r2
    
    ' finish output formulas
    WeeklyPlan_Formula = Left(WeeklyPlan_Formula, Len(WeeklyPlan_Formula) - 1)
    WeeklyActual_Alt = Replace(WeeklyActual_Formula, "+", "),ISBLANK(")
    WeeklyActual_Alt = Right(WeeklyActual_Alt, Len(WeeklyActual_Alt) - 2) & ")"
    WeeklyActual_Formula = Right(WeeklyActual_Formula, Len(WeeklyActual_Formula) - 1)
    WeeklyActual_Formula = "=IF(AND(" & WeeklyActual_Alt & "), """"" & ", " & WeeklyActual_Formula & ")"
    
    ' resize number of output rows should probably be a different macro
    ResizeOutputTable
    
    ' apply formulas
    For Each r3 In OutputTable.ListColumns("Primary Areas").DataBodyRange
        r3.Formula = PrimaryAreas_Formula
    Next r3
    For Each r4 In OutputTable.ListColumns("Weekly Plan").DataBodyRange
        r4.Formula = WeeklyPlan_Formula
    Next r4
    For Each r5 In OutputTable.ListColumns("Weekly Actual").DataBodyRange
        r5.Formula = WeeklyActual_Formula
    Next r5
    
    AddLog ("OutputTable Output_" & CurrentSheetName & " initialized.")
End Sub

Sub ResizeOutputTable()
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    Dim FirstReportDate As String: FirstReportDate = Range("S4").Value
    Dim LastReportDate As String: LastReportDate = Range("S5").Value
    Dim weekDifference As Long: weekDifference = DateDiff("d", FirstReportDate, LastReportDate) / 7
    Dim OutputTable As ListObject
    Set OutputTable = ActiveSheet.ListObjects("Output_" & CurrentSheetName)
    ' Dim InputTable As ListObject
    ' Set InputTable = ActiveSheet.ListObjects("Input_" & CurrentSheetName)
    Dim NumRowsToAdd As Long: NumRowsToAdd = weekDifference - OutputTable.ListColumns(1).DataBodyRange.Count
    Dim NumRowsToDelete As Long: NumRowsToDelete = Abs(weekDifference - OutputTable.ListColumns(1).DataBodyRange.Count + 2)
    Dim i As Long
    Dim lastrow As Long
    
    If NumRowsToAdd >= 0 Then
        For i = 0 To NumRowsToAdd
            OutputTable.ListRows.Add
        Next i
        AddLog (NumRowsToAdd + 1 & " rows added to the OutputTable.")
    End If
    If NumRowsToAdd < 0 Then
        For i = 0 To NumRowsToDelete
            lastrow = OutputTable.Range.Rows.Count - 1
            OutputTable.ListRows(lastrow).Delete
        Next i
        AddLog (NumRowsToDelete + 1 & " rows deleted from the OutputTable on sheet " & CurrentSheetName)
    End If
End Sub

Function PrimaryAreas(numberOfAreas As Long)
    'Dim numberOfAreas As Long: numberOfAreas = 3
    Application.Volatile
    
    Dim tb As ListObject
    Dim outputBuilder As Variant
    Set tb = Application.Caller.Worksheet.ListObjects("Output_" & Application.Caller.Worksheet.Range("S2").Value)
    Dim i As Long, listCol As Long
    listCol = 7 ' This is the first column you need to check for
    
    For i = 1 To numberOfAreas
        If Application.Caller.Offset(0, listCol - 2).Value <> "" Then ' change to active cell if you want to run it manually to test
            outputBuilder = outputBuilder + Right(tb.ListColumns(listCol).Name, Len(tb.ListColumns(listCol).Name) - 3) & ", "
        End If
        listCol = listCol + 2 ' This is how I get it to go every other
    Next i
    
    If outputBuilder <> "" Then outputBuilder = Left(outputBuilder, Len(outputBuilder) - 2)
    PrimaryAreas = outputBuilder
End Function
