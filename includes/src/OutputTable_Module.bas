Attribute VB_Name = "OutputTable_Module"
Option Explicit
Sub AddColumnsOutputTable()
    Application.ScreenUpdating = False
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    AddLog ("Starting AddColumnsOutputTable Macro on " & CurrentSheetName)
    Dim e1, e2, e3, e4
    Dim r1, r2, r3, r4, r5, r6, r7, r8
    Dim r7_Counter As Long, r8_Counter As Long
    Dim itwc, itwc2
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
    
    Dim InputTable_WeeklyColumns As Collection
    Set InputTable_WeeklyColumns = New Collection
   
    ' Check if macro has been run before
    If OutputTable.ListColumns.Count <> 6 Then
        e2 = MsgBox("You can only initilize the sheet once. If you need to add an area you will have to create a new trade sheet, copy over the areas and info on the colored table, initilize that worksheet then copy over any actual production to the gray table.", vbExclamation)
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
    
    '=IF([@[WP_A1]]="", 0,[@[WP_A1]]) +IF([@[WP_A2]]="", 0,[@[WP_A2]])+IF([@[WP_A3]]="", 0,[@[WP_A3]])
    ' TODO UPDATE FORMULAS TO MAKE THE DIFFERENT TABLE COLUMN VALUES A COLLECTION SO I CAN MAKE AN IF STATEMENT OUT OF THEM

    ' Add output table columns & continue building output formulas
    For Each r2 In InputTable.ListColumns("Short Description").DataBodyRange
        OutputTable.ListColumns.Add.Name = "WP_" & r2
        OutputTable.ListColumns.Add.Name = "WA_" & r2
        
        InputTable_WeeklyColumns.Add r2.Value
    Next r2
    
    ' output formulas
    PrimaryAreas_Formula = "=PrimaryAreas(" & InputTable.ListColumns("Short Description").DataBodyRange.Count & ")"
    
    WeeklyPlan_Formula = "="
    WeeklyActual_Formula = "="
    For itwc = 1 To InputTable_WeeklyColumns.Count
        WeeklyPlan_Formula = WeeklyPlan_Formula & "IF([@[WP_" & InputTable_WeeklyColumns(itwc) & "]]="""", 0, [@[WP_" & InputTable_WeeklyColumns(itwc) & "]])+"
        WeeklyActual_Formula = WeeklyActual_Formula & "IF([@[WA_" & InputTable_WeeklyColumns(itwc) & "]]="""", 0, [@[WA_" & InputTable_WeeklyColumns(itwc) & "]])+"
    Next itwc
    
    WeeklyPlan_Formula = Left(WeeklyPlan_Formula, Len(WeeklyPlan_Formula) - 1)
    WeeklyActual_Formula = Left(WeeklyActual_Formula, Len(WeeklyActual_Formula) - 1)
    
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
    ' =WeeklyPlanned(Output_Template[[#Headers],[WP_A1]], [@Date])
    For itwc2 = 1 To InputTable_WeeklyColumns.Count
        r6 = Null
        For Each r6 In OutputTable.ListColumns("WP_" & InputTable_WeeklyColumns(itwc2)).DataBodyRange
            r6.Formula = "=WeeklyPlanned(Output_" & CurrentSheetName & "[[#Headers],[WP_" & InputTable_WeeklyColumns(itwc2) & "]], [@Date])"
        Next r6
    Next itwc2
    r7_Counter = 14
    For Each r7 In OutputTable.ListColumns("Accumulated Plan").DataBodyRange
        r7.Formula = "=SUM($AC$14:AC" & r7_Counter & ")"
        r7_Counter = r7_Counter + 1
    Next r7
    r8_Counter = 14
    For Each r8 In OutputTable.ListColumns("Accumulated Actual").DataBodyRange
        r8.Formula = "=IF(AD" & r8_Counter & "=0,NA(),SUM($AD$14:AD" & r8_Counter & "))"
        r8_Counter = r8_Counter + 1
    Next r8
    
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
        AddLog (NumRowsToAdd + 1 & " rows added to the OutputTable on sheet " & CurrentSheetName)
    End If
    If NumRowsToAdd < -1 Then
        For i = 0 To NumRowsToDelete
            lastrow = OutputTable.Range.Rows.Count - 1
            OutputTable.ListRows(lastrow).Delete
        Next i
        AddLog (NumRowsToDelete + 1 & " rows deleted from the OutputTable on sheet " & CurrentSheetName)
    End If
End Sub
Function WeeklyPlanned(ColumnHeader As String, RowDate As Date)
    'Dim ColumnHeader As String: ColumnHeader = Range("AG13").Value
    'Dim RowDate As Date: RowDate = Range("AA24").Value
    '
    ' Written By Camron 2020-05-20
    ''''''''''''''''''''''''''''''''''''''''''''''
    Application.Volatile
    Dim CurrentSheetName As String: CurrentSheetName = Application.Caller.Parent.Name ' ActiveSheet.Range("S2").Value '
    Dim OutputTable As ListObject
    Set OutputTable = Sheets(CurrentSheetName).ListObjects("Output_" & CurrentSheetName)
    Dim InputTable As ListObject
    Set InputTable = Sheets(CurrentSheetName).ListObjects("Input_" & CurrentSheetName)
    Dim HolidayTable As ListObject
    Set HolidayTable = Sheets("Settings").ListObjects("Holidays_Table")
    Dim reportArea As String: reportArea = Right(ColumnHeader, Len(ColumnHeader) - 3)
    Dim InputTable_DataRow As Variant
    Dim DataRow_DateDiff As Variant
    Dim d_Counter As Long
    Dim loopDate As Date
    Dim loopHolidayResult, weekLoopHolidayResult
    Dim workDaysArray As Variant: workDaysArray = Sheets(CurrentSheetName).Range("U2:U8").Value
    Dim d As Variant
    Dim loopDayOfWeek As Long, weekLoopDayOfWeek As Long
    Dim loopWorkingDaysCounter As Long
    Dim weekLoopDate As Date
    Dim i As Long
    Dim dailyProductionValue As Double
    Dim daysThisWeekCounter As Long
    
    InputTable_DataRow = Application.Match(reportArea, Sheets(CurrentSheetName).Range("Input_Template[Short Description]").Value, 0) ' use application.caller instead of active sheet when not testing
    DataRow_DateDiff = DateDiff("d", InputTable.DataBodyRange(InputTable_DataRow, 4), InputTable.DataBodyRange(InputTable_DataRow, 5))
    
    loopDate = InputTable.DataBodyRange(InputTable_DataRow, 4)
    loopWorkingDaysCounter = 0
    For d_Counter = 0 To DataRow_DateDiff
        loopHolidayResult = Application.Match(loopDate, Range("Holidays_Table[Holidays]").Value, 0)
        If IsError(loopHolidayResult) Then
            loopDayOfWeek = Application.WorksheetFunction.Weekday(loopDate, 2)
            If workDaysArray(loopDayOfWeek, 1) = True Then
                loopWorkingDaysCounter = loopWorkingDaysCounter + 1
            End If
        End If
        loopDate = loopDate + 1
        loopDayOfWeek = 0
    Next d_Counter
    weekLoopDate = RowDate - 7
    dailyProductionValue = InputTable.DataBodyRange(InputTable_DataRow, 6) / loopWorkingDaysCounter ' 6 for total row
    
    daysThisWeekCounter = 0
    For i = 1 To 7
        If weekLoopDate >= InputTable.DataBodyRange(InputTable_DataRow, 4) And weekLoopDate <= InputTable.DataBodyRange(InputTable_DataRow, 5) Then
            weekLoopHolidayResult = Application.Match(weekLoopDate, Range("Holidays_Table[Holidays]").Value, 0)
            If IsError(weekLoopHolidayResult) Then ' if there's an error it's not found on the holiday list
                weekLoopDayOfWeek = Application.WorksheetFunction.Weekday(weekLoopDate, 2)
                If workDaysArray(weekLoopDayOfWeek, 1) = True Then
                    daysThisWeekCounter = daysThisWeekCounter + 1
                End If
            End If
        End If
        weekLoopDate = weekLoopDate + 1
    Next i
    
    If daysThisWeekCounter * dailyProductionValue = 0 Then
        WeeklyPlanned = ""
    Else
        WeeklyPlanned = daysThisWeekCounter * dailyProductionValue
    End If
    
End Function

Function GetRow(TableName As String, ColumnNum As Long, Key As Variant) As Range
    ' https://stackoverflow.com/questions/6249039/find-a-row-from-excel-table-using-vba
    On Error Resume Next
    Set GetRow = Range(TableName) _
        .Rows(WorksheetFunction.Match(Key, Range(TableName).Columns(ColumnNum), 0))
    If Err.Number <> 0 Then
        Err.Clear
        Set GetRow = Nothing
    End If
End Function

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

Sub UpdateTrade()
    Application.ScreenUpdating = False
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    Dim CurrentReportDate As Date: CurrentReportDate = Range("S3").Value
    
    AddLog ("Start Trade update on " & CurrentSheetName)
    
    Dim OutputTable As ListObject
    Set OutputTable = ActiveSheet.ListObjects("Output_" & CurrentSheetName)
    Dim InputTable As ListObject
    Set InputTable = ActiveSheet.ListObjects("Input_" & CurrentSheetName)
    
    Dim r1 As Long, r2 As Long, r3 As Long 'r for row
    Dim c1 As Long 'c for column
    Dim loopShortDescription As String
    Dim loopColumnHeaders As String
    Dim loopOutputTableTotal As Double
    Dim loopProductionDifference As Double
    
    'Count Row for current report date
    For r3 = 1 To OutputTable.DataBodyRange.Rows.Count
        If OutputTable.DataBodyRange(r3, 1) = CurrentReportDate Then GoTo exitR3loop
    Next r3
exitR3loop: ' exit loop with r3 as the row number for the current date
      
    For r1 = 1 To InputTable.DataBodyRange.Rows.Count
        loopShortDescription = InputTable.DataBodyRange(r1, 3)
        For c1 = 1 To OutputTable.ListColumns.Count
            loopColumnHeaders = OutputTable.HeaderRowRange(1, c1).Value
            If loopColumnHeaders = "WA_" & loopShortDescription Then GoTo exitC1Loop
        Next c1
exitC1Loop: ' I escape the C1 loop with the c1 value which is the column number of the output table that needs to be updated
        
        If OutputTable.DataBodyRange(r3, c1).Value <> "" Then AddLog ("WA_" & loopShortDescription & " already had a value of " & OutputTable.DataBodyRange(r3, c1).Value & " before the trade update. Value Cleared.")
        OutputTable.DataBodyRange(r3, c1).Value = "" ' remove value from current week's production
        
        For r2 = 1 To OutputTable.DataBodyRange.Rows.Count
            loopOutputTableTotal = loopOutputTableTotal + OutputTable.DataBodyRange(r2, c1).Value
        Next r2
        
        loopProductionDifference = InputTable.DataBodyRange(r1, 7).Value - loopOutputTableTotal
        If loopProductionDifference < 0 Then AddLog ("There was negative production in area WA_" & loopShortDescription & " = " & loopProductionDifference & ". Production wasn't updated to the negative value because that breaks the graph.")
        If loopProductionDifference > 0 Then OutputTable.DataBodyRange(r3, c1).Value = loopProductionDifference
        
        ' clear loop values
        loopProductionDifference = 0
        loopOutputTableTotal = 0
    Next r1
    
    AddLog ("Finished Trade update on " & CurrentSheetName)
End Sub
