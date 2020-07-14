Attribute VB_Name = "TradeReport_Module"
'''''''''''''''''''''''''''''''
' UpdateTrade()
' CreateAssembleBackupFile()
' CreateMergedPDFBackupFile()
' MyMkDir(sPath As String)
' FileExists(FilePath As String) As Boolean

Sub UpdateTrade()
    Application.ScreenUpdating = False
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    Dim CurrentReportDate As Date: CurrentReportDate = Range("S3").Value
    Dim ReportUpdateMethod As String: ReportUpdateMethod = Range("S8").Value
    
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
        If loopProductionDifference >= 0 Then OutputTable.DataBodyRange(r3, c1).Value = loopProductionDifference
        
        ' clear loop values
        loopProductionDifference = 0
        loopOutputTableTotal = 0
    Next r1
    
    'create backup file
    If ReportUpdateMethod = "Assemble Addin" Then CreateAssembleBackupFile
    
    AddLog ("Finished Trade update on " & CurrentSheetName)
End Sub

Sub CreateAssembleBackupFile()
    
    On Error GoTo ErrorHandler
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    Dim nameFilePath As String
    Dim nameFile As String
    
    assembSheetName = "Assemb_Template"
    nameFilePath = Application.ActiveWorkbook.Path & "\includes\assets\tradebackup\"
    nameFile = nameFilePath & CurrentSheetName & "_Backup - " & WorksheetFunction.Text(Sheets(CurrentSheetName).Range("S3").Value, "yyyy-mm-dd") & ".pdf"
    
    Sheets("Assemb_Template").Range("A1").Formula = "=UPPER(""" & Sheets(CurrentSheetName).Range("C7").Value & " " & WorksheetFunction.Text(Sheets(CurrentSheetName).Range("S3").Value, "mm/dd/yyyy") & """)"
    Sheets("Assemb_Template").Range("A34").Value = Sheets(CurrentSheetName).Range("U9").Value
    
    MyMkDir (nameFilePath)
    
    If FileExists(nameFile) = True Then
        AddLog ("TradeBackup_" & WorksheetFunction.Text(Sheets(CurrentSheetName).Range("S3").Value, "yyyy-mm-dd") & ".pdf already exists. Exiting Sub")
        ActiveWorkbook.FollowHyperlink nameFile
        Exit Sub
    End If
    
    ThisWorkbook.Worksheets(assembSheetName).ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        nameFile, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True
    
    Exit Sub
ErrorHandler:
    AddLog ("CreateAssembleBackupFile had an error. Error: " & Err.Number & " -- " & Err.Description)
    e = MsgBox("Error: " & Err.Number & vbNewLine & vbNewLine & Err.Description, vbExclamation, "Error")
    Err.Clear
End Sub

Sub CreateMergedPDFBackupFile()
    On Error GoTo ErrorHandler
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    Dim TakeoffFileLocations As String: TakeoffFileLocations = Range("U10").Value
    Dim TakeoffFile_Array As Variant
    Dim nameFilePath As String
    Dim nameFile As String
    
    nameFilePath = Application.ActiveWorkbook.Path & "\includes\assets\tradebackup\"
    nameFile = nameFilePath & CurrentSheetName & "_Backup - " & WorksheetFunction.Text(Sheets(CurrentSheetName).Range("S3").Value, "yyyy-mm-dd") & ".pdf"
    TakeoffFile_Array = Split(TakeoffFileLocations, "----")
    
    MyMkDir (nameFilePath)
    
    Call CombinePDFs(TakeoffFile_Array, nameFile, False)
    
    Exit Sub
ErrorHandler:
    AddLog ("CreateMergedPDFBackupFile had an error. Error: " & Err.Number & " -- " & Err.Description)
    e = MsgBox("Error: " & Err.Number & vbNewLine & vbNewLine & Err.Description, vbExclamation, "Error")
    Err.Clear
End Sub

Public Sub MyMkDir(sPath As String)
'https://www.devhut.net/2011/09/15/vba-create-directory-structurecreate-multiple-directories/

    Dim iStart          As Integer
    Dim aDirs           As Variant
    Dim sCurDir         As String
    Dim i               As Integer
 
    If sPath <> "" Then
        aDirs = Split(sPath, "\")
        If Left(sPath, 2) = "\\" Then
            iStart = 3
        Else
            iStart = 1
        End If
 
        sCurDir = Left(sPath, InStr(iStart, sPath, "\"))
 
        For i = iStart To UBound(aDirs)
            sCurDir = sCurDir & aDirs(i) & "\"
            If Dir(sCurDir, vbDirectory) = vbNullString Then
                MkDir sCurDir
            End If
        Next i
    End If
End Sub
Function FileExists(FilePath As String) As Boolean
    Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function
