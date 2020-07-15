Attribute VB_Name = "GenerateReports_Module"
Sub GenerateReports()

    Dim CurrentReportDate As Date: CurrentReportDate = Range("Report_Date").Value
    Dim excelbackupFolderPath As String
    Dim excelbackupFile As String
        
    ' check if trades are ready
    For rr = 11 To 250
        If Cells(rr, 8).Value <> "" Then
            If Cells(rr, 10).Value = "Yes" Then
                If Cells(rr, 9).Value = "Not Ready" Then
                    AddLog (Cells(rr, 8).Value & " was not ready for report generation")
                    e = MsgBox(Cells(rr, 8).Value & " was not ready for report generation. Would you like to continue anyway?", vbYesNo)
                    If e = vbYes Then GoTo nexttradecheck
                    Exit Sub
                End If
            End If
        End If
nexttradecheck:
    Next rr
    
    
    ' Get unique subs
    ' export for each sub
    ' export main report


    'save a workbook copy
    excelbackupFolderPath = Application.ActiveWorkbook.Path & "\includes\excelbackup\"
    excelbackupFile = excelbackupFolderPath & Range("Project_Number").Value & " - UIP Excel Backup File_" & WorksheetFunction.Text(CurrentReportDate, "yyyy-mm-dd") & ".xlsx"
    MyMkDir (excelbackupFolderPath)
    
    ActiveWorkbook.SaveCopyAs (excelbackupFile)
    
End Sub
