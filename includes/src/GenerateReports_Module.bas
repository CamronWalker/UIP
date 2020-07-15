Attribute VB_Name = "GenerateReports_Module"
Sub GenerateReports()

    Dim CurrentReportDate As Date: CurrentReportDate = Range("Report_Date").Value
    Dim excelbackupFolderPath As String
    Dim excelbackupFile As String
    
    
    
    
    ' check if trades are ready
    ' Get unique subs
    ' export for each sub
    ' export main report


    'save a workbook copy
    excelbackupFolderPath = Application.ActiveWorkbook.Path & "\includes\excelbackup\"
    excelbackupFile = nameFilePath & Range("Project_Number").Value & " - Excel Backup File_" & WorksheetFunction.Text(CurrentReportDate, "yyyy-mm-dd") & ".pdf"
    MyMkDir (excelbackupFolderPath)
    
    ActiveWorkbook.SaveCopyAs (excelbackupFile)
    

End Sub
