Attribute VB_Name = "GenerateReports_Module"
Sub GenerateReports()

    Dim CurrentReportDate As Date: CurrentReportDate = Range("Report_Date").Value
    Dim excelbackupFolderPath As String
    Dim excelbackupFile As String
    Dim uniqueSubsCollection As Collection
    Dim subDocsPathsCollection As Collection
    Set subDocsPathsCollection = New Collection
        
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

    Set uniqueSubsCollection = CollectUniques(Range("C11:C250"))
        
    ' export for each sub
    For Each subcontractor In uniqueSubsCollection
        For sr = 11 To 250
            If Cells(sr, 3).Value = subcontractor Then
                If Cells(sr, 10).Value = "Yes" Then
                    subDocsPathsCollection.Add = Application.ActiveWorkbook.Path & "\includes\assets\tradecovers\" & Cells(sr, 8).Value & "_Cover - " & WorksheetFunction.Text(CurrentReportDate, "yyyy-mm-dd") & ".pdf"
                    subDocsPathsCollection.Add = Application.ActiveWorkbook.Path & "\includes\assets\tradebackup\" & Cells(sr, 8).Value & "_Backup - " & WorksheetFunction.Text(CurrentReportDate, "yyyy-mm-dd") & ".pdf"
                    
                End If
            End If
        Next sr
        
        
    Next subcontractor
    
        
    ' export main report


    'save a workbook copy
    excelbackupFolderPath = Application.ActiveWorkbook.Path & "\includes\excelbackup\"
    excelbackupFile = excelbackupFolderPath & Range("Project_Number").Value & " - UIP Excel Backup File_" & WorksheetFunction.Text(CurrentReportDate, "yyyy-mm-dd") & ".xlsx"
    MyMkDir (excelbackupFolderPath)
    
    ActiveWorkbook.SaveCopyAs (excelbackupFile)
    
End Sub


Public Function CollectUniques(rng As Range) As Collection
    
    Dim varArray As Variant, var As Variant
    Dim col As Collection
    
    'Guard clause - if Range is nothing, return a Nothing collection
    'Guard clause - if Range is empty, return a Nothing collection
    If rng Is Nothing Or WorksheetFunction.CountA(rng) = 0 Then
        Set CollectUniques = col
        Exit Function
    End If
        
    If rng.Count = 1 Then '<~ check for a single cell range
        Set col = New Collection
        col.Add Item:=CStr(rng.Value), Key:=CStr(rng.Value)
    Else '<~ otherwise the range contains multiple cells
        
        'Convert the passed-in range to a Variant array for SPEED and bind the Collection
        varArray = rng.Value
        Set col = New Collection
        
        'Ignore errors temporarily, as each attempt to add a repeat
        'entry to the collection will cause an error
        On Error Resume Next
        
            'Loop through everything in the variant array, adding
            'to the collection if it's not an empty string
            For Each var In varArray
                If CStr(var) <> vbNullString Then
                    col.Add Item:=CStr(var), Key:=CStr(var)
                End If
            Next var
    
        On Error GoTo 0
    End If
    
    'Return the contains-uniques-only collection
    Set CollectUniques = col
    
End Function
