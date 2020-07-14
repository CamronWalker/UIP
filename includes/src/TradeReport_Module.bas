Attribute VB_Name = "TradeReport_Module"
Sub CreateAssembleBackupFile()
    
    'On Error GoTo ErrorHandler
    Dim CurrentSheetName As String: CurrentSheetName = Range("S2").Value
    Dim nameFilePath As String
    Dim nameFile As String
    
    assembSheetName = "Assemb_Template"
    nameFilePath = Application.ActiveWorkbook.Path & "\includes\assets\tradebackup\" & CurrentSheetName & "\"
    nameFile = nameFilePath & "TradeBackup_" & WorksheetFunction.Text(Sheets(CurrentSheetName).Range("S3").Value, "yyyy-mm-dd") & ".pdf"
    
    Sheets("Assemb_Template").Range("A1").Formula = "=UPPER(""" & Sheets(CurrentSheetName).Range("C7").Value & " " & WorksheetFunction.Text(Sheets(CurrentSheetName).Range("S3").Value, "mm/dd/yyyy") & """)"
    Sheets("Assemb_Template").Range("A34").Value = Sheets(CurrentSheetName).Range("U9").Value
    
    MyMkDir (nameFilePath)
    
    If FileExists(nameFile) = True Then
        AddLog ("TradeBackup_" & WorksheetFunction.Text(Sheets(CurrentSheetName).Range("S3").Value, "yyyy-mm-dd") & ".pdf already exists. Exiting Sub")
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
