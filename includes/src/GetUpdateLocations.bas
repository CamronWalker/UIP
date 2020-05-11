Attribute VB_Name = "GetUpdateLocations"
Sub UpdatePowerBIExportFolderPath()
    ' Camron Walker 2020-05-11
 
    Dim fldr As FileDialog
    Dim sItem As String
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select Power BI Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
    Range("Power_BI_Export_Folder").Value = sItem
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Sub
