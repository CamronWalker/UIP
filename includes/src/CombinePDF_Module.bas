Attribute VB_Name = "CombinePDF_Module"
Sub CombinePDFs(inputPathArray As Variant, outputPath As String, waitOnReturn As Boolean)
    ' This is my trying to adapt the old combinepdfs from
    ' https://www.mrexcel.com/forum/excel-questions/870539-combine-pdfs-using-vba-2.html
    ' and
    ' https://stackoverflow.com/questions/15951837/wait-for-shell-command-to-complete
    ' and turn them into something that uses sejda-console because it is really fast
    ' and because some of the new pdf files versions weren't working with pdftk
    ' Written by Camron Walker 01/07/2019
    ' https://gitlab.com/camronwalker/uip-master
    
    'Dim inputPathArray
    'Dim outputPath As String: outputPath = "C:\Users\camron\code\uip-master\QuantityLinks\Output PDF File.pdf"
    'Dim waitOnReturn As Boolean: waitOnReturn = True
    'inputPathArray = Array("C:\Users\camron\code\uip-master\QuantityLinks\Test PDF File.pdf", "C:\Users\camron\code\uip-master\QuantityLinks\Test PDF File.pdf", "C:\Users\camron\code\uip-master\QuantityLinks\Test PDF File.pdf")
    
    Dim dejdaPath As String: dejdaPath = Application.ActiveWorkbook.Path & "\includes\assets\sejda-console\bin\sejda-console.bat"
    Dim strShell As String
    Dim i As Long
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim windowStyle As Integer: windowStyle = 7
    
    strShell = """" & dejdaPath & """" & " merge -f"
    
    For i = LBound(inputPathArray) To UBound(inputPathArray)
        strShell = strShell & " """ & inputPathArray(i) & """"
    Next i
    
    strShell = strShell & " -o """ & outputPath & """ -a flatten --overwrite -b one_entry_each_doc"
    Debug.Print strShell
    
    wsh.Run strShell, windowStyle, waitOnReturn
    
End Sub
