VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddNewTrade_Form 
   Caption         =   "Add New Trade"
   ClientHeight    =   4290
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4410
   OleObjectBlob   =   "AddNewTrade_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddNewTrade_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Create_Button_Click()
    ' TradeDec_TextBox
    ' SubName_TextBox
    ' Division_ComboBox
    Dim tradeRowNumber As Long
    Dim tradeID As String
    
    ' Verify required fields
    If TradeDec_TextBox = "" Then
        With TradeDec_Label
            .Caption = TradeDec_Label.Caption & " *Required"
            .ForeColor = RGB(255, 0, 0)
        End With
        Exit Sub
    End If
    If SubName_TextBox = "" Then
        With SubName_Label
            .Caption = SubName_Label.Caption & " *Required"
            .ForeColor = RGB(255, 0, 0)
        End With
        Exit Sub
    End If
    If Division_ComboBox = "" Then
        With Division_Label
            .Caption = Division_Label.Caption & " *Required"
            .ForeColor = RGB(255, 0, 0)  '"&H000000FF&"
        End With
        Exit Sub
    End If
    
    ' end verify start create
    
    Call TurnOffFunctionality
    Rows.EntireRow.Hidden = False
    
    
    For i = 11 To 250
        If Cells(i, 2).Value = Division_ComboBox.Value Then
            For tr = 1 To 100 ' determine how many rows after the trade division to create the trade tr should also be the the second half of the trade id
                If Cells(i + tr, 2) = "" Then
                    tradeRowNumber = i + tr
                    GoTo afterTRcount
                End If
            Next tr
afterTRcount:

            Rows(tradeRowNumber).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
            GoTo aftericount
        End If
    Next i
aftericount:
    
    AddLog ("Adding trade: " & TradeDec_TextBox & " to row " & tradeRowNumber)
    
    tradeID = Left(Cells(i, 2).Value, 2) & WorksheetFunction.Text(tr, "00")
    
    'Copy and Rename Sheet
    Sheets("Template").Copy After:=Sheets(Sheets.Count)
    Set coppiedSheet = ActiveSheet
    coppiedSheet.Name = tradeID
    
    'Update Links
    coppiedSheet.Range("C6").Formula = "=Main!C" & tradeRowNumber 'Subcontractor
    ' =RIGHT(Main!B33, LEN(Main!B33) - 4)
    coppiedSheet.Range("C7").Formula = "=RIGHT(Main!B" & tradeRowNumber & ", LEN(Main!B" & tradeRowNumber & ") - 4)"  'Trade
    
    'Table name changes
    For tblCount = 1 To ActiveSheet.ListObjects.Count
        Select Case Left(ActiveSheet.ListObjects(tblCount).Name, 13)
            Case Is = "Output_Templa"
                ActiveSheet.ListObjects(tblCount).Name = "Output_" & tradeID
            Case Is = "Input_Templat"
                ActiveSheet.ListObjects(tblCount).Name = "Input_" & tradeID
        End Select
    Next tblCount
    
    Sheets("Main").Activate
    Cells(tradeRowNumber, 2).Value = WorksheetFunction.Text(tr, "00") & "  " & TradeDec_TextBox
    Cells(tradeRowNumber, 3).Value = SubName_TextBox
    ' ='0501'!N5
    Cells(tradeRowNumber, 4).Formula = "='" & tradeID & "'!N5"
    ' ='0501'!O5
    Cells(tradeRowNumber, 5).Formula = "='" & tradeID & "'!O5"
    Cells(tradeRowNumber, 8).NumberFormat = "General"
    Cells(tradeRowNumber, 8).Formula = "=HYPERLINK(" & """#" & tradeID & "!A1""" & ", " & """" & tradeID & """" & ")"
    ' =IF(Report_Date =@ INDIRECT(H16 & "!S9"), "Ready", "Not Ready")
    Cells(tradeRowNumber, 9).Formula = "=IF(Report_Date =@ INDIRECT(H" & tradeRowNumber & " & ""!S9""" & " ), """ & "Ready""" & ", """ & "Not Ready""" & ")"
    Cells(tradeRowNumber, 10).Formula = "No"
    Cells(tradeRowNumber, 11).Formula = "No"
    
    ' call finishing / formatting functions
    Call Hide_Unused_Trades(False)
    Call TurnOnFunctionality
    Unload Me
End Sub

Private Sub Cancel_Button_Click()
    Unload Me
    End
End Sub


Private Sub UserForm_Initialize()
    Dim DivTable_Array As Variant
    DivTable_Array = Sheets("Settings").ListObjects("Divisions_Table").DataBodyRange.Value
    Division_ComboBox.List = DivTable_Array
    
End Sub
