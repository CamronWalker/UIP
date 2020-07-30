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
    
    ' TODO: add checks to make sure the form is filled out
    
    Call TurnOffFunctionality
    
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
        End If
    Next i
    
    AddLog ("Adding trade: " & TradeDec_TextBox & " to row " & tradeRowNumber)
    
    Cells(tradeRowNumber, 2).Value = WorksheetFunction.Text(tr, "00") & "  " & TradeDec_TextBox
    Cells(tradeRowNumber, 3).Value = SubName_TextBox
    'TODO: Finish filling out trade row
    
    
    'TODO: Copy template and configure
    
    
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
