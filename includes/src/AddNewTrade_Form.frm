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


Private Sub UserForm_Initialize()
    Dim DivTable_Array As Variant
    DivTable_Array = Sheets("Settings").ListObjects("Divisions_Table").DataBodyRange.Value
    Division_ComboBox.List = DivTable_Array
    
End Sub


Private Sub Create_Button_Click()

End Sub

Private Sub Cancel_Button_Click()
    Unload Me
    End
End Sub
