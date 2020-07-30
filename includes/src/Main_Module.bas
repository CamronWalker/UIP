Attribute VB_Name = "Main_Module"
Sub Add_New_Trade_Sub()
    AddNewTrade_Form.Show
    
End Sub

Sub Hide_Unused_Trades(Optional disableWorkbookFunctionality As Boolean = True)
    If disableWorkbookFunctionality = True Then TurnOffFunctionality
    Dim DivTable_Array As Variant
    DivTable_Array = Sheets("Settings").ListObjects("Divisions_Table").DataBodyRange.Value
    
    Rows.EntireRow.Hidden = False
    For i = 11 To 250
        If IsInArray(Cells(i, 2).Value, DivTable_Array) = True Then
            If Cells(i + 1, 2).Value = "" Then
                Rows(i).Hidden = True
                Rows(i + 1).Hidden = True
            End If
        End If
    Next i
    If disableWorkbookFunctionality = True Then TurnOnFunctionality
End Sub

Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'DEVELOPER: Ryan Wells https://wellsr.com/vba/2016/excel/check-if-value-is-in-array-vba/
'DESCRIPTION: Function to check if a value is in an array of values
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function
