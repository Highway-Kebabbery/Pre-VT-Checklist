Attribute VB_Name = "MsgBoxes"
'Written by: Nathan Whisman 08Mar2024
'
'*****************************************************************************************

Public Sub NoDataEntered()
    MsgBox "No data found. Enter object names in column A beginning with cell A7."
End Sub
Public Sub WrongParentObject()
    MsgBox "No objects found in the D3 database. Make sure the correct object type was selected in cell A4."
End Sub
Public Sub DbConnFail()
    MsgBox "Could not connect to the database. Try checking your connection manually."
End Sub
