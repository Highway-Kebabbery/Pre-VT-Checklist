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
Public Sub DbConnDriverCheck(Bitness As String)
    MsgBox "Failed to connect to the SQLite database." & vbCrLf & _
           "It appears you are using the " & Bitness & " version of Microsoft Office." & vbCrLf & _
           "Please ensure you are using the appropriate version of the SQLite ODBC driver for your Office installation.", _
           vbCritical, "Connection Error"
End Sub

