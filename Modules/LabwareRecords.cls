VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabwareRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by: D.Guichard daniel.guichard@pfizer.com 04Jan2018
'Requirements: This class uses reference "Microsoft ActiveX Data Objects"
'How to Use:
'   1. Define database by using function SelectDabase with the database name (Prod, QA or Dev)
'   2. Create a recordset and set it equal to GetRecords with SQL query string argument
'
'
'Nathan Whisman: Reassigned recSet as a class attribute which removed the need for a few lines of code and fixed an error I encountered. 07Mar2024
'*************************************************************************************************************************

'Declare global class variables:
Private connectionString As String
Private Conn As ADODB.Connection
Public recSet As ADODB.Recordset

Public Function GetRecords(sqlString As String) As ADODB.Recordset
'Returns records as a ADODB.Recordset object
'Move through records with RecordSet.MoveNext

Set Me.recSet = New ADODB.Recordset

'Default connection is prod, set global variable automatically if not set by SelectDatabase function
If connectionString = "" Then
    SelectDatabase "V5"
End If
CreateConnection

'Error handler exists if recordset is empty
On Error GoTo ErrHandler
Me.recSet.Open sqlString, Conn, adOpenStatic

Exit Function
ErrHandler:
CloseConnection
End Function
Public Function SelectDatabase(DataBaseName As String)
'Allows selection of database by name. Use "V5" for V5 Prod, "D3", or "V2" as string argument

Dim UpperDbName As String
UpperDbName = UCase(DataBaseName)

Select Case UpperDbName
Case "V5"
    connectionString = "DRIVER=SQLite3 ODBC Driver;Database=" & ThisWorkbook.Path & "\Databases\V5.db;"
Case "D3"
    connectionString = "DRIVER=SQLite3 ODBC Driver;Database=" & ThisWorkbook.Path & "\Databases\D3.db;"
Case "V2"
    connectionString = "DRIVER=SQLite3 ODBC Driver;Database=" & ThisWorkbook.Path & "\Databases\V2.db;"
Case Else
    connectionString = "DRIVER=SQLite3 ODBC Driver;Database=" & ThisWorkbook.Path & "\Databases\V5.db;"
End Select
End Function
Private Function CreateConnection()
'Creates a new connection to database

'ensure there is not a current database or recordset connection running
If Not (Conn Is Nothing) Then
    Set Conn = Nothing
End If
'create connection with connection string and open connection
Set Conn = New ADODB.Connection
Conn.connectionString = connectionString
On Error GoTo ConnectionError
Conn.Open
Exit Function

'Demo version of project uses SQLite ODBC Driver to facilitate database connection.
'It's possible for people to install the wrong bit version of the driver, so I added this to help them if this happens.
ConnectionError:
    Dim Bitness As String
    #If Win64 Then
        Bitness = "64-bit"
    #Else
        Bitness = "32-bit"
    #End If
    MsgBoxes.DbConnDriverCheck Bitness
    End
End Function
Private Function CloseConnection()
'Closes database connection if open

If Not Conn Is Nothing Then
    Conn.Close
End If
End Function
