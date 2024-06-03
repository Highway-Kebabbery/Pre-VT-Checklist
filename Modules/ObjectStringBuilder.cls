VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by: Nathan Whisman 25Feb2024
'Requirements: Objects must exist contiguously in a column.
'This class contains methods to build and store strings of object names formatted for use in an SQL query.
'
'How to Use:
'   1. Instantiate object and then run .Init.
'       *Objects must be contiguous.
'   2. Call the appropriate method.
'*****************************************************************************************

Public Objects As String
Private WS As Worksheet
Public LastRow As Long
Public ProductObjectRange As Range
Public DataBase As String

Public Function Init(SheetName As String, Col As String, StartRow As Double, Optional LastRo As Integer = -1, Optional DbName As String)
'VBA doesn't support constructors, so it's important to run this method immediately after instantiating a class of this type.
    
    Set WS = ThisWorkbook.Sheets(SheetName)
    
    'Find last row containing data
    If LastRo = -1 Then
        Me.LastRow = WS.Cells(WS.Rows.Count, Col).End(xlUp).Row
        If Me.LastRow = 6 Then
            MsgBoxes.NoDataEntered
            End
        End If
    Else
        Me.LastRow = LastRo
    End If

    Set ProductObjectRange = ActiveSheet.Range(Col & StartRow & ":" & Col & Me.LastRow)
    
    Me.DataBase = DbName
End Function
Public Function CreateD3CompObjectString()
    Me.Objects = "('"

    If Me.DataBase = "V5" Then
    'Create list for V5 objects, which require prefix truncation.
        'Not all V8 objects have the "SFW_" prefix, as in Item Codes or global objects. Don't truncate those without "SFW_".
        For Each Cell In Me.ProductObjectRange
        'Append objects to list with conversion from V8 format to V5 format
            If Cell.value <> "" Then
                If Left(Cell.value, 4) = "SFW_" Then
                    Me.Objects = Me.Objects & Right(Cell.value, Len(Cell.value) - 4) & "', '"
                Else
                    Me.Objects = Me.Objects & Cell.value & "', '"
                End If
            End If
        Next Cell
        Me.Objects = Left(Me.Objects, Len(Me.Objects) - 3) & ")"
    ElseIf Me.DataBase = "V2" Then
        Me.CreateV8ObjectString
    End If
End Function
Public Function CreateV8ObjectString()
    Me.Objects = "('"
    
    For Each Cell In Me.ProductObjectRange
        If Cell.value <> "" Then
            Me.Objects = Me.Objects & Cell.value & "', '"
        End If
    Next Cell
    
    If Len(Objects) > 3 Then
        Me.Objects = Left(Me.Objects, Len(Me.Objects) - 3) & ")"
    Else
        Me.Objects = ""
        MsgBoxes.NoDataEntered
        End
    End If
End Function
