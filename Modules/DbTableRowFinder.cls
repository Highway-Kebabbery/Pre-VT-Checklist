VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbTableRowFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by: Nathan Whisman 25Feb2024
'Requirements: Objects must exist contiguously in a column.
'This class contains methods to find and store the row numbers where contiguos objects of a specified parent table begin and end in a column.
'
'*****************************************************************************************

Public MinRow As Double
Public MaxRow As Integer
Public TableFound As Boolean

Public Sub FindSupportingRowNums(SheetName As String, TableName As String, ObjNameCol As String)
    Dim LastRow As Long
    
    Worksheets(SheetName).Activate
    Me.MinRow = 1E+300
    Me.MaxRow = -1
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, ObjNameCol).End(xlUp).Row

    For rowNum = 7 To LastRow
        If ActiveSheet.Cells(rowNum, 3) = TableName Then
            If rowNum < Me.MinRow Then
                Me.MinRow = rowNum
            End If
            If rowNum > Me.MaxRow Then
                Me.MaxRow = rowNum
            End If
        End If
    Next rowNum
    
    If Me.MinRow = 1E+300 And Me.MaxRow = -1 Then
        Me.TableFound = False
    Else
        Me.TableFound = True
    End If
End Sub
