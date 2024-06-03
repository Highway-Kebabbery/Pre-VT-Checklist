VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by: Nathan Whisman 07Apr2024
'Requirements: Objects must exist contiguously in a column. This class uses reference "Microsoft ActiveX Data Objects"
'This class contains methods to query calculations given a string of analyses, scrape the subroutines from the calculations, and build a string of subroutines that can be used in a query.
'
'How to Use:
'   1. Instantiate object and then run .Init.
'       *Objects must be contiguous.
'   2. Call FindSourceCode.
'   3. Call ScrapeSubroutines.
'   4. Me.Objects is now a string of subroutines formatted for direct input into a query. e.g."('<SUBROUTINE NAME 1>', '<SUBROUTINE NAME 2>', '<SUBROUTINE NAME 3')"
'*******************************************************************************************************************************************************************

Public Objects As String
Public rs As ADODB.Recordset
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
    Set Me.rs = Nothing
    
End Function
Public Function FindSourceCode(DataBase As String, AnalysisStr As String)
    Dim GetCalcsQuery As String
    Dim rs As New ADODB.Recordset
    
    'For some ridiculous reason, using UPPER() on the column in my SELECT statement leads to pulling corrupted data when using the demo SQLite database, so I capitalize the result using VBA in ScrapeSubroutines()...
    GetCalcsQuery = _
    "SELECT ca.source_code " & _
    "FROM calculation ca " & _
        "INNER JOIN versions v " & _
            "ON v.table_name = 'ANALYSIS' " & _
            "AND ca.analysis = v.name " & _
            "AND ca.version = v.version " & _
    "WHERE ca.analysis IN " & AnalysisStr & " " & _
        "AND ca.source_code IS NOT NULL " & _
        "AND (upper(ca.source_code) LIKE '%GOSUB%' OR upper(ca.source_code) LIKE '%SUBROUTINE%') " & _
        "ORDER BY ca.analysis, ca.component "

    ThisWorkbook.GetRecords DataBase, GetCalcsQuery, rs
    Set Me.rs = rs
End Function
Public Function ScrapeSubroutines()
    Dim RawSourceCodeStr As String
    Dim SourceCodeStr As String
    Dim SplitSourceCodeArr() As String
    Dim TrimmedCodeLine As String
    Dim i As Variant
    Dim SplitCodeLineArr() As String
    
    Me.Objects = "('"
    
    With Me.rs
        If .EOF And .BOF Then
            Exit Function
        Else
            .MoveFirst
            Do While Not .EOF
                'Break code into an array containing one line of code per element.
                RawSourceCodeStr = UCase(ReadClobToString(rs("SOURCE_CODE")))
                SourceCodeStr = Replace(RawSourceCodeStr, Chr(10), Chr(13))
                SplitSourceCodeArr = Split(SourceCodeStr, Chr(13))
                
                'Scrape subroutines from each line.
                For i = 0 To UBound(SplitSourceCodeArr)
                    TrimmedCodeLine = Trim(SplitSourceCodeArr(i))
                    If Not (InStr(TrimmedCodeLine, "'") = 1) Then   'Exclude comments from search
                        If Not (InStr(TrimmedCodeLine, "GOSUB") = 0) Then
                            SplitCodeLineArr = Split(TrimmedCodeLine, "GOSUB")  'Using "GOSUB" as the delimiter places the subroutine name in the second element of the array.
                            TrimmedCodeLine = Trim(Replace(SplitCodeLineArr(1), """", " "))
                            If InStr(Me.Objects, TrimmedCodeLine) = 0 Then
                                Me.Objects = Me.Objects & TrimmedCodeLine & "', '"
                            End If
                            Erase SplitCodeLineArr
                        ElseIf Not (InStr(TrimmedCodeLine, "SUBROUTINE(") = 0) Then
                            SplitCodeLineArr = Split(TrimmedCodeLine, """") 'Using '"' as the delimiter always places the subroutine name in the second element with no whitespace as quotations around the subroutine name are required syntax
                            If InStr(Me.Objects, SplitCodeLineArr(1)) = 0 Then
                                Me.Objects = Me.Objects & SplitCodeLineArr(1) & "', '"
                            End If
                            If (Trim(SplitCodeLineArr(2)) = ",") And (InStr(Me.Objects, SplitCodeLineArr(2)) = 0) Then 'Subroutine() allows for an optional second subroutine as an error-handler with further options after that. This checks to see whether another subroutine exists and accounts for the use of the first sub-routine, omitting the second, and using one of the options after that.
                                Me.Objects = Me.Objects & SplitCodeLineArr(3) & "', '"
                            End If
                            Erase SplitCodeLineArr
                        End If
                    End If
                Next i
                
                .MoveNext
                Erase SplitSourceCodeArr
            Loop
            If Not Len(Me.Objects) = 1 Then 'Nested line of code produces errors when no subroutine found.
                Me.Objects = Left(Me.Objects, Len(Me.Objects) - 3) & ")"
            End If
        End If
    End With
End Function
Private Function ReadClobToString(ClobField As Variant) As String
    Dim strm As New ADODB.Stream
    
    With strm
        .Charset = "UTF-8"
        .Type = 2
        .Open
        .WriteText ClobField.value
        .Position = 0
        ReadClobToString = .ReadText
    End With

    strm.Close
    Set strm = Nothing
End Function

