VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Written by: Nathan Whisman 25Feb2024
'Requirements: 1) Some of these functions reference "Microsoft ActiveX Data Objects." They were written with version 6.1, but best practice is to use the latest version.
'              2) Some of these functions reference the LabwareRecords class module written by D.Guichard daniel.guichard@pfizer.com
'
'*****************************************************************************************

Public Sub ClearReformat(SheetName As String, Optional ResultCellRange As String = "C7:L10000", Optional EntryCellRange As String = "A7:A1006")
    Dim ResultRang As Range
    Dim EntryRang As Range
    
    Worksheets(SheetName).Activate
    Set ResultRang = ActiveSheet.Range(ResultCellRange)
    Set EntryRang = ActiveSheet.Range(EntryCellRange)
    
    With ResultRang
        .ClearContents
        .Interior.Color = xlNone
    End With
    
    With EntryRang
        .Interior.Color = RGB(255, 217, 102)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = False
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
    ActiveSheet.Range("K1:L1").Interior.Color = RGB(217, 217, 217)
    ActiveSheet.Range("K1:L1").value = ""
End Sub
Public Sub AutoFit(SheetName As String, Optional ColRang As String = "C:L")
    Worksheets(SheetName).Activate
    ActiveSheet.Columns(ColRang).AutoFit
    'The instructions above this range mess up the broad-sweeping autofit command above.
    ActiveSheet.Range("C7:C1006").Columns.AutoFit
End Sub
Public Sub GetRecords(DataBase As String, Query As String, rs As ADODB.Recordset)
    Dim lw As New LabwareRecords
    
    lw.SelectDatabase DataBase
    lw.GetRecords Query
    Set rs = lw.recSet
End Sub
Public Sub PopulateComparisonsAgainstD3(ParentTable As String, SheetName As String, DataBase As String, TableName As String, ObjNameCol As String, MinRow As Double, MaxRow As Integer)
    Dim rs As ADODB.Recordset
    Dim D3CompObjectString As New ObjectStringBuilder
    Dim ProductQuery As New QueryBuilder
    
    D3CompObjectString.Init SheetName, ObjNameCol, MinRow, MaxRow, DataBase
    D3CompObjectString.CreateD3CompObjectString
    
    Set rs = Nothing
    If ParentTable = "PRODUCT" Then
        ProductQuery.BuildProductQuery DataBase, TableName, D3CompObjectString.Objects
        ThisWorkbook.GetRecords DataBase, ProductQuery.sqlQuery, rs
        Sheet1.SetProductRecords rs, DataBase, TableName, MinRow, MaxRow
    ElseIf ParentTable = "ANALYSIS" Then
        ProductQuery.BuildAnalysisQuery DataBase, TableName, D3CompObjectString.Objects
        ThisWorkbook.GetRecords DataBase, ProductQuery.sqlQuery, rs
        Sheet1.SetAnalysisRecords rs, DataBase, TableName, MinRow, MaxRow
    End If
End Sub
Public Sub FillInV5Blanks(SheetName As String, StartRow As Double, ObjNameCol As String)
    Dim LastRow As Long
    Dim V5BlankCols
    
    Worksheets(SheetName).Activate
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, ObjNameCol).End(xlUp).Row
    V5BlankCols = Array("D", "F", "I")
    
    For Each Col In V5BlankCols
        For Each Cell In ActiveSheet.Range(Col & StartRow & ":" & Col & LastRow)
            If IsEmpty(Cell.value) Then
                Cell.value = "N/A"
            End If
        Next Cell
    Next Col
End Sub
Public Sub SetStatus(ParentObjectCell As String, SheetName As String, ParentTable As String, SheetStartRow As Double, ObjNameCol As String, EntryCol As String)
    Dim LastResultRow As Long
    Dim LastEntryRow As Long
    
    Worksheets(SheetName).Activate
    LastResultRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, ObjNameCol).End(xlUp).Row
    LastEntryRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, EntryCol).End(xlUp).Row
    
    If ActiveSheet.Range(ParentObjectCell) = "ANALYSIS" Then
        ActiveSheet.Range("O13") = "Object was not queried in V5. V5 is only queried for ANALYSIS, BATCH_LINK, and STD_REAG_TEMP."
        ActiveSheet.Range("O21") = "Object was not queried in V5. V5 is only queried for ANALYSIS, BATCH_LINK, and STD_REAG_TEMP."
    ElseIf ActiveSheet.Range(ParentObjectCell) = "PRODUCT" Then
        ActiveSheet.Range("O13") = "Object was not queried in V5. V5 is only queried for PRODUCT, TEST_LIST, SAMPLING_POINT, and ANALYSIS."
        ActiveSheet.Range("O21") = "Object was not queried in V5. V5 is only queried for PRODUCT, TEST_LIST, SAMPLING_POINT, and ANALYSIS."
    End If
    
    'Lazy way to check and see if number of parent objects queried is number of parent objects returned (e.g. PRODUCT, ANALYSIS).
    If Not ((ActiveSheet.Cells(LastEntryRow, 3) = ParentTable) And (Not (ActiveSheet.Cells(LastEntryRow + 1, 3) = ParentTable))) Then
        ActiveSheet.Cells(1, 11) = "WARNING: Did not find all " & ParentTable & " objects in D3."
        ActiveSheet.Range("K1:L1").Interior.Color = RGB(255, 0, 0)
    End If
    
    For rowNum = SheetStartRow To LastResultRow
    'Note that all of the following evaluations are short-circuited.
        
        'Capture status of existence in each database.
        For colNum = 6 To 8
            If (Not (ActiveSheet.Cells(rowNum, colNum) = "Yes")) And (Not (ActiveSheet.Cells(rowNum, colNum) = "N/A")) Then
                ActiveSheet.Cells(rowNum, 12) = ActiveSheet.Cells(rowNum, 12) & ActiveSheet.Cells(rowNum, colNum) & ", "
            End If
        Next colNum
        
        'Compare changed_on dates between V5-D3 and D3-V2. Any fails between V5-V2 also have a fail for one of these comparisons.
        If (InStr(ActiveSheet.Cells(rowNum, 3), "LIST_ENTRY") = 0) And (InStr(ActiveSheet.Cells(rowNum, 4), "LIST_ENTRY") = 0) Then
            'No need to evaluate the "D3 changed_on" field because it will only ever contain a date.
            If (InStr(ActiveSheet.Cells(rowNum, 9), "Not found") = 0) And (InStr(ActiveSheet.Cells(rowNum, 9), "N/A") = 0) Then
                If CDate(ActiveSheet.Cells(rowNum, 9)) > CDate(ActiveSheet.Cells(rowNum, 10)) Then  'Force date format for comparison. Encountered a bug where it randomly failed the date comparison and this fixed it.
                    ActiveSheet.Cells(rowNum, 12) = ActiveSheet.Cells(rowNum, 12) & "Object is newer in V5 than in D3, "
                End If
            End If
            If (InStr(ActiveSheet.Cells(rowNum, 11), "Not found") = 0) And (InStr(ActiveSheet.Cells(rowNum, 11), "N/A") = 0) Then
                If CDate(ActiveSheet.Cells(rowNum, 10)) > CDate(ActiveSheet.Cells(rowNum, 11)) Then
                    ActiveSheet.Cells(rowNum, 12) = ActiveSheet.Cells(rowNum, 12) & "Object is newer in D3 than in V2, "
                End If
            End If
        End If
        
        'Final formatting of Status field
        If Not ActiveSheet.Cells(rowNum, 12) = "" Then
            'Truncate trailing ", "
            ActiveSheet.Cells(rowNum, 12) = Left(ActiveSheet.Cells(rowNum, 12), Len(ActiveSheet.Cells(rowNum, 12)) - 2)
            ActiveSheet.Cells(rowNum, 12).Interior.Color = RGB(255, 0, 0)
        End If
    Next rowNum
    
    'Status messages for supporting objects aren't critical for PRODUCT parent objects.
    If ParentTable = "PRODUCT" Then
        For rowNum = 7 To 1006
            If (ActiveSheet.Cells(rowNum, 12).Interior.Color = RGB(255, 0, 0)) And (Not ((ActiveSheet.Cells(rowNum, 3) = "PRODUCT") Or (ActiveSheet.Cells(rowNum, 3) = "T_PH_SAMPLE_PLAN"))) Then
                ActiveSheet.Cells(rowNum, 12).Interior.Color = RGB(255, 255, 255)
            End If
        Next rowNum
    End If
End Sub
Public Sub FindRelatedAnalyses(InitialObjectList As String, ObjNameCol As String, DataBase As String)
    Dim FindBatchLinks As String
    Dim FindRelatedAnalyses As String
    Dim BatchLinksStr As String
    Dim LastRow As Long
    Dim rs As New ADODB.Recordset
    Dim NextRow As Integer
    Dim InitialObjectRange As String
    Dim ObjNameColNum As Integer
    
    BatchLinksStr = "('"
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, ObjNameCol).End(xlUp).Row
    NextRow = LastRow + 1
    InitialObjectRange = ObjNameCol & "6:" & ObjNameCol & LastRow
    'Convert ObjNameCol to Integer for use in .Cells()
    ObjNameColNum = Range(ObjNameCol & 1).Column
    
    'Find batch links used by analyses posted in column A
    FindBatchLinks = _
        "SELECT DISTINCT a.batch_link " & _
        "FROM analysis a " & _
            "INNER JOIN versions v " & _
                "ON v.table_name = 'ANALYSIS' " & _
                "AND a.name = v.name " & _
                "AND a.version = v.version " & _
        "WHERE a.name IN " & InitialObjectList & " " & _
            "AND a.batch_link IS NOT NULL " & _
        "ORDER BY a.batch_link "
    
    GetRecords DataBase, FindBatchLinks, rs
    
    With rs
        If .EOF And .BOF Then
            Exit Sub
        Else
            .MoveFirst
            Do While Not .EOF
                BatchLinksStr = BatchLinksStr & .Fields("BATCH_LINK") & "', '"
                .MoveNext
            Loop
            BatchLinksStr = Left(BatchLinksStr, Len(BatchLinksStr) - 3) & ")"
        End If
    End With
    
    'Use returned batch links to append all related analyses to column A
    FindRelatedAnalyses = _
        "SELECT DISTINCT a.name " & _
        "FROM analysis a " & _
            "INNER JOIN versions v " & _
                "ON v.table_name = 'ANALYSIS' " & _
                "AND a.name = v.name " & _
                "AND a.version = v.version " & _
        "WHERE a.batch_link IN " & BatchLinksStr & " " & _
        "ORDER BY a.name "
    
    Set rs = Nothing
    GetRecords DataBase, FindRelatedAnalyses, rs
    
    With rs
        If .EOF And .BOF Then
            Exit Sub
        Else
            .MoveFirst
            Do While Not .EOF
                If Application.WorksheetFunction.CountIf(ActiveSheet.Range(InitialObjectRange), .Fields("NAME")) = 0 Then
                    ActiveSheet.Cells(NextRow, ObjNameColNum) = .Fields("NAME")
                    NextRow = NextRow + 1
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
