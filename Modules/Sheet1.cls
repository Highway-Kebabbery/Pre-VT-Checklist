VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Written by: Nathan Whisman 08Mar2024 v1.0 finished
'Requirements: 1) These functions reference "Microsoft ActiveX Data Objects." They were written with version 6.1, but best practice is to use the latest version.
'              2) Some of these funcitons reference functions referencing the LabwareRecords class module written by D.Guichard daniel.guichard@pfizer.com
'
'This Worksheet, in conjunction with functions in ThisWorkbook, identifies all supporting objects for the specified parent objects, tells you whether the objects
'   exist in V5/V2, are flagged as removed/inactive in V5/D3/V2, and whether they're newer in V5/D3 than in V2.
'
'Future updates: For analysis objects, the first step will be to use batch_link to identify all related samples, find all analyses on those samples, append all of
'those analyses to column A, and THEN begin normak execution so that all supporting objects for an entire worklist can be checked. I will make this an optional feature
'(that uses a checkbox or something to turn it on and off) because I don't know yet whether it's always required.
'
'v1.1 updated 08Mar2024: Error handling to ensure data was entered flagged any empty recordset returned from D3. Updated to only check for ParentTable.
'
'A note, given that this is strictly a portfolio piece now:
'I created this on my own time, outside of work hours, as a side project that would help me do my job faster. It wasn't sanctioned by the company.
'As such, I moved faster than I would've liked to move because the goal was to build this and finish my main task faster than it would've taken rather than to have
'build time + task time even out or take longer than without this tool. This means that I made some compromises in areas where I know code could be better, but I just needed
'to get it done. The main area that I want to come back and improve later (when I don't need to build other pieces for a portfolio so that I can get a job) is to clean up
'my class modules so that they're all properly encapsulated.
'***********************************************************************************************************************************************************************

Private Sub BrowseButton_Click()
    Dim ColTitles As New ColTitles
    Dim ParentObjectCell As String: ParentObjectCell = "A4"
    Dim SheetName As String: SheetName = "Pre-VT Object Check"
    
    Worksheets(SheetName).Activate
    
    If ActiveSheet.Range(ParentObjectCell) = "ANALYSIS" Then
        BrowseAnalysis ParentObjectCell, SheetName
    ElseIf ActiveSheet.Range(ParentObjectCell) = "PRODUCT" Then
        BrowseProduct ParentObjectCell, SheetName
    End If
End Sub
Private Sub BrowseAnalysis(ParentObjectCell As String, SheetName As String)
    Dim AnalysisObjectNameCol As String
    Dim SupportingObjectNameCol As String
    Dim SheetStartRow As Double
    Dim V8ObjectString As New ObjectStringBuilder
    Dim V8SubroutineString As New SubStringBuilder
    Dim DataBase As String
    Dim V5TableNames As New TableBuilder
    Dim V8TableNames As New TableBuilder
    Dim TableName As String
    Dim ParentTable As String
    Dim rs As ADODB.Recordset
    Dim AnalysisQuery As New QueryBuilder
    Dim SupportingTableRows As New DbTableRowFinder
    Dim ColTitles As New ColTitles
    
    ColTitles.Init
    AnalysisObjectNameCol = ColTitles.GetStrParentObjs
    SupportingObjectNameCol = ColTitles.GetStrV8ObjName
    SheetStartRow = 7
        
    Worksheets(SheetName).Activate
    ThisWorkbook.ClearReformat SheetName
    ParentTable = ActiveSheet.Range(ParentObjectCell).value
    
    'Locate and add all analyses related through batch links to the objects entered on the sheet
    DataBase = "D3"
    V8ObjectString.Init SheetName, AnalysisObjectNameCol, SheetStartRow
    V8ObjectString.CreateV8ObjectString
    
    ThisWorkbook.FindRelatedAnalyses V8ObjectString.Objects, AnalysisObjectNameCol, DataBase
    
    'Re-build the string of object names after updating column A
    V8ObjectString.Init SheetName, AnalysisObjectNameCol, SheetStartRow
    V8ObjectString.CreateV8ObjectString
    
    'Query calculations table and scrape the names of all subroutines used in all calculations related to the analyses in column A.
    V8SubroutineString.Init SheetName, AnalysisObjectNameCol, SheetStartRow
    V8SubroutineString.FindSourceCode DataBase, V8ObjectString.Objects
    V8SubroutineString.ScrapeSubroutines
    
    'Use ANALYSIS object list to search D3 for all objects and supporting objects and populate them in the sheet.
    V8TableNames.BuildAnalysisTableArr DataBase
    
    For Each i In V8TableNames.TableArr
        TableName = i
        Set rs = Nothing
        
        If Not (TableName = "SUBROUTINE") Then
            AnalysisQuery.BuildAnalysisQuery DataBase, TableName, V8ObjectString.Objects
        Else
            AnalysisQuery.BuildAnalysisQuery DataBase, TableName, V8SubroutineString.Objects
        End If
        ThisWorkbook.GetRecords DataBase, AnalysisQuery.sqlQuery, rs
        
        If rs.BOF And rs.EOF And TableName = ParentTable Then
            MsgBoxes.WrongParentObject
            End
        End If
        
        SetAnalysisRecords rs, DataBase, TableName
    Next i
    
    'Check V5 against select populated D3 objects to ensure all objects exist and are not flagged as removed.
    DataBase = "V5"
    V5TableNames.BuildAnalysisTableArr DataBase
    
    For Each i In V5TableNames.TableArr
        'Function call wouldn't accept variant data type, so I assign "i" to a variable of type compatible with function call.
        TableName = i
        SupportingTableRows.FindSupportingRowNums SheetName, TableName, SupportingObjectNameCol
        If SupportingTableRows.TableFound = True Then
            ThisWorkbook.PopulateComparisonsAgainstD3 ParentTable, SheetName, DataBase, TableName, SupportingObjectNameCol, SupportingTableRows.MinRow, SupportingTableRows.MaxRow
        End If
    Next i
    ThisWorkbook.FillInV5Blanks SheetName, SheetStartRow, SupportingObjectNameCol
    
    'Check V2 against all D3 objects to ensure all objects exist in V2 and are not flagged as removed.
    DataBase = "V2"
    For Each i In V8TableNames.TableArr
        TableName = i
        SupportingTableRows.FindSupportingRowNums SheetName, TableName, SupportingObjectNameCol
        If SupportingTableRows.TableFound = True Then
            ThisWorkbook.PopulateComparisonsAgainstD3 ParentTable, SheetName, DataBase, TableName, SupportingObjectNameCol, SupportingTableRows.MinRow, SupportingTableRows.MaxRow
        End If
    Next i
    
    ThisWorkbook.SetStatus ParentObjectCell, SheetName, ParentTable, SheetStartRow, SupportingObjectNameCol, AnalysisObjectNameCol
    ThisWorkbook.AutoFit SheetName
End Sub
Private Sub BrowseProduct(ParentObjectCell As String, SheetName As String)
    Dim ProductObjectNameCol As String
    Dim SupportingObjectNameCol As String
    Dim SheetStartRow As Double
    Dim V8ObjectString As New ObjectStringBuilder
    Dim DataBase As String
    Dim V5TableNames As New TableBuilder
    Dim V8TableNames As New TableBuilder
    Dim TableName As String
    Dim ParentTable As String
    Dim rs As ADODB.Recordset
    Dim ProductQuery As New QueryBuilder
    Dim SupportingTableRows As New DbTableRowFinder
    Dim ColTitles As New ColTitles
    
    ColTitles.Init
    ProductObjectNameCol = ColTitles.GetStrParentObjs
    SupportingObjectNameCol = ColTitles.GetStrV8ObjName
    SheetStartRow = 7
    
    Worksheets(SheetName).Activate
    ThisWorkbook.ClearReformat SheetName
    ParentTable = ActiveSheet.Range(ParentObjectCell).value
    
    'Use PRODUCT object list to search D3 for all objects and supporting objects and populate them in the sheet.
    DataBase = "D3"
    V8ObjectString.Init SheetName, ProductObjectNameCol, SheetStartRow
    V8ObjectString.CreateV8ObjectString
    V8TableNames.BuildProductTableArr DataBase
    
    For Each i In V8TableNames.TableArr
        TableName = i
        Set rs = Nothing
        ProductQuery.BuildProductQuery DataBase, TableName, V8ObjectString.Objects
        ThisWorkbook.GetRecords DataBase, ProductQuery.sqlQuery, rs
        
        If rs.BOF And rs.EOF And TableName = ParentTable Then
            MsgBoxes.WrongParentObject
            End
        End If
         
        SetProductRecords rs, DataBase, TableName
    Next i
    
    'Check V5 against select populated D3 objects to ensure all objects exist and are not flagged as removed.
    DataBase = "V5"
    V5TableNames.BuildProductTableArr DataBase
    
    For Each i In V5TableNames.TableArr
        'Function call wouldn't accept variant data type, so I assign "i" to a variable of type compatible with function call.
        TableName = i
        SupportingTableRows.FindSupportingRowNums SheetName, TableName, SupportingObjectNameCol
        If SupportingTableRows.TableFound = True Then
            ThisWorkbook.PopulateComparisonsAgainstD3 ParentTable, SheetName, DataBase, TableName, SupportingObjectNameCol, SupportingTableRows.MinRow, SupportingTableRows.MaxRow
        End If
    Next i
    ThisWorkbook.FillInV5Blanks SheetName, SheetStartRow, SupportingObjectNameCol
    
    'Check V2 against all D3 objects to ensure all objects exist in V2 and are not flagged as removed.
    
    DataBase = "V2"
    For Each i In V8TableNames.TableArr
        TableName = i
        SupportingTableRows.FindSupportingRowNums SheetName, TableName, SupportingObjectNameCol
        If SupportingTableRows.TableFound = True Then
            ThisWorkbook.PopulateComparisonsAgainstD3 ParentTable, SheetName, DataBase, TableName, SupportingObjectNameCol, SupportingTableRows.MinRow, SupportingTableRows.MaxRow
        End If
    Next i
    
    ThisWorkbook.SetStatus ParentObjectCell, SheetName, ParentTable, SheetStartRow, SupportingObjectNameCol, ProductObjectNameCol
    ThisWorkbook.AutoFit SheetName
End Sub
Public Sub SetAnalysisRecords(rs As ADODB.Recordset, DataBase As String, TableName As String, Optional MinRow As Double = -1, Optional MaxRow As Integer = -1, Optional ObjNameCol As String = -1)
    Dim i As Integer
    Dim LastRow As Long
    Dim ColTitles As New ColTitles
    Dim V8TableColNum As Long
    Dim V5TableColNum As Long
    Dim ObjNameColNum As Long
    Dim V5ExistsColNum As Long
    Dim D3ExistsColNum As Long
    Dim V2ExistsColNum As Long
    Dim V5ChangedColNum As Long
    Dim D3ChangedColNum As Long
    Dim V2ChangedColNum As Long
    Dim rowNum As Long
    
    ColTitles.Init
    V8TableColNum = ColTitles.GetNumV8TableName
    V5TableColNum = ColTitles.GetNumV5TableName
    ObjNameColNum = ColTitles.GetNumV8ObjName
    V5ExistsColNum = ColTitles.GetNumExistsV5
    D3ExistsColNum = ColTitles.GetNumExistsD3
    V2ExistsColNum = ColTitles.GetNumExistsV2
    V5ChangedColNum = ColTitles.GetNumV5Changed
    D3ChangedColNum = ColTitles.GetNumD3Changed
    V2ChangedColNum = ColTitles.GetNumV2Changed
    
    Select Case DataBase
    Case "V5"
        For rowNum = MinRow To MaxRow
            'Determine appropriate search value given that not all V8 objects begin with "SFW_".
            If Left(ActiveSheet.Cells(rowNum, ObjNameColNum), 4) = "SFW_" Then
                SearchValue = Right(ActiveSheet.Cells(rowNum, ObjNameColNum), (Len(ActiveSheet.Cells(rowNum, ObjNameColNum)) - 4))
            Else
                SearchValue = ActiveSheet.Cells(rowNum, ObjNameColNum)
            End If
            If rs.EOF = True And rs.BOF = True Then
                'Set fields when entire record is empty.
                ActiveSheet.Cells(rowNum, V5ExistsColNum) = "Not found in V5"
                ActiveSheet.Cells(rowNum, V5ChangedColNum) = "Not found in V5"
            Else
                With rs
                    .MoveFirst
                    .Find "NAME = '" & SearchValue & "'"
                    If .EOF = False Then
                        ActiveSheet.Cells(rowNum, V5TableColNum) = TableName
                        SetExistsField rs, rowNum, DataBase, V5ExistsColNum
                        ActiveSheet.Cells(rowNum, V5ChangedColNum) = .Fields("CHANGED_ON")
                    Else
                        'Set fields when individual records don't exist.
                        ActiveSheet.Cells(rowNum, V5ExistsColNum) = "Not found in V5"
                        ActiveSheet.Cells(rowNum, V5ChangedColNum) = "Not found in V5"
                    End If
                End With
            End If
        Next rowNum
    Case "D3"
        Dim NameCol As String
        If rs.BOF And rs.EOF Then
            'Biggest area of improvement that I only realized upon completion: What if an object in their list doesn't exist in D3?
            'Temp fix - I'll set a special status box. If there aren't the same number of product objects returned as queried, it'll trigger.
            'Lazy, but better than nothing given that I'm ready to move on from this.
            Exit Sub
        Else
            LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row
            rowNum = LastRow + 1
            rs.MoveFirst
            
            If Not TableName = "UNITS" Then
                NameCol = "NAME"
            Else
                NameCol = "UNIT_CODE"
            End If
            
            With rs
                If Not TableName = "LIST_ENTRY" Then
                    Do While Not .EOF
                        ActiveSheet.Cells(rowNum, V8TableColNum) = TableName
                        ActiveSheet.Cells(rowNum, ObjNameColNum) = .Fields(NameCol)
                        SetExistsField rs, rowNum, DataBase, D3ExistsColNum
                        ActiveSheet.Cells(rowNum, D3ChangedColNum) = .Fields("CHANGED_ON")
                        .MoveNext
                        rowNum = rowNum + 1
                    Loop
                ElseIf TableName = "LIST_ENTRY" Then
                    Do While Not .EOF
                        ActiveSheet.Cells(rowNum, V8TableColNum) = TableName
                        'I suspect LIST_ENTRY has a composite key. Query is built to assume le.list = 'INST_GRPS'.
                        ActiveSheet.Cells(rowNum, ObjNameColNum) = .Fields(NameCol)
                        ActiveSheet.Cells(rowNum, D3ExistsColNum) = "Yes"
                        ActiveSheet.Cells(rowNum, D3ChangedColNum) = "N/A"
                        .MoveNext
                        rowNum = rowNum + 1
                    Loop
                End If
            End With
        End If
    Case "V2"
    'Goal: Query V2, sort through record sets as in V5 recordsets above, check for existence, and post fields to cells.
        For rowNum = MinRow To MaxRow
            'Determine appropriate search value given that not all V8 objects begin with "SFW_".
            SearchValue = ActiveSheet.Cells(rowNum, ObjNameColNum)
            If rs.EOF = True And rs.BOF = True Then
                'Set fields when entire record is empty.
                ActiveSheet.Cells(rowNum, V2ExistsColNum) = "Not found in V2"
                ActiveSheet.Cells(rowNum, V2ChangedColNum) = "Not found in V2"
            Else
                With rs
                    .MoveFirst
                    
                    If TableName = "UNITS" Then
                        .Find "UNIT_CODE = '" & SearchValue & "'"
                    Else
                        .Find "NAME = '" & SearchValue & "'"
                    End If
                    
                    If .EOF = False Then
                        If TableName = "LIST_ENTRY" Then
                            If .EOF = False Then
                                ActiveSheet.Cells(rowNum, V2ExistsColNum) = "Yes"
                                ActiveSheet.Cells(rowNum, V2ChangedColNum) = "N/A"
                            Else
                                'Set fields when individual records don't exist.
                                ActiveSheet.Cells(rowNum, V2ExistsColNum) = "Not found in V2"
                                ActiveSheet.Cells(rowNum, V2ChangedColNum) = "Not found in V2"
                            End If
                        Else
                        SetExistsField rs, rowNum, DataBase, V2ExistsColNum
                            ActiveSheet.Cells(rowNum, V2ChangedColNum) = .Fields("CHANGED_ON")
                        End If
                    Else
                        'Set fields when individual records don't exist.
                        ActiveSheet.Cells(rowNum, V2ExistsColNum) = "Not found in V2"
                        ActiveSheet.Cells(rowNum, V2ChangedColNum) = "Not found in V2"
                    End If
                End With
            End If
        Next rowNum
    Case Else
        Exit Sub
    End Select
End Sub
Public Sub SetProductRecords(rs As ADODB.Recordset, DataBase As String, TableName As String, Optional MinRow As Double = -1, Optional MaxRow As Integer = -1, Optional ObjNameCol As String = -1)
    Dim i As Integer
    Dim LastRow As Long
    Dim ColTitles As New ColTitles
    Dim V8TableColNum As Long
    Dim V5TableColNum As Long
    Dim ObjNameColNum As Long
    Dim V5ExistsColNum As Long
    Dim D3ExistsColNum As Long
    Dim V2ExistsColNum As Long
    Dim V5ChangedColNum As Long
    Dim D3ChangedColNum As Long
    Dim V2ChangedColNum As Long
    Dim rowNum As Long
    
    ColTitles.Init
    V8TableColNum = ColTitles.GetNumV8TableName
    V5TableColNum = ColTitles.GetNumV5TableName
    ObjNameColNum = ColTitles.GetNumV8ObjName
    V5ExistsColNum = ColTitles.GetNumExistsV5
    D3ExistsColNum = ColTitles.GetNumExistsD3
    V2ExistsColNum = ColTitles.GetNumExistsV2
    V5ChangedColNum = ColTitles.GetNumV5Changed
    D3ChangedColNum = ColTitles.GetNumD3Changed
    V2ChangedColNum = ColTitles.GetNumV2Changed
    
    Select Case DataBase
    Case "V5"
        For rowNum = MinRow To MaxRow
            'Determine appropriate search value given that not all V8 objects begin with "SFW_".
            If Left(ActiveSheet.Cells(rowNum, ObjNameColNum), 4) = "SFW_" Then
                SearchValue = Right(ActiveSheet.Cells(rowNum, ObjNameColNum), (Len(ActiveSheet.Cells(rowNum, ObjNameColNum)) - 4))
            Else
                SearchValue = ActiveSheet.Cells(rowNum, ObjNameColNum)
            End If
            If rs.EOF = True And rs.BOF = True Then
                'Set fields when entire record is empty.
                ActiveSheet.Cells(rowNum, V5ExistsColNum) = "Not found in V5"
                ActiveSheet.Cells(rowNum, V5ChangedColNum) = "Not found in V5"
            Else
                With rs
                    .MoveFirst
                    .Find "NAME = '" & SearchValue & "'"
                    If .EOF = False Then
                        ActiveSheet.Cells(rowNum, V5TableColNum) = TableName
                        SetExistsField rs, rowNum, DataBase, V5ExistsColNum
                        ActiveSheet.Cells(rowNum, V5ChangedColNum) = .Fields("CHANGED_ON")
                    Else
                        ActiveSheet.Cells(rowNum, V5ExistsColNum) = "Not found in V5"
                        ActiveSheet.Cells(rowNum, V5ChangedColNum) = "Not found in V5"
                    End If
                End With
            End If
        Next rowNum
    Case "D3"
        Dim NameCol As String
        If rs.BOF And rs.EOF Then
            Exit Sub
        Else
            LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row
            rowNum = LastRow + 1
            rs.MoveFirst
            
            If Not TableName = "UNITS" Then
                NameCol = "NAME"
            Else
                NameCol = "UNIT_CODE"
            End If
            
            With rs
                Do While Not rs.EOF
                    ActiveSheet.Cells(rowNum, V8TableColNum) = TableName
                    ActiveSheet.Cells(rowNum, ObjNameColNum) = rs.Fields(NameCol)
                    SetExistsField rs, rowNum, DataBase, D3ExistsColNum
                    ActiveSheet.Cells(rowNum, D3ChangedColNum) = rs.Fields("CHANGED_ON")
                    rs.MoveNext
                    rowNum = rowNum + 1
                Loop
            End With
        End If
    Case "V2"
    'Goal: Query V2, sort through record sets as in V5 recordsets above, check for existence, and post fields to cells.
        For rowNum = MinRow To MaxRow
            'Determine appropriate search value given that not all V8 objects begin with "SFW_".
            SearchValue = ActiveSheet.Cells(rowNum, ObjNameColNum)
            If rs.EOF = True And rs.BOF = True Then
                'Set fields when entire record is empty.
                ActiveSheet.Cells(rowNum, V2ExistsColNum) = "Not found in V2"
                ActiveSheet.Cells(rowNum, V2ChangedColNum) = "Not found in V2"
            Else
                With rs
                    .MoveFirst
                    If TableName = "UNITS" Then
                        .Find "UNIT_CODE = '" & SearchValue & "'"
                    Else
                        .Find "NAME = '" & SearchValue & "'"
                    End If
                    
                    If .EOF = False Then
                        SetExistsField rs, rowNum, DataBase, V2ExistsColNum
                        ActiveSheet.Cells(rowNum, V2ChangedColNum) = .Fields("CHANGED_ON")
                    Else
                        ActiveSheet.Cells(rowNum, V2ExistsColNum) = "Not found in V2"
                        ActiveSheet.Cells(rowNum, V2ChangedColNum) = "Not found in V2"
                    End If
                End With
            End If
        Next rowNum
    Case Else
        Exit Sub
    End Select
End Sub
Public Function SetExistsField(rs As ADODB.Recordset, rowNum As Long, DataBase As String, ExistsColNum As Long)
    Dim ActiveFieldExist: Set ActiveFieldExist = Nothing
    
    'Test whether recordset contains "ACTIVE" field and set "Exists in <DATABASE NAME>?" column.
    On Error Resume Next
    Set ActiveFieldExist = rs("ACTIVE")
    On Error GoTo 0
    
    With rs
        If ActiveFieldExist Is Nothing Then
            If (.Fields("REMOVED") = "F") Then
                ActiveSheet.Cells(rowNum, ExistsColNum) = "Yes"
            ElseIf (.Fields("REMOVED") = "T") Then
                ActiveSheet.Cells(rowNum, ExistsColNum) = "Removed in " & DataBase
            End If
        Else
            If (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "T") Then
                'Most likely scenario, and most likely to be short-circuited by "REMOVED".
                ActiveSheet.Cells(rowNum, ExistsColNum) = "Yes"
            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "T") Then
                ActiveSheet.Cells(rowNum, ExistsColNum) = "Removed in " & DataBase
            ElseIf (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "F") Then
                ActiveSheet.Cells(rowNum, ExistsColNum) = "Inactive in " & DataBase
            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "F") Then
                ActiveSheet.Cells(rowNum, ExistsColNum) = "Removed AND Inactive in " & DataBase
            End If
        End If
    End With
End Function
