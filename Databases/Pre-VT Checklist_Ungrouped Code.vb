'Written by: N.Whisman nathan.whisman@pfizer.com 08Mar2024 v1.0 finished
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
'***********************************************************************************************************************************************************************

Private Sub BrowseButton_Click()
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
    
    AnalysisObjectNameCol = "A"
    SupportingObjectNameCol = "E"
    SheetStartRow = 7
        
    Worksheets(SheetName).Activate
    ThisWorkbook.ClearReformat SheetName
    ParentTable = ActiveSheet.Range(ParentObjectCell).Value
    
    'Locate and add all analyses related through batch links to the objects entered on the sheet
    DataBase = "D3"
    V8ObjectString.Init SheetName, AnalysisObjectNameCol, SheetStartRow
    V8ObjectString.CreateV8ObjectString
    
    ThisWorkbook.FindRelatedAnalyses V8ObjectString.Objects, AnalysisObjectNameCol, DataBase
    
    'Re-build the string of object names after updating column A
    V8ObjectString.Init SheetName, AnalysisObjectNameCol, SheetStartRow
    V8ObjectString.CreateV8ObjectString
    
    'THIS IS WHERE I'LL USE THE OBJECT LIST TO FIND ANY SUBROUTINES I MAY NEED TO PULL. MAKE A NEW CLASS WITH AN ATTRIBUTE THAT IS THE STRING.
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
        
    ProductObjectNameCol = "A"
    SupportingObjectNameCol = "E"
    SheetStartRow = 7
    
    Worksheets(SheetName).Activate
    ThisWorkbook.ClearReformat SheetName
    ParentTable = ActiveSheet.Range(ParentObjectCell).Value
    
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
    Dim ObjNameColNum As Long
    Dim ActiveFieldExist: Set ActiveFieldExist = Nothing
    
    Select Case DataBase
    Case "V5"
        'Convert ObjNameCol to Long for use in .Cells()
        ObjNameColNum = Range(ObjNameCol & 1).Column
        For rowNum = MinRow To MaxRow
            'Determine appropriate search value given that not all V8 objects begin with "SFW_".
            If Left(ActiveSheet.Cells(rowNum, ObjNameColNum), 4) = "SFW_" Then
                SearchValue = Right(ActiveSheet.Cells(rowNum, ObjNameColNum), (Len(ActiveSheet.Cells(rowNum, ObjNameColNum)) - 4))
            Else
                SearchValue = ActiveSheet.Cells(rowNum, ObjNameColNum)
            End If
            If rs.EOF = True And rs.BOF = True Then
                'Set fields when entire record is empty.
                ActiveSheet.Cells(rowNum, 6) = "Not found in V5"
                ActiveSheet.Cells(rowNum, 9) = "Not found in V5"
            Else
                With rs
                    .MoveFirst
                    .Find "NAME = '" & SearchValue & "'"
                    If .EOF = False Then
                        ActiveSheet.Cells(rowNum, 4) = TableName
                        
                        'Test whether recordset contains "ACTIVE" field and set "Exists in V5?" column.
                        On Error Resume Next
                        Set ActiveFieldExist = rs("ACTIVE")
                        On Error GoTo 0
                        If ActiveFieldExist Is Nothing Then
                            If (.Fields("REMOVED") = "F") Then
                                ActiveSheet.Cells(rowNum, 6) = "Yes"
                            ElseIf (.Fields("REMOVED") = "T") Then
                                ActiveSheet.Cells(rowNum, 6) = "Removed in V5"
                            End If
                        Else
                            If (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "T") Then
                                'Most likely scenario, and most likely to be short-circuited by "REMOVED".
                                ActiveSheet.Cells(rowNum, 6) = "Yes"
                            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "T") Then
                                ActiveSheet.Cells(rowNum, 6) = "Removed in V5"
                            ElseIf (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "F") Then
                                ActiveSheet.Cells(rowNum, 6) = "Inactive in V5"
                            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "F") Then
                                ActiveSheet.Cells(rowNum, 6) = "Removed AND Inactive in V5"
                            End If
                        End If
                        
                        ActiveSheet.Cells(rowNum, 9) = .Fields("CHANGED_ON")
                    Else
                        'Set fields when individual records don't exist.
                        ActiveSheet.Cells(rowNum, 6) = "Not found in V5"
                        ActiveSheet.Cells(rowNum, 9) = "Not found in V5"
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
            i = LastRow + 1
            rs.MoveFirst
            
            If Not TableName = "UNITS" Then
                NameCol = "NAME"
            Else
                NameCol = "UNIT_CODE"
            End If
            
            With rs
                If Not TableName = "LIST_ENTRY" Then
                    Do While Not .EOF
                        ActiveSheet.Cells(i, 3) = TableName
                        ActiveSheet.Cells(i, 5) = .Fields(NameCol)
                        
                        'Test whether recordset contains "ACTIVE" field and set "Exists in V2?" column.
                        On Error Resume Next
                        Set ActiveFieldExist = rs("ACTIVE")
                        On Error GoTo 0
                        If ActiveFieldExist Is Nothing Then
                            If (.Fields("REMOVED") = "F") Then
                                ActiveSheet.Cells(i, 7) = "Yes"
                            ElseIf (.Fields("REMOVED") = "T") Then
                                ActiveSheet.Cells(i, 7) = "Removed in D3"
                            End If
                        Else
                            If (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "T") Then
                                'Most likely scenario, and most likely to be short-circuited by "REMOVED".
                                ActiveSheet.Cells(i, 7) = "Yes"
                            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "T") Then
                                ActiveSheet.Cells(i, 7) = "Removed in D3"
                            ElseIf (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "F") Then
                                ActiveSheet.Cells(i, 7) = "Inactive in D3"
                            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "F") Then
                                ActiveSheet.Cells(i, 7) = "Removed AND Inactive in D3"
                            End If
                        End If
                        ActiveSheet.Cells(i, 10) = .Fields("CHANGED_ON")
                        .MoveNext
                        i = i + 1
                    Loop
                ElseIf TableName = "LIST_ENTRY" Then
                    Do While Not .EOF
                        ActiveSheet.Cells(i, 3) = TableName
                        'I suspect LIST_ENTRY has a composite key. Query is built to assume le.list = 'INST_GRPS'.
                        ActiveSheet.Cells(i, 5) = .Fields(NameCol)
                        ActiveSheet.Cells(i, 7) = "Yes"
                        ActiveSheet.Cells(i, 10) = "N/A"
                        .MoveNext
                        i = i + 1
                    Loop
                End If
            End With
        End If
    Case "V2"
    'Goal: Query V2, sort through record sets as in V5 recordsets above, check for existence, and post fields to cells.
        ObjNameColNum = Range(ObjNameCol & 1).Column
        For rowNum = MinRow To MaxRow
            'Determine appropriate search value given that not all V8 objects begin with "SFW_".
            SearchValue = ActiveSheet.Cells(rowNum, ObjNameColNum)
            If rs.EOF = True And rs.BOF = True Then
                'Set fields when entire record is empty.
                ActiveSheet.Cells(rowNum, 8) = "Not found in V2"
                ActiveSheet.Cells(rowNum, 11) = "Not found in V2"
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
                                ActiveSheet.Cells(rowNum, 8) = "Yes"
                                ActiveSheet.Cells(rowNum, 11) = "N/A"
                            Else
                                'Set fields when individual records don't exist.
                                ActiveSheet.Cells(rowNum, 8) = "Not found in V2"
                                ActiveSheet.Cells(rowNum, 11) = "Not found in V2"
                            End If
                        Else
                        'Test whether recordset contains "ACTIVE" field and set "Exists in V2?" column.
                            On Error Resume Next
                            Set ActiveFieldExist = rs("ACTIVE")
                            On Error GoTo 0
                            If ActiveFieldExist Is Nothing Then
                                If (.Fields("REMOVED") = "F") Then
                                    ActiveSheet.Cells(rowNum, 8) = "Yes"
                                ElseIf (.Fields("REMOVED") = "T") Then
                                    ActiveSheet.Cells(rowNum, 8) = "Removed in V2"
                                End If
                            Else
                                If (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "T") Then
                                    'Most likely scenario, and most likely to be short-circuited by "REMOVED".
                                    ActiveSheet.Cells(rowNum, 8) = "Yes"
                                ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "T") Then
                                    ActiveSheet.Cells(rowNum, 8) = "Removed in V2"
                                ElseIf (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "F") Then
                                    ActiveSheet.Cells(rowNum, 8) = "Inactive in V2"
                                ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "F") Then
                                    ActiveSheet.Cells(rowNum, 8) = "Removed AND Inactive in V2"
                                End If
                            End If
                            
                            ActiveSheet.Cells(rowNum, 11) = .Fields("CHANGED_ON")
                        End If
                    Else
                        'Set fields when individual records don't exist.
                        ActiveSheet.Cells(rowNum, 8) = "Not found in V2"
                        ActiveSheet.Cells(rowNum, 11) = "Not found in V2"
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
    Dim ObjNameColNum As Long
    Dim ActiveFieldExist: Set ActiveFieldExist = Nothing
    
    Select Case DataBase
    Case "V5"
        'Convert ObjNameCol to Long for use in .Cells()
        ObjNameColNum = Range(ObjNameCol & 1).Column
        For rowNum = MinRow To MaxRow
            'Determine appropriate search value given that not all V8 objects begin with "SFW_".
            If Left(ActiveSheet.Cells(rowNum, ObjNameColNum), 4) = "SFW_" Then
                SearchValue = Right(ActiveSheet.Cells(rowNum, ObjNameColNum), (Len(ActiveSheet.Cells(rowNum, ObjNameColNum)) - 4))
            Else
                SearchValue = ActiveSheet.Cells(rowNum, ObjNameColNum)
            End If
            If rs.EOF = True And rs.BOF = True Then
                'Set fields when entire record is empty.
                ActiveSheet.Cells(rowNum, 6) = "Not found in V5"
                ActiveSheet.Cells(rowNum, 9) = "Not found in V5"
            Else
                With rs
                    .MoveFirst
                    .Find "NAME = '" & SearchValue & "'"
                    If .EOF = False Then
                        ActiveSheet.Cells(rowNum, 4) = TableName
                        
                        'Test whether recordset contains "ACTIVE" field and set "Exists in V5?" column.
                        On Error Resume Next
                        Set ActiveFieldExist = rs("ACTIVE")
                        On Error GoTo 0
                        If ActiveFieldExist Is Nothing Then
                            If (.Fields("REMOVED") = "F") Then
                                ActiveSheet.Cells(rowNum, 6) = "Yes"
                            ElseIf (.Fields("REMOVED") = "T") Then
                                ActiveSheet.Cells(rowNum, 6) = "Removed in V5"
                            End If
                        Else
                            If (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "T") Then
                                'Most likely scenario, and most likely to be short-circuited by "REMOVED".
                                ActiveSheet.Cells(rowNum, 6) = "Yes"
                            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "T") Then
                                ActiveSheet.Cells(rowNum, 6) = "Removed in V5"
                            ElseIf (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "F") Then
                                ActiveSheet.Cells(rowNum, 6) = "Inactive in V5"
                            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "F") Then
                                ActiveSheet.Cells(rowNum, 6) = "Removed AND Inactive in V5"
                            End If
                        End If
                        
                        ActiveSheet.Cells(rowNum, 9) = .Fields("CHANGED_ON")
                    Else
                        ActiveSheet.Cells(rowNum, 6) = "Not found in V5"
                        ActiveSheet.Cells(rowNum, 9) = "Not found in V5"
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
            i = LastRow + 1
            rs.MoveFirst
            
            If Not TableName = "UNITS" Then
                NameCol = "NAME"
            Else
                NameCol = "UNIT_CODE"
            End If
            
            Do While Not rs.EOF
                ActiveSheet.Cells(i, 3) = TableName
                ActiveSheet.Cells(i, 5) = rs.Fields(NameCol)
                If rs.Fields("REMOVED") = "F" Then
                    ActiveSheet.Cells(i, 7) = "Yes"
                ElseIf rs.Fields("REMOVED") = "T" Then
                    ActiveSheet.Cells(i, 7) = "Removed in D3"
                End If
                ActiveSheet.Cells(i, 10) = rs.Fields("CHANGED_ON")
                rs.MoveNext
                i = i + 1
            Loop
        End If
    Case "V2"
    'Goal: Query V2, sort through record sets as in V5 recordsets above, check for existence, and post fields to cells.
        ObjNameColNum = Range(ObjNameCol & 1).Column
        For rowNum = MinRow To MaxRow
            'Determine appropriate search value given that not all V8 objects begin with "SFW_".
            SearchValue = ActiveSheet.Cells(rowNum, ObjNameColNum)
            If rs.EOF = True And rs.BOF = True Then
                'Set fields when entire record is empty.
                ActiveSheet.Cells(rowNum, 8) = "Not found in V2"
                ActiveSheet.Cells(rowNum, 11) = "Not found in V2"
            Else
                With rs
                    .MoveFirst
                    If TableName = "UNITS" Then
                        .Find "UNIT_CODE = '" & SearchValue & "'"
                    Else
                        .Find "NAME = '" & SearchValue & "'"
                    End If
                    
                    If .EOF = False Then
                        'Test whether recordset contains "ACTIVE" field and set "Exists in V2?" column.
                        On Error Resume Next
                        Set ActiveFieldExist = rs("ACTIVE")
                        On Error GoTo 0
                        If ActiveFieldExist Is Nothing Then
                            If (.Fields("REMOVED") = "F") Then
                                ActiveSheet.Cells(rowNum, 8) = "Yes"
                            ElseIf (.Fields("REMOVED") = "T") Then
                                ActiveSheet.Cells(rowNum, 8) = "Removed in V2"
                            End If
                        Else
                            If (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "T") Then
                                'Most likely scenario, and most likely to be short-circuited by "REMOVED".
                                ActiveSheet.Cells(rowNum, 8) = "Yes"
                            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "T") Then
                                ActiveSheet.Cells(rowNum, 8) = "Removed in V2"
                            ElseIf (.Fields("REMOVED") = "F") And (.Fields("ACTIVE") = "F") Then
                                ActiveSheet.Cells(rowNum, 8) = "Inactive in V2"
                            ElseIf (.Fields("REMOVED") = "T") And (.Fields("ACTIVE") = "F") Then
                                ActiveSheet.Cells(rowNum, 8) = "Removed AND Inactive in V2"
                            End If
                        End If
                        
                        ActiveSheet.Cells(rowNum, 11) = .Fields("CHANGED_ON")
                    Else
                        ActiveSheet.Cells(rowNum, 8) = "Not found in V2"
                        ActiveSheet.Cells(rowNum, 11) = "Not found in V2"
                    End If
                End With
            End If
        Next rowNum
    Case Else
        Exit Sub
    End Select
End Sub
'Written by: N.Whisman nathan.whisman@pfizer.com 25Feb2024
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
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
    ActiveSheet.Range("K1:L1").Interior.Color = RGB(217, 217, 217)
    ActiveSheet.Range("K1:L1").Value = ""
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
        Sheet4.SetProductRecords rs, DataBase, TableName, MinRow, MaxRow, ObjNameCol
    ElseIf ParentTable = "ANALYSIS" Then
        ProductQuery.BuildAnalysisQuery DataBase, TableName, D3CompObjectString.Objects
        ThisWorkbook.GetRecords DataBase, ProductQuery.sqlQuery, rs
        Sheet4.SetAnalysisRecords rs, DataBase, TableName, MinRow, MaxRow, ObjNameCol
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
            If IsEmpty(Cell.Value) Then
                Cell.Value = "N/A"
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
                If ActiveSheet.Cells(rowNum, 9) > ActiveSheet.Cells(rowNum, 10) Then
                    ActiveSheet.Cells(rowNum, 12) = ActiveSheet.Cells(rowNum, 12) & "Object is newer in V5 than in D3, "
                End If
            End If
            If (InStr(ActiveSheet.Cells(rowNum, 11), "Not found") = 0) And (InStr(ActiveSheet.Cells(rowNum, 11), "N/A") = 0) Then
                If ActiveSheet.Cells(rowNum, 10) > ActiveSheet.Cells(rowNum, 11) Then
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
'Written by: N.Whisman nathan.whisman@pfizer.com 08Mar2024
'
'*****************************************************************************************

Public Sub NoDataEntered()
    MsgBox "No data found. Enter object names in column A beginning with cell A7."
End Sub
Public Sub WrongParentObject()
    MsgBox "No objects found in D3. Make sure the correct object type was selected in cell A4."
End Sub
Public Sub DbConnFail()
    MsgBox "Could not connect to the database. Try checking your connection manually."
End Sub
'Written by: N.Whisman nathan.whisman@pfizer.com 25Feb2024
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
    SelectDatabase "D3"
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
    connectionString
Case "D3"
    connectionString
Case "V2"
    connectionString
Case Else
    connectionString
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
Conn.Open
End Function
Private Function CloseConnection()
'Closes database connection if open

If Not Conn Is Nothing Then
    Conn.Close
End If
End Function
'Written by: N.Whisman nathan.whisman@pfizer.com 25Feb2024
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
            If Cell.Value <> "" Then
                If Left(Cell.Value, 4) = "SFW_" Then
                    Me.Objects = Me.Objects & Right(Cell.Value, Len(Cell.Value) - 4) & "', '"
                Else
                    Me.Objects = Me.Objects & Cell.Value & "', '"
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
        If Cell.Value <> "" Then
            Me.Objects = Me.Objects & Cell.Value & "', '"
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
'Written by: N.Whisman nathan.whisman@pfizer.com 25Feb2024
'Class contains methods for selecting a query and storing it.
'
'*****************************************************************************************
Public sqlQuery As String

Public Function BuildProductQuery(DataBase As String, TableName As String, ObjectList As String)

DataBase = UCase(DataBase)
TableName = UCase(TableName)

If DataBase = "V5" Then
'If the number of supported queries is expanded, make sure all fields in the chosen query exist on the new table.
    If TableName = "PRODUCT" Or TableName = "ANALYSIS" Or TableName = "T_PH_SAMPLE_PLAN" Then
        Me.sqlQuery = _
        "SELECT DISTINCT " & TableName & ".name, " & TableName & ".removed, " & TableName & ".changed_on, " & TableName & ".active  " & _
        "FROM " & TableName & "  " & _
            "INNER JOIN versions v  " & _
                "ON v.table_name = '" & TableName & "'  " & _
                "AND " & TableName & ".name = v.name  " & _
                "AND " & TableName & ".version = v.version  " & _
        "WHERE " & TableName & ".name IN " & ObjectList & "  " & _
        "ORDER BY " & TableName & ".name "
    ElseIf TableName = "TEST_LIST" Or TableName = "SAMPLING_POINT" Then
        Me.sqlQuery = _
        "SELECT DISTINCT " & TableName & ".name, " & TableName & ".removed, " & TableName & ".changed_on  " & _
        "FROM " & TableName & "  " & _
        "WHERE " & TableName & ".name IN " & ObjectList & "  " & _
        "ORDER BY " & TableName & ".name "
    End If
ElseIf DataBase = "D3" Then
    Select Case TableName
    'Simplification would be a pain because all queries run a sub-query of PRODUCT, which has a different field name than the table name for half of the object types needed
    Case "PRODUCT"
        Me.sqlQuery = _
        "SELECT DISTINCT p.name, p.removed, p.changed_on, p.active  " & _
        "FROM product p " & _
            "INNER JOIN versions v  " & _
                "ON v.table_name = 'PRODUCT'  " & _
                "AND p.name = v.name  " & _
                "AND p.version = v.version  " & _
        "WHERE p.name IN " & ObjectList & "  " & _
        "ORDER BY p.name "
    Case "T_PH_ITEM_CODE"
        Me.sqlQuery = _
        "SELECT DISTINCT ic.name, ic.removed, ic.changed_on, ic.active " & _
        "FROM t_ph_item_code ic " & _
            "INNER JOIN versions v " & _
                "ON v.table_name = 'T_PH_ITEM_CODE' " & _
                "AND ic.name = v.name " & _
                "AND ic.version = v.version " & _
        "WHERE ic.display_as IN ( " & _
            "SELECT p.code " & _
            "FROM product p " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND p.name = v.name " & _
                "AND p.version = v.version " & _
            "WHERE p.name IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY ic.name "
    Case "TEST_LIST"
        Me.sqlQuery = _
        "SELECT DISTINCT tl.name, tl.removed, tl.changed_on " & _
        "FROM test_list tl " & _
        "WHERE tl.name IN ( " & _
            "SELECT p.test_list " & _
            "FROM product p " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND p.name = v.name " & _
                "AND p.version = v.version " & _
            "WHERE p.name IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY tl.name "
    Case "SAMPLING_POINT"
        Me.sqlQuery = _
        "SELECT DISTINCT sp.name, sp.removed, sp.changed_on " & _
        "FROM sampling_point sp " & _
        "WHERE sp.name IN ( " & _
            "SELECT pgs.sampling_point " & _
            "FROM prod_grade_stage pgs " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND pgs.product = v.name " & _
                "AND pgs.version = v.version " & _
            "WHERE pgs.product IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY sp.name "
    Case "T_PH_GRADE"
        Me.sqlQuery = _
        "SELECT DISTINCT g.name, g.removed, g.changed_on " & _
        "FROM t_ph_grade g " & _
        "WHERE g.name IN ( " & _
            "SELECT pgs.grade " & _
            "FROM prod_grade_stage pgs " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND pgs.product = v.name " & _
                "AND pgs.version = v.version " & _
            "WHERE pgs.product IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY g.name "
    Case "T_PH_STAGE"
        Me.sqlQuery = _
        "SELECT DISTINCT s.name, s.removed, s.changed_on " & _
        "FROM t_ph_stage s " & _
        "WHERE s.name IN ( " & _
            "SELECT pgs.stage " & _
            "FROM prod_grade_stage pgs " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND pgs.product = v.name " & _
                "AND pgs.version = v.version " & _
            "WHERE pgs.product IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY s.name "
    Case "ANALYSIS"
        Me.sqlQuery = _
        "SELECT DISTINCT a.name, a.removed, a.changed_on, a.active " & _
        "FROM analysis a " & _
        "WHERE a.name IN ( " & _
            "SELECT pgs.analysis " & _
            "FROM prod_grade_stage pgs " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND pgs.product = v.name " & _
                "AND pgs.version = v.version " & _
            "WHERE pgs.product IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY a.name "
    Case "T_PH_SPEC_TYPE"
        Me.sqlQuery = _
        "SELECT DISTINCT st.name, st.removed, st.changed_on " & _
        "FROM t_ph_spec_type st " & _
        "WHERE st.name IN ( " & _
            "SELECT pgs.spec_type " & _
            "FROM prod_grade_stage pgs " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND pgs.product = v.name " & _
                "AND pgs.version = v.version " & _
            "WHERE pgs.product IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY st.name "
    Case "TEST_LOCATION"
        Me.sqlQuery = _
        "SELECT DISTINCT tl.name, tl.removed, tl.changed_on " & _
        "FROM test_location tl " & _
        "WHERE tl.name IN ( " & _
            "SELECT pgs.test_location " & _
            "FROM prod_grade_stage pgs " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND pgs.product = v.name " & _
                "AND pgs.version = v.version " & _
            "WHERE pgs.product IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY tl.name "
    Case "CONDITION"
        Me.sqlQuery = _
        "SELECT DISTINCT c.name, c.removed, c.changed_on " & _
        "FROM condition c " & _
        "WHERE c.name IN ( " & _
            "SELECT pgs.c_stor_cond " & _
            "FROM prod_grade_stage pgs " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND pgs.product = v.name " & _
                "AND pgs.version = v.version " & _
            "WHERE pgs.product IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY c.name "
    Case "UNITS"
        Me.sqlQuery = _
        "SELECT DISTINCT u.unit_code, u.removed, u.changed_on " & _
        "FROM units u " & _
        "WHERE u.unit_code IN ( " & _
            "SELECT pgs.c_units " & _
            "FROM prod_grade_stage pgs " & _
                "INNER JOIN versions v " & _
                "ON v.table_name = 'PRODUCT' " & _
                "AND pgs.product = v.name " & _
                "AND pgs.version = v.version " & _
            "WHERE pgs.product IN " & ObjectList & "  " & _
            ") " & _
        "ORDER BY u.unit_code "
    Case "T_PH_SAMPLE_PLAN"
        Me.sqlQuery = _
        "SELECT sp.name, sp.removed, sp.changed_on, sp.active " & _
        "FROM t_ph_sample_plan sp " & _
            "INNER JOIN versions v " & _
                "ON v.table_name = 'T_PH_SAMPLE_PLAN' " & _
                "AND sp.name = v.name " & _
                "AND sp.version = v.version " & _
        "WHERE sp.name IN " & ObjectList & "  " & _
        "ORDER BY sp.name "
    End Select
ElseIf DataBase = "V2" Then
    If TableName = "PRODUCT" Or TableName = "T_PH_ITEM_CODE" Or TableName = "ANALYSIS" Or TableName = "T_PH_SAMPLE_PLAN" Then
    'Query for versioned objects from their native table where "NAME" is the key field
        Me.sqlQuery = _
        "SELECT DISTINCT " & TableName & ".name, " & TableName & ".removed, " & TableName & ".changed_on, " & TableName & ".active " & _
        "FROM " & TableName & " " & _
            "INNER JOIN versions v " & _
                "ON v.table_name = '" & TableName & "' " & _
                "AND " & TableName & ".name = v.name " & _
                "AND " & TableName & ".version = v.version " & _
        "WHERE " & TableName & ".name IN " & ObjectList & "  " & _
        "ORDER BY " & TableName & ".name "
    ElseIf TableName = "TEST_LIST" Or TableName = "SAMPLING_POINT" Or TableName = "T_PH_GRADE" Or TableName = "T_PH_STAGE" _
    Or TableName = "T_PH_SPEC_TYPE" Or TableName = "TEST_LOCATION" Or TableName = "CONDITION" Then
    'Query for non-versioned objects from their native table where "NAME" is the key field
        Me.sqlQuery = _
        "SELECT DISTINCT " & TableName & ".name, " & TableName & ".removed, " & TableName & ".changed_on " & _
        "FROM " & TableName & " " & _
        "WHERE " & TableName & ".name IN " & ObjectList & "  " & _
        "ORDER BY " & TableName & ".name "
    ElseIf TableName = "UNITS" Then
    'Query for UNITS; "UNIT_CODE" is the key field (as opposed to "NAME")
        Me.sqlQuery = _
        "SELECT DISTINCT u.unit_code, u.removed, u.changed_on " & _
        "FROM units u " & _
        "WHERE u.unit_code IN " & ObjectList & "  " & _
        "ORDER BY u.unit_code "
    End If
End If
End Function
Public Function BuildAnalysisQuery(DataBase As String, TableName As String, ObjectList As String)

DataBase = UCase(DataBase)
TableName = UCase(TableName)

If DataBase = "V5" Then
'If the number of supported queries is expanded, make sure all fields in the chosen query exist on the new table.
    If TableName = "ANALYSIS" Then
        Me.sqlQuery = _
        "SELECT DISTINCT " & TableName & ".name, " & TableName & ".removed, " & TableName & ".changed_on, " & TableName & ".active  " & _
        "FROM " & TableName & "  " & _
            "INNER JOIN versions v  " & _
                "ON v.table_name = '" & TableName & "'  " & _
                "AND " & TableName & ".name = v.name  " & _
                "AND " & TableName & ".version = v.version  " & _
        "WHERE " & TableName & ".name IN " & ObjectList & "  " & _
        "ORDER BY " & TableName & ".name "
    ElseIf TableName = "BATCH_LINK" Or TableName = "STD_REAG_TEMP" Then
        Me.sqlQuery = _
        "SELECT DISTINCT " & TableName & ".name, " & TableName & ".removed, " & TableName & ".changed_on  " & _
        "FROM " & TableName & "  " & _
        "WHERE " & TableName & ".name IN " & ObjectList & "  " & _
        "ORDER BY " & TableName & ".name "
    End If
ElseIf DataBase = "D3" Then
    Select Case TableName
    Case "ANALYSIS"
        Me.sqlQuery = _
        "SELECT DISTINCT a.name, a.removed, a.active, a.changed_on " & _
        "FROM analysis a " & _
            "INNER JOIN versions v " & _
                "ON v.table_name = 'ANALYSIS' " & _
                "AND a.name = v.name " & _
                "AND a.version = v.version " & _
        "WHERE a.name IN " & ObjectList & " " & _
        "ORDER BY a.name "
    Case "COMMON_NAME"
        Me.sqlQuery = _
        "SELECT cn.name, cn.removed, cn.changed_on " & _
        "FROM common_name cn " & _
        "WHERE cn.name IN ( " & _
            "SELECT DISTINCT a.common_name " & _
            "FROM analysis a " & _
                "INNER JOIN versions v " & _
                    "ON v.table_name = 'ANALYSIS' " & _
                    "AND a.name = v.name " & _
                    "AND a.version = v.version " & _
            "WHERE a.name IN " & ObjectList & " " & _
            ") " & _
        "ORDER BY cn.name "
    Case "ANALYSIS_TYPES"
        Me.sqlQuery = _
        "SELECT at.name, at.removed, at.changed_on " & _
        "FROM analysis_types at " & _
        "WHERE at.name IN ( " & _
            "SELECT DISTINCT a.analysis_type " & _
            "FROM analysis a " & _
                "INNER JOIN versions v " & _
                    "ON v.table_name = 'ANALYSIS' " & _
                    "AND a.name = v.name " & _
                    "AND a.version = v.version " & _
            "WHERE a.name IN " & ObjectList & " " & _
            ") " & _
        "ORDER BY at.name "
    Case "BATCH_LINK"
        Me.sqlQuery = _
        "SELECT bl.name, bl.removed, bl.changed_on " & _
        "FROM ( " & _
            "SELECT bl.name, bl.removed, bl.changed_on " & _
            "FROM batch_link bl " & _
            "WHERE bl.name IN ( " & _
                "SELECT DISTINCT a.batch_link " & _
                "FROM analysis a " & _
                    "INNER JOIN versions v " & _
                        "ON v.table_name = 'ANALYSIS' AND a.name = v.name AND a.version = v.version " & _
                "WHERE a.name IN " & ObjectList & " " & _
                ") " & _
            "Union ALL " & _
            "SELECT bl.name, bl.removed, bl.changed_on " & _
            "FROM batch_link bl " & _
            "WHERE bl.name IN ( " & _
                "SELECT DISTINCT av.batch_link " & _
                "FROM analysis_variation av " & _
                    "INNER JOIN versions v " & _
                        "ON v.table_name = 'ANALYSIS' AND av.analysis = v.name AND av.version = v.version " & _
                "WHERE av.analysis IN " & ObjectList & " " & _
                ") " & _
        ") bl " & _
        "ORDER BY bl.name "
    Case "UNITS"
        Me.sqlQuery = _
        "SELECT u.unit_code, u.removed, u.changed_on " & _
        "FROM ( " & _
            "SELECT u.unit_code, u.removed, u.changed_on " & _
            "FROM units u " & _
            "WHERE u.unit_code IN ( " & _
                "SELECT DISTINCT cv.units " & _
                "FROM comp_variation cv " & _
                    "INNER JOIN versions v " & _
                        "ON v.table_name = 'ANALYSIS' AND cv.analysis = v.name AND cv.version = v.version " & _
                "WHERE cv.analysis IN " & ObjectList & " " & _
                ") " & _
            "Union ALL " & _
            "SELECT u.unit_code, u.removed, u.changed_on " & _
            "FROM units u " & _
            "WHERE u.unit_code IN ( " & _
                "SELECT DISTINCT c.units " & _
                "FROM component c " & _
                    "INNER JOIN versions v " & _
                        "ON v.table_name = 'ANALYSIS' AND c.analysis = v.name AND c.version = v.version " & _
                "WHERE c.analysis IN " & ObjectList & " " & _
                ") " & _
        ") u " & _
        "ORDER BY u.unit_code "
    Case "FORMAT_CALCULATION"
        Me.sqlQuery = _
        "SELECT fc.name, fc.removed, fc.changed_on " & _
        "FROM format_calculation fc " & _
        "WHERE fc.name IN ( " & _
            "SELECT DISTINCT c.format_calculation " & _
            "FROM component c " & _
                "INNER JOIN versions v " & _
                    "ON v.table_name = 'ANALYSIS' " & _
                    "AND c.analysis = v.name " & _
                    "AND c.version = v.version " & _
            "WHERE c.analysis IN " & ObjectList & " " & _
            ") " & _
        "ORDER BY fc.name "
    Case "LIST"
        Me.sqlQuery = _
        "SELECT l.name, l.removed, l.changed_on " & _
        "FROM list l " & _
        "WHERE l.name IN ( " & _
            "SELECT DISTINCT c.list_key " & _
            "FROM component c " & _
                "INNER JOIN versions v " & _
                    "ON v.table_name = 'ANALYSIS' " & _
                    "AND c.analysis = v.name " & _
                    "AND c.version = v.version " & _
            "WHERE c.analysis IN " & ObjectList & " " & _
            ") " & _
        "ORDER BY l.name "
    Case "STD_REAG_TEMP"
        Me.sqlQuery = _
        "SELECT s.name, s.removed, s.changed_on " & _
        "FROM std_reag_temp s " & _
        "WHERE s.name IN ( " & _
            "SELECT DISTINCT c.std_reag_template " & _
            "FROM component c " & _
                "INNER JOIN versions v " & _
                    "ON v.table_name = 'ANALYSIS' " & _
                    "AND c.analysis = v.name " & _
                    "AND c.version = v.version " & _
            "WHERE c.analysis IN " & ObjectList & " " & _
            ") " & _
        "ORDER BY s.name "
    Case "LIST_ENTRY"
        Me.sqlQuery = _
        "SELECT le.name " & _
        "FROM list_entry le " & _
        "WHERE le.list = 'INST_GRPS' " & _
            "AND le.name IN ( " & _
            "SELECT DISTINCT a.inst_group " & _
            "FROM analysis a " & _
            "WHERE a.name IN " & ObjectList & " " & _
            ") " & _
        "ORDER BY le.name "
    Case "SUBROUTINE"
        Me.sqlQuery = _
        "SELECT s.name, s.removed, s.changed_on " & _
        "FROM subroutine s " & _
        "WHERE s.name IN " & ObjectList & "  " & _
        "ORDER BY s.name "
    End Select
ElseIf DataBase = "V2" Then
    If TableName = "ANALYSIS" Then
    'Query for versioned objects from their native table where "NAME" is the key field
        Me.sqlQuery = _
        "SELECT DISTINCT " & TableName & ".name, " & TableName & ".removed, " & TableName & ".changed_on, " & TableName & ".active " & _
        "FROM " & TableName & " " & _
            "INNER JOIN versions v " & _
                "ON v.table_name = '" & TableName & "' " & _
                "AND " & TableName & ".name = v.name " & _
                "AND " & TableName & ".version = v.version " & _
        "WHERE " & TableName & ".name IN " & ObjectList & "  " & _
        "ORDER BY " & TableName & ".name "
    ElseIf TableName = "COMMON_NAME" Or TableName = "ANALYSIS_TYPES" Or TableName = "BATCH_LINK" Or TableName = "FORMAT_CALCULATION" _
    Or TableName = "LIST" Or TableName = "STD_REAG_TEMP" Or TableName = "SUBROUTINE" Then
'    'Query for non-versioned objects from their native table where "NAME" is the key field
        Me.sqlQuery = _
        "SELECT DISTINCT " & TableName & ".name, " & TableName & ".removed, " & TableName & ".changed_on " & _
        "FROM " & TableName & " " & _
        "WHERE " & TableName & ".name IN " & ObjectList & "  " & _
        "ORDER BY " & TableName & ".name "
    ElseIf TableName = "UNITS" Then
    'Query for UNITS; "UNIT_CODE" is the key field (as opposed to "NAME")
        Me.sqlQuery = _
        "SELECT DISTINCT u.unit_code, u.removed, u.changed_on " & _
        "FROM units u " & _
        "WHERE u.unit_code IN " & ObjectList & "  " & _
        "ORDER BY u.unit_code "
    ElseIf TableName = "LIST_ENTRY" Then
    'Hard-coded to select the le.list = 'INST_GRPS' object because I suspect this table has a composite key.
        Me.sqlQuery = _
        "SELECT le.name " & _
        "FROM list_entry le " & _
        "WHERE le.list = 'INST_GRPS' " & _
            "AND le.name IN " & ObjectList & "  " & _
        "ORDER BY le.name "
    End If
End If

'Diagnostic MsgBox
'MsgBox Me.sqlQuery
End Function
'Written by: N.Whisman nathan.whisman@pfizer.com 07Apr2024
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
    
    GetCalcsQuery = _
    "SELECT UPPER(ca.source_code) " & _
    "FROM component  c " & _
        "LEFT OUTER JOIN calculation ca " & _
            "ON c.analysis = ca.analysis " & _
            "AND c.version = ca.version " & _
            "AND c.name = ca.component " & _
        "INNER JOIN versions v " & _
            "ON v.table_name = 'ANALYSIS' " & _
            "AND c.analysis = v.name " & _
            "AND c.version = v.version " & _
    "WHERE c.analysis IN " & AnalysisStr & " " & _
        "AND ca.source_code IS NOT NULL " & _
        "AND (UPPER(ca.source_code) LIKE '%GOSUB%' OR UPPER(ca.source_code) LIKE '%SUBROUTINE%') " & _
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
                RawSourceCodeStr = ReadClobToString(rs("UPPER(CA.SOURCE_CODE)"))
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
                        ElseIf Not (InStr(TrimmedCodeLine, "SUBROUTINE") = 0) Then
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
            Me.Objects = Left(Me.Objects, Len(Me.Objects) - 3) & ")"
        End If
    End With
End Function
Private Function ReadClobToString(ClobField As Variant) As String
    Dim strm As New ADODB.Stream
    
    With strm
        .Charset = "UTF-8"
        .Type = 2
        .Open
        .WriteText ClobField.Value
        .Position = 0
        ReadClobToString = .ReadText
    End With

    strm.Close
    Set strm = Nothing
End Function
'Written by: N.Whisman nathan.whisman@pfizer.com 25Feb2024
'Class contains methods for building and storing arrays of table names needed for verifying objects in V5/D3/V2 databases.
'
'*****************************************************************************************

Public TableArr As Variant

Public Function BuildAnalysisTableArr(DataBase As String)
    'Supports V5/D3/V2 as DataBase

    DataBase = UCase(DataBase)
    
    If DataBase = "V5" Then
        ReDim TableArr(2)
        'COMMON_NAME, ANALYSIS_TYPES, UNITS, FORMAT_CALCULATION, LIST, and LIST_ENTRY->INST_GRPS
        'are not set up to be checked in V5 because many of them represent a mix of global and site objects used
        'in D3, and the names were changed for a majority of the objects which makes it hard to program.
        TableArr(0) = "ANALYSIS"
        TableArr(1) = "BATCH_LINK"
        TableArr(2) = "STD_REAG_TEMP"
    ElseIf DataBase = "D3" Or DataBase = "V2" Then
        ReDim TableArr(8)
        TableArr(0) = "ANALYSIS"
        TableArr(1) = "ANALYSIS_TYPES"
        TableArr(2) = "BATCH_LINK"
        TableArr(3) = "UNITS"
        TableArr(4) = "FORMAT_CALCULATION"
        TableArr(5) = "LIST"
        TableArr(6) = "STD_REAG_TEMP"
        TableArr(7) = "LIST_ENTRY"
        TableArr(8) = "SUBROUTINE"
    End If
End Function

Public Function BuildProductTableArr(DataBase As String)
    'Supports V5/D3/V2 as DataBase
    
    DataBase = UCase(DataBase)
    
    If DataBase = "V5" Then
        ReDim TableArr(3)
        'Doesn't look like we had ITEM_CODEs in V5; don't see a GRADE table in V5 database;
        'STAGE and SPEC_TYPE came from list_entry, but no point comparing to V5 with no changed_on or removed fields;
        'For TEST_LOCATION, CONDITION, and UNITS we use global objects, too, so I'm only checking for existence in V2.
        TableArr(0) = "PRODUCT"
        TableArr(1) = "TEST_LIST"
        TableArr(2) = "SAMPLING_POINT"
        TableArr(3) = "ANALYSIS"
    ElseIf DataBase = "D3" Or DataBase = "V2" Then
        ReDim TableArr(11)
        TableArr(0) = "PRODUCT"
        TableArr(1) = "T_PH_ITEM_CODE"
        TableArr(2) = "TEST_LIST"
        TableArr(3) = "SAMPLING_POINT"
        TableArr(4) = "T_PH_GRADE"
        TableArr(5) = "T_PH_STAGE"
        TableArr(6) = "ANALYSIS"
        TableArr(7) = "T_PH_SPEC_TYPE"
        TableArr(8) = "TEST_LOCATION"
        TableArr(9) = "CONDITION"
        TableArr(10) = "UNITS"
        TableArr(11) = "T_PH_SAMPLE_PLAN"
    End If
End Function
