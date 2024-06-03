VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by: Nathan Whisman 25Feb2024
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
'Diagnostic MsgBox
'MsgBox Me.sqlQuery
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

