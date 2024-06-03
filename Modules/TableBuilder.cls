VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by: Nathan Whisman 25Feb2024
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
