VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColTitles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by: Nathan Whisman 01Jun2024
'This class simply serves to remove magic numbers by assigning column numbers and letters to class attributes.
'
'How to Use:
'   1. Instantiate object and then run .Init.
'   2. Call methods if needed.
'
'A note to viewers given that this is  now purely a portfolio piece:
'I had more time to play around with the class here as the project was complete and I had no deadline.
'In other areas of this project, I use public class attributes (not perfectly encapsulated, but no worse than
'Python) because I seemingly have to choose between a familiar class syntax for changing class attributes
'(e.g. Me.attribute = value) or otherwise being able to use truly encapsulated, private class attributes.
'I don't know why yet, but the syntax I'd expect for a class (using "Me.attribute = value" inside of a method
'to set the attribute) isn't able to see or work with private class attributes. That's why you see me do things
'one way in this module but another way in other class modules.
'**************************************************************************************************************


Private me_StrParentObjs As String
Private me_StrV8TableName As String
Private me_StrV5TableName As String
Private me_StrV8ObjName As String
Private me_StrExistsV5 As String
Private me_StrExistsD3 As String
Private me_StrExistsV2 As String
Private me_StrV5Changed As String
Private me_StrD3Changed As String
Private me_StrV2Changed As String
Private me_StrStatus As String
Private me_NumParentObjs As Long
Private me_NumV8TableName As Long
Private me_NumV5TableName As Long
Private me_NumV8ObjName As Long
Private me_NumExistsV5 As Long
Private me_NumExistsD3 As Long
Private me_NumExistsV2 As Long
Private me_NumV5Changed As Long
Private me_NumD3Changed As Long
Private me_NumV2Changed As Long
Private me_NumStatus As Long

Public Sub Init()
    me_StrParentObjs = "A"
    me_StrV8TableName = "C"
    me_StrV5TableName = "D"
    me_StrV8ObjName = "E"
    me_StrExistsV5 = "F"
    me_StrExistsD3 = "G"
    me_StrExistsV2 = "H"
    me_StrV5Changed = "I"
    me_StrD3Changed = "J"
    me_StrV2Changed = "K"
    me_StrStatus = "L"
    me_NumParentObjs = 0
    me_NumV8TableName = 3
    me_NumV5TableName = 4
    me_NumV8ObjName = 5
    me_NumExistsV5 = 6
    me_NumExistsD3 = 7
    me_NumExistsV2 = 8
    me_NumV5Changed = 9
    me_NumD3Changed = 10
    me_NumV2Changed = 11
    me_NumStatus = 12
End Sub

' Getter and setter for StrParentObjs
Public Property Get GetStrParentObjs() As String
    GetStrParentObjs = me_StrParentObjs
End Property

Public Property Let SetStrParentObjs(value As String)
    me_StrParentObjs = value
End Property

Public Sub DeleteStrParentObjs()
    me_StrParentObjs = ""
End Sub

' Getter and setter for StrV8TableName
Public Property Get GetStrV8TableName() As String
    GetStrV8TableName = me_StrV8TableName
End Property

Public Property Let SetStrV8TableName(value As String)
    me_StrV8TableName = value
End Property

Public Sub DeleteStrV8TableName()
    me_StrV8TableName = ""
End Sub

' Getter and setter for StrV5TableName
Public Property Get GetStrV5TableName() As String
    GetStrV5TableName = me_StrV5TableName
End Property

Public Property Let SetStrV5TableName(value As String)
    me_StrV5TableName = value
End Property

Public Sub DeleteStrV5TableName()
    me_StrV5TableName = ""
End Sub

' Getter and setter for StrV8ObjName
Public Property Get GetStrV8ObjName() As String
    GetStrV8ObjName = me_StrV8ObjName
End Property

Public Property Let SetStrV8ObjName(value As String)
    me_StrV8ObjName = value
End Property

Public Sub DeleteStrV8ObjName()
    me_StrV8ObjName = ""
End Sub

' Getter and setter for StrExistsV5
Public Property Get GetStrExistsV5() As String
    GetStrExistsV5 = me_StrExistsV5
End Property

Public Property Let SetStrExistsV5(value As String)
    me_StrExistsV5 = value
End Property

Public Sub DeleteStrExistsV5()
    me_StrExistsV5 = ""
End Sub

' Getter and setter for StrExistsD3
Public Property Get GetStrExistsD3() As String
    GetStrExistsD3 = me_StrExistsD3
End Property

Public Property Let SetStrExistsD3(value As String)
    me_StrExistsD3 = value
End Property

Public Sub DeleteStrExistsD3()
    me_StrExistsD3 = ""
End Sub

' Getter and setter for StrExistsV2
Public Property Get GetStrExistsV2() As String
    GetStrExistsV2 = me_StrExistsV2
End Property

Public Property Let SetStrExistsV2(value As String)
    me_StrExistsV2 = value
End Property

Public Sub DeleteStrExistsV2()
    me_StrExistsV2 = ""
End Sub

' Getter and setter for StrV5Changed
Public Property Get GetStrV5Changed() As String
    GetStrV5Changed = me_StrV5Changed
End Property

Public Property Let SetStrV5Changed(value As String)
    me_StrV5Changed = value
End Property

Public Sub DeleteStrV5Changed()
    me_StrV5Changed = ""
End Sub

' Getter and setter for StrD3Changed
Public Property Get GetStrD3Changed() As String
    GetStrD3Changed = me_StrD3Changed
End Property

Public Property Let SetStrD3Changed(value As String)
    me_StrD3Changed = value
End Property

Public Sub DeleteStrD3Changed()
    me_StrD3Changed = ""
End Sub

' Getter and setter for StrV2Changed
Public Property Get GetStrV2Changed() As String
    GetStrV2Changed = me_StrV2Changed
End Property

Public Property Let SetStrV2Changed(value As String)
    me_StrV2Changed = value
End Property

Public Sub DeleteStrV2Changed()
    me_StrV2Changed = ""
End Sub

' Getter and setter for StrStatus
Public Property Get GetStrStatus() As String
    GetStrStatus = me_StrStatus
End Property

Public Property Let SetStrStatus(value As String)
    me_StrStatus = value
End Property

Public Sub DeleteStrStatus()
    me_StrStatus = ""
End Sub

' Getter and setter for NumParentObjs
Public Property Get GetNumParentObjs() As Long
    GetNumParentObjs = me_NumParentObjs
End Property

Public Property Let SetNumParentObjs(value As Long)
    me_NumParentObjs = value
End Property

Public Sub DeleteNumParentObjs()
    me_NumParentObjs = 0
End Sub

' Getter and setter for NumV8TableName
Public Property Get GetNumV8TableName() As Long
    GetNumV8TableName = me_NumV8TableName
End Property

Public Property Let SetNumV8TableName(value As Long)
    me_NumV8TableName = value
End Property

Public Sub DeleteNumV8TableName()
    me_NumV8TableName = 0
End Sub

' Getter and setter for NumV5TableName
Public Property Get GetNumV5TableName() As Long
    GetNumV5TableName = me_NumV5TableName
End Property

Public Property Let SetNumV5TableName(value As Long)
    me_NumV5TableName = value
End Property

Public Sub DeleteNumV5TableName()
    me_NumV5TableName = 0
End Sub

' Getter and setter for NumV8ObjName
Public Property Get GetNumV8ObjName() As Long
    GetNumV8ObjName = me_NumV8ObjName
End Property

Public Property Let SetNumV8ObjName(value As Long)
    me_NumV8ObjName = value
End Property

Public Sub DeleteNumV8ObjName()
    me_NumV8ObjName = 0
End Sub

' Getter and setter for NumExistsV5
Public Property Get GetNumExistsV5() As Long
    GetNumExistsV5 = me_NumExistsV5
End Property

Public Property Let SetNumExistsV5(value As Long)
    me_NumExistsV5 = value
End Property

Public Sub DeleteNumExistsV5()
    me_NumExistsV5 = 0
End Sub

' Getter and setter for NumExistsD3
Public Property Get GetNumExistsD3() As Long
    GetNumExistsD3 = me_NumExistsD3
End Property

Public Property Let SetNumExistsD3(value As Long)
    me_NumExistsD3 = value
End Property

Public Sub DeleteNumExistsD3()
    me_NumExistsD3 = 0
End Sub

' Getter and setter for NumExistsV2
Public Property Get GetNumExistsV2() As Long
    GetNumExistsV2 = me_NumExistsV2
End Property

Public Property Let SetNumExistsV2(value As Long)
    me_NumExistsV2 = value
End Property

Public Sub DeleteNumExistsV2()
    me_NumExistsV2 = 0
End Sub

' Getter and setter for NumV5Changed
Public Property Get GetNumV5Changed() As Long
    GetNumV5Changed = me_NumV5Changed
End Property

Public Property Let SetNumV5Changed(value As Long)
    me_NumV5Changed = value
End Property

Public Sub DeleteNumV5Changed()
    me_NumV5Changed = 0
End Sub

' Getter and setter for NumD3Changed
Public Property Get GetNumD3Changed() As Long
    GetNumD3Changed = me_NumD3Changed
End Property

Public Property Let SetNumD3Changed(value As Long)
    me_NumD3Changed = value
End Property

Public Sub DeleteNumD3Changed()
    me_NumD3Changed = 0
End Sub

' Getter and setter for NumV2Changed
Public Property Get GetNumV2Changed() As Long
    GetNumV2Changed = me_NumV2Changed
End Property

Public Property Let SetNumV2Changed(value As Long)
    me_NumV2Changed = value
End Property

Public Sub DeleteNumV2Changed()
    me_NumV2Changed = 0
End Sub

' Getter and setter for NumStatus
Public Property Get GetNumStatus() As Long
    GetNumStatus = me_NumStatus
End Property

Public Property Let SetNumStatus(value As Long)
    me_NumStatus = value
End Property

Public Sub DeleteNumStatus()
    me_NumStatus = 0
End Sub

