VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Encoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarTheString As String 'local copy
Public Sub Decode(CodeKey As Byte)
    Dim MsgLength As Long
    Dim i As Long
    Dim NewChar As Integer
    
    MsgLength = Len(mvarTheString)
    For i = 1 To MsgLength
        NewChar = (Asc(Mid(mvarTheString, i, 1)) + CodeKey) Mod 256
        Mid(mvarTheString, i, 1) = Chr(NewChar)
    Next i
End Sub

Public Sub Encode(CodeKey As Byte)
    Dim MsgLength As Long
    Dim i As Long
    Dim NewChar As Integer
    
    MsgLength = Len(mvarTheString)
    For i = 1 To MsgLength
        NewChar = (Asc(Mid(mvarTheString, i, 1)) - CodeKey)
        If NewChar < 0 Then NewChar = 256 - Abs(NewChar)
        Mid(mvarTheString, i, 1) = Chr(NewChar)
    Next i
End Sub

Public Property Let TheString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TheString = 5
    mvarTheString = vData
End Property


Public Property Get TheString() As String
Attribute TheString.VB_Description = "The string to be encoded/decoded"
Attribute TheString.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TheString
    TheString = mvarTheString
End Property



