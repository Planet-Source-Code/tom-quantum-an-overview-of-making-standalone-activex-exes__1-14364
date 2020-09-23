VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AXEncoder Test"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   2205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Encode ""Hello, World!"""
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTest_Click()
    Dim cls As Object
    
    Set cls = CreateObject("AXEncoder.Encoder")
    cls.TheString = "Hello, World!"
    cls.Encode 5
    MsgBox cls.TheString
End Sub
