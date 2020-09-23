VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEncode 
   Caption         =   "Encoder"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "frmEncode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin MSComCtl2.UpDown updKey 
      Height          =   285
      Left            =   736
      TabIndex        =   2
      Top             =   2760
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtKey"
      BuddyDispid     =   196610
      OrigLeft        =   960
      OrigTop         =   2760
      OrigRight       =   1200
      OrigBottom      =   3015
      Max             =   255
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblPreview 
      Caption         =   "AaBbYyZz"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frmEncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GlobalClass As New Encoder

Private Sub cmdDecode_Click()
    GlobalClass.TheString = txtText.Text
    GlobalClass.Decode updKey.Value
    txtText.Text = GlobalClass.TheString
    GlobalClass.TheString = "AaBbYyZz"
    GlobalClass.Decode updKey.Value
    lblPreview.Caption = GlobalClass.TheString
End Sub

Private Sub cmdEncode_Click()
    GlobalClass.TheString = txtText.Text
    GlobalClass.Encode updKey.Value
    txtText.Text = GlobalClass.TheString
    GlobalClass.TheString = "AaBbYyZz"
    GlobalClass.Encode updKey.Value
    lblPreview.Caption = GlobalClass.TheString
End Sub
