VERSION 5.00
Begin VB.Form Test 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkMode        =   1  'Source
   LinkTopic       =   "Test"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDDE 
      Caption         =   "Enable DDE"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      LinkItem        =   "txtSource"
      LinkTopic       =   "VB_Test|Form1"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDDE_Click()
  
   If chkDDE.Value = 1 Then
      Text2.LinkMode = 1
   Else
      Text2.LinkMode = 0
   End If

End Sub
