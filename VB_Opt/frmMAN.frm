VERSION 5.00
Begin VB.Form frmMAN 
   Caption         =   "MANUAL mode"
   ClientHeight    =   7365
   ClientLeft      =   -75
   ClientTop       =   345
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10740
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Submit selected items for execution."
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdAUTO 
      Caption         =   "Return to AUTO Mode"
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Busway Plant-Selmer,TN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   1200
      X2              =   10800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      Caption         =   "Remmele Bar Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "GE Industrial Systems"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "GELogoFont"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   -360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "The PLC is currently in MANUAL mode!"
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3000
      Width           =   5775
   End
   Begin VB.Label lblSelectTitle 
      Caption         =   "Manual Mode Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
End
Attribute VB_Name = "frmMAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAUTO_Click()
   Unload Me
   frmRun.Show
End Sub

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
   End
End Sub
