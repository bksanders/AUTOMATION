VERSION 5.00
Begin VB.Form frmPrevBatch 
   Caption         =   "Prev Batch Detected!"
   ClientHeight    =   3795
   ClientLeft      =   2310
   ClientTop       =   2655
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   7065
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Finish"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdSuspend 
      Caption         =   "Suspend"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Batch incomplete.  Finish Batch."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Batch incomplete.  Suspend Batch."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Batch complete.  Save to History Table. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmPrevBatch.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmPrevBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFinish_Click()
   OPTIMIZED = True
   frmRun.RefreshExecQue
   Unload Me
End Sub

Private Sub cmdSave_Click()
   procBatch
   Unload Me
End Sub

Private Sub cmdSuspend_Click()
   procSusp
   Unload Me
End Sub
