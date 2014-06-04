VERSION 5.00
Begin VB.Form frmChangeQty 
   Caption         =   "Change Quantity"
   ClientHeight    =   2595
   ClientLeft      =   2490
   ClientTop       =   3225
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   5175
   Begin VB.TextBox txtSeq 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   12
      ToolTipText     =   "Enter an Item number for search."
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtBldQTY 
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Enter an Item number for search."
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtOpenQTY 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "Enter an Item number for search."
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtJob 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Enter a Job number for search."
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtRel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Enter a Release number for search."
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtItem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Enter an Item number for search."
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Bld Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Open Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblJob 
      Caption         =   "Job"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblRelease 
      Caption         =   "Release"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblItem 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmChangeQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccept_Click()

Dim bldQTY As Integer
Dim openQTY As Integer

   bldQTY = Val(txtBldQTY.Text)
   openQTY = Val(txtOpenQTY.Text)
   
   If bldQTY < 0 Then
      Beep
      txtBldQTY.Text = usrMACHlist(intMachIDX).BldQnty
      Exit Sub
   End If
   
   If bldQTY = 0 Then
      frmBatch.lbxMach.SELECTED(intMachIDX - 1) = False      'deselect the part
   End If
   
   If bldQTY > openQTY Then
      txtBldQTY.Text = usrMACHlist(intMachIDX).BldQnty
      Beep
      Exit Sub
   Else
      usrMACHlist(intMachIDX).BldQnty = Val(txtBldQTY.Text)
      If usrBatch.InitPick = False Then                      'chk/make init pick
         usrBatch.Mat = usrMACHlist(intMachIDX).Material
         usrBatch.Width = usrMACHlist(intMachIDX).BarWidth
         usrBatch.InitPick = True
      End If
      frmBatch.refMACHlist
      displayButtons
      autoOPT
   End If
   
   Unload Me
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
   i = intMachIDX
   txtJob.Text = usrMACHlist(i).Job
   txtRel.Text = usrMACHlist(i).Rel
   txtItem.Text = usrMACHlist(i).Item
   txtSeq.Text = usrMACHlist(i).Seq
   txtOpenQTY.Text = usrMACHlist(i).Qnty
   txtBldQTY.Text = usrMACHlist(i).BldQnty
End Sub
