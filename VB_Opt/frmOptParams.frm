VERSION 5.00
Begin VB.Form frmOptParams 
   Caption         =   "Optimization Paramaters"
   ClientHeight    =   2670
   ClientLeft      =   2865
   ClientTop       =   3225
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   Begin VB.TextBox txtFillDays 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Projected # of days to search when filling off-fall."
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtMaxRel 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Max # of releases allowed in a single batch."
      Top             =   960
      Width           =   615
   End
   Begin VB.CheckBox chkAutoPick 
      Caption         =   "Enable Auto Pick"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Enable automatic selection for initial batch."
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "# of days for fill search:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Max # of Releases:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmOptParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnLoading As Boolean        'loading paramaters flag

Private Sub chkAutoPick_Click()
   If blnLoading = False Then
      If chkAutoPick.Value = 1 Then
         blnAutoPICK = True
      Else
         blnAutoPICK = False
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   Me.Hide
   Unload Me
End Sub

Private Sub Form_Load()

   blnLoading = True                      'set loading paramaters flag
   
   If blnAutoPICK = True Then
      chkAutoPick.Value = 1
   Else
      chkAutoPick.Value = 0
   End If
   
   txtMaxRel = intMaxRel
   txtFillDays = intFillDays
   
   blnLoading = False
End Sub

Private Sub txtFillDays_Change()
Dim intValue As Integer
   
   If blnLoading = False Then
      intValue = Val(txtFillDays.Text)
      If intValue > 0 And intValue < 31 Then
         intFillDays = intValue
      Else
         MsgBox ("Illegal value for # of fill days!")
      End If
   End If
End Sub

Private Sub txtMaxRel_Change()
Dim intValue As Integer
   
   If blnLoading = False Then
      intValue = Val(txtMaxRel.Text)
      If intValue > 0 And intValue < 11 Then
         intMaxRel = intValue
      Else
         MsgBox ("Illegal value for max # of releases!")
      End If
   End If
End Sub
