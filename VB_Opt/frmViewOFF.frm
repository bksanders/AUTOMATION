VERSION 5.00
Begin VB.Form frmViewOFF 
   Caption         =   "View Off-Fall"
   ClientHeight    =   6000
   ClientLeft      =   3240
   ClientTop       =   3030
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   5250
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   18
      Left            =   1320
      TabIndex        =   37
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   17
      Left            =   1320
      TabIndex        =   35
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   16
      Left            =   1320
      TabIndex        =   33
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   15
      Left            =   1320
      TabIndex        =   31
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   14
      Left            =   1320
      TabIndex        =   29
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   13
      Left            =   1320
      TabIndex        =   27
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   25
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   12
      Left            =   1320
      TabIndex        =   23
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   11
      Left            =   1320
      TabIndex        =   21
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   19
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   9
      Left            =   1320
      TabIndex        =   17
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3720
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   8
      Left            =   4080
      TabIndex        =   15
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   7
      Left            =   4080
      TabIndex        =   13
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   3
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtOFF 
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   ">=100"":"
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
      Index           =   18
      Left            =   600
      TabIndex        =   38
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "95 - 99.999"":"
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
      Index           =   17
      Left            =   120
      TabIndex        =   36
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "90 - 94.999"":"
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
      Index           =   16
      Left            =   120
      TabIndex        =   34
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "85 - 89.999"":"
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
      Index           =   15
      Left            =   120
      TabIndex        =   32
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "80 - 84.999"":"
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
      Index           =   14
      Left            =   120
      TabIndex        =   30
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "75 - 79.999"":"
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
      Index           =   13
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "< 15"":"
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
      Index           =   0
      Left            =   3360
      TabIndex        =   26
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "70 - 75.999"":"
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
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "65 - 69.999"":"
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
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "60 - 64.999"":"
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
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "55 - 59.999"":"
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
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "50 - 55.999"":"
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
      Index           =   8
      Left            =   2880
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "45 -49.999"":"
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
      Index           =   7
      Left            =   2880
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "40 - 44.999"":"
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
      Index           =   6
      Left            =   2880
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "35 - 39.999"":"
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
      Index           =   5
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "30 - 34.999"":"
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
      Index           =   4
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "25 - 29.999"":"
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
      Index           =   3
      Left            =   2880
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "20 - 24.999"":"
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
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "15 - 19.999"":"
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
      Index           =   1
      Left            =   2880
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewOFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
   Me.Hide
   Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer        'loop index
   For i = 0 To 18
      txtOFF(i).Text = arrOffFall(i)
   Next i
End Sub
