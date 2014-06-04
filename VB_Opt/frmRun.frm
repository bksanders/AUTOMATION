VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRun 
   Caption         =   "Run Screen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRun.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmRun"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.Timer tmrPrevBatch 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11160
      Top             =   10560
   End
   Begin VB.TextBox txtFDRTruck 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   79
      Top             =   960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtPITruck 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   78
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPIGTruck 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   77
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdGOTOBatch 
      Caption         =   "    Batch         Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   76
      ToolTipText     =   "Go To Batch Scren"
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox txtCount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   75
      Top             =   10080
      Width           =   855
   End
   Begin VB.CommandButton cmdClrReq 
      Caption         =   "Clr Req"
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
      Left            =   10800
      TabIndex        =   74
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtServerStat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   72
      Top             =   9960
      Width           =   375
   End
   Begin VB.TextBox txtBatchStat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   8880
      TabIndex        =   70
      Top             =   7920
      Width           =   1815
   End
   Begin VB.TextBox txtStkLength 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   6000
      TabIndex        =   69
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSuspend 
      Caption         =   "Suspend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   67
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   65
      Top             =   9120
      Width           =   855
   End
   Begin VB.CommandButton cmdACKCmp 
      Caption         =   "ACK Cmp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   64
      Top             =   9600
      Width           =   855
   End
   Begin VB.CommandButton cmdACKReq 
      Caption         =   "ACK Req"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   63
      Top             =   9360
      Width           =   855
   End
   Begin VB.Timer tmrTCPStat 
      Interval        =   250
      Left            =   11040
      Top             =   9120
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   8280
      Top             =   9240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtClientStat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   61
      Top             =   9240
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   9000
      TabIndex        =   58
      Top             =   10320
      Width           =   1935
      Begin VB.OptionButton optMub 
         Caption         =   "link to Mubea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optSim 
         Caption         =   "link to Sim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.TextBox txtTCPRecv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   55
      Top             =   10080
      Width           =   3375
   End
   Begin VB.TextBox txtTCPSend 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   52
      Top             =   9240
      Width           =   3375
   End
   Begin VB.CommandButton cmdOptimize 
      Caption         =   "Optimize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   50
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1080
      TabIndex        =   49
      Top             =   1575
      Width           =   1095
   End
   Begin VB.TextBox txtRemBlanks 
      Height          =   405
      Left            =   8400
      TabIndex        =   42
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "Deselect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   40
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtTotBlanks 
      Height          =   405
      Left            =   7320
      TabIndex        =   38
      Top             =   1560
      Width           =   735
   End
   Begin MSAdodcLib.Adodc AdodcRun 
      Height          =   375
      Left            =   5640
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=dsnMBLocal"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "dsnMBLocal"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblRun"
      Caption         =   "AdodcRun"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGOTOHist 
      Caption         =   "   History         Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   19
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdGOTOSelect 
      Caption         =   " Select Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   18
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdSendBack 
      Caption         =   "Send Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   17
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtMubMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   3360
      Width           =   7095
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   12
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtMat 
      Height          =   405
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
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
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Exit Mubea HMI."
      Top             =   8040
      Width           =   975
   End
   Begin VB.ListBox lbxExec 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   4320
      Width           =   10455
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   8040
      Width           =   975
   End
   Begin MSAdodcLib.Adodc AdodcExec 
      Height          =   375
      Left            =   3600
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=dsnMBLocal"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "dsnMBLocal"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblExecQue Where Priority > 0 ORDER BY Priority ASC"
      Caption         =   "AdodcExec"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox lbxRun 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   240
      TabIndex        =   33
      Top             =   2280
      Width           =   8895
   End
   Begin VB.Frame FrameRunMode 
      Caption         =   "Run Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9240
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
      Begin VB.OptionButton optStop 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optPause 
         Caption         =   "Pause"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optSingle 
         Caption         =   "Single Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Auto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   8280
      Top             =   9960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblFDR 
      Alignment       =   1  'Right Justify
      Caption         =   "FDR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10035
      TabIndex        =   82
      Top             =   1005
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblPI 
      Alignment       =   1  'Right Justify
      Caption         =   "PI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10275
      TabIndex        =   81
      Top             =   1365
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblPIG 
      Alignment       =   1  'Right Justify
      Caption         =   "PI Gnd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9840
      TabIndex        =   80
      Top             =   1725
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label40 
      Caption         =   "Server Status:"
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
      Left            =   9120
      TabIndex        =   73
      Top             =   10005
      Width           =   1335
   End
   Begin VB.Label Label25 
      Caption         =   "Batch:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   71
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Suspend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10830
      TabIndex        =   68
      Top             =   6300
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label24 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   66
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label23 
      Caption         =   "Client Status:"
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
      Left            =   9240
      TabIndex        =   62
      Top             =   9285
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   8160
      X2              =   8160
      Y1              =   9000
      Y2              =   10560
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   8160
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Label Label39 
      Caption         =   "to Mubea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "from Mubea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Label Label32 
      Caption         =   "TCP Recv:"
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
      Left            =   2160
      TabIndex        =   54
      Top             =   10125
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "TCP Send:"
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
      Left            =   2160
      TabIndex        =   53
      Top             =   9285
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "TCP Comm: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label Label37 
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
      TabIndex        =   47
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   1080
      X2              =   12840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label28 
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
      TabIndex        =   45
      Top             =   360
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   9225
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label35 
      Caption         =   "Stk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   44
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label31 
      Caption         =   "Remain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   43
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label30 
      Caption         =   "Stk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   41
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label Label29 
      Caption         =   "Length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   39
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label27 
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   37
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label26 
      Caption         =   "Leg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   36
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label21 
      Caption         =   "BlankLength"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   35
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "PH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   34
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "PH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7530
      TabIndex        =   32
      Top             =   4080
      Width           =   390
   End
   Begin VB.Label Label18 
      Caption         =   "Pr#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   30
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "BlankLength"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   29
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   28
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "Mat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label Label13 
      Caption         =   "BldQty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6135
      TabIndex        =   26
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5445
      TabIndex        =   25
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Leg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8085
      TabIndex        =   24
      Top             =   4080
      Width           =   420
   End
   Begin VB.Label Label10 
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4035
      TabIndex        =   23
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3315
      TabIndex        =   22
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Rel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2670
      TabIndex        =   21
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Job"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1965
      TabIndex        =   20
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Mubea MSG:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblSelectTitle 
      Caption         =   "RUN  SCREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Rel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Job"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Execution Que:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label38 
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
      Left            =   120
      TabIndex        =   48
      Top             =   -480
      Width           =   1215
   End
   Begin VB.Label Label36 
      Caption         =   "Mubea Bar Machine"
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
      Left            =   9240
      TabIndex        =   46
      Top             =   360
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   10800
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblTruckHDR 
      Alignment       =   1  'Right Justify
      Caption         =   "Truck #:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   83
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usrEQParts(300) As part      'Exec Listbox Storage
Dim intPartSel As Integer        'pointer to which part is selected

Sub RefreshRunQue()

'--- globals used
'usrRUNParts                     'Run Que Listbox Storage = Run Parts array
'intRunCnt                       '# of parts in Run Que

'--- variable declarations
Dim strTemp As String
Dim bldOK As Boolean             'bld runpiece result flag
Dim i As Integer                 'loop index

lbxRun.Clear
AdodcRun.Refresh
If Not AdodcRun.Recordset.EOF Then        'if NOT at EOF, there are parts in que
   
   '--- prep for 1st part
   AdodcRun.Recordset.MoveFirst
   i = 1                                  'init parts array index
   
   '-------------------- build the temp parts array ------------------------
   'loop thru the Run Que and build a parts array
   Do Until AdodcRun.Recordset.EOF
      '--- build array entry for part
      'usrRUNParts(i).FullJobNum = AdodcRun.Recordset.Fields("FullOrder")
      usrRUNParts(i).Job = AdodcRun.Recordset.Fields("Order Number")
      usrRUNParts(i).Rel = AdodcRun.Recordset.Fields("Release")
      usrRUNParts(i).Item = AdodcRun.Recordset.Fields("Item")
      usrRUNParts(i).Seq = AdodcRun.Recordset.Fields("Sequence Number")
      'usrRUNParts(i).ShipDate = AdodcRun.Recordset.Fields("Scheduled Ship Date")
      'usrRUNParts(i).Qnty = AdodcRun.Recordset.Fields("Quantity")
      'usrRUNParts(i).BldQnty = AdodcRun.Recordset.Fields("BldQnty")
      usrRUNParts(i).Phase = AdodcRun.Recordset.Fields("Phase")
      usrRUNParts(i).Leg = AdodcRun.Recordset.Fields("Leg")
      usrRUNParts(i).Stack = AdodcRun.Recordset.Fields("Stack")
      usrRUNParts(i).BarType = AdodcRun.Recordset.Fields("BarType")
      usrRUNParts(i).Material = AdodcRun.Recordset.Fields("Material")
      usrRUNParts(i).BarWidth = AdodcRun.Recordset.Fields("BarWidth")
      usrRUNParts(i).BlankLength = AdodcRun.Recordset.Fields("BlankLength")
      'usrRUNParts(i).RunLength = AdodcRun.Recordset.Fields("RunLength")
      usrRUNParts(i).E1fig = AdodcRun.Recordset.Fields("E1figure")
      usrRUNParts(i).E1dim = AdodcRun.Recordset.Fields("E1dimension")
      usrRUNParts(i).E2fig = AdodcRun.Recordset.Fields("E2figure")
      usrRUNParts(i).E2dim = AdodcRun.Recordset.Fields("E2dimension")
      usrRUNParts(i).Cdim = AdodcRun.Recordset.Fields("Cdimension")
      usrRUNParts(i).C1dim = AdodcRun.Recordset.Fields("C1dimension")
      usrRUNParts(i).Ddim = AdodcRun.Recordset.Fields("Ddimension")
      usrRUNParts(i).D1dim = AdodcRun.Recordset.Fields("D1dimension")
      'usrRUNParts(i).Build = AdodcRun.Recordset.Fields("Build")
      'usrRUNParts(i).Status = AdodcRun.Recordset.Fields("Status")
      usrRUNParts(i).Priority = AdodcRun.Recordset.Fields("Priority")
             
      '--- construct string and add to listbox for display
      strTemp = Format(usrRUNParts(i).Job, "0000") & " " & _
                Format(usrRUNParts(i).Rel, "0000") & " " & _
                Format(usrRUNParts(i).Item, "0000") & " " & _
                Format(usrRUNParts(i).Seq, "0000") & "        " & _
                Format(usrRUNParts(i).Phase, "@") & "   " & _
                Format(usrRUNParts(i).Leg, "@") & "   " & _
                Format(usrRUNParts(i).Stack, "@") & "       " & _
                Format(usrRUNParts(i).BlankLength, "000000")
      lbxRun.AddItem (strTemp)
              
      AdodcRun.Recordset.MoveNext      'incr for next part
      i = i + 1
   Loop 'end of recordset loop
   
   intRunCnt = i - 1                   'get #of parts in run que
      
   lbxRun.Refresh
   frmRun.Refresh
   
Else  'no parts in RUN que
   intRunCnt = 0
End If

lbxRun.Refresh

End Sub  'RefreshRunQue()

'--- This sub deselects all items in the Exec Que

Public Sub deSelExec()

Dim i, j As Integer

   j = lbxExec.ListCount - 1                 'get loop range

   For i = 0 To j                            'loop thru listbox contents
      lbxExec.SELECTED(i) = False            'deselect each part
   Next i   'listbox loop
   
   intPartSel = -1                           'reset select pointer
   
   RefreshExecQue                            'refresh the que
   
End Sub  'deSelExec

Private Sub cmdACKCmp_Click()
Dim strMSG As String          'TCP message string
   
On Error GoTo errorHandler

   'strMSG = "GEAS_CmplAck"
   'tcpConn.SendData strMSG                   'send the ACK string
   'txtTCPSend.Text = strMSG                  'copy to display

   txtTCPSend.Text = "GEAS_CmplAck"
   tcpClient.Connect

Exit Sub

errorHandler:
   MsgBox Err.Description
End Sub

Private Sub cmdACKReq_Click()
Dim strMSG As String          'TCP message string
      
On Error GoTo errorHandler

   'strMSG = "GEAS_ReqAck"
   'tcpConn.SendData strMSG                   'send the ACK string
   'txtTCPSend.Text = strMSG                  'copy to display
   
   txtTCPSend.Text = "GEAS_ReqAck"
   tcpClient.Connect

Exit Sub

errorHandler:
   MsgBox Err.Description
End Sub

Private Sub cmdClrReq_Click()
   txtTCPRecv.Text = ""
End Sub

Private Sub cmdConnect_Click()
   If tcpClient.State <> sckListening Then
      tcpClient.Close
   End If
   tcpClient.Connect                         'initiate a tcp connection
End Sub

Private Sub cmdDeselect_Click()
   deSelExec                                 'deselct exec que
End Sub

Private Sub cmdExit_Click(Index As Integer)
   AdodcExec.Recordset.Close                 'close recordsets
   AdodcRun.Recordset.Close
   'Print #1, "Close Capture file: " & Now
   'Close #1                                 'close capture file-debug only
   'Shell_NotifyIcon NIM_DELETE, nid          'remove icon from sys tray
   Unload Me                                 'unload form
   End
End Sub

Private Sub cmdGOTOBatch_Click()
   frmBatch.Show
End Sub

Private Sub cmdGOTOHist_Click()
   frmHist.Show
End Sub

Private Sub cmdGOTOSelect_Click()
   frmSelect.Show                            'display Execution form
End Sub

Private Sub cmdOptimize_Click()
'--- global variables
'OPTIMIZED                    Optimized flag
'BATCH                        batch in progress flag

'--- variable declarations
Dim tempResult As Boolean
   
   '------------------------------------------------- check for empty execque
   'don't optimize if ExecQue Empty
   If intExecCnt < 1 Then                             'ExecQue empty
      Beep
      Exit Sub
   End If
    
   If BATCH = False Then   '------------------------- NO Batch in progress
      clrBatch tempResult                             'clear the batch
      bldBatch                                        'generate batch
      optBatch                                        'optimize batch
      OPTIMIZED = True                                'set optimize flag
      RefreshExecQue
      'MsgBox ("Batch Optimized!")
   Else
      MsgBox ("Batch in progress. Can't optimize!")
   End If   'BATCH
End Sub

'------------------------------------------------------------------------
' Descrip:  This sub sends an item back to the select que.
' Notes:    -identifies the selected part.
'           -Selects all parts for this item.
'           -Deletes all parts in item from exec que
'           -reprioritises any parts w/ higher priority
'           -adds the item to the select que.
'------------------------------------------------------------------------

Private Sub cmdSendBack_Click()

'--- global variables
'BATCH                              'Batch(in Prog) flag
'intAsgnCnt                         '# of assigned parts(for machine list update)

'--- local variables
Dim i, j, intDone As Integer
Dim SELECTED As Boolean             'part selected flag
Dim MATCH As Boolean
Dim intRmvParts As Integer          'total # of parts to remove
Dim intRmvCnt As Integer            'count of removed parts
Dim intRmvPri As Integer            'current priority being removed
Dim intPriCnt As Integer            '# of parts w/ current priority
Dim intTemp As Integer              'temp integer variable
Dim longTemp As Long                'temp long variable
Dim intAddRes As Integer            'add result
Dim conLocal As Connection          'local connect
Dim adorsRMV As ADODB.Recordset     'remove parts recordset
Dim adorsPRI As ADODB.Recordset     'remove priority recordset
Dim conCamdata As Connection        'camdata connection
Dim adorsCamdata As ADODB.Recordset 'camdata RS
Dim usrTempPart As part             'temporary part variable

   '--------------------------- check Batch Status ----------------------------
   'Only allow sendback if NO batch in progress
   If BATCH = True Then                               'batch in progress
      Beep
      Exit Sub
   End If
     
   '------------------------ identify the selected part -----------------------
   i = 0
   j = lbxExec.ListCount
   SELECTED = False
   Do Until SELECTED Or i = j                         'loop thru listbox contents
      If lbxExec.SELECTED(i) = True Then              'if row is selected
         usrRecID.Job = usrEQParts(i + 1).Job         'set record id
         usrRecID.Rel = usrEQParts(i + 1).Rel
         usrRecID.Item = usrEQParts(i + 1).Item
         usrRecID.Seq = usrEQParts(i + 1).Seq
         SELECTED = True
      End If
      i = i + 1
   Loop  'end of listbox loop
   
   intRmvParts = 0                                    'reset remove parts
   If SELECTED Then                                   'if a part was selected
      '-------------------- select all parts w/ this JRI -------------------
      For i = 1 To j                                  'loop thru EQ parts array
         If usrEQParts(i).Job = usrRecID.Job And _
            usrEQParts(i).Rel = usrRecID.Rel And _
            usrEQParts(i).Item = usrRecID.Item Then
            lbxExec.SELECTED(i - 1) = True            'select the part
            intRmvParts = intRmvParts + 1             'incr # of parts to remove
         End If
      Next i
      lbxExec.Refresh
      
      '------------------------- are you sure prompt -----------------------
      'set up msg box
      strMSG = "Are you sure you want to send selected item back?"   ' Define msgbox test
      intStyle = vbYesNo + vbDefaultButton2                          ' Define msgbox buttons
      strTitle = "Sending Item back to Select Que..."                ' Define msgbox title
      intResponse = MsgBox(strMSG, intStyle, strTitle)
      deSelExec                                       'unselect everything
      If intResponse = vbYes Then  '--- OP chose Yes
         '----------------- Remove Parts from the Exec Que -----------------
              
         '---------- remove parts loop
         intRmvCnt = 0                       'reset remove count
         Do Until intRmvCnt = intRmvParts    '--- loop until all parts removed
            
            '--- get priority of 1st part to be removed
            With AdodcExec
               If .Recordset.RecordCount > 0 Then        'skip if ExecQue empty
                  .Recordset.MoveFirst
                  Do Until .Recordset.EOF Or MATCH
                     If .Recordset.Fields("Order Number") = usrRecID.Job And _
                        .Recordset.Fields("Release") = usrRecID.Rel And _
                        .Recordset.Fields("Item") = usrRecID.Item Then
                        intRmvPri = .Recordset.Fields("Priority")
                        MATCH = True
                     End If
                     .Recordset.MoveNext
                  Loop
               Else  ' Nothing to remove
                  MsgBox ("Sendback: Could NOT locate parts for Removal!")
                  Exit Sub
               End If
            End With
            
            If MATCH = False Then
               MsgBox ("Sendback: Could NOT locate parts for Removal!")
               Exit Sub
            End If
         
            '--- delete all parts w/ this pri
            If blnEnOPT Then
               clrLocalTBL "tblHoldAssign"               'clear hold assign table
               intAsgnCnt = 0
            End If
            With AdodcExec
               If .Recordset.RecordCount > 0 Then        'skip if ExecQue empty
                  .Recordset.MoveFirst
                  Do Until .Recordset.EOF
                     intTemp = .Recordset.Fields("Priority")
                     If intTemp = intRmvPri Then
                        'If removing a part from the Exec Que and
                        'optimization is enabled...then we must capture this part
                        'to update its assignment in the machine table. NO need to
                        'assign remakes.
                        If blnEnOPT And .Recordset.Fields("Build") <> "R" Then
                           usrTempPart.Job = .Recordset.Fields("Order Number")
                           usrTempPart.Rel = .Recordset.Fields("Release")
                           usrTempPart.Item = .Recordset.Fields("Item")
                           usrTempPart.Seq = .Recordset.Fields("Sequence Number")
                           usrTempPart.Qnty = .Recordset.Fields("Quantity")
                           intAddRes = 0
                           addHoldAssign usrTempPart, intAddRes
                           intAsgnCnt = intAsgnCnt + 1
                        End If
                        
                        .Recordset.Delete                'remove the part
                        .Recordset.Update
                        intRmvCnt = intRmvCnt + 1        'update remove count
                     End If
                     .Recordset.MoveNext
                  Loop
               End If
            End With
            If blnEnOPT And intAsgnCnt > 0 Then
               runQuery0 "qryManAssignUpdateRMV"         'run the update query
            End If
            
            '--- refresh the RemoveRS & Exec Que
            OPTIMIZED = False                            'reset optimized flag
            RefreshExecQue
            
            '--- decr priorities in Exec Que
            'loop thru the Exec Que if a part has a priority > that the priority
            'just removed then decrement the parts priority.
            With AdodcExec
               If .Recordset.RecordCount > 0 Then     'skip if ExecQue empty
                  .Recordset.MoveFirst
                  Do Until .Recordset.EOF
                     intTemp = .Recordset.Fields("Priority")
                     If intTemp > intRmvPri Then
                        .Recordset.Fields("Priority") = intTemp - 1
                        .Recordset.Update
                     End If
                     .Recordset.MoveNext
                  Loop
               End If
            End With
            
            RefreshExecQue
            
         Loop  'remove parts
         
         '------------------- Put Item Back in Select Que ------------------
         '--- make connection to camdata db
         Set conCamdata = New Connection
         conCamdata.Open "PROVIDER=MSDASQL;dsn=dsnMBCamdata;uid=;pwd=;"
                                    
         '--- make recordset for item
         Set adorsCamdata = New ADODB.Recordset
         If usrRecID.Rel = "000" Then
            strSQL = "SELECT [Order Number],Release,Item,[Scheduled Ship Date]," & _
                     "Quantity,Material,BarWidth " & _
                     "FROM mubbarff " & _
                     "WHERE ([Order Number] = '" & usrRecID.Job & "'" & _
                     "AND Release Is Null " & _
                     "AND Item = " & usrRecID.Item & ")"
         Else
            strSQL = "SELECT [Order Number],Release,Item,[Scheduled Ship Date]," & _
                     "Quantity,Material,BarWidth " & _
                     "FROM mubbarff " & _
                     "WHERE ([Order Number] = '" & usrRecID.Job & "'" & _
                     "AND Release = '" & usrRecID.Rel & "'" & _
                     "AND Item = " & usrRecID.Item & ")"
         End If
         adorsCamdata.Open strSQL, conCamdata, adOpenStatic, adLockOptimistic
   
         If adorsCamdata.RecordCount > 0 Then         'if a Record is found
            adorsCamdata.MoveFirst
            usrSelItem.Job = adorsCamdata.Fields("Order Number")
            If IsNull(adorsCamdata.Fields("Release")) Then
               usrSelItem.Rel = "000"
            Else
               usrSelItem.Rel = adorsCamdata.Fields("Release")
            End If
            usrSelItem.Item = adorsCamdata.Fields("Item")
            usrSelItem.ShipDate = adorsCamdata.Fields("Scheduled Ship Date")
            usrSelItem.Qnty = adorsCamdata.Fields("Quantity")
            usrSelItem.Mat = adorsCamdata.Fields("Material")
            usrSelItem.Width = adorsCamdata.Fields("BarWidth")
         Else                                          'if not found
            usrSelItem.Job = usrRecID.Job
            usrSelItem.Rel = usrRecID.Rel
            usrSelItem.Item = usrRecID.Item
            usrSelItem.ShipDate = Date
            usrSelItem.Qnty = 0
            usrSelItem.Mat = " "
            usrSelItem.Width = " "
         End If
         
         '--- clean up
         adorsCamdata.Close                        'close recordset
         Set adorsCamdata = Nothing                'unload recordset
         conCamdata.Close                          'close connection
         Set conCamdata = Nothing                  'unload connection
         
         '--- put item in Select que
         intAddRes = 0
         addSel usrSelItem, intAddRes              'add to select
         If intAddRes > 2 Then
            MsgBox ("Sendback: Item NOT added back to Select Que!")
         End If
              
      Else  '--- OP chose No
         'place holder for NO code
      End If
   Else                                            'no part selected
      'MsgBox ("No part selected!")
      Beep
   End If
     
End Sub  'cmdSendBack_Click()

Private Sub cmdRefresh_Click()

RefreshRunQue
RefreshExecQue

End Sub

Private Sub cmdSuspend_Click()
   If BATCH = True Then    '------------------------- Batch in Progress
      SUSPEND = True                                  'set the SUSPEND flag
      cmdSuspend.Visible = False                      'hide button
      Shape2.Visible = True                           'reveal reminder
      Label33.Visible = True
   
      If BARinPROG = True Then
         '------------------------- Bar in Prog prompt -----------------------
         'set up msg box
         strMSG = "A bar is in progress." & Chr$(13) & _
                  "Do you want to wait for it to complete?"             ' Define msgbox test
         intStyle = vbYesNo + vbDefaultButton2                          ' Define msgbox buttons
         strTitle = "Bar in progress..."                                ' Define msgbox title
         intResponse = MsgBox(strMSG, intStyle, strTitle)
         If intResponse = vbYes Then                                    '--- OP chose Yes
            Exit Sub
         Else                                                           '--- OP chose Yes
            procSusp
         End If
      Else
         procSusp
      End If
   Else
      Beep
   End If   'BATCH
End Sub

Private Sub cmdTest_Click()
   procBatch
End Sub

Private Sub cmdView_Click()
   Dim i As Integer
   
   i = lbxExec.ListIndex                              'get index of selected part
   
   If i > -1 Then                                     'if a part was selected
      usrRecID.Pri = usrEQParts(i + 1).Priority
      frmViewPart.Show                                'display form
   Else                                               'no part selected
      'MsgBox ("No part selected for viewing!")
      Beep
   End If
   
End Sub

Private Sub Form_Load()

   If STARTUP = False Then                            'when 1st starting up
      setup
      FirstPass = True
   End If
   
   '--------------------------- set visibility for tracking --------------------
   If blnEnTRACK Then
      lblTruckHDR.Visible = True
      If EnFDRTruck Then                              'fdr truck
         lblFDR.Visible = True
         txtFDRTruck.Visible = True
      Else
         lblFDR.Visible = False
         txtFDRTruck.Visible = False
      End If
      If EnPITruck Then                               'plug-in truck
         lblPI.Visible = True
         txtPITruck.Visible = True
      Else
         lblPI.Visible = False
         txtPITruck.Visible = False
      End If
      If EnPIGTruck Then                              'plug-in truck
         lblPIG.Visible = True
         txtPIGTruck.Visible = True
      Else
         lblPIG.Visible = False
         txtPIGTruck.Visible = False
      End If
   End If

   '------------------------ set visibility for Optimization --------------------
   If blnEnOPT = True Then
      cmdGOTOBatch.Visible = True
   Else
      cmdGOTOBatch.Visible = False
   End If
   
   intPartSel = -1                                    'clear the select pointer
      
   RefreshRunQue
   RefreshExecQue
   
   '-------------------------- detect/handle prev batch -------------------------
   'turn on the prevBatch timer...so that prev batch screen is displayed once the
   'run screen is up.
   If intExecCnt > 0 Then
      tmrPrevBatch.Enabled = True
   End If
   
   '------------------------ Setup to Run App in Sys Tray ---------------------
   'Me.Show
   'Me.Refresh
   'With nid
   '  .cbSize = Len(nid)
   '  .hwnd = Me.hwnd
   '  .uId = vbNull
   '  .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   '  .uCallBackMessage = WM_MOUSEMOVE
   '  .hIcon = Me.Icon
   '  .szTip = "Mubea HMI" & vbNullChar
   'End With
   'Shell_NotifyIcon NIM_ADD, nid
   'Me.WindowState = vbMinimized               'starts form in minimized mode
   'Me.WindowState = vbNormal                  'starts form in normal mode
   
End Sub

Sub RefreshExecQue()

Dim strTemp As String
Dim i As Integer

'--- display batch status
txtBatchStat.Text = ""
If OPTIMIZED = True Then txtBatchStat.Text = "Optimized."
If BATCH = True Then txtBatchStat.Text = "In Progress!"

lbxExec.Clear
AdodcExec.Refresh
If Not AdodcExec.Recordset.EOF Then '---------------- Que NOT empty
   
   '--- Display Mat/Width/BlankCounts
   strBatchMat = AdodcExec.Recordset.Fields("Material")
   txtMat.Text = strBatchMat
   strBatchWidth = AdodcExec.Recordset.Fields("BarWidth")
   txtWidth.Text = strBatchWidth
   txtTotBlanks = intTotBlanks
   txtRemBlanks = intRemBlanks
   
     
   '--- prep for 1st part
   AdodcExec.Recordset.MoveFirst
   i = 1                                              'init parts array index
   intHiPri = 0                                       'init HI prior. variable
   intTempPri = -1                                    'init temp prior. holder
   '--- loop thru recordset for Exec table and build a display array for Exec Que Listbox.
   Do Until AdodcExec.Recordset.EOF
      '--- build array entry for part
      usrEQParts(i).FullJobNum = AdodcExec.Recordset.Fields("FullOrder")
      usrEQParts(i).Job = AdodcExec.Recordset.Fields("Order Number")
      usrEQParts(i).Rel = AdodcExec.Recordset.Fields("Release")
      usrEQParts(i).Item = AdodcExec.Recordset.Fields("Item")
      usrEQParts(i).Seq = AdodcExec.Recordset.Fields("Sequence Number")
      usrEQParts(i).ShipDate = AdodcExec.Recordset.Fields("Scheduled Ship Date")
      usrEQParts(i).Qnty = AdodcExec.Recordset.Fields("Quantity")
      
      
      
      usrEQParts(i).BldQnty = AdodcExec.Recordset.Fields("BldQnty")
      usrEQParts(i).Phase = AdodcExec.Recordset.Fields("Phase")
      usrEQParts(i).Leg = AdodcExec.Recordset.Fields("Leg")
      usrEQParts(i).Stack = AdodcExec.Recordset.Fields("Stack")
      usrEQParts(i).BarType = AdodcExec.Recordset.Fields("BarType")
      usrEQParts(i).Material = AdodcExec.Recordset.Fields("Material")
      usrEQParts(i).BarWidth = AdodcExec.Recordset.Fields("BarWidth")
      usrEQParts(i).BlankLength = AdodcExec.Recordset.Fields("BlankLength")
      
      
      
      usrEQParts(i).RunLength = AdodcExec.Recordset.Fields("RunLength")
      usrEQParts(i).E1fig = AdodcExec.Recordset.Fields("E1figure")
      usrEQParts(i).E1dim = AdodcExec.Recordset.Fields("E1dimension")
      usrEQParts(i).E2fig = AdodcExec.Recordset.Fields("E2figure")
      usrEQParts(i).E2dim = AdodcExec.Recordset.Fields("E2dimension")
      usrEQParts(i).Cdim = AdodcExec.Recordset.Fields("Cdimension")
      usrEQParts(i).C1dim = AdodcExec.Recordset.Fields("C1dimension")
      usrEQParts(i).Ddim = AdodcExec.Recordset.Fields("Ddimension")
      usrEQParts(i).D1dim = AdodcExec.Recordset.Fields("D1dimension")
      usrEQParts(i).Build = AdodcExec.Recordset.Fields("Build")
      usrEQParts(i).Status = AdodcExec.Recordset.Fields("Status")
      usrEQParts(i).Priority = AdodcExec.Recordset.Fields("Priority")
          
      '--- capture Highest priority in Que
      If usrEQParts(i).Priority > intHiPri Then
         intHiPri = usrEQParts(i).Priority
      End If
          
      '--- construct string for display
      strTemp = Format(usrEQParts(i).Priority, "00") & "  " & _
                Format(usrEQParts(i).Status, "@@") & "-" & _
                Format(usrEQParts(i).Build, "@") & "     " & _
                Format(usrEQParts(i).Job, "00000") & "  " & _
                Format(usrEQParts(i).Rel, "@@@") & "  " & _
                Format(usrEQParts(i).Item, "0000") & "  " & _
                Format(usrEQParts(i).Seq, "0000") & "       " & _
                Format(usrEQParts(i).Qnty, "00000") & "  " & _
                Format(usrEQParts(i).BldQnty, "00000") & "       " & _
                Format(usrEQParts(i).Phase, "@") & "    " & _
                Format(usrEQParts(i).Leg, "@") & "    " & _
                Format(usrEQParts(i).Stack, "@") & "      " & _
                Format(usrEQParts(i).BlankLength, "000000")

      lbxExec.AddItem (strTemp)                       'add string to listbox
      
      '--- incr for next part
      AdodcExec.Recordset.MoveNext
      i = i + 1
   Loop 'end of recordset loop
         
   intExecCnt = i - 1                                 'get #of parts in Exec que
   
   
   
         
   If intPartSel > -1 Then                            'check for a prev. selected part
      lbxExec.SELECTED(intPartSel) = True             'reselect the part
   End If
     
   lbxExec.Refresh
   frmRun.Refresh
   cmdView.Visible = True                             'make view button visible
Else  '---------------------------------------------- Que Empty
   cmdView.Visible = False                            'hide view button
   strBatchMat = ""
   txtMat.Text = ""
   strBatchWidth = ""
   txtWidth.Text = ""
   txtTotBlanks = 0
   txtRemBlanks = 0
   intExecCnt = 0
End If

If FirstPass = True Then     'jng -- Calc opt parmeters for recovered batch
     If intExecCnt > 0 Then     'jng-- make sure execute que not empty
        
        dblTotPartsLen = 0                  'jng
   
          AdodcExec.Recordset.MoveFirst                             'jng
           Do Until AdodcExec.Recordset.EOF                        'jng
             intPartQty = AdodcExec.Recordset.Fields("Quantity")  'jng
             dblPartLen = AdodcExec.Recordset.Fields("BlankLength")   'jng
             dblTotPartsLen = dblTotPartsLen + (dblPartLen * intPartQty) 'jng 3/31/08
             '--- jng - next record
             AdodcExec.Recordset.MoveNext                              'jng
           Loop                         '  jng end of Total Parts Length loop
      
   
                                      'jng
         intBatchBlanks = intTotBlanks                    'jng 03/31/08
         dblBatchLngth = (intBatchBlanks * StockLength)   'jng 03/31/08
         dblBatchOpt = dblTotPartsLen / dblBatchLngth     'jng  3/31/08
         dblBatchOpt = dblBatchOpt * 100                  'jng 3/31/08
         dblLongScrap = 0      'jng -- offall not availible for recovered Batch
         intOptType = 9        'jng- recovered batch
      End If                   'jng
      FirstPass = False
    End If                      'jng  end FirstPass calculations

lbxExec.Refresh      '????????????? why again ??????????????

End Sub

'This procedure allows selecting of parts by group.
' When an part is selected, this sub id's the part, gets its priority and
' and selects all parts in the que with that priority(ie all parts in this group)

Private Sub lbxExec_Click()
   Dim i, j As Integer
   Dim intIndex As Integer
   
   Dim intPri As Integer
   
   intIndex = lbxExec.ListIndex                       'index selected part
   intPri = usrEQParts(intIndex + 1).Priority         'get its priority
   
   j = lbxExec.ListCount - 1                          'range of loop
   For i = 0 To j                                     'loop thru listbox contents
      If usrEQParts(i + 1).Priority = intPri Then
         lbxExec.SELECTED(i) = True                   'select each part w/ this priority
      End If
   Next i   'listbox loop
   
End Sub

Private Sub optAuto_Click()
   If OPTIMIZED = True Then                  'only enter AUTO if batch optimized
      SINGLEGRP = False                               'reset SINGLEGRP flag
      PAUSE = False                                   'reset PAUSE flag
      Shape1.FillColor = &HC000&                      'run indicator = green
      startExec                                       'start auto execution
   Else                                               'batch NOT optimized
      MsgBox ("Batch must be optimized!")
      NoStopPROMPT = True                             'skip STOP prompt
      optStop.Value = True
   End If   'optimized
End Sub

Private Sub optPAUSE_Click()
   If OPTIMIZED = True Then                  'only enter AUTO if batch optimized
      SINGLEGRP = False                               'set SINGLEGRP flag
      PAUSE = True                                    'reset PAUSE flag
      Shape1.FillColor = &HFFFF&                      'run indicator-yellow
      startExec                                       'start auto execution
   Else                                               'batch NOT optimized
      MsgBox ("Batch must be optimized!")
      NoStopPROMPT = True                             'skip STOP prompt
      optStop.Value = True
   End If   'optimized
End Sub

Private Sub optSingle_Click()
   If OPTIMIZED = True Then                  'only enter AUTO if batch optimized
      SINGLEGRP = True                                'set SINGLEGRP flag
      Shape1.FillColor = &HFFFF00                     'run indicator-blue
      PAUSE = False                                   'reset PAUSE flag
      startExec                                       'start auto execution
   Else                                               'batch NOT optimized
      MsgBox ("Batch must be optimized!")
      NoStopPROMPT = True                             'skip STOP prompt
      optStop.Value = True
   End If   'optimized
End Sub

Private Sub optStop_Click()
   Dim strMSG, strTitle As String         'message box variables
   Dim intStyle, intResponse As Integer
   
   If ESTOP = True Or NoStopPROMPT = True Then   '------ NO prompt
      stopExec                                           'stop execution
   Else  '---------------------------------------------- "Are you sure" prompt
      '--- set up msgbox
      strMSG = "Are you sure you want to STOP Execution?"
      intStyle = vbYesNo + vbDefaultButton2              'Define msgbox buttons
      strTitle = "Stoping Execution...."                 'Define msgbox title
      '--- msgbox response
      intResponse = MsgBox(strMSG, intStyle, strTitle)
      If intResponse = vbYes Then  '-------------------- OP chose Yes.
         stopExec                                        'stop execution
      Else  '------------------------------------------- OP chose No
         Exit Sub
      End If
   End If   'NO prompt
   
End Sub  'optStop_Click()

Private Sub tcpServer_Close()
   If tcpServer.State <> sckClosed Then
      tcpServer.Close
   End If
   tcpServer.Listen
   
End Sub

'------------------------------------------------------------------------------
'Name:      tcpServer_DataArrival
'Accepts:   none
'Returns:   none
'Requires:
'Discrip:   gets data from the tcp connection and handles signals from MubeaOI
'Notes:
'------------------------------------------------------------------------------
Public Sub tcpServer_DataArrival(ByVal bytesTotal As Long)


   '--- variable declarations
   Dim strData As String      'message string
   Dim strMSG As String       'message code string
   Dim intMSG As Integer      'message code
   Dim intMSGLen As Integer   'message length
   Dim intTemp As Integer     'temp variable
   '---------------------------------------------------------------------------
   
   If pendMSG = True Then   '------------------------- handle pending message
      strData = txtTCPRecv.Text                       'pending message in txtbox
   Else
      tcpServer.GetData strData                       'get message data
      If strData <> "" Then
        intTemp = InStr(1, strData, Chr$(13), 1)
        strData = Left(strData, intTemp - 1)
        txtTCPRecv.Text = strData                       'copy to display
      End If
   End If   'pendMSG
   
   If AUTO = True Then  '---------------------------- ID the message (if in Auto)
      pendMSG = False                                 'reset pending flag
      Select Case strData                             'parse message
      Case "ESWO_ReqGroup"    '--- group request
         procReq                                      'process the request
      Case "ESWO_CmplGroup"   '--- group complete
         procCmpl                                     'process the complete signal
      Case Else               '--- other
         If Left$(strData, 10) = "ESWO_Mesg:" Then    'message string detected
            intMSGLen = Len(strData)                  'get str len
            strMSG = Right$(strData, (intMSGLen - 10)) 'get msg code
            intMSG = Val(strMSG)                      'conv to integer
            procMSG intMSG                            'process the message
            txtTCPRecv.Text = ""                      'clear the recv text
         Else                                         'unrecongnizable message
            MsgBox ("tcpServer: Did NOT recognize message from Mubea-OI!")
         End If
      End Select              '--- strData
   End If   '--- auto
End Sub  'tcpServer_DataArrival

Private Sub tcpClient_Connect()
Dim strData As String
   strData = txtTCPSend.Text
   tcpClient.SendData strData                           'send the message
End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   'Select Case Number
   'Case 10061
   '   Beep
   'Case Else
   '   MsgBox ("Error#:" & Number & " = " & Description)
   'End Select
End Sub

Private Sub tcpClient_SendComplete()
   tcpClient.Close
End Sub

Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
   If tcpServer.State <> sckClosed Then               'check if State is closed
      tcpServer.Close                                 'if NOT, close before accepting
   End If
   tcpServer.Accept requestID                         'accept the request
End Sub

Private Sub tmrPrevBatch_Timer()
   tmrPrevBatch.Enabled = False           'turn off the prev Batch display timer
   frmPrevBatch.Show vbModal              'show the prev batch screen
End Sub

Private Sub tmrTCPStat_Timer()
   txtClientStat = tcpClient.State
   txtServerStat = tcpServer.State
   'Select Case tcpConn.State
   'Case 2, 3, 4, 5, 6
   '   txtConn.Text = "Connecting..."
   'Case 7
   '   txtConn.Text = "Connected."
   'Case 8, 9
   '   txtConn.Text = "Disconnected!"
   'End Select
End Sub

Private Sub txtFDRTruck_DblClick()
   strTruckType = "FDR"
   frmTruck.Show
End Sub

Private Sub txtPIGTruck_DblClick()
   strTruckType = "PIG"
   frmTruck.Show
End Sub

Private Sub txtPITruck_DblClick()
   strTruckType = "PI"
   frmTruck.Show
End Sub

Private Sub txtStkLength_Change()
   StockLength = Val(txtStkLength.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Me.WindowState = vbMinimized
 Cancel = True
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As _
'   Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
'Dim result As Long
'Dim Msg As Long
 'the value of X will vary depending upon the scalemode setting
' If Me.ScaleMode = vbPixels Then
'  Msg = X
' Else
' Msg = X / Screen.TwipsPerPixelX
' End If
' Select Case Msg
'  Case WM_LBUTTONUP        '514 restore form window
'   Me.WindowState = vbNormal
'   result = SetForegroundWindow(Me.hwnd)
'   Me.Show
'  Case WM_LBUTTONDBLCLK    '515 restore form window
'   Me.WindowState = vbNormal
'   result = SetForegroundWindow(Me.hwnd)
'  Me.Show
'  Case WM_RBUTTONUP        '517 display popup menu
'   result = SetForegroundWindow(Me.hwnd)
'   If Y = 0 Then
'     Me.PopupMenu Me.mPopupSys
'  End If
' End Select
'End Sub

'Private Sub Form_Resize()
 'this is necessary to assure that the minimized window is hidden
' If Me.WindowState = vbMinimized Then Me.Hide
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
 'this removes the icon from the system tray
' Shell_NotifyIcon NIM_DELETE, nid
'End Sub
'Private Sub mPopExit_Click()
'called when user clicks the popup menu Exit command
'Dim Msg   ' Declare variable.
'   Msg = "Do you really want to stop the Mubea HMI? "
'   'If user clicks the No button, stop QueryUnload.
'   If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
'   AdodcExec.Recordset.Close                 'close recordsets
'   AdodcRun.Recordset.Close
'   Shell_NotifyIcon NIM_DELETE, nid
'   End
'End Sub

'Private Sub mPopRestore_Click()
 'called when the user clicks the popup menu Restore command
' Me.WindowState = vbNormal
' result = SetForegroundWindow(Me.hwnd)
' Me.Show
'End Sub


