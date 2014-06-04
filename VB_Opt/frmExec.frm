VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmExec 
   Caption         =   "Run Screen"
   ClientHeight    =   11115
   ClientLeft      =   810
   ClientTop       =   540
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   13320
   Begin VB.TextBox txtSubQty 
      Height          =   285
      Left            =   6600
      TabIndex        =   56
      Top             =   960
      Width           =   495
   End
   Begin VB.Timer tmrExec 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   12600
      Top             =   1320
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "Deselect"
      Height          =   375
      Left            =   12000
      TabIndex        =   54
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtPcBldQty 
      Height          =   285
      Left            =   7560
      TabIndex        =   51
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtPcQty 
      Height          =   285
      Left            =   5640
      TabIndex        =   50
      Top             =   960
      Width           =   495
   End
   Begin MSAdodcLib.Adodc AdodcRun 
      Height          =   375
      Left            =   2400
      Top             =   10680
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
      Connect         =   "DSN=dsnLocal_RemHsg"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "dsnLocal_RemHsg"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblExecQue Where Priority = 0"
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
   Begin VB.CommandButton cmdGOTOComp 
      Caption         =   "  Complete       Screen"
      Height          =   495
      Left            =   12000
      TabIndex        =   27
      Top             =   9360
      Width           =   975
   End
   Begin VB.CommandButton cmdGOTOMain 
      Caption         =   "  Main Screen"
      Height          =   495
      Left            =   12000
      TabIndex        =   26
      Top             =   10080
      Width           =   975
   End
   Begin VB.CommandButton cmdGOTOSelect 
      Caption         =   " Select Screen"
      Height          =   495
      Left            =   12000
      TabIndex        =   25
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton cmdSendBack 
      Caption         =   "Send Back"
      Height          =   375
      Left            =   12000
      TabIndex        =   24
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtPcMachQty 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10800
      TabIndex        =   22
      Text            =   "0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtPLCMessage 
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
      TabIndex        =   19
      Text            =   "How 'bout dis!!!!!"
      Top             =   3120
      Width           =   6015
   End
   Begin VB.CommandButton cmdSuspend 
      Caption         =   "Suspend"
      Height          =   375
      Left            =   8280
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   12000
      TabIndex        =   16
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame FrameRunMode 
      Caption         =   "Run Mode"
      Height          =   1455
      Left            =   8280
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
      Begin VB.OptionButton optStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optPause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optSingle 
         Caption         =   "Single Part"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Auto"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdDecPriority 
      Caption         =   "DN"
      Height          =   495
      Left            =   12240
      TabIndex        =   11
      Top             =   7080
      Width           =   495
   End
   Begin VB.CommandButton cmdIncPriority 
      Caption         =   "UP"
      Height          =   495
      Left            =   12240
      TabIndex        =   10
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox txtJob 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtRel 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Index           =   1
      Left            =   12120
      TabIndex        =   4
      ToolTipText     =   "Submit selected items for execution."
      Top             =   240
      Width           =   975
   End
   Begin VB.ListBox lbxExec 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   4440
      Width           =   11535
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   11880
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdodcExec 
      Height          =   375
      Left            =   240
      Top             =   10680
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
      Connect         =   "DSN=dsnLocal_RemHsg"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "dsnLocal_RemHsg"
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   240
      TabIndex        =   41
      Top             =   1680
      Width           =   7815
   End
   Begin VB.Label Label31 
      Caption         =   "SubQty"
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
      Left            =   6480
      TabIndex        =   57
      Top             =   720
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
      Left            =   6480
      TabIndex        =   55
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label29 
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
      Left            =   5640
      TabIndex        =   53
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label28 
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
      Left            =   7440
      TabIndex        =   52
      Top             =   720
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
      Left            =   240
      TabIndex        =   49
      Top             =   1440
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
      Left            =   2160
      TabIndex        =   48
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label25 
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
      Left            =   2880
      TabIndex        =   47
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label24 
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
      Left            =   3480
      TabIndex        =   46
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label23 
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
      Left            =   4440
      TabIndex        =   45
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label22 
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
      Left            =   5160
      TabIndex        =   44
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "RunLength"
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
      TabIndex        =   43
      Top             =   1440
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
      Left            =   1560
      TabIndex        =   42
      Top             =   1440
      Width           =   495
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
      Left            =   5160
      TabIndex        =   40
      Top             =   4200
      Width           =   495
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
      TabIndex        =   39
      Top             =   4200
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
      Left            =   840
      TabIndex        =   38
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "RunLength"
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
      Left            =   10200
      TabIndex        =   37
      Top             =   4200
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
      Left            =   9360
      TabIndex        =   36
      Top             =   4200
      Width           =   735
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
      Left            =   8640
      TabIndex        =   35
      Top             =   4200
      Width           =   495
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
      Left            =   7680
      TabIndex        =   34
      Top             =   4200
      Width           =   855
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
      Left            =   7080
      TabIndex        =   33
      Top             =   4200
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
      Left            =   5760
      TabIndex        =   32
      Top             =   4200
      Width           =   495
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
      Left            =   4080
      TabIndex        =   31
      Top             =   4200
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
      Left            =   3480
      TabIndex        =   30
      Top             =   4200
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
      Left            =   2880
      TabIndex        =   29
      Top             =   4200
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
      Left            =   2160
      TabIndex        =   28
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "MachQty:"
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
      Left            =   9720
      TabIndex        =   21
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "PLC Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   3120
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
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblSelectTitle 
      Caption         =   "Remelle Housing Run Screen"
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   5535
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
      Left            =   1320
      TabIndex        =   3
      Top             =   720
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
      Left            =   240
      TabIndex        =   2
      Top             =   720
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
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "frmExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usrEQParts(100) As Part      'Exec Listbox Storage
Dim intPartSel As Integer        'pointer to which part is selected

Sub RefreshRunQue()
'--- globals used
'usrRUNParts                     'Run Que Listbox Storage = Run Parts array
'intRunCnt                       '# of parts in Run Que
'--- variable declarations
Dim strTemp As String
Dim i As Integer

lbxRun.Clear
AdodcRun.Refresh
If Not AdodcRun.Recordset.EOF Then        'if NOT at EOF, there are parts in que
   
   '--- prep for 1st part
   AdodcRun.Recordset.MoveFirst
   i = 1                                  'init parts array index
   
   '--- Display Job/Rel/Item
   txtJob.Text = AdodcRun.Recordset.Fields("Order Number")
   txtRel.Text = AdodcRun.Recordset.Fields("Release")
   txtItem.Text = AdodcRun.Recordset.Fields("Item")
   
   '--- loop thru recordset for Exec table and build a display array for Exec Que Listbox.
   Do Until AdodcRun.Recordset.EOF
      '--- build array entry for part
      usrRUNParts(i).FullJobNum = AdodcRun.Recordset.Fields("FullOrder")
      usrRUNParts(i).Job = AdodcRun.Recordset.Fields("Order Number")
      usrRUNParts(i).Rel = AdodcRun.Recordset.Fields("Release")
      usrRUNParts(i).Item = AdodcRun.Recordset.Fields("Item")
      usrRUNParts(i).Seq = AdodcRun.Recordset.Fields("Sequence Number")
      usrRUNParts(i).ShipDate = AdodcRun.Recordset.Fields("Scheduled Ship Date")
      usrRUNParts(i).Qnty = AdodcRun.Recordset.Fields("Quantity")
      usrRUNParts(i).BldQnty = AdodcRun.Recordset.Fields("BldQnty")
      usrRUNParts(i).Phase = AdodcRun.Recordset.Fields("Phase")
      usrRUNParts(i).Leg = AdodcRun.Recordset.Fields("Leg")
      usrRUNParts(i).Stack = AdodcRun.Recordset.Fields("Stack")
      usrRUNParts(i).HsgType = AdodcRun.Recordset.Fields("HsgType")
      usrRUNParts(i).Material = AdodcRun.Recordset.Fields("Material")
      usrRUNParts(i).BarWidth = AdodcRun.Recordset.Fields("BarWidth")
      usrRUNParts(i).BlankLength = AdodcRun.Recordset.Fields("BlankLength")
      usrRUNParts(i).RunLength = AdodcRun.Recordset.Fields("RunLength")
      usrRUNParts(i).E1fig = AdodcRun.Recordset.Fields("E1figure")
      usrRUNParts(i).E1dim = AdodcRun.Recordset.Fields("E1dimension")
      usrRUNParts(i).E2fig = AdodcRun.Recordset.Fields("E2figure")
      usrRUNParts(i).E2dim = AdodcRun.Recordset.Fields("E2dimension")
      usrRUNParts(i).Cdim = AdodcRun.Recordset.Fields("Cdimension")
      usrRUNParts(i).C1dim = AdodcRun.Recordset.Fields("C1dimension")
      usrRUNParts(i).Ddim = AdodcRun.Recordset.Fields("Ddimension")
      usrRUNParts(i).D1dim = AdodcRun.Recordset.Fields("D1dimension")
      usrRUNParts(i).Build = AdodcRun.Recordset.Fields("Build")
      usrRUNParts(i).Status = AdodcRun.Recordset.Fields("Status")
      usrRUNParts(i).Priority = AdodcRun.Recordset.Fields("Priority")
          
      '--- construct string and add to listbox for display
      strTemp = Format(usrRUNParts(i).Seq, "@@@") & "                          " & _
                Format(usrRUNParts(i).Phase, "@") & "            " & _
                Format(usrRUNParts(i).Leg, "@") & "           " & _
                Format(usrRUNParts(i).Qnty, "@@@@@") & "               " & _
                Format(usrRUNParts(i).BldQnty, "@@@@@") & "          " & _
                Format(usrRUNParts(i).Material, "@") & "               " & _
                Format(usrRUNParts(i).BarWidth, "@") & "                    " & _
                Format(usrRUNParts(i).RunLength, "@@@@@@@@@")
      lbxRun.AddItem (strTemp)
      
      '--- incr for next part
      AdodcRun.Recordset.MoveNext
      i = i + 1
   Loop 'end of recordset loop
   
   intRunCnt = i - 1                   'get #of parts in run que
   lbxRun.Refresh
   frmExec.Refresh
   
Else  'no parts in RUN que
   txtJob.Text = ""                    'clear the Job/Rel/Item boxes
   txtRel.Text = ""
   txtItem.Text = ""
   
End If

lbxRun.Refresh      '????????????? why again ??????????????

End Sub


'--- This sub deselects all items in the Exec Que

Public Sub deSelExec()

Dim i, j As Integer

   j = lbxExec.ListCount - 1                 'get loop range

   For i = 0 To j                            'loop thru listbox contents
      lbxExec.Selected(i) = False            'deselect each part
   Next i   'listbox loop
   
   intPartSel = -1                           'reset select pointer
   
   RefreshExecQue                            'refresh the que
   
End Sub  'deSelExec



Private Sub cmdDecPriority_Click()
   Dim i As Integer                          'list box index
   Dim j As Integer                          'loop range variable
   Dim intTemp As Integer                    'temp priority variable
   Dim intTHISpri As Integer                 'priority for selected part/group
   Dim intNEXTpri As Integer                 'priority for prev part/group
   Dim intNOSwap  As Integer                 'flag for swapping this/prev priorities
   Dim conRHLocal As Connection
   Dim adorsTHISGroup As ADODB.Recordset
   Dim adorsNEXTGroup As ADODB.Recordset
   
   '--- get index of selected part
   j = lbxExec.ListCount - 1                          'set loop range
   For i = 0 To j                                     'loop thru listbox
      If lbxExec.Selected(i) = True Then
         Exit For
      End If
   Next 'listbox loop
  
   If lbxExec.SelCount > 0 Then                       'if a part was selected
      intTHISpri = usrEQParts(i + 1).Priority         'get the priority for the selected part
      intNEXTpri = intTHISpri + 1                     'get the priority for the prev part/group
      
      If (i + lbxExec.SelCount) > j Then                                 'already at bottom of que
         'MsgBox ("Already at the Bottom of the que.")
         Beep
      Else
         intPartSel = i                               'mark part to maint. select
         
         '--- make connection to local database
         Set conRHLocal = New Connection
         conRHLocal.Open "PROVIDER=MSDASQL;dsn=dsnRHLocal;uid=;pwd=;"
                                    
         '--- create a recordset for all parts in selected group
         Set adorsTHISGroup = New ADODB.Recordset     'init recordset
                                                      'build the SQL string
         strSQL = "SELECT * FROM tblExecQue WHERE " & _
                  "[Priority] = " & intTHISpri
                                                      'open recordset
         adorsTHISGroup.Open strSQL, conRHLocal, adOpenStatic, adLockOptimistic
         
         If adorsTHISGroup.RecordCount < 1 Then       'No records found
            MsgBox ("No records with this priority found!")
            adorsTHISGroup.Close                      'close recordset
            Set adorsTHISGroup = Nothing              'unload recordset
            conRHLocal.Close                          'close connection
            Set conRHLocal = Nothing                  'unload connection
            Exit Sub
         End If
             
         '--- create a recordset for all parts in Next group
         Set adorsNEXTGroup = New ADODB.Recordset     'init recordset
                                                      'build the SQL string
         strSQL = "SELECT * FROM tblExecQue WHERE " & _
                  "[Priority] = " & intNEXTpri
                                                      'open recordset
         adorsNEXTGroup.Open strSQL, conRHLocal, adOpenStatic, adLockOptimistic
         
         If adorsNEXTGroup.RecordCount < 1 Then       'No records found,must have skipped a priority#
            intNOSwap = 1                             'don't change priority for prev item
         End If
         
         
         '--- loop to sub 1 from priority of this group
         With adorsTHISGroup
            .MoveFirst
            Do Until .EOF
               .Fields("Priority") = intTHISpri + 1
               .Update
               .MoveNext
            Loop  'this group
         End With
            
         '--- loop to add 1 to priority of next group
         If intNOSwap <> 1 Then                       'don't exec this loop if NOSwap flag set
            With adorsNEXTGroup
               .MoveFirst
               Do Until .EOF
                  .Fields("Priority") = intNEXTpri - 1
                  .Update
                  .MoveNext
               Loop  'this group
            End With
            intPartSel = intPartSel + adorsNEXTGroup.RecordCount  'make sure part remains selected
         End If
          
         '--- clean up
         adorsTHISGroup.Close                         'close recordset
         Set adorsTHISGroup = Nothing                 'unload recordset
         adorsNEXTGroup.Close                         'close recordset
         Set adorsNEXTGroup = Nothing                 'unload recordset
         conRHLocal.Close                             'close connection
         Set conRHLocal = Nothing                     'unload connection
                   
         RefreshExecQue                               'refresh the que
              
      End If
   Else                                               'no part selected
      'MsgBox ("No part selected!")
      Beep
   End If
End Sub  'cmdDecPriority_Click()

Private Sub cmdDeselect_Click()

   deSelExec                                 'deselct exec que
   
End Sub

Private Sub cmdExit_Click(Index As Integer)
       
  AdodcRun.Recordset.Close         'close recordset
  Unload Me                         'unload form

End Sub

Private Sub cmdGOTOComp_Click()
   Me.Hide
   frmHist.Show
End Sub

Private Sub cmdGOTOSelect_Click()
   Me.Hide                                   'hide the Select Screen
   frmSelect.Show                        'display Execution form
End Sub

' Increment Priority (move up) Notes
'------------------------------------------------------------------------
' To move a part up in the Exec Que, ie raise its priority
'  - ID selected Item
'  - If Items priority is not already 1, ie not at top of que, then
      '- sub 1 from its priority
      '- move to prev record
      '- add 1 to prev record's priority
      '- update the recordset
      '- refresh the que
'------------------------------------------------------------------------

Private Sub cmdIncPriority_Click()
   Dim i As Integer                          'list box index
   Dim j As Integer                          'loop range variable
   Dim intTemp As Integer                    'temp priority variable
   Dim intTHISpri As Integer                 'priority for selected part/group
   Dim intPREVpri As Integer                 'priority for prev part/group
   Dim intNOSwap  As Integer                 'flag for swapping this/prev priorities
   Dim conRHLocal As Connection
   Dim adorsTHISGroup As ADODB.Recordset
   Dim adorsPREVGroup As ADODB.Recordset
   
   '--- get index of selected part
   j = lbxExec.ListCount - 1                          'set loop range
   For i = 0 To j                                     'loop thru listbox
      If lbxExec.Selected(i) = True Then
         Exit For
      End If
   Next 'listbox loop
  
   If lbxExec.SelCount > 0 Then                       'if a part was selected
      intTHISpri = usrEQParts(i + 1).Priority         'get the priority for the selected part
      intPREVpri = intTHISpri - 1                     'get the priority for the prev part/group
      
      If intTHISpri < 2 Then                          'already at top of que
         'MsgBox ("Already at the top of the que.")
         Beep
      Else
         intPartSel = i                               'mark part to maint. select
         
         '--- make connection to local database
         Set conRHLocal = New Connection
         conRHLocal.Open "PROVIDER=MSDASQL;dsn=dsnRHLocal;uid=;pwd=;"
                                    
         '--- create a recordset for all parts in selected group
         Set adorsTHISGroup = New ADODB.Recordset     'init recordset
                                                      'build the SQL string
         strSQL = "SELECT * FROM tblExecQue WHERE " & _
                  "[Priority] = " & intTHISpri
                                                      'open recordset
         adorsTHISGroup.Open strSQL, conRHLocal, adOpenStatic, adLockOptimistic
         
         If adorsTHISGroup.RecordCount < 1 Then       'No records found
            MsgBox ("No records with this priority found!")
            adorsTHISGroup.Close                      'close recordset
            Set adorsTHISGroup = Nothing              'unload recordset
            conRHLocal.Close                          'close connection
            Set conRHLocal = Nothing                  'unload connection
            Exit Sub
         End If
             
         '--- create a recordset for all parts in prev group
         Set adorsPREVGroup = New ADODB.Recordset     'init recordset
                                                      'build the SQL string
         strSQL = "SELECT * FROM tblExecQue WHERE " & _
                  "[Priority] = " & intPREVpri
                                                      'open recordset
         adorsPREVGroup.Open strSQL, conRHLocal, adOpenStatic, adLockOptimistic
         
         If adorsPREVGroup.RecordCount < 1 Then       'No records found,must have skipped a priority#
            intNOSwap = 1                             'don't change priority for prev item
         End If
         
         
         '--- loop to sub 1 from priority of this group
         With adorsTHISGroup
            .MoveFirst
            Do Until .EOF
               .Fields("Priority") = intTHISpri - 1
               .Update
               .MoveNext
            Loop  'this group
         End With
            
         '--- loop to add 1 to priority of next group
         If intNOSwap <> 1 Then                       'don't exec this loop if NOSwap flag set
            With adorsPREVGroup
               .MoveFirst
               Do Until .EOF
                  .Fields("Priority") = intPREVpri + 1
                  .Update
                  .MoveNext
               Loop  'this group
            End With
            
            intPartSel = intPartSel - adorsPREVGroup.RecordCount  'make sure part remains selected
         End If
          
         '--- clean up
         adorsTHISGroup.Close                         'close recordset
         Set adorsTHISGroup = Nothing                 'unload recordset
         adorsPREVGroup.Close                         'close recordset
         Set adorsPREVGroup = Nothing                 'unload recordset
         conRHLocal.Close                             'close connection
         Set conRHLocal = Nothing                     'unload connection
                   
         RefreshExecQue                               'refresh the que
              
      End If
   Else                                               'no part selected
      'MsgBox ("No part selected!")
      Beep
   End If
End Sub  'cmdIncPriority_Click()

'Send Back Notes
'------------------------------------------------------------------------
' To send a part from the Exec Que back to the Select Que

'

'  - build a item based on info for selected part
'
'  - loop thru recordset and delete all parts w/ this JRI
'  - if any part fails to delete, then notify the operator and abort sendback
'  - if all parts delete, update adodc
'
'  - add the built item to the sel que
'  - if item fails to add notify operator
'  - update sel que
'------------------------------------------------------------------------

Private Sub cmdSendBack_Click()




Dim i, j, intDone As Integer
    
   '---loop set up
   i = 0
   j = lbxExec.ListCount
   intDone = 0
   
   '--- indentify the selected part
   Do Until intDone = 1 Or i = j                   'loop thru listbox contents
      If lbxExec.Selected(i) = True Then           'if row is selected
         usrRecID.Job = usrEQParts(i + 1).Job      'set record id
         usrRecID.Rel = usrEQParts(i + 1).Rel
         usrRecID.Item = usrEQParts(i + 1).Item
         usrRecID.Seq = usrEQParts(i + 1).Seq
         intDone = 1
      End If
      i = i + 1
   Loop  'end of listbox loop
   
   If intDone = 1 Then                             'if a part was selected
      '--- select all parts w/ this JRI
      For i = 1 To j
         If usrEQParts(i).Job = usrRecID.Job And _
            usrEQParts(i).Rel = usrRecID.Rel And _
            usrEQParts(i).Item = usrRecID.Item Then
            lbxExec.Selected(i - 1) = True         'select the part
         End If
      Next i
      
      '--- are you sure prompt
      'set up msg box
      strMSG = "Are you sure you want to send selected item back?"      ' Define msgbox test
      intStyle = vbYesNo + vbDefaultButton2        ' Define msgbox buttons
      strTitle = "Sending Item back to Select Que..."             ' Define msgbox title
      intResponse = MsgBox(strMSG, intStyle, strTitle)
      If intResponse = vbYes Then  '--- OP chose Yes.
         'place holder for yes code
      
      
      
      
      Else  '--- OP chose No, deselect and exit
         deSelExec
         Exit Sub
      End If
      
   Else                                            'no part selected
      'MsgBox ("No part selected!")
      Beep
   End If
     
End Sub

Private Sub cmdRefresh_Click()

RefreshExecQue

End Sub



Private Sub cmdSuspend_Click()
   Suspend = True                                  'set the suspend flag
End Sub

Private Sub cmdView_Click()
   Dim i As Integer
   
   i = lbxExec.ListIndex                              'get index of selected part
   
   If i > -1 Then                                     'if a part was selected
      usrRecID.Pri = usrEQParts(i + 1).Priority
      frmViewPart.Show                                'display form
   Else                                               'no part selected
      MsgBox ("No part selected for viewing!")
   End If
   
End Sub

Private Sub Form_Load()

If StartUp = False Then                            'when 1st starting up
   setup
End If

intPartSel = -1                                    'clear the select pointer

RefreshRunQue
RefreshExecQue

End Sub

Sub RefreshExecQue()

Dim strTemp As String
Dim i As Integer

lbxExec.Clear
AdodcExec.Refresh
If Not AdodcExec.Recordset.EOF Then       'don't display if nothing in que
   
   '--- prep for 1st part
   AdodcExec.Recordset.MoveFirst
   i = 1                                  'init parts array index
   
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
      usrEQParts(i).HsgType = AdodcExec.Recordset.Fields("HsgType")
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
          
      '--- construct string for display
      strTemp = Format(usrEQParts(i).Priority, "@@@@") & "        " & _
                Format(usrEQParts(i).Status, "@@") & " - " & _
                Format(usrEQParts(i).Build, "@") & "                " & _
                Format(usrEQParts(i).Job, "@@@@@") & "     " & _
                Format(usrEQParts(i).Rel, "@@@") & "   " & _
                Format(usrEQParts(i).Item, "@@@@@") & "        " & _
                Format(usrEQParts(i).Seq, "@@@") & "                       " & _
                Format(usrEQParts(i).Phase, "@") & "           " & _
                Format(usrEQParts(i).Leg, "@") & "           " & _
                Format(usrEQParts(i).Stack, "@") & "         " & _
                Format(usrEQParts(i).Qnty, "@@@@@") & "           " & _
                Format(usrEQParts(i).BldQnty, "@@@@@") & "             " & _
                Format(usrEQParts(i).Material, "@") & "                 " & _
                Format(usrEQParts(i).BarWidth, "@") & "                 " & _
                Format(usrEQParts(i).RunLength, "@@@@@@@@@")

      lbxExec.AddItem (strTemp)                   'add string to listbox
      
      '--- incr for next part
      AdodcExec.Recordset.MoveNext
      i = i + 1
   Loop 'end of recordset loop
         
   If intPartSel > -1 Then                         'check for a prev. selected part
      lbxExec.Selected(intPartSel) = True          'reselect the part
   End If
     
   lbxExec.Refresh
   frmExec.Refresh
   
End If

lbxExec.Refresh      '????????????? why again ??????????????

End Sub

Private Sub cmdSubmitPriority_Click(Index As Integer)
    MsgBox "This function not implemented."
End Sub

Private Sub cmdSubmitAll_Click(Index As Integer)
    MsgBox "This function not implemented."
End Sub


'This procedure allows selecting of parts by group.
' When an part is selected, this sub id's the part, gets its priority and
' and selects all parts in the que with that prioroty(ie all parts in this group)

Private Sub lbxExec_Click()
   Dim i, j As Integer
   Dim intIndex As Integer
   
   Dim intPri As Integer
   
   intIndex = lbxExec.ListIndex                       'index selected part
   intPri = usrEQParts(intIndex + 1).Priority         'get its priority
   
   j = lbxExec.ListCount - 1                          'range of loop
   For i = 0 To j                                     'loop thru listbox contents
      If usrEQParts(i + 1).Priority = intPri Then
         lbxExec.Selected(i) = True                   'select each part w/ this priority
      End If
   Next i   'listbox loop
   
End Sub

'---------------------------------------------------------------------------
'Name:      tmrExec_Timer()
'Accepts:   none
'Returns:   none
'Requires:
'Discrip:   This is the done timer routine for the exec timer.
'Notes:  (1)The sub immediatly dispables the exec timer to ensure that the
'           exec sub is not called again while still executing.
'---------------------------------------------------------------------------

Private Sub tmrExec_Timer()
   tmrExec.Enabled = False                            'turn off Execution timer
   Exec                                               'run the execution sub
End Sub

