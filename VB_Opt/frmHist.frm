VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHist 
   Caption         =   "History Screen"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   12180
   Begin VB.Frame Frame1 
      Caption         =   "Search by: "
      Height          =   615
      Left            =   8880
      TabIndex        =   39
      Top             =   960
      Width           =   3225
      Begin VB.OptionButton OptDate 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton OptJRI 
         Caption         =   "JRI"
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optTruck 
         Caption         =   "Truck"
         Height          =   195
         Left            =   2295
         TabIndex        =   40
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.TextBox txtTruck 
      Height          =   375
      Left            =   5370
      TabIndex        =   37
      Top             =   1830
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Fin. Batch"
      Height          =   375
      Left            =   10920
      TabIndex        =   35
      ToolTipText     =   "Remake Selected Item"
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdRemake 
      Caption         =   "Remake"
      Height          =   375
      Left            =   10920
      TabIndex        =   34
      ToolTipText     =   "Remake Selected Item"
      Top             =   3840
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5400
      TabIndex        =   33
      ToolTipText     =   "Search Date"
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22806529
      CurrentDate     =   37642
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   10920
      TabIndex        =   26
      ToolTipText     =   "Perform Search"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   10920
      TabIndex        =   16
      ToolTipText     =   "Refresh the History Que"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdGOTORun 
      Caption         =   "  Run  Screen"
      Height          =   495
      Left            =   10920
      TabIndex        =   3
      ToolTipText     =   "Go to Run Screen"
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   10920
      TabIndex        =   1
      ToolTipText     =   "View Selected Part"
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox lbxHist 
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
      Height          =   5130
      Left            =   225
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   2880
      Width           =   10575
   End
   Begin VB.TextBox txtJob 
      Height          =   285
      Left            =   4680
      TabIndex        =   27
      ToolTipText     =   "Enter Job# to search for"
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Left            =   6600
      TabIndex        =   28
      ToolTipText     =   "Enter Item to search for"
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtRel 
      Height          =   285
      Left            =   5760
      TabIndex        =   29
      ToolTipText     =   "Enter Release to search for"
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label28 
      Caption         =   "Truck#"
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
      Left            =   9810
      TabIndex        =   43
      Top             =   2625
      Width           =   825
   End
   Begin VB.Label Label27 
      Caption         =   "Truck#:"
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
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label26 
      Caption         =   "Sta"
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
      TabIndex        =   36
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label22 
      Caption         =   "HISTORY  SCREEN"
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
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label20 
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
      TabIndex        =   23
      Top             =   240
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   1080
      X2              =   12240
      Y1              =   840
      Y2              =   840
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
      TabIndex        =   21
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "S"
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
      Left            =   5895
      TabIndex        =   20
      Top             =   2625
      Width           =   165
   End
   Begin VB.Label Label3 
      Caption         =   " Blank"
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
      Left            =   7485
      TabIndex        =   19
      Top             =   2385
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Build"
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
      Left            =   4380
      TabIndex        =   18
      Top             =   2400
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
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
      Left            =   8520
      TabIndex        =   17
      Top             =   2625
      Width           =   585
   End
   Begin VB.Label Label19 
      Caption         =   "P"
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
      Left            =   5190
      TabIndex        =   15
      Top             =   2625
      Width           =   165
   End
   Begin VB.Label Label17 
      Caption         =   "Bld"
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
      Left            =   690
      TabIndex        =   14
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label Label16 
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
      Left            =   7485
      TabIndex        =   13
      Top             =   2625
      Width           =   735
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
      Left            =   6645
      TabIndex        =   12
      Top             =   2625
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "M"
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
      Left            =   6255
      TabIndex        =   11
      Top             =   2625
      Width           =   240
   End
   Begin VB.Label Label13 
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
      Left            =   4455
      TabIndex        =   10
      Top             =   2640
      Width           =   390
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
      Left            =   3750
      TabIndex        =   9
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label Label11 
      Caption         =   "L"
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
      Left            =   5550
      TabIndex        =   8
      Top             =   2625
      Width           =   255
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
      Left            =   2955
      TabIndex        =   7
      Top             =   2640
      Width           =   450
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
      Left            =   2340
      TabIndex        =   6
      Top             =   2640
      Width           =   435
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
      Left            =   1845
      TabIndex        =   5
      Top             =   2640
      Width           =   420
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
      Left            =   1215
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Completed Parts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label21 
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
      TabIndex        =   24
      Top             =   -360
      Width           =   1215
   End
   Begin VB.Label Label18 
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
      Left            =   9000
      TabIndex        =   22
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label25 
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
      Left            =   4680
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label24 
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
      Left            =   5760
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label23 
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
      Left            =   6600
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usrHSTParts(1000) As part          'Exec Listbox Storage
Dim intSearchMode As Integer           'search mode: 0=date,1=JRI,2=Truck#
Dim strJob, strRel As String
Dim intItem As Integer

Private Sub cmdExit_Click(Index As Integer)
  'AdodcHist.Recordset.Close                     'close recordset
  Unload Me                                     'unload form
  End
End Sub

Private Sub cmdFinish_Click()

'---------------------------------------------------- global variables
'BATCH                        batch (in prog) flag
'OPTIMIZED                    optimized flag

'---------------------------------------------------- local variables
Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String
Dim i, j As Integer           'loop counters
Dim SELECTED As Boolean       'record select flag
Dim strDTS As String          'date/time stamp of batch
Dim usrTemp As part           'temp user part
Dim intAddOK As Integer       'add result

   '------------------------------------------------- check Batch Status
   'Only allow Finish if NO batch in progress
   If BATCH = True Then                               'batch in progress
      Beep
      Exit Sub
   End If

   '------------------------------------------------- identify the selected part
   i = 0
   j = lbxHist.ListCount
   SELECTED = False
   Do Until SELECTED Or i = j                         'loop thru listbox contents
      If lbxHist.SELECTED(i) = True Then              'if row is selected
         strDTS = usrHSTParts(i + 1).DTS              'capture batch DTS
         SELECTED = True
      End If
      i = i + 1
   Loop  'end of listbox loop
   
   If SELECTED Then                                   'if a part was selected
      '---------------------------------------------- are you sure prompt
      '--- set up msg box
      strMSG = "Are you sure you want to Finish the selected batch?"
      intStyle = vbYesNo + vbDefaultButton2
      strTitle = "Finishing batch..."
      intResponse = MsgBox(strMSG, intStyle, strTitle)
      '--- msgbox response
      If intResponse = vbNo Then  '------------------ OP chose NO
         Exit Sub
      Else  '---------------------------------------- OP chose YES
         '--- make connection to local database
         Set conLocal = New Connection
         conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
         '--- create the recordset
         strSQL = "SELECT * FROM tblHist " & _
                  "WHERE DTStamp = '" & strDTS & "' " & _
                  "AND Status = ' S' " & _
                  "AND Quantity <> BldQnty "
         Set adoRS = New ADODB.Recordset                 'init record set
         adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
         
         With adoRS
            If .RecordCount > 0 Then '------------------ supspended parts found
               .MoveFirst
               Do Until .EOF                             'loop thru recordset
                  '--- build temp part
                  usrTemp.FullJobNum = .Fields("FullOrder")
                  usrTemp.Job = .Fields("Order Number")
                  usrTemp.Rel = .Fields("Release")
                  usrTemp.Item = .Fields("Item")
                  usrTemp.Seq = .Fields("Sequence Number")
                  usrTemp.ShipDate = .Fields("Scheduled Ship Date")
                  usrTemp.Qnty = .Fields("Quantity")
                  usrTemp.BldQnty = .Fields("BldQnty")
                  usrTemp.Phase = .Fields("Phase")
                  usrTemp.Leg = .Fields("Leg")
                  usrTemp.Stack = .Fields("Stack")
                  usrTemp.BarType = .Fields("BarType")
                  usrTemp.Material = .Fields("Material")
                  usrTemp.BarWidth = .Fields("BarWidth")
                  usrTemp.BlankLength = .Fields("BlankLength")
                  usrTemp.RunLength = .Fields("RunLength")
                  usrTemp.E1fig = .Fields("E1figure")
                  usrTemp.E1dim = .Fields("E1dimension")
                  usrTemp.E2fig = .Fields("E2figure")
                  usrTemp.E2dim = .Fields("E2dimension")
                  usrTemp.Cdim = .Fields("Cdimension")
                  usrTemp.C1dim = .Fields("C1dimension")
                  usrTemp.Ddim = .Fields("Ddimension")
                  usrTemp.D1dim = .Fields("D1dimension")
                  usrTemp.Build = .Fields("Build")
                  usrTemp.Status = "EQ"                        'set status
                  usrTemp.Priority = .Fields("Priority")
                  If IsNull(.Fields("Truck")) Then
                     usrTemp.Truck = "None"
                  Else
                     usrTemp.Truck = .Fields("Truck")
                  End If
                  usrTemp.Machine = .Fields("Machine")
                  usrTemp.DTS = .Fields("DTStamp")             'set DTS
               
                  addExec usrTemp, intAddOK                    'add to ExecQue
                  
                  .MoveNext
               Loop  '.EOF
               Me.Hide
               Unload Me
               Exit Sub
            Else                                         'NO SUSPENDed parts
               Beep
               MsgBox ("cmdFinish: No unfinished parts in this batch!")
            End If 'supspended parts
         End With 'adoRS
         
         adoRS.Close                                     'unload recordset
         Set adoRS = Nothing
         conLocal.Close                                  'unload connection
         Set conLocal = Nothing
         
      End If   'intResponse
   Else                                                  'no part selected
      'MsgBox ("No part selected!")
      Beep
   End If   'selected
End Sub  'cmdFinish_Click()

Private Sub cmdGOTORun_Click()
   Me.Hide                             'hide the Select Screen
   Unload Me
End Sub

Private Sub cmdRefresh_Click()
   RefreshHistQue
End Sub

'---------------------------------------------------------------------------
' This sub checks for a selected part, ID's the Item of which it is a part,
' then builds a select item for use w/ the partial screen....then displays
' the partial screen.
'---------------------------------------------------------------------------
Private Sub cmdRemake_Click()
   Dim i, j, intDone As Integer
   Dim SELECTED As Boolean             'part selected flag
   Dim conCamdata As Connection        'camdata connection
   Dim adorsCamdata As ADODB.Recordset 'camdata RS
   
   '---------------------- identify the selected part ----------------------
   i = 0
   j = lbxHist.ListCount
   SELECTED = False
   Do Until SELECTED Or i = j                         'loop thru listbox contents
      If lbxHist.SELECTED(i) = True Then              'if row is selected
         usrRecID.Job = usrHSTParts(i + 1).Job        'set record id
         usrRecID.Rel = usrHSTParts(i + 1).Rel
         usrRecID.Item = usrHSTParts(i + 1).Item
         usrRecID.Seq = usrHSTParts(i + 1).Seq
         SELECTED = True
      End If
      i = i + 1
   Loop  'end of listbox loop
   
   If SELECTED Then                                   'if a part was selected
      '------------------- Put Item Back in Partial Que ------------------
      '--- make connection to camdata db
      Set conCamdata = New Connection
      conCamdata.Open "PROVIDER=MSDASQL;dsn=dsnMBCamdata;uid=;pwd=;"
                                 
      '--- make recordset for item
      Set adorsCamdata = New ADODB.Recordset
      If usrRecID.Job = "000" Then
         strSQL = "SELECT [Order Number],Release,Item,[Scheduled Ship Date],Quantity " & _
                  "FROM mubbarff " & _
                  "WHERE ([Order Number] = '" & usrRecID.Job & "'" & _
                  "AND Release Is Null " & _
                  "AND Item = " & usrRecID.Item & ")"
      Else
         strSQL = "SELECT [Order Number],Release,Item,[Scheduled Ship Date],Quantity " & _
                  "FROM mubbarff " & _
                  "WHERE ([Order Number] = '" & usrRecID.Job & "'" & _
                  "AND Release = '" & usrRecID.Rel & "'" & _
                  "AND Item = " & usrRecID.Item & ")"
      End If
      adorsCamdata.Open strSQL, conCamdata, adOpenStatic, adLockOptimistic

      '--- build the select item
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
         usrSelItem.Build = "R"
      Else                                         'if not found
         usrSelItem.Job = usrRecID.Job
         usrSelItem.Rel = usrRecID.Rel
         usrSelItem.Item = usrRecID.Item
         usrSelItem.ShipDate = Date
         usrSelItem.Qnty = 0
         usrSelItem.Build = "R"
      End If
      
      '--- clean up
      adorsCamdata.Close                           'close recordset
      Set adorsCamdata = Nothing                   'unload recordset
      conCamdata.Close                             'close connection
      Set conCamdata = Nothing                     'unload connection
      
      'could add item to select que here (2-10-03)
      
      Unload Me                                    'unload hist screen
      frmPartial.Show                              'display the partial screen
      
   Else                                            'no part selected
      'MsgBox ("No part selected!")
      Beep
   End If   'selected
End Sub

Private Sub cmdSearch_Click()
   Select Case intSearchMode
   
   Case 0                              'date search
      
   
   Case 1                              'JRI search
      If txtJob.Text = "" Or txtRel.Text = "" Or txtItem.Text = "" Then
         MsgBox ("Search: You must enter a valid Job,Rel & Item!")
         Exit Sub
      Else
         strJob = txtJob.Text
         strRel = txtRel.Text
         intItem = Val(txtItem.Text)
      End If
   
   Case 2                              'truck search
   
   End Select
   
   RefreshHistQue
   Beep
End Sub

Private Sub cmdView_Click()
   Dim i As Integer     'listbox index
   
   i = lbxHist.ListIndex                           'get index to selected part
           
   If i > -1 Then                                  'if a part was selected
      usrRecID.Job = usrHSTParts(i + 1).Job        'set record id
      usrRecID.Rel = usrHSTParts(i + 1).Rel
      usrRecID.Item = usrHSTParts(i + 1).Item
      usrRecID.Seq = usrHSTParts(i + 1).Seq
      frmViewHist.Show                             'display form
   Else                                            'no part selected
      'MsgBox ("No part selected for viewing!")
      Beep
   End If
End Sub

Private Sub Form_Load()
   DTPicker1.Value = Date                             'set DT picker
   intSearchMode = 0                                  'set mode
   RefreshHistQue
End Sub

Sub RefreshHistQue()
   Dim conDB As Connection
   Dim adoRS As ADODB.Recordset
   Dim strConn As String
   Dim strTable As String
   Dim strSQL As String
   Dim strDate As String
   Dim strTruck As String
   Dim strTemp As String
   Dim strDisplay As String
   Dim intCDBok As Integer
   Dim i As Integer
   
   lbxHist.Clear                                      'clear the listbox
   
   '--------------------------- Construct Query ------------------------
   strTable = "tblHist"
   
   Select Case intSearchMode
   Case 0                                             'search by Date
      strDate = DTPicker1.Value
                                                      'build the SQL string
      strSQL = "SELECT * FROM " & strTable & " " & _
               "WHERE Mid([DTStamp],1,8) = '" & strDate & "' " & _
               "ORDER BY DTStamp DESC"
   Case 1                                             'search by JRI
      strJob = txtJob.Text
      strRel = txtRel.Text
      intItem = Val(txtItem.Text)
                                                      'build the SQL string
      strSQL = "SELECT * FROM " & strTable & " " & _
               "WHERE " & _
               "[Order Number] = '" & strJob & "' AND " & _
               "Release = '" & strRel & "' AND " & _
               "Item = " & intItem & " " & _
               "ORDER BY DTStamp DESC"
   Case 2                                             'search by truck#
      strTruck = txtTruck.Text
                                                      'build the SQL string
      strSQL = "SELECT * FROM " & strTable & " " & _
               "WHERE Truck = '" & strTruck & "'"
   End Select
   
   '--------------------- Create Recordset for Hist Que --------------------
   '--- make connection to local database
   Set conDB = New Connection
   conDB.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                              
   '--- create the recordset
   Set adoRS = New ADODB.Recordset                'init record set
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic

   If adoRS.RecordCount > 0 Then                  'if a Record is found
      adoRS.MoveFirst
      i = 1                                           'init parts array index
      
      '--- loop thru recordset for Exec table and build a display array for Exec Que Listbox.
      Do Until adoRS.EOF
         '--- build array entry for part
         usrHSTParts(i).FullJobNum = adoRS.Fields("FullOrder")
         usrHSTParts(i).Job = adoRS.Fields("Order Number")
         usrHSTParts(i).Rel = adoRS.Fields("Release")
         usrHSTParts(i).Item = adoRS.Fields("Item")
         usrHSTParts(i).Seq = adoRS.Fields("Sequence Number")
         usrHSTParts(i).ShipDate = adoRS.Fields("Scheduled Ship Date")
         usrHSTParts(i).Qnty = adoRS.Fields("Quantity")
         usrHSTParts(i).BldQnty = adoRS.Fields("BldQnty")
         usrHSTParts(i).Phase = adoRS.Fields("Phase")
         usrHSTParts(i).Leg = adoRS.Fields("Leg")
         usrHSTParts(i).Stack = adoRS.Fields("Stack")
         usrHSTParts(i).BarType = adoRS.Fields("BarType")
         usrHSTParts(i).Material = adoRS.Fields("Material")
         usrHSTParts(i).BarWidth = adoRS.Fields("BarWidth")
         usrHSTParts(i).BlankLength = adoRS.Fields("BlankLength")
         usrHSTParts(i).E1fig = adoRS.Fields("E1figure")
         usrHSTParts(i).E1dim = adoRS.Fields("E1dimension")
         usrHSTParts(i).E2fig = adoRS.Fields("E2figure")
         usrHSTParts(i).E2dim = adoRS.Fields("E2dimension")
         usrHSTParts(i).Cdim = adoRS.Fields("Cdimension")
         usrHSTParts(i).C1dim = adoRS.Fields("C1dimension")
         usrHSTParts(i).Ddim = adoRS.Fields("Ddimension")
         usrHSTParts(i).D1dim = adoRS.Fields("D1dimension")
         usrHSTParts(i).Build = adoRS.Fields("Build")
         usrHSTParts(i).Status = adoRS.Fields("Status")
         usrHSTParts(i).Priority = adoRS.Fields("Priority")
         If IsNull(adoRS.Fields("Truck")) Then
            usrHSTParts(i).Truck = "None"
         Else
            usrHSTParts(i).Truck = adoRS.Fields("Truck")
         End If
         usrHSTParts(i).Machine = adoRS.Fields("Machine")
         usrHSTParts(i).DTS = adoRS.Fields("DTStamp")
                   
         '--- construct string for display
         strDisplay = Format(usrHSTParts(i).Status, "@@") & "  " & _
                     Format(usrHSTParts(i).Build, "@") & "  " & _
                     Format(usrHSTParts(i).Job, "@@@@@") & " " & _
                     Format(usrHSTParts(i).Rel, "@@@") & " " & _
                     Format(usrHSTParts(i).Item, "0000") & " " & _
                     Format(usrHSTParts(i).Seq, "0000") & "  " & _
                     Format(usrHSTParts(i).Qnty, "00000") & " " & _
                     Format(usrHSTParts(i).BldQnty, "00000") & "  " & _
                     Format(usrHSTParts(i).Phase, "@") & "  " & _
                     Format(usrHSTParts(i).Leg, "@") & "  " & _
                     Format(usrHSTParts(i).Stack, "@") & "  " & _
                     Format(usrHSTParts(i).Material, "@") & "  " & _
                     Format(usrHSTParts(i).BarWidth, "00000") & "  " & _
                     Format(usrHSTParts(i).BlankLength, "000000") & ""
         Select Case intSearchMode
         Case 0
            strTemp = Mid$(usrHSTParts(i).DTS, 9)
            strTemp = Format(strTemp, "@@@@@@@@@@@") & "" & _
                      Format(usrHSTParts(i).Truck, "@@@@@@@@")
            strDisplay = strDisplay & strTemp
         Case 1
            strTemp = "  " & Mid$(usrHSTParts(i).DTS, 1, 8)
            strTemp = Format(strTemp, "@@@@@@@@") & "  " & _
                      Format(usrHSTParts(i).Truck, "@@@@@@@@")
            strDisplay = strDisplay & strTemp
         Case 2
            strTemp = Format(usrHSTParts(i).DTS, "@@@@@@@@@@@@@@@@@@@@@")
            strDisplay = strDisplay & strTemp
         End Select
                  
         lbxHist.AddItem (strDisplay)                 'add string to listbox
         
         adoRS.MoveNext                               'incr for next part
         i = i + 1
      
      Loop 'end of recordset loop
            
      lbxHist.Refresh
      frmHist.Refresh
      cmdView.Visible = True                          'show view button
   Else
      cmdView.Visible = False                         'hide view button
   End If

   '------------------------------- Clean UP -------------------------------
   adoRS.Close                                    'close recordset
   Set adoRS = Nothing                            'unload recordset
   conDB.Close                                   'close connection
   Set conDB = Nothing                           'unload connection

   lbxHist.Refresh
   
End Sub  'RefreshHistQue





Private Sub OptDate_Click()
   txtJob.Visible = False                             'hide JRI entry controls
   txtRel.Visible = False
   txtItem.Visible = False
   Label23.Visible = False
   Label24.Visible = False
   Label25.Visible = False
   
   Label27.Visible = False                            'hide Truck entry controls
   txtTruck.Visible = False
   
   DTPicker1.Visible = True                           'show DTpicker
   Label28.Visible = True                             'show truck header
   
   intSearchMode = 0                                  'set search mode = date
   
   DTPicker1.Value = Date                             'set default date to today
   
   RefreshHistQue
End Sub
Private Sub OptJRI_Click()
   DTPicker1.Visible = False                          'hide DTpicker
   
   Label27.Visible = False                            'hide Truck entry controls
   txtTruck.Visible = False
   Label28.Visible = True                             'show truck header
   
   txtJob.Visible = True                              'show JRI entry controls
   txtRel.Visible = True
   txtItem.Visible = True
   Label23.Visible = True
   Label24.Visible = True
   Label25.Visible = True
   
   intSearchMode = 1                                  'set search mode = JRI
   
   RefreshHistQue
   
   txtJob.SetFocus
   
End Sub


Private Sub optTruck_Click()
   txtJob.Visible = False                             'hide JRI entry controls
   txtRel.Visible = False
   txtItem.Visible = False
   Label23.Visible = False
   Label24.Visible = False
   Label25.Visible = False
   
   DTPicker1.Visible = False                          'hide DTpicker
   
   Label28.Visible = False                            'hide truck header
   Label27.Visible = True                             'show Truck entry controls
   txtTruck.Visible = True
   
   intSearchMode = 2                                  'set search mode = JRI
   
   RefreshHistQue
   
   txtTruck.SetFocus
   
End Sub

'---------------------------------------------------------------------------
'This sub checks for input from the Barcode Scanner and parces it.
'The barcode scanner is connected via keyboard wedge.
'Input from the barcode scanner goes directly into
'  txtJob.Text, since this textbox is given focus when the form is opened.
'---------------------------------------------------------------------------
Private Sub txtJob_Change()
   Dim strBarCode As String
   Dim strTemp As String
   
   strBarCode = txtJob.Text
   
   If Len(strBarCode) > 19 Then
   
      strTemp = Mid$(strBarCode, 2, 2)
      
      If strTemp = "SO" Or strTemp = "so" Then     'barcode input parce it
         txtJob.Text = Mid$(strBarCode, 7, 5)      'get Job#
         txtRel.Text = Mid$(strBarCode, 13, 4)     'get Rel#
         txtItem.Text = Mid$(strBarCode, 17, 4)    'get item#
      End If
   End If
End Sub
