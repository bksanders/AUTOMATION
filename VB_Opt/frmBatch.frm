VERSION 5.00
Begin VB.Form frmBatch 
   Caption         =   "Batch Screen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11850
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Item"
      Height          =   495
      Left            =   9840
      TabIndex        =   47
      ToolTipText     =   "Refresh Screen"
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdParams 
      Caption         =   "     Opt Paramaters"
      Height          =   495
      Left            =   9840
      TabIndex        =   46
      ToolTipText     =   "Display Setup Screen."
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   495
      Left            =   4560
      TabIndex        =   45
      ToolTipText     =   "Submit Batch for Execution"
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   495
      Left            =   3480
      TabIndex        =   44
      ToolTipText     =   "Accept Batch"
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill Search"
      Height          =   495
      Left            =   2400
      TabIndex        =   43
      ToolTipText     =   "Search to fill off-fall"
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "  Change           Qty"
      Height          =   495
      Left            =   9840
      TabIndex        =   41
      ToolTipText     =   "Change Qty of select part"
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "  Remove         Item"
      Height          =   495
      Left            =   9840
      TabIndex        =   40
      ToolTipText     =   "Remove Selected Item"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Item"
      Height          =   495
      Left            =   9840
      TabIndex        =   39
      ToolTipText     =   "Add Selected Item"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   5640
      TabIndex        =   37
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtMaterial 
      Height          =   285
      Left            =   2640
      TabIndex        =   35
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtTotBlnks 
      Height          =   285
      Left            =   8640
      TabIndex        =   30
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdInitBatch 
      Caption         =   "    Make        Init Batch"
      Height          =   495
      Left            =   1320
      TabIndex        =   29
      ToolTipText     =   "Use selected seq to generate init batch."
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox txtLongOff 
      Height          =   285
      Left            =   5640
      TabIndex        =   28
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtOPT 
      Height          =   285
      Left            =   8640
      TabIndex        =   27
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtTotOff 
      Height          =   285
      Left            =   2640
      TabIndex        =   26
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Batch"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "Clear Batch"
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   9840
      TabIndex        =   13
      ToolTipText     =   "Refresh Screen"
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdGOTORun 
      Caption         =   "  Run  Screen"
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      ToolTipText     =   "Go to Run Screen"
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdView 
      Caption         =   " View Off-Fall"
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      ToolTipText     =   "View off-fall"
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox lbxMach 
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
      Height          =   5685
      Left            =   225
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   2160
      Width           =   9420
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status Text"
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
      Left            =   5640
      TabIndex        =   42
      Top             =   7920
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label27 
      Caption         =   "Width:"
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
      Left            =   4680
      TabIndex        =   38
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label26 
      Caption         =   "Material:"
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
      Left            =   1680
      TabIndex        =   36
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "Sel'd"
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
      Left            =   120
      TabIndex        =   34
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label24 
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
      Left            =   7200
      TabIndex        =   33
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label23 
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
      Left            =   7320
      TabIndex        =   32
      Top             =   1920
      Width           =   390
   End
   Begin VB.Label Label17 
      Caption         =   "Total Blanks:"
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
      Left            =   7200
      TabIndex        =   31
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "% Optimization:"
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
      Left            =   6960
      TabIndex        =   25
      Top             =   1305
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "Longest Off-Fall: Off-Fall:"
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
      Left            =   3840
      TabIndex        =   24
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label Label6 
      Caption         =   "Total Off-Fall:"
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
      Left            =   1200
      TabIndex        =   23
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "BUILD BATCH"
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
      Left            =   4440
      TabIndex        =   21
      Top             =   240
      Width           =   2805
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
      TabIndex        =   19
      Top             =   240
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   1080
      X2              =   10800
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
      TabIndex        =   17
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "St"
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
      Left            =   3975
      TabIndex        =   16
      Top             =   1920
      Width           =   375
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
      Left            =   5550
      TabIndex        =   15
      Top             =   1680
      Width           =   735
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
      Left            =   8280
      TabIndex        =   14
      Top             =   1920
      Width           =   585
   End
   Begin VB.Label Label19 
      Caption         =   "Ph"
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
      Left            =   3210
      TabIndex        =   12
      Top             =   1920
      Width           =   375
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
      Left            =   5565
      TabIndex        =   11
      Top             =   1920
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
      Left            =   4740
      TabIndex        =   10
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label Label14 
      Caption         =   "Mt"
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
      Left            =   4350
      TabIndex        =   9
      Top             =   1920
      Width           =   495
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
      Left            =   6600
      TabIndex        =   8
      Top             =   1920
      Width           =   390
   End
   Begin VB.Label Label11 
      Caption         =   "Lg"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   1920
      Width           =   465
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
      Left            =   2505
      TabIndex        =   6
      Top             =   1905
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
      Left            =   1950
      TabIndex        =   5
      Top             =   1920
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
      Left            =   1455
      TabIndex        =   4
      Top             =   1920
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
      Left            =   825
      TabIndex        =   3
      Top             =   1920
      Width           =   495
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
      TabIndex        =   20
      Top             =   -360
      Width           =   1215
   End
   Begin VB.Label Label18 
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
      Left            =   8280
      TabIndex        =   18
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flgREFRESH As Boolean                 'refresh flag, true when bld/ref mach list
Dim flgITEM As Boolean                    'item change flag, true when adding/removing to batch

Private Sub cmdExit_Click(Index As Integer)
  Unload Me                                           'unload form
  End
End Sub

Private Sub cmdAccept_Click()
   
   '--- global variables
   'intMachCnt                # of seq's in usrMachList()
   'usrMachList()             machine list array
   
   '--- local variables
   Dim i As Integer           'loop index
   Dim intAddOK As Integer    'add result
   Dim intPri As Integer      'priority counter
   
   frmBatch.lblStatus.Caption = "Accepting Batch..."
   frmBatch.lblStatus.Visible = True

   '--------------------------- build the batch que ---------------------------
   clrLocalTBL "tblBatchQue"                          'clear batch que
   intPri = 1
   For i = 1 To intMachCnt
      If usrMACHlist(i).BldQnty > 0 Then
               
         usrMACHlist(i).Priority = intPri             'set priority(order in que)
         intPri = intPri + 1
         addBatchQue i, intAddOK                      'add to batch que
         
         If intAddOK < 1 Then
            MsgBox ("Error adding to tblBatchQue! Batch NOT accepted!")
            Exit Sub
         End If
      End If
   Next i
   
   '------------------------ update the machine table -------------------------
   runQuery0 "qryAssignUpdateADD"
   
   '----------------------------- lock the batch ------------------------------
   usrBatch.Lock = True                               'lock the batch
   
   '----------------------------- Refresh display -----------------------------
   frmBatch.lblStatus.Visible = False
   bldMACHlist
   displaySTATS
   displayButtons
   
End Sub  'cmdAccept_Click()

Private Sub cmdChange_Click()
   Dim i As Integer     'listbox index
   
   i = lbxMach.ListIndex                           'get index to selected part
   intMachIDX = i + 1                              'calc array index
   If i > -1 Then                                  'if a part was selected
      If usrBatch.InitPick = True Then
         If usrMACHlist(intMachIDX).Material <> usrBatch.Mat Then
            MsgBox ("Can't add: material mismatch!")
            Exit Sub
         End If
         If usrMACHlist(intMachIDX).BarWidth <> usrBatch.Width Then
            MsgBox ("Can't add: width mismatch!")
            Exit Sub
         End If
      End If
      frmChangeQty.Show
   Else                                            'no part selected
      'MsgBox ("No part selected!")
      Beep
   End If
End Sub

'---------------------------------------------------------------------------
'This sub searches for matching parts to fill in the batch off-fall.
'This is the old version...that filled off-fall w/ partion stacks.
'This version is not in use. 6/24/05.
'---------------------------------------------------------------------------
Private Sub cmdFillOLD_Click()
Dim ofIDX As Integer       'off-fall array index
Dim rID As Integer         'release array loop index
Dim unMatched As Integer   'unmatched quantity
Dim relSTART As Integer    'starting point for release loop
Dim LoLength As Long       'lower length range
Dim idxMatch As Integer    'index of current match
Dim qtyMatch As Integer    'quantity of current match
Dim i As Integer           'loop index
Dim intGroup As Integer    'temporary integer
Dim addOK As Boolean

   frmBatch.lblStatus.Caption = "Filling Batch..."
   frmBatch.lblStatus.Visible = True

   usrBatch.Lock = True                               'lock the batch
   
   If usrBatch.InitPick = True Then                   'No need to check init rel
      relSTART = 2                                    'if everything in init rel
   Else                                               'is autoselected.
      relSTART = 1
   End If
   
   'intMinOFidx = calcMinOFidx(24000)
   For ofIDX = 18 To intMinOFidx Step -1              'loop thru off-fall array
      
      DoEvents                                        'check events
      
      LoLength = 15 + (ofIDX - 2) * 5                 'calc lower length range
      LoLength = LoLength * 1000
         
      If arrOffFall(ofIDX) > 0 Then
         unMatched = arrOffFall(ofIDX)                'get quantity to be matched
         
         '----------------------------- Search Loop ------------------------------
         Do Until unMatched = 0
            '----------------------------------------- search w/in release list
            idxMatch = 0                              'reset the match index
            For rID = relSTART To intRelCnt           'loop thru release list
               'search for = qty & = length
               idxMatch = fillsearch(rID, "=", unMatched, LoLength)
               If idxMatch > 0 Then
                  qtyMatch = unMatched
                  Exit For
               End If
               'search for < qty < & = length
               idxMatch = fillsearch(rID, "<", unMatched, LoLength)
               If idxMatch > 0 Then
                  qtyMatch = usrMACHlist(idxMatch).Qnty
                  Exit For
               End If
               'search for > qty & = length
               idxMatch = fillsearch(rID, ">", unMatched, LoLength)
               If idxMatch > 0 Then
                  qtyMatch = unMatched
                  Exit For
               End If
            Next rID   'release loop
            
            '----------------------------------------- search outside release list
            If idxMatch = 0 And intRelCnt < intMaxRel Then
               'search for = qty & = length
               idxMatch = fillsearch(0, "=", unMatched, LoLength)
               qtyMatch = unMatched
               
               'search for < qty < & = length
               If idxMatch = 0 Then
                  idxMatch = fillsearch(0, "<", unMatched, LoLength)
                  qtyMatch = usrMACHlist(idxMatch).Qnty
               End If
               
               'search for > qty & = length
               If idxMatch = 0 Then
                  idxMatch = fillsearch(0, ">", unMatched, LoLength)
                  qtyMatch = unMatched
               End If
            End If
            
            '----------------------------------------- handle match/no match
            If idxMatch > 0 Then    '----------------- match found
               
               '--- add to release to list
               rID = addRelease(usrMACHlist(idxMatch).Job, usrMACHlist(idxMatch).Rel)
               If rID = 0 Then
                  MsgBox ("cmdFill:Error while adding to release list!")
                  Exit Sub
               End If
               
               '--- select as part of batch
               usrMACHlist(idxMatch).BldQnty = qtyMatch   'set build quantity
               
               '--- add to batch table
               For i = 1 To qtyMatch
                  intGroup = locOffFall(LoLength)     'locate off-fall group
                  If intGroup > 0 Then
                     usrMACHlist(idxMatch).Priority = intGroup
                        
                     addBatch2 usrMACHlist(idxMatch), "tblBatch2", addOK
                     If addOK = False Then
                        MsgBox ("cmdFill: An error occured when Adding a record " & _
                        Chr$(13) & "to the Batch Table. Batch fill aborted!")
                        Exit Sub
                     Else
                        'update batch stats
                        updateSTATS usrMACHlist(idxMatch).BlankLength
                     End If
                  Else
                     MsgBox ("cmdFill: Could NOT locate off-fall!" & _
                        Chr$(13) & "Batch fill aborted!")
                     Exit Sub
                  End If
               Next i
               
               '--- add to batch que here? or could do in seperate sub
                  '--- ID & mark build...partial/full
                  'If usrMACHlist(idxMatch).BldQnty = usrMACHlist(idxMatch).Qnty Then
                  '   usrMACHlist(idxMatch).Build = "F"
                  'Else
                  '    usrMACHlist(idxMatch).Build = "P"
                  'End If
                  
               unMatched = unMatched - qtyMatch       'update unmatched quantity
            Else                    '----------------- no match found
               'since no match exist...add unmatched qty to next smaller off-fall group
               arrOffFall(ofIDX - 1) = arrOffFall(ofIDX - 1) + unMatched
               unMatched = 0                          'reset unmatch to exit loop
            End If   'match
         Loop  'matched loop
      End If   'off-fall qty > 0
   Next ofIDX  'off-fall loop
   
   frmBatch.lblStatus.Visible = False
   refMACHlist
   displaySTATS
   displayButtons
   
End Sub     'cmdFillOLD_Click()

'---------------------------------------------------------------------------
'This sub searches for matching parts to fill in the batch off-fall.
'Version 2:  forces filling w/ full stacks

'A search pass loops thru the machine list..from its beginning until the fill
'date limit is reached.   Each item in the machine list is then checked to see
'if that item(or at least 1 stack of the item) can fit into the current off-
'fall.   If so the match paramaters are returned.  The match paramaters are
'compared to the prev match, and the better match is retained.  Thus, the
'search pass will select the best fitting item.

'1st the routine performs a search pass thru each release, in the release list.
'In this way, priority is given to those releases already in the batch.  If No
'match is found w/in the release already in the batch... the search is widened to
'search for items in new releases.  A paramater limits the # of releases that
'may be added to a batch.

'Once a match is found.  The matching stacks are added to the batch. The off-
'fall table is updated, each seq in the item(along w/ the appropiate build qty)
'is marked as part of the batch.  And the release is added to the release list.

'The routine will continue to search for a match, until it completes a search
'pass w/out locating a match.
'---------------------------------------------------------------------------
Private Sub cmdFill_Click()

'--- global variables
'usrBatch      'batch paramater variable

'--- local variables
Dim rID As Integer            'release array loop index
Dim relSTART As Integer       'starting point for release loop
Dim idx As Integer            'index of current match
Dim i As Integer              'loop index
Dim intGroup As Integer       'temporary integer
Dim addOK As Boolean
Dim fillDate As Date          'date range for search
Dim NoMatchFound As Boolean   'no match found flag
Dim strJob As String          'Job #
Dim strRel As String          'Rel #
Dim intItem As Integer        'Item #
Dim strJRI As String          'current item's JRI
Dim prevJRI As String         'JRI of next seq
Dim usrMatch As MatchParams   'match paramaters
Dim prevMatch As MatchParams  'paramaters of prev match
Dim nullMatch As MatchParams  'null match variable
Dim intResult As Integer      'result integer

   frmBatch.lblStatus.Caption = "Filling Batch..."
   frmBatch.lblStatus.Visible = True
   
   If usrBatch.InitPick = True Then                   'No need to check init rel
      relSTART = 2                                    'if everything in init rel
   Else                                               'is autoselected.
      relSTART = 1
   End If
   
   fillDate = usrBatch.SDate + intFillDays                'calc the fill date
   
   Do Until NoMatchFound
      DoEvents                                        'check events
                 
      prevMatch = nullMatch                           'reset prev match
                 
      '===================== search w/in release list =========================
      'this code section performs a search pass thru each rel in the batch
      For rID = relSTART To intRelCnt                 'loop thru release list
         strJob = arrRelease(rID).Job                 'get release info
         strRel = arrRelease(rID).Rel
         i = 1                                        'start at beginning of mach list
         intItem = 0
         Do Until usrMACHlist(i).SchedDate > fillDate
            If usrMACHlist(i).Job = strJob And _
               usrMACHlist(i).Rel = strRel And _
               usrMACHlist(i).Item <> intItem And _
               usrMACHlist(i).BldQnty <> usrMACHlist(i).Qnty Then
               
               usrMatch = nullMatch                   'reset user match
               usrMatch.Index = i                     'set match index
            
               chkItemFit usrMatch                    'check the item
               
               If usrMatch.Stacks > 0 Then            'match found
                  'compare this match w/ prev match
                  intResult = compareMatch(usrMatch, prevMatch)
                  If intResult = 1 Then               'if this match is better
                     prevMatch = usrMatch             'keep it
                  End If
               End If
            End If
            intItem = usrMACHlist(i).Item             'capture this item#
            i = i + 1                                 'incr loop index
         Loop  'machine list loop
      Next rID   'release loop
                     
                     
      '======================== search outside release list ========================
      If usrMatch.Stacks = 0 And intRelCnt < intMaxRel Then
         i = 1                                        'start at beginning of mach list
         prevJRI = ""
         Do Until usrMACHlist(i).SchedDate > fillDate
            strJRI = usrMACHlist(i).Job & usrMACHlist(i).Rel & usrMACHlist(i).Item
            If strJRI <> prevJRI And _
               usrMACHlist(i).BldQnty <> usrMACHlist(i).Qnty Then
               
               usrMatch = nullMatch                   'reset user match
               usrMatch.Index = i                     'set match index
            
               chkItemFit usrMatch                    'check the item
               
               If usrMatch.Stacks > 0 Then            'match found
                  'compare this match w/ prev match
                  intResult = compareMatch(usrMatch, prevMatch)
                  If intResult = 1 Then               'if this match is better
                     prevMatch = usrMatch             'keep it
                  End If
               End If
            End If
            prevJRI = strJRI                          'capture seq's JRI
            i = i + 1                                 'incr loop index
         Loop  'machine list loop
      End If   'new release search
                
            
      '========================== handle match/no match ===========================
      'usrMatch = prevMatch                                           'recover best match
      usrMatch.Index = prevMatch.Index                               'broken out for
      usrMatch.Complete = prevMatch.Complete                         'troubleshooting
      usrMatch.Stacks = prevMatch.Stacks                             '& testing
      usrMatch.SeqCnt = prevMatch.SeqCnt
      usrMatch.Length = prevMatch.Length
      usrMatch.MaxSeqLen = prevMatch.MaxSeqLen
      
      If usrMatch.Stacks > 0 Then    '------------------------------- match found
         idx = usrMatch.Index                                        'get index of match
         
         '------------------------- add to batch table  ---------------------------
         For i = 1 To usrMatch.Stacks
            For j = idx To (idx + usrMatch.SeqCnt - 1)
               intGroup = fillOffFall(usrMACHlist(j).BlankLength)    'locate off-fall group
               If intGroup > 0 Then
                  usrMACHlist(j).Priority = intGroup
                  addBatch2 usrMACHlist(j), "tblBatch2", addOK
                  If addOK = False Then
                     MsgBox ("cmdFill: An error occured when Adding a record " & _
                     Chr$(13) & "to the Batch Table. Batch fill aborted!")
                     Exit Sub
                  Else
                     updateSTATS usrMACHlist(j).BlankLength          'update batch stats
                  End If
               Else
                  MsgBox ("cmdFill: Could NOT locate off-fall!" & _
                     Chr$(13) & "Batch fill aborted!")
                  Exit Sub
               End If
            Next j
         Next i
         
         '--------------------- mark seq's as part of batch -----------------------
         For i = idx To (idx + usrMatch.SeqCnt - 1)
            usrMACHlist(i).BldQnty = usrMatch.Stacks                 'set build quantity
         Next i
         
         '----------------------- add to release to list --------------------------
         rID = addRelease(usrMACHlist(idx).Job, usrMACHlist(idx).Rel)
         If rID = 0 Then
            MsgBox ("cmdFill:Error while adding to release list!")
            Exit Sub
         End If
                 
      Else                    '------------------------------------- no match found
         NoMatchFound = True
      End If   'match
      
   Loop 'match loop
   
   usrBatch.Fill = True                                              'set fill flag
   frmBatch.lblStatus.Visible = False
   refMACHlist
   displaySTATS
   displayButtons
   
End Sub     'cmdFill2_Click()

Private Sub cmdFind_Click()
   frmFindItem.Show
End Sub

Private Sub cmdParams_Click()
   frmOptParams.Show
End Sub

Private Sub cmdInitBatch_Click()
   Dim i As Integer     'listbox index
   
   i = lbxMach.ListIndex                           'get index to selected part
           
   If i > -1 Then                                  'if a part was selected
      usrBatch.Job = usrMACHlist(i + 1).Job        'set initial batch params
      usrBatch.Rel = usrMACHlist(i + 1).Rel
      usrBatch.SDate = usrMACHlist(i + 1).SchedDate
      usrBatch.Mat = usrMACHlist(i + 1).Material
      usrBatch.Width = usrMACHlist(i + 1).BarWidth
      usrBatch.InitBatch = True
      usrBatch.InitPick = True
      
      intRelCnt = 1                                'log the initial release
      arrRelease(intRelCnt).Job = usrBatch.Job
      arrRelease(intRelCnt).Rel = usrBatch.Rel
      
      bldMACHlist                                  'build mach list
      displayButtons
      autoOPT                                      'auto optimize
   Else                                            'no part selected
      'MsgBox ("No part selected for Initial Pick!")
      Beep
   End If
End Sub

Private Sub cmdGOTORun_Click()
   Me.Hide
   Unload Me
End Sub

Private Sub cmdRefresh_Click()
   refMACHlist
   displaySTATS
   displayButtons
End Sub

'------------------------------------------------------------------------------
'this sub sumits the pre-built batch for execution
'------------------------------------------------------------------------------
Private Sub cmdSubmit_Click()

'--- global variables
'intExecCnt          '# of seq's in Exec Que
'BATCH               'batch in progress flag

'--- local variables
   
   '--- check for batch in progress
   If BATCH Then
      MsgBox ("A batch is in progress. Submit Aborted!")
      Exit Sub
   End If
   
   '--- check for unempty exec que
   If intExecCnt > 0 Then
      MsgBox ("The Exec Que is NOT empty. Submit Aborted!")
      Exit Sub
   End If
   
   '--- copy tblBatchQue to tblExecQue
   runQuery0 "qryCopyBatchQue"                     'copy batch que
   runQuery0 "qrySetExecStatus"                    'set status = "EQ"
   
   '--- copy tblBatch2 to tblBatch
   runQuery0 "qryCopyBatchTable"
   
   '--- set up batch
   OPTIMIZED = True                                'set optimize flag
   intTotBlanks = usrBatch.Blanks                  'capture total blanks in batch
   intRemBlanks = intTotBlanks                     'set remaining blanks
   
   dblBatchOpt = dblBatchOpt2       'jng 03/31/08-added to keep optimization for history tracking
   dblTotPartsLen = dblTotPartsLen2      'jng 03/31/08-added to keep total parts length for history tracking
   dblBatchLngth = dblBatchLngth2      'jng 03/31/08-added to keep total Stock length for history tracking
   intBatchBlanks = intBatchBlanks2     'jng 3/31/08
   dblLongScrap = dblLongScrap2                    'jng
   intOptType = intOptType2                     'jng
   
   '--- reset build batch
   clrLocalTBL "tblBatchQue"
   clrLocalTBL "tblBatch2"
   usrBatch = usrNullBatch                         'reset batch params
   clrOffFall                                      'clear off-fall array
   clrRelArr                                       'clear the rel array
      
   '--- refresh the exec que
   frmRun.RefreshExecQue
   
   '--- close batch build screen
   Me.Hide
   Unload Me
   
End Sub

Private Sub cmdView_Click()
   frmViewOFF.Show
End Sub

Private Sub cmdClear_Click()
   
   'if batch has been locked...then the machine table was updated...it must
   're-updated to reflect the clearing of the batch.
   If usrBatch.Lock = True Then
      runQuery0 "qryAssignUpdateRMV"
      clrLocalTBL "tblBatchQue"
      clrLocalTBL "tblBatch2"
   End If
   
   usrBatch = usrNullBatch                'reset user batch params
   
   clrOffFall                             'clear off-fall array
   clrRelArr                              'clear the rel array
   
   bldMACHlist
   displaySTATS
   displayButtons
End Sub



Private Sub Form_Load()
   
   If usrBatch.Fill = False And usrBatch.Lock = False Then   'don't mess w/ batch if filled or locked
      clrRelArr
      clrOffFall
      If blnAutoPICK = True And usrBatch.InitPick = False Then
         autoPick                            'automatically make the init pick
      End If
      bldMACHlist                            'build mach list
      displayButtons
      If usrBatch.InitPick = True Then
         autoOPT                             'auto optimize
      End If
   Else
      refMACHlist
      displaySTATS
      displayButtons
   End If
   
End Sub

Sub bldMACHlist()
   Dim conDB As Connection
   Dim adoRS As ADODB.Recordset
   Dim strSQL As String
   Dim strConn As String
   Dim strTemp As String
   Dim strPrevJRI As String
   Dim i As Integer
   
   flgREFRESH = True                                        'set refresh flag
   lblStatus.Caption = "Building Machine List..."
   lblStatus.Visible = True
   DoEvents
   
   lbxMach.Clear                                            'clear the listbox
      
   '------------------------ Construct Search Criteria ---------------------
   If usrBatch.Lock = True Then
      strConn = "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
      
      strSQL = "SELECT [Order Number], Release, Item, [Sequence Number], " & _
                  "[Scheduled Ship Date], ShipDate, Quantity, BldQnty, Leg, Phase, Stack, " & _
                  "BarType, Material, BarWidth, BlankLength, E1figure, E1dimension, " & _
                  "E2figure, E2dimension, Cdimension, C1dimension, Ddimension, " & _
                  "D1dimension, FullOrder, Build, Status, DTStamp " & _
                  "From tblBatchQue " & _
                  "GROUP BY [Order Number], Release, Item, [Sequence Number], " & _
                  "[Scheduled Ship Date], ShipDate, Quantity, BldQnty, Leg, Phase, Stack, " & _
                  "BarType, Material, BarWidth, BlankLength, E1figure, E1dimension, " & _
                  "E2figure, E2dimension, Cdimension, C1dimension, Ddimension, " & _
                  "D1dimension, FullOrder, Build, Status, DTStamp " & _
                  "ORDER BY ShipDate"
   Else
      strConn = "PROVIDER=MSDASQL;dsn=dsnCentralDB;uid=;pwd=;"
      If usrBatch.InitPick = True Then
         strSQL = "SELECT [Order Number], Release, Item, [Sequence Number], " & _
                  "[Scheduled Ship Date], ShipDate, Quantity, Leg, Phase, Stack, " & _
                  "BarType, Material, BarWidth, BlankLength, E1figure, E1dimension, " & _
                  "E2figure, E2dimension, Cdimension, C1dimension, Ddimension, " & _
                  "D1dimension, FullOrder, Machine, Status, OpenQty " & _
                  "From tblMachine " & _
                  "GROUP BY [Order Number], Release, Item, [Sequence Number], " & _
                  "[Scheduled Ship Date], ShipDate, Quantity, Leg, Phase, Stack, " & _
                  "BarType, Material, BarWidth, BlankLength, E1figure, E1dimension, " & _
                  "E2figure, E2dimension, Cdimension, C1dimension, Ddimension, " & _
                  "D1dimension, FullOrder, Machine, Status, OpenQty " & _
                  "Having (((Machine) = 'M' OR (Machine) = 'E') AND ((OpenQty) > 0) " & _
                  "AND (Material = '" & usrBatch.Mat & "') " & _
                  "AND (BarWidth = '" & usrBatch.Width & "')) " & _
                  "ORDER BY ShipDate"
      Else
         strSQL = "SELECT [Order Number], Release, Item, [Sequence Number], " & _
                  "[Scheduled Ship Date], ShipDate, Quantity, Leg, Phase, Stack, " & _
                  "BarType, Material, BarWidth, BlankLength, E1figure, E1dimension, " & _
                  "E2figure, E2dimension, Cdimension, C1dimension, Ddimension, " & _
                  "D1dimension, FullOrder, Machine, Status, OpenQty " & _
                  "From tblMachine " & _
                  "GROUP BY [Order Number], Release, Item, [Sequence Number], " & _
                  "[Scheduled Ship Date], ShipDate, Quantity, Leg, Phase, Stack, " & _
                  "BarType, Material, BarWidth, BlankLength, E1figure, E1dimension, " & _
                  "E2figure, E2dimension, Cdimension, C1dimension, Ddimension, " & _
                  "D1dimension, FullOrder, Machine, Status, OpenQty " & _
                  "Having (((Machine) = 'M' Or (Machine) = 'E') And ((OpenQty) > 0)) " & _
                  "ORDER BY ShipDate"
      End If
   End If
   '--------------------- Create Recordset for Machine Que --------------------
   '--- make connection to central database"
   Set conDB = New Connection
   conDB.Open strConn
                              
   '--- create the recordset
   Set adoRS = New ADODB.Recordset                          'init record set
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then                              'if a Record is found
         .MoveFirst
         i = 1                                              'init parts array index
         
         '--- loop thru recordset and build a display array for Listbox.
         Do Until .EOF
            
            '--- build array entry for part
            If usrBatch.Lock = True Then
               usrMACHlist(i).FullJobNum = .Fields("FullOrder")
               usrMACHlist(i).Job = .Fields("Order Number")
               usrMACHlist(i).Rel = .Fields("Release")
               usrMACHlist(i).Item = .Fields("Item")
               usrMACHlist(i).Seq = .Fields("Sequence Number")
               usrMACHlist(i).ShipDate = .Fields("Scheduled Ship Date")
               usrMACHlist(i).SchedDate = .Fields("ShipDate")
               usrMACHlist(i).Qnty = .Fields("Quantity")
               usrMACHlist(i).BldQnty = .Fields("BldQnty")
               usrMACHlist(i).Phase = .Fields("Phase")
               usrMACHlist(i).Leg = .Fields("Leg")
               usrMACHlist(i).Stack = .Fields("Stack")
               usrMACHlist(i).BarType = .Fields("BarType")
               usrMACHlist(i).Material = .Fields("Material")
               usrMACHlist(i).BarWidth = .Fields("BarWidth")
               usrMACHlist(i).BlankLength = .Fields("BlankLength")
               usrMACHlist(i).E1fig = .Fields("E1figure")
               usrMACHlist(i).E1dim = .Fields("E1dimension")
               usrMACHlist(i).E2fig = .Fields("E2figure")
               usrMACHlist(i).E2dim = .Fields("E2dimension")
               usrMACHlist(i).Cdim = .Fields("Cdimension")
               usrMACHlist(i).C1dim = .Fields("C1dimension")
               usrMACHlist(i).Ddim = .Fields("Ddimension")
               usrMACHlist(i).D1dim = .Fields("D1dimension")
               usrMACHlist(i).Build = .Fields("Build")
               usrMACHlist(i).Status = .Fields("Status")
               'usrMACHlist(i).Priority = .Fields("Priority")
               'usrMACHlist(i).Priority = 0
               usrMACHlist(i).DTS = .Fields("DTStamp")
            Else
               usrMACHlist(i).FullJobNum = .Fields("FullOrder")
               usrMACHlist(i).Job = .Fields("Order Number")
               usrMACHlist(i).Rel = .Fields("Release")
               usrMACHlist(i).Item = .Fields("Item")
               usrMACHlist(i).Seq = .Fields("Sequence Number")
               usrMACHlist(i).ShipDate = .Fields("Scheduled Ship Date")
               usrMACHlist(i).SchedDate = .Fields("ShipDate")
               usrMACHlist(i).Qnty = .Fields("OpenQty")
               usrMACHlist(i).BldQnty = 0
               usrMACHlist(i).Phase = .Fields("Phase")
               usrMACHlist(i).Leg = .Fields("Leg")
               usrMACHlist(i).Stack = .Fields("Stack")
               usrMACHlist(i).BarType = .Fields("BarType")
               usrMACHlist(i).Material = .Fields("Material")
               usrMACHlist(i).BarWidth = .Fields("BarWidth")
               usrMACHlist(i).BlankLength = .Fields("BlankLength")
               usrMACHlist(i).E1fig = .Fields("E1figure")
               usrMACHlist(i).E1dim = .Fields("E1dimension")
               usrMACHlist(i).E2fig = .Fields("E2figure")
               usrMACHlist(i).E2dim = .Fields("E2dimension")
               usrMACHlist(i).Cdim = .Fields("Cdimension")
               usrMACHlist(i).C1dim = .Fields("C1dimension")
               usrMACHlist(i).Ddim = .Fields("Ddimension")
               usrMACHlist(i).D1dim = .Fields("D1dimension")
               'usrMACHlist(i).Build = .Fields("Build")
               'usrMACHlist(i).Status = .Fields("Status")
               'usrMACHlist(i).Priority = .Fields("Priority")
               usrMACHlist(i).Priority = 0
               'usrMACHlist(i).DTS = .Fields("DTStamp")
            End If
            
            '--- if initBatch then make init batch selection
            If usrBatch.InitBatch = True And usrBatch.Lock = False And _
               usrMACHlist(i).Job = usrBatch.Job And _
               usrMACHlist(i).Rel = usrBatch.Rel And _
               usrMACHlist(i).Material = usrBatch.Mat And _
               usrMACHlist(i).BarWidth = usrBatch.Width Then
               usrMACHlist(i).BldQnty = usrMACHlist(i).Qnty       'set bld qty
            End If
             
            '--- construct string for display
            strTemp = Format(usrMACHlist(i).Job, "@@@@@") & " " & _
                      Format(usrMACHlist(i).Rel, "@@@") & " " & _
                      Format(usrMACHlist(i).Item, "0000")
            If strTemp = strPrevJRI Then
               strTemp = "              "        '14 spaces
            Else
               strPrevJRI = strTemp
            End If
            strTemp = "  " & strTemp & " " & _
                      Format(usrMACHlist(i).Seq, "0000") & "  " & _
                      Format(usrMACHlist(i).Phase, "@") & "  " & _
                      Format(usrMACHlist(i).Leg, "@") & "  " & _
                      Format(usrMACHlist(i).Stack, "@") & "  " & _
                      Format(usrMACHlist(i).Material, "@") & "  " & _
                      Format(usrMACHlist(i).BarWidth, "00000") & "  " & _
                      Format(usrMACHlist(i).BlankLength, "000000") & "  " & _
                      Format(usrMACHlist(i).Qnty, "00000") & "  " & _
                      Format(usrMACHlist(i).BldQnty, "00000") & "  " & _
                      Format(usrMACHlist(i).SchedDate, "@@@@@@@@")
            lbxMach.AddItem (strTemp)                       'add string to listbox
            
            '--- mark selected parts
            If usrMACHlist(i).BldQnty > 0 Then              'if bldqty >0
               lbxMach.SELECTED(i - 1) = True               'then select part
            End If
                   
            .MoveNext                                       'incr for next part
            i = i + 1
         
         Loop 'end of recordset loop
         intMachCnt = i - 1
         lbxMach.Refresh
         cmdView.Visible = True                             'show view button
      Else
         cmdView.Visible = False                            'hide view button
      End If
      .Close                                                'close recordset
   End With
   
   '---------------------------- Clean UP Recordset -------------------------
   Set adoRS = Nothing                                      'unload recordset
   conDB.Close                                              'close connection
   Set conDB = Nothing                                      'unload connection
   
   lbxMach.ListIndex = 0                                    'reset the index
   lbxMach.Refresh
   
   lblStatus.Visible = False
   flgREFRESH = False
   
End Sub  'bldMACHlist

Sub refMACHlist()
   Dim conDB As Connection
   Dim adoRS As ADODB.Recordset
   Dim strSQL As String
   Dim strConn As String
   Dim strTemp As String
   Dim strPrevJRI As String
   Dim i As Integer
   
   flgREFRESH = True                                        'set refresh flag
   
   lblStatus.Caption = "Refreshing Machine List..."
   lblStatus.Visible = True
   DoEvents
   
   lbxMach.Clear                                            'clear the listbox
   
   For i = 1 To intMachCnt
                    
      '----------------- construct string for display ------------------------
      
      strTemp = Format(usrMACHlist(i).Job, "@@@@@") & " " & _
                Format(usrMACHlist(i).Rel, "@@@") & " " & _
                Format(usrMACHlist(i).Item, "0000")
      If strTemp = strPrevJRI Then                          'check for matching JRI
         strTemp = "              "                         '14 spaces
      Else
         strPrevJRI = strTemp
      End If
      
      strTemp = "  " & strTemp & " " & _
                Format(usrMACHlist(i).Seq, "0000") & "  " & _
                Format(usrMACHlist(i).Phase, "@") & "  " & _
                Format(usrMACHlist(i).Leg, "@") & "  " & _
                Format(usrMACHlist(i).Stack, "@") & "  " & _
                Format(usrMACHlist(i).Material, "@") & "  " & _
                Format(usrMACHlist(i).BarWidth, "00000") & "  " & _
                Format(usrMACHlist(i).BlankLength, "000000") & "  " & _
                Format(usrMACHlist(i).Qnty, "00000") & "  " & _
                Format(usrMACHlist(i).BldQnty, "00000") & "  " & _
                Format(usrMACHlist(i).SchedDate, "@@@@@@@@")
                
      lbxMach.AddItem (strTemp)                       'add string to listbox
      
      If usrMACHlist(i).BldQnty > 0 Then              'if bldqty >0
         lbxMach.SELECTED(i - 1) = True               'then select part
      End If
      
   Next 'machine list loop
         
   lbxMach.Refresh
   lblStatus.Visible = False
   flgREFRESH = False
   
End Sub  'refMACHlist

Private Sub cmdAdd_Click()
   Dim i As Integer     'listbox index
   Dim j As Integer     'mach array index
   Dim strJRI As String
   Dim strCurJRI As String
   Dim intTemp As Integer
   
   flgITEM = True                                  'set item flag
   i = lbxMach.ListIndex                           'get index to selected part
   j = i + 1                                       'calc the array index
   
   If usrBatch.InitPick = True Then
      If usrMACHlist(j).Material <> usrBatch.Mat Then
         MsgBox ("Can't add: material mismatch!")
         Exit Sub
      End If
      If usrMACHlist(j).BarWidth <> usrBatch.Width Then
         MsgBox ("Can't add: width mismatch!")
         Exit Sub
      End If
   End If
   
   If i > -1 Then                                  'if a part was selected
      strJRI = usrMACHlist(j).Job & usrMACHlist(j).Rel & usrMACHlist(j).Item
      strCurJRI = strJRI
      Do Until strJRI <> strCurJRI
         lbxMach.SELECTED(j - 1) = True            'select the part
         j = j + 1
         strCurJRI = usrMACHlist(j).Job & usrMACHlist(j).Rel & usrMACHlist(j).Item
      Loop
   Else                                            'no part selected
      Beep
   End If
   
   refMACHlist
   displayButtons
   autoOPT                                         're-optimize
   
   lbxMach.ListIndex = i                        'reset the index
   lbxMach.TopIndex = i                         'move to top of listbox
   
   flgITEM = False                                 'reset item flag
   
End Sub     'cmdAdd_Click
Private Sub cmdRemove_Click()
   Dim i As Integer     'listbox index
   Dim j As Integer     'mach array index
   Dim strJRI As String
   Dim strCurJRI As String
   
   flgITEM = True                                  'set item flag
   
   i = lbxMach.ListIndex                           'get index to selected part
   j = i + 1                                       'calc the array index
   
   If i > -1 Then                                  'if a part was selected
      strJRI = usrMACHlist(j).Job & usrMACHlist(j).Rel & usrMACHlist(j).Item
      strCurJRI = strJRI
      Do Until strJRI <> strCurJRI
         lbxMach.SELECTED(j - 1) = False           'deselect the part
         j = j + 1
         strCurJRI = usrMACHlist(j).Job & usrMACHlist(j).Rel & usrMACHlist(j).Item
      Loop
      lbxMach.ListIndex = i                        'reset the index
   Else                                            'no part selected
      'MsgBox ("No part selected!")
      Beep
   End If
   
   refMACHlist                                     'refresh the list
   autoOPT                                         're-optimize
   flgITEM = False                                 'reset item flag
End Sub


'---------------------------------------------------------------------------
'When a part in the list is checked/unchecked...set/reset the build quantity
'---------------------------------------------------------------------------
Private Sub lbxMach_ItemCheck(Item As Integer)
   
Dim intTemp As Integer

   If flgREFRESH = False Then                            'do nothing if refreshing que
      If lbxMach.SELECTED(Item) = True Then              'selecting
         
         If usrBatch.InitPick = True Then                'check against init pick
            If usrMACHlist(Item + 1).Material <> usrBatch.Mat Then
               MsgBox ("Can't add: material mismatch!")
               refMACHlist                                     'refresh the list
               Exit Sub
            End If
            If usrMACHlist(Item + 1).BarWidth <> usrBatch.Width Then
               MsgBox ("Can't add: width mismatch!")
               refMACHlist                                     'refresh the list
               Exit Sub
            End If
         End If
             
         'add release to array
         intTemp = addRelease(usrMACHlist(Item + 1).Job, usrMACHlist(Item + 1).Rel)
         If intTemp > 0 Then
            usrMACHlist(Item + 1).BldQnty = usrMACHlist(Item + 1).Qnty
            If usrBatch.InitPick = False Then            'make init pick
               usrBatch.Mat = usrMACHlist(Item + 1).Material
               usrBatch.Width = usrMACHlist(Item + 1).BarWidth
               usrBatch.InitPick = True
            End If
         Else
            MsgBox ("Seq Not Added. To many Releases already in Batch!")
         End If
      Else
         usrMACHlist(Item + 1).BldQnty = 0               'deselecting
      End If
      
      If flgITEM = False Then                            'if not add/rem an entire item
         refMACHlist                                     'refresh the list
         displayButtons                                  'display cmd buttons
         autoOPT                                         're-optimize
         lbxMach.ListIndex = Item                        'reset the index
         lbxMach.TopIndex = Item                         'move to top of listbox
      End If
   End If   'refresh flag
End Sub
