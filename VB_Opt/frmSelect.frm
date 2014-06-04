VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSelect 
   Caption         =   "Item Select"
   ClientHeight    =   8400
   ClientLeft      =   7215
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   4665
   Begin VB.CommandButton cmdSubmitRemake 
      Caption         =   "Remake"
      Height          =   375
      Left            =   3360
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSubmitPartial 
      Caption         =   "Partial"
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSubmitFull 
      Caption         =   "Full"
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtMat 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   975
   End
   Begin VB.ListBox lbxSelect 
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
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
   End
   Begin VB.CommandButton cmdClearItems 
      Caption         =   "Clear"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Clear All Items in Que."
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdRemoveItem 
      Caption         =   "Remove"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "Remove Selected Items from Que."
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "Add an Item to the Que."
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   7920
      Width           =   975
   End
   Begin MSAdodcLib.Adodc AdodcSelect 
      Height          =   375
      Left            =   2280
      Top             =   8640
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "SELECT * FROM tblSelQue"
      Caption         =   "AdodcSelect"
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
   Begin VB.Frame ItemSelectFrame 
      Caption         =   "Select"
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   4335
   End
   Begin VB.Frame SubmitItemFrame 
      Caption         =   "Submit for Execution "
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Width           =   4335
   End
   Begin VB.Label Label7 
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
      Left            =   2400
      TabIndex        =   17
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label6 
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
      Left            =   1440
      TabIndex        =   16
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label5 
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
      Left            =   3600
      TabIndex        =   13
      Top             =   1560
      Width           =   375
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
      TabIndex        =   12
      Top             =   1560
      Width           =   495
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
      Left            =   990
      TabIndex        =   11
      Top             =   1560
      Width           =   420
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
      Left            =   1635
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "ShipDate"
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
      Left            =   2295
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblSelectTitle 
      Caption         =   "SELECT  SCREEN"
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
      Left            =   480
      TabIndex        =   5
      Top             =   165
      Width           =   3615
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim usrSelQue(20) As SelectItem        'select que array


Private Sub cmdClose_Click()
   AdodcSelect.Recordset.Close      'close the recordset
   Unload Me                        'unload the Select Screen
   frmRun.Show                      'display Execution form
End Sub


Private Sub cmdSubmitPartial_Click()
   Dim i As Integer
   
   i = lbxSelect.ListIndex                               'get index of selected part
   
   If i > -1 Then                                        'if a part was selected
      usrSelItem = usrSelQue(i + 1)                      'capture selected item
      usrSelItem.Build = "P"                             'submit as Partial
      frmPartial.Show                                    'display form
   Else                                                  'no part selected
      'MsgBox ("No item selected!")
   End If
End Sub

Private Sub cmdSubmitRemake_Click()
   Dim i As Integer
   
   i = lbxSelect.ListIndex                               'get index of selected part
   
   If i > -1 Then                                        'if a part was selected
      usrSelItem = usrSelQue(i + 1)                      'capture selected item
      usrSelItem.Build = "R"                             'submit as Partial
      frmPartial.Show                                    'display form
   Else                                                  'no part selected
      'MsgBox ("No item selected!")
   End If
End Sub

'------------------------------------------------------------------------------
'This sub submits all selected items as a full build
'------------------------------------------------------------------------------
Private Sub cmdSubmitFull_Click()
'---------------------------------------------------- global variables
'subType
'usrSelQue
'usrRecId
'BATCH                                 BATCH(in prog) flag

'---------------------------------------------------- local variables
Dim i, j As Integer
Dim intSubmitOK As Integer
Dim blnRemOK As Boolean
Dim intBldOpt As Integer

   '------------------------------------------------- check Batch Status
   'Only allow Submit if NO batch in progress
   If BATCH = True Then                               'batch in progress
      Beep
      Exit Sub
   End If
      
   intBldOpt = 1                                      'set build option
                                                      'force build w/o grounds
   If blnEnOPT Then
      clrLocalTBL "tblHoldAssign"                     'clear hold assign table
      intAsgnCnt = 0
   End If
   
   j = lbxSelect.ListCount - 1                        'set range of loop
   For i = 0 To j                                     'loop thru listbox contents
      If lbxSelect.SELECTED(i) = True Then            'if row is selected
         
         'if qnty >0 then the item was found in the camdb and its ok to
         'process it as a full build
         If usrSelQue(i + 1).Qnty > 0 Then            'OK to submit as FULL
            
            usrSelQue(i + 1).Build = "F"              'submit as full
            intSubmitOK = 0
            subItemF usrSelQue(i + 1), intBldOpt, intSubmitOK   'submit the item as Full
            
            If intSubmitOK = 1 Then                   'item submited ok
               'MsgBox ("Item submitted!")
                  
               '--- remove the item from the Select que
               usrRecID.Job = usrSelQue(i + 1).Job    'build a record id
               usrRecID.Rel = usrSelQue(i + 1).Rel
               usrRecID.Item = usrSelQue(i + 1).Item
                   
               blnRemOK = False
               rmvSel usrRecID, blnRemOK              'remove it
               
               If blnRemOK = True Then                'item removed ok
                  'MsgBox ("Item Removed!")
               Else                                   'item not removed
                  MsgBox ("Item NOT Removed!")
               End If
               
            Else                                      'item not submitted
               MsgBox ("Item NOT submitted!")
            End If   'submit ok
         Else                                         'NOT ok to submit as FULL
            MsgBox ("This item has was NOT found in the Camdata DB." & Chr$(13) & _
                   "It must be submitted as a Partial or Remake.")
         End If   'Full OK
      End If   'selected
   Next i  'list box loop
        
   If blnEnOPT And intAsgnCnt > 0 Then
      runQuery0 "qryManAssignUpdateADD"               'run the update query
   End If
        
  RefreshSelectQue                                    'call refresh sub
    
End Sub  'cmdSubmitFull_Click()

Private Sub cmdRefresh_Click()

RefreshSelectQue

End Sub

Private Sub Form_Load()

RefreshSelectQue

End Sub

Sub RefreshSelectQue()

Dim strTemp As String
Dim i As Integer

lbxSelect.Clear
AdodcSelect.Refresh
If Not AdodcSelect.Recordset.EOF Then     'If que NOT empty
   
   '--- get material & width
   AdodcSelect.Recordset.MoveFirst
   strSelMat = AdodcSelect.Recordset.Fields("Mat")
   txtMat.Text = strSelMat
   strSelWidth = AdodcSelect.Recordset.Fields("Width")
   txtWidth.Text = getWidth(strSelWidth)
   
   i = 1                                  'reset loop index
   Do Until AdodcSelect.Recordset.EOF     'loop thru recordset for Item Select table
      usrSelQue(i).Job = AdodcSelect.Recordset.Fields("Order Number")
      usrSelQue(i).Rel = AdodcSelect.Recordset.Fields("Release")
      usrSelQue(i).Item = AdodcSelect.Recordset.Fields("Item")
      usrSelQue(i).ShipDate = AdodcSelect.Recordset.Fields("Scheduled Ship Date")
      usrSelQue(i).Qnty = AdodcSelect.Recordset.Fields("Quantity")
      usrSelQue(i).Mat = AdodcSelect.Recordset.Fields("Mat")
      usrSelQue(i).Width = AdodcSelect.Recordset.Fields("Width")
      
      '--- construct string for display
      strTemp = Format(usrSelQue(i).Job, "@@@@@") & "  " & _
                Format(usrSelQue(i).Rel, "@@@") & "  " & _
                Format(usrSelQue(i).Item, "0000") & "   " & _
                Format(usrSelQue(i).ShipDate, "@@@@@@") & "   " & _
                Format(usrSelQue(i).Qnty, "0000")
      lbxSelect.AddItem (strTemp)         'add item to listbox
      
      '--- incr to next record
      AdodcSelect.Recordset.MoveNext
      i = i + 1
   Loop  'recordset loop
   
   intSelCnt = i - 1                      'capture the select count
   
   lbxSelect.Refresh
   frmSelect.Refresh
   
Else                                      'que is empty
   strSelMat = ""
   txtMat.Text = ""
   strSelWidth = ""
   txtWidth.Text = ""
   intSelCnt = 0
End If

lbxSelect.Refresh

'--- hide submit buttons if a batch is in progress
If BATCH = True Then                      'batch IN progress
   cmdSubmitFull.Visible = False          'hide buttons
   cmdSubmitPartial.Visible = False
   cmdSubmitRemake.Visible = False
Else                                      'batch NOT in progress
   cmdSubmitFull.Visible = True           'show buttons
   cmdSubmitPartial.Visible = True
   cmdSubmitRemake.Visible = True
End If

End Sub


Private Sub cmdAddItem_Click(index As Integer)
    frmAddItem.Show                    'display the form
End Sub

Private Sub cmdClearItems_Click(index As Integer)
   Dim i, j  As Integer
   Dim blnRemOK As Boolean
   Dim usrItem As RecordID
    
   j = lbxSelect.ListCount - 1                  'determine range of loop
     
   blnRemOK = False                             'reset flag
   For i = 0 To j        'loop thru listbox contents
      usrItem.Job = usrSelQue(i + 1).Job        'build record id
      usrItem.Rel = usrSelQue(i + 1).Rel
      usrItem.Item = usrSelQue(i + 1).Item
      rmvSel usrItem, blnRemOK                  'remove the item
   Next i  'list box loop
  
  RefreshSelectQue
  
End Sub

Private Sub cmdRemoveItem_Click(index As Integer)
   Dim i, j  As Integer
   Dim blnRemOK As Boolean
   Dim usrItem As RecordID
    
   j = lbxSelect.ListCount - 1                  'determine range of loop
     
   blnRemOK = False                             'reset remove flag
   For i = 0 To j        'loop thru listbox contents
      If lbxSelect.SELECTED(i) = True Then      'if row is selected
         usrItem.Job = usrSelQue(i + 1).Job     'build record id
         usrItem.Rel = usrSelQue(i + 1).Rel
         usrItem.Item = usrSelQue(i + 1).Item
         rmvSel usrItem, blnRemOK               'remove the item
      End If
   Next i  'list box loop
  
  RefreshSelectQue
  
End Sub




