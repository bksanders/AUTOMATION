VERSION 5.00
Begin VB.Form frmPartial 
   Caption         =   "Partial Item Build"
   ClientHeight    =   5295
   ClientLeft      =   945
   ClientTop       =   1665
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10935
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdManAdd 
      Caption         =   " MANUAL"
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.ListBox lbxPartial 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Label Label3 
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
      Left            =   4845
      TabIndex        =   20
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Des'd"
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
      Left            =   9030
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Def"
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
      Left            =   8325
      TabIndex        =   17
      Top             =   840
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
      Left            =   960
      TabIndex        =   16
      Top             =   1080
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
      Left            =   1680
      TabIndex        =   15
      Top             =   1080
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
      Left            =   2325
      TabIndex        =   14
      Top             =   1080
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
      Left            =   3045
      TabIndex        =   13
      Top             =   1080
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
      Left            =   4335
      TabIndex        =   12
      Top             =   1080
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
      Left            =   8325
      TabIndex        =   11
      Top             =   1080
      Width           =   375
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
      Left            =   9150
      TabIndex        =   10
      Top             =   1080
      Width           =   375
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
      Left            =   5310
      TabIndex        =   9
      Top             =   1080
      Width           =   495
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
      Left            =   5910
      TabIndex        =   8
      Top             =   1080
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
      Left            =   6705
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label17 
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
      TabIndex        =   6
      Top             =   1080
      Width           =   615
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
      Left            =   3885
      TabIndex        =   5
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label lblTitle 
      Caption         =   "Build Partial/Remake Screen"
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
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmPartial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flgPrtlRef As Boolean           'semafor flag true when refreshing partial que

Sub RefreshPartialQue()

Dim strTemp As String
Dim i As Integer

flgPrtlRef = True                                     'set refresh flag

lbxPartial.Clear

If intParts > 0 Then                                  'don't display if nothing in que
           
   '--- loop thru parts array and build a display array for Partial Que Listbox.
   For i = 1 To intParts
      '--- construct string for display
      strTemp = "    " & _
                Format(arrParts(i).Job, "@@@@@") & "  " & _
                Format(arrParts(i).Rel, "@@@") & "  " & _
                Format(arrParts(i).Item, "0000") & "  " & _
                Format(arrParts(i).Seq, "0000") & "    " & _
                Format(arrParts(i).Phase, "@") & "   " & _
                Format(arrParts(i).Leg, "@") & "   " & _
                Format(arrParts(i).Stack, "@") & "   " & _
                Format(arrParts(i).Material, "@") & "   " & _
                Format(arrParts(i).BarWidth, "00000") & "    " & _
                Format(arrParts(i).BlankLength, "000000") & "    " & _
                Format(arrParts(i).Qnty, "00000") & "  " & _
                Format(arrParts(i).BldQnty, "00000")

      lbxPartial.AddItem (strTemp)                    'add string to listbox
      
      If arrParts(i).BldQnty > 0 Then                 'if desired Qty > 0
         lbxPartial.SELECTED(i - 1) = True            'then select part
      End If
      
   Next i   'end of array loop
                
   '--- check for a prev. selected part
   'If intPartSel > -1 Then                           'prev selected part?
      'lbxPartial.Selected(intPartSel) = True         'reselect the part
   'End If
     
   lbxPartial.Refresh
   frmPartial.Refresh
   
End If   'no parts in Partial Que

lbxPartial.Refresh      '????????????? why again ??????????????

flgPrtlRef = False                                    'clear refresh flag

End Sub

Private Sub cmdClose_Click()
   Me.Hide                                               'close/unload the form
   Unload Me
End Sub

Private Sub cmdEdit_Click()
   Dim i As Integer
   
   i = lbxPartial.ListIndex                              'get index of selected part
   
   If i > -1 Then                                        'if a part was selected
      intIndex = i + 1                                   'capture selected item
      frmEditPart.Show                                   'display form
   Else                                                  'no part selected
      MsgBox ("No Part selected!")
   End If

End Sub


Private Sub cmdManAdd_Click()
   frmManEntry.Show
End Sub

Private Sub cmdSubmit_Click()

'---------------------------------------------------- global variables
'BATCH                                 BATCH(in prog) flag

'---------------------------------------------------- local variables
Dim i As Integer
Dim intSubOk As Integer                'submit results
Dim rmvOK As Boolean                   'remove results
Dim usrRmvID As RecordID               'record ID
   
   '------------------------------------------------- check Batch Status
   'Only allow Submit if NO batch in progress
   If BATCH = True Then                               'batch in progress
      Beep
      Exit Sub
   End If

   '------------------------------------------------- prep for submit
   For i = 1 To intParts                              'loop thru parts array
      arrParts(i).Qnty = arrParts(i).BldQnty          'move bldqty to qty
      arrParts(i).BldQnty = 0                         'zero the bldqty
      arrParts(i).Status = "EQ"                       'set the status to EQ
   Next i
   
   '------------------------------------------------- Submit the Item
   If intParts > 0 Then                               'must have valid parts array
      If blnEnOPT Then
         clrLocalTBL "tblHoldAssign"                  'clear hold assign table
         intAsgnCnt = 0
      End If
      intSubOk = 0                                    'reset results
      subItemPR intSubOk                              'submit item in parts array
      If intSubOk = 1 Then                            'item submitted
         'if the item was submitted then we must update the assignemnt in
         'the machine table
         If blnEnOPT And intAsgnCnt > 0 Then
            runQuery0 "qryManAssignUpdateADD"         'run the update query
         End If
         usrRmvID.Job = usrSelItem.Job                'prep Rec ID
         usrRmvID.Rel = usrSelItem.Rel
         usrRmvID.Item = usrSelItem.Item
         rmvSel usrRmvID, rmvOK                       'remove the record
         frmSelect.RefreshSelectQue
      Else                                            'NOT submitted
         MsgBox ("Submit: Item NOT submitted!")
      End If
   End If
   
   '------------------------------------------------- Clean UP
   Me.Hide                                            'close/unload the form
   Unload Me
   frmSelect.Hide
   Unload frmSelect
   
End Sub

Private Sub Form_Load()
   bldPartsArray usrSelItem, 0                        'normal build option
   RefreshPartialQue
End Sub

'---------------------------------------------------------------------------
'When a part in the que is checked/unchecked...set/reset the desired quantity
'  - for a remake the default qty = 1
'  - for a partial the default qty = the qty specified in the camdata db
'---------------------------------------------------------------------------
Private Sub lbxPartial_ItemCheck(Item As Integer)
   If flgPrtlRef = False Then                         'do nothing if refreshing que
      If lbxPartial.SELECTED(Item) = True Then        'selecting
         If arrParts(Item + 1).Build = "R" Then       'remake
            arrParts(Item + 1).BldQnty = 1
         Else                                         'partial
            arrParts(Item + 1).BldQnty = arrParts(Item + 1).Qnty
         End If
      Else
         arrParts(Item + 1).BldQnty = 0               'deselecting
      End If
      RefreshPartialQue                               'refresh the que
   End If
End Sub
