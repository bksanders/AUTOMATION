VERSION 5.00
Begin VB.Form frmFindItem 
   Caption         =   "Find an item "
   ClientHeight    =   2145
   ClientLeft      =   3615
   ClientTop       =   2835
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   3885
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Enter an Item number for search."
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtRel 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Enter a Release number for search."
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtJob 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Enter a Job number for search."
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblMSG 
      Caption         =   "Item NOT Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
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
      TabIndex        =   5
      Top             =   240
      Width           =   735
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
      TabIndex        =   3
      Top             =   240
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
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmFindItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intFound As Integer

Private Sub cmdClose_Click()
   Unload Me
End Sub

'------------------------------------------------------------------------------
'This sub searches the machine table to locate an item.
'it sets focus to this item in the listbox.
'------------------------------------------------------------------------------
Private Sub cmdSearch_Click()
Dim intSearchType As Integer        'type of search to conduct
Dim i As Integer                    'loop variable
Dim strJob As String                'job #
Dim strRel As String                'rel #
Dim intItem As Integer              'item #
Dim blnFOUND As Boolean             'item found flag

   '------------------------- verify JOB/REL/ITEM are valid -------------------
   If txtJob.Text = "" Then
      MsgBox "You must enter a valid Job!"
      Exit Sub
   End If
   
   If txtJob.Text <> "" And txtItem.Text <> "" And txtRel.Text = "" Then
      MsgBox "You must enter a valid Release!"
      Exit Sub
   End If
   
   '--------------------------- determine search type -------------------------
   '1 = by job# only
   '2 = by job & rel
   '3 = by job, rel & item
   
   intSearchType = 1                'default to search by job only
   
   If txtJob.Text <> "" And txtRel.Text <> "" And txtItem.Text = "" Then
      intSearchType = 2             'search by job & rel
   End If
   
   If txtJob.Text <> "" And txtRel.Text <> "" And txtItem.Text <> "" Then
      intSearchType = 3             'search by job,rel,& item
   End If
   
   strJob = txtJob.Text             'set search variables
   strRel = txtRel.Text
   intItem = Val(txtItem.Text)
   
   '------------------------------ perform search -----------------------------
   
'Public usrMACHlist(7000) As part          'listbox Storage array
'Public intMachCnt As Integer              '# of parts in machine array
'????? Public intMachIDX As Integer              'index for selected part
   
   If intMachCnt > 0 Then                    'no need to search if machine list empty
      blnFOUND = False                       'reset found flag
      For i = 1 To intMachCnt                'loop thru machine list
         Select Case intSearchType
         Case 1
            If usrMACHlist(i).Job = strJob Then
               blnFOUND = True
            End If
         Case 2
            If usrMACHlist(i).Job = strJob And _
               usrMACHlist(i).Rel = strRel Then
               blnFOUND = True
            End If
         Case 3
            If usrMACHlist(i).Job = strJob And _
               usrMACHlist(i).Rel = strRel And _
               usrMACHlist(i).Item = intItem Then
               blnFOUND = True
            End If
         End Select
         
         If blnFOUND = True Then
            lblMSG.Visible = False
            frmBatch.lbxMach.ListIndex = i - 1                    'set the index
            frmBatch.lbxMach.TopIndex = i - 1                     'bring item to top
            Exit For
         Else
            lblMSG.Visible = True
         End If
      Next
   End If   'list NOT empty
   
   If blnFOUND = True Then                                        'if item found
      Me.Hide                                                     'close find form
      Unload Me
   End If
End Sub  'cmdSearch_Click

Private Sub txtItem_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then                'if enter hit, run the search
      cmdSearch_Click
   End If
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
         txtJob.Text = Trim(Mid$(strBarCode, 7, 5))      'get Job#
         txtRel.Text = Trim(Mid$(strBarCode, 13, 4))     'get Rel#
         txtItem.Text = Trim(Mid$(strBarCode, 17, 4))    'get item#
      End If
   End If
   
   txtJob.SetFocus                                 'set focus to job field
   
End Sub

