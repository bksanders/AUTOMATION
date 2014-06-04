VERSION 5.00
Begin VB.Form frmAddItem 
   Caption         =   "Add an Item to Select Que"
   ClientHeight    =   2145
   ClientLeft      =   4905
   ClientTop       =   4740
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4365
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to Que"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
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
      Caption         =   "Item Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
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
Attribute VB_Name = "frmAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intFound As Integer

Private Sub cmdAdd_Click()
   
'--------------------- variable declarations --------------------------
Dim conLocal As Connection                      'dbase variables
Dim adoRS As ADODB.Recordset
Dim strSQL As String
Dim intAddRes As Integer
Dim intRScnt As Integer                         'temp recordcount variable
Dim strMSG, strTitle As String                  'message box variables
Dim intStyle, intResponse As Integer
        
   '----------------------- Check if Item was Found ----------------------
   ' if item was not found in the prev search...check w/ the OP to make
   ' sure he/she still wants to add this item.   If so, it must be added as
   ' a remake.  If not, just exit.
   If intFound = 0 Then  '------------------------ item NOT Found by search
      '--- set up msg box
      strMSG = "You are trying to add an UNFOUND item.  Are you sure?"
      intStyle = vbYesNo + vbDefaultButton2
      strTitle = "Adding Unfound Item"
      '--- OP response
      intResponse = MsgBox(strMSG, intStyle, strTitle)
      If intResponse = vbYes Then  '-------------- OP chose Yes.
         'place holder for yes code
      Else  '------------------------------------- OP chose No, reset and exit
         cmdAdd.Visible = False                    'turn off ADD button
         lblMSG.Caption = "Item NOT Added."        'OP feedback
         Exit Sub
      End If
   End If   'NOT found
                  
   '------------------------ check against History Table ----------------------
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
   
   '--- create the recordset
   strSQL = "SELECT * FROM tblHist " & _
            "WHERE [Order Number] = '" & usrSelItem.Job & "' " & _
            "AND Release = '" & usrSelItem.Rel & "' " & _
            "AND Item = " & usrSelItem.Item
   Set adoRS = New ADODB.Recordset                    'init record set
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   intRScnt = adoRS.RecordCount                       'get recordcount
   adoRS.Close                                        'unload recordset
   Set adoRS = Nothing
   conLocal.Close                                     'unload connection
   Set conLocal = Nothing
   
   If intRScnt > 0 Then '---------------------------- part found in History
     '--- set up msg box
      strMSG = "You are trying to add an item " & _
               "which is already in the History table." & Chr$(13) & Chr$(13) & _
               "                                         " & _
               "Are you sure?"
      intStyle = vbYesNo + vbDefaultButton2
      strTitle = "Item in History"
      '--- OP response
      intResponse = MsgBox(strMSG, intStyle, strTitle)
      If intResponse = vbYes Then  '-------------- OP chose Yes.
         'place holder for yes code
      Else  '------------------------------------- OP chose No, reset and exit
         cmdAdd.Visible = False                    'turn off ADD button
         lblMSG.Caption = "Item NOT Added."        'OP feedback
         Exit Sub
      End If
   End If
               
   '-------------------- Add Item to the Select Que ------------------------
   intAddRes = 0
   addSel usrSelItem, intAddRes
    
   Select Case intAddRes
   Case 1
      lblMSG.Caption = "Item Added."               'OP feedback
      txtJob.Text = ""                             'clear search boxes
      txtRel.Text = ""
      txtItem.Text = ""
      txtJob.SetFocus
   Case 2
      lblMSG.Caption = "Item Already in Que."
      txtJob.Text = ""                             'clear search boxes
      txtRel.Text = ""
      txtItem.Text = ""
      txtJob.SetFocus
   Case 3
      lblMSG.Caption = "Material misMatch."
      txtJob.Text = ""                             'clear search boxes
      txtRel.Text = ""
      txtItem.Text = ""
      txtJob.SetFocus
   Case 4
      lblMSG.Caption = "Barwidth misMatch."
      txtJob.Text = ""                             'clear search boxes
      txtRel.Text = ""
      txtItem.Text = ""
      txtJob.SetFocus
   Case 5
      lblMSG.Caption = "Que is Full."
      txtJob.Text = ""                             'clear search boxes
      txtRel.Text = ""
      txtItem.Text = ""
      txtJob.SetFocus
   Case Else
      lblMSG.Caption = "Item NOT Added."           'OP feedback
      txtJob.SetFocus
   End Select
   
   cmdAdd.Visible = False                          'turn off ADD button
   
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Dim conCamdata As Connection
   Dim adorsCamdata As ADODB.Recordset
   Dim strSQL, strMSG As String
   Dim varNull As Variant
   '--- setup for search
   intFound = 0                                 'reset found flag
   usrSelItem = usrNULLItem                     'clear item variable
   
   '--- verify JOB/REL/ITEM are valid
   If txtJob.Text = "" Or txtItem.Text = "" Then
      MsgBox "You must enter a valid Job & Item!"
   
   Else  'job/rel/item entered, ok to process
      '--- make connection to camdata database
      Set conCamdata = New Connection
      conCamdata.Open "PROVIDER=MSDASQL;dsn=dsnMBCamdata;uid=;pwd=;"
                                 
      '--- make recordset and search for item
      lblMSG.Caption = "searching..."        'OP feedback
      lblMSG.Visible = True
   
      Set adorsCamdata = New ADODB.Recordset
      If txtRel.Text = "" Then
         strSQL = "SELECT [Order Number],Release,Item," & _
                         "[Scheduled Ship Date],Quantity, " & _
                         "Material,Barwidth " & _
                  "FROM mubbarff " & _
                  "WHERE ([Order Number] = '" & txtJob.Text & "' " & _
                  "AND Release Is Null " & _
                  "AND Item = " & txtItem.Text & ")"
         txtRel.Text = "000"
      Else
         strSQL = "SELECT [Order Number],Release,Item," & _
                         "[Scheduled Ship Date],Quantity, " & _
                         "Material,Barwidth " & _
                  "FROM mubbarff " & _
                  "WHERE ([Order Number] = '" & txtJob.Text & "'" & _
                  "AND Release = '" & txtRel.Text & "'" & _
                  "AND Item = " & txtItem.Text & ")"
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
         
         lblMSG.Caption = "Item Found"             'OP feedback
         intFound = 1
      Else                                          'if not found
         usrSelItem.Job = txtJob.Text
         usrSelItem.Rel = txtRel.Text
         usrSelItem.Item = txtItem.Text
         usrSelItem.ShipDate = 100000
         usrSelItem.Qnty = 0
         usrSelItem.Mat = " "
         usrSelItem.Width = " "
         
         lblMSG.Caption = "Item NOT Found!"
         intFound = 0
      End If
                 
      cmdAdd.Visible = True                     'make add button visible
      cmdAdd.SetFocus                           'set focus on the ADD pb
      
      adorsCamdata.Close                        'close recordset
      Set adorsCamdata = Nothing                'unload recordset
      conCamdata.Close                          'close connection
      Set conCamdata = Nothing                  'unload connection
      
      If intFound = 1 Then
         cmdAdd_Click
      End If
    End If  'valid JRI
    
End Sub

Private Sub txtItem_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then                'if enter hit run the search
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
         txtJob.Text = Mid$(strBarCode, 7, 5)      'get Job#
         txtRel.Text = Mid$(strBarCode, 13, 4)     'get Rel#
         txtItem.Text = Mid$(strBarCode, 17, 4)    'get item#
      End If
   End If
   
   txtJob.SetFocus                                 'set focus to job field
   
End Sub

