VERSION 5.00
Begin VB.Form frmManEntry 
   Caption         =   "Manual Part Entry"
   ClientHeight    =   8430
   ClientLeft      =   4230
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   3750
   Begin VB.TextBox txtFields 
      DataField       =   "E2dimension"
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   45
      Top             =   7400
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E2dimension"
      Height          =   285
      Index           =   20
      Left            =   2040
      TabIndex        =   44
      Top             =   7050
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E2dimension"
      Height          =   285
      Index           =   19
      Left            =   2040
      TabIndex        =   43
      Top             =   6700
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E2dimension"
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   42
      Top             =   6350
      Width           =   1335
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   7
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E2dimension"
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   40
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Order Number"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   33
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   8
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2100
      Width           =   495
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   6
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   4950
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   5
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   4260
      Width           =   855
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   4
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   3900
      Width           =   495
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   3
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   3540
      Width           =   495
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   2
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3180
      Width           =   495
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   1
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   2820
      Width           =   495
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   0
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2460
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2640
      TabIndex        =   24
      Top             =   7905
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   7905
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Release"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   450
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Item"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   780
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Sequence Number"
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Text            =   "1"
      Top             =   1110
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Scheduled Ship Date"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Quantity"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   2
      Text            =   "1"
      Top             =   1770
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BlankLength"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   1
      Top             =   4620
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E1dimension"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   0
      Top             =   5310
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "D1dimension:"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   39
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ddimension:"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   38
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "C1dimension:"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   37
      Top             =   6735
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cdimension:"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   36
      Top             =   6375
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E2dimension:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   35
      Top             =   6030
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Order Number:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Release:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   465
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   795
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sequence Number:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   1125
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Scheduled Ship Date:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Quantity:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   1785
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Phase:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   2490
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Leg:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2850
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stack:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   3210
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bar Type:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   14
      Top             =   3570
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Material:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   13
      Top             =   3930
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BarWidth:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   4290
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BlankLength:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   11
      Top             =   4635
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E1figure:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   10
      Top             =   4980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E1dimension:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   9
      Top             =   5325
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E2figure:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   8
      Top             =   5670
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Build:"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   7
      Top             =   2130
      Width           =   1815
   End
End
Attribute VB_Name = "frmManEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
   Dim longTemp As Long          'temp variable
      
   '------------------ validate the part -------------------
   '--- Blanklength
   longTemp = Val(txtFields(13).Text)
   If longTemp < BlankMin Or longTemp > BlankMax Then
      MsgBox ("Invalid BlankLength! " & BlankMin & " >= BlankLength <= " & BlankMax)
      txtFields(13).Text = " "
      Exit Sub
   End If
   '--- E1Dim
   longTemp = Val(txtFields(15).Text)
   If longTemp < E1DimMin Or longTemp > E1DimMax Then
      MsgBox ("Invalid E1Dim! " & E1DimMin & " >= E1Dim <= " & E1DimMax)
      txtFields(15).Text = " "
      Exit Sub
   End If
   '--- E2Dim
   longTemp = Val(txtFields(17).Text)
   If longTemp < E2DimMin Or longTemp > E2DimMax Then
      MsgBox ("Invalid E2Dim! " & E2DimMin & " >= E2Dim <= " & E2DimMax)
      txtFields(17).Text = " "
      Exit Sub
   End If
   '--- CDim
   longTemp = Val(txtFields(18).Text)
   If longTemp <> 0 And longTemp < CDimMin Or longTemp > CDimMax Then
      MsgBox ("Invalid CDim! " & CDimMin & " >= CDim <= " & CDimMax)
      txtFields(18).Text = " "
      Exit Sub
   End If
   '--- C1Dim
   longTemp = Val(txtFields(19).Text)
   If longTemp <> 0 And longTemp < C1DimMin Or longTemp > C1DimMax Then
      MsgBox ("Invalid C1Dim! " & C1DimMin & " >= C1Dim <= " & C1DimMax)
      txtFields(19).Text = " "
      Exit Sub
   End If
   '--- DDim
   longTemp = Val(txtFields(20).Text)
   If longTemp <> 0 And longTemp < DDimMin Or longTemp > DDimMax Then
      MsgBox ("Invalid DDim! " & DDimMin & " >= DDim <= " & DDimMax)
      txtFields(20).Text = " "
      Exit Sub
   End If
   '--- D1Dim
   longTemp = Val(txtFields(21).Text)
   If longTemp <> 0 And longTemp < D1DimMin Or longTemp > D1DimMax Then
      MsgBox ("Invalid D1Dim! " & D1DimMin & " >= D1Dim <= " & D1DimMax)
      txtFields(21).Text = " "
      Exit Sub
   End If
   '--- Build
   If Combo(8).Text = "F" Then                     'can't have a full build
      MsgBox ("Invalid Build:  Must be Partial or Remake!")
      Combo(8).Text = "R"                          'reset to default
      Exit Sub
   End If
  
   '------------------- build the part -------------------
   intParts = intParts + 1                         'increment the # of parts in array
   
   '--- build an entry in the parts array for this part
   arrParts(intParts).Job = txtFields(0).Text
   arrParts(intParts).Rel = txtFields(1).Text
   arrParts(intParts).Item = txtFields(2).Text
   arrParts(intParts).Seq = intParts
   arrParts(intParts).ShipDate = txtFields(4).Text
   arrParts(intParts).Qnty = 0
   arrParts(intParts).BldQnty = txtFields(5).Text
   arrParts(intParts).Phase = Combo(0).Text
   arrParts(intParts).Leg = Combo(1).Text
   arrParts(intParts).Stack = Combo(2).Text
   arrParts(intParts).BarType = Combo(3).Text
   arrParts(intParts).Material = Combo(4).Text
   arrParts(intParts).BarWidth = Combo(5).Text
   arrParts(intParts).BlankLength = txtFields(13).Text
   arrParts(intParts).RunLength = -1
   arrParts(intParts).E1fig = Combo(6).Text
   arrParts(intParts).E1dim = txtFields(15).Text
   arrParts(intParts).E2fig = Combo(7).Text
   arrParts(intParts).E2dim = txtFields(17).Text
   arrParts(intParts).Cdim = txtFields(18).Text
   arrParts(intParts).C1dim = txtFields(19).Text
   arrParts(intParts).Ddim = txtFields(20).Text
   arrParts(intParts).D1dim = txtFields(21).Text
   arrParts(intParts).FullJobNum = txtFields(0).Text
   arrParts(intParts).Build = Combo(8).Text
   arrParts(intParts).Status = " *"
   arrParts(intParts).Priority = -1
   arrParts(intParts).Group = -1
   arrParts(intParts).DTS = "               "
   
   '-------------------- Clean up ---------------------
   frmPartial.RefreshPartialQue                    'refresh the que
   'MsgBox ("Part added to Partial Que!")           'OP feedback
   txtFields(3).Text = intParts + 1                'incr the displayed seq#
   
End Sub  'cmdAdd_Click()

Private Sub cmdClose_Click()
   Me.Hide
   Unload Me
End Sub

Private Sub Form_Load()
   Dim conLocal As Connection
   Dim adoRS As ADODB.Recordset
   Dim strSQL As String
   Dim strTemp As String
   Dim dblTemp As Double
   
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- construct Phase combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblPhase "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(0).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         strTemp = adoRS.Fields("Phase")        'construct string
         Combo(0).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(0).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local DB Error: Could not find Phase info!")
   End If
    
   '--- construct Leg combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblLeg "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(1).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         strTemp = adoRS.Fields("Leg")          'construct string
         Combo(1).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(1).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local DB Error: Could not find Leg info!")
   End If
    
   '--- construct Stack combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblStack "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(2).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         strTemp = adoRS.Fields("Stack")        'construct string
         Combo(2).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(2).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local DB Error: Could not find Stack info!")
   End If
    
   '--- construct BarType combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblBarType "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(3).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         strTemp = adoRS.Fields("Type")         'construct string
         Combo(3).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(3).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local DB Error: Could not find BarType info!")
   End If
    
  '--- construct Material combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblMaterial "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(4).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         strTemp = adoRS.Fields("Material")     'construct string
         Combo(4).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(4).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local DB Error: Could not find Material info!")
   End If
    
   '--- construct BarWidth combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblBarWidth "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(5).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         dblTemp = adoRS.Fields("BarWidth")     'construct string
         dblTemp = dblTemp * 1000               'convert to mils
         strTemp = Str(dblTemp)                 'convert to string
         Combo(5).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(5).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local DB Error: Could not find BarWidth info!")
   End If
    
   '--- construct E1Fig combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblE1Fig "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(6).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         strTemp = adoRS.Fields("E1Fig")        'construct string
         Combo(6).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(6).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local Db Error: Could not find E1Fig info!")
   End If
    
   '--- construct E2Fig combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblE2Fig "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(7).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         strTemp = adoRS.Fields("E2Fig")        'construct string
         Combo(7).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(7).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local DB Error: Could not find E2Fig info!")
   End If
   
   '--- construct Build combo box
   Set adoRS = New ADODB.Recordset              'make recordset
   strSQL = "SELECT * FROM tblBuild "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                'if records found
      adoRS.MoveFirst                           'go to 1st record
      Combo(8).Clear                            'clear the combo box
      
      Do Until adoRS.EOF     'loop thru recordset
         strTemp = adoRS.Fields("Build")        'construct string
         Combo(8).AddItem (strTemp)             'add item to combo box
         adoRS.MoveNext                         'incr to next record
      Loop  'recordset loop
      
      Combo(8).Refresh                          'refresh combo
      
   Else     'no records found
      MsgBox ("Local DB Error: Could not find Build info!")
   End If
   
   '--- unload recordset & connection
   adoRS.Close
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing

   '--- set default values
   'use the JRI & ShipDate from the item selected on the Select Screen
   txtFields(0).Text = usrSelItem.Job           'Job#
   txtFields(1).Text = usrSelItem.Rel           'Rel#
   txtFields(2).Text = usrSelItem.Item          'Item#
   txtFields(4).Text = usrSelItem.ShipDate      'ShipDate
   Combo(8).Text = "R"                          'set default to R = Remake
   Combo(4).Text = "A"                          'set default to 3 = ????
   txtFields(5).Text = "1"
   
End Sub  'Form_Load


Private Sub txtFields_Change(Index As Integer)

   '--- check to see if all required values are entered
   If txtFields(0).Text <> "" And _
      txtFields(1).Text <> "" And _
      txtFields(2).Text <> "" And _
      txtFields(4).Text <> "" And _
      txtFields(5).Text <> "" And _
      txtFields(13).Text <> "" And _
      txtFields(15).Text <> "" And _
      txtFields(17).Text <> "" And _
      txtFields(18).Text <> "" And _
      txtFields(19).Text <> "" And _
      txtFields(20).Text <> "" And _
      txtFields(21).Text <> "" And _
      Combo(0).Text <> "" And _
      Combo(1).Text <> "" And _
      Combo(2).Text <> "" And _
      Combo(3).Text <> "" And _
      Combo(4).Text <> "" And _
      Combo(5).Text <> "" And _
      Combo(6).Text <> "" And _
      Combo(7).Text <> "" And _
      Combo(8).Text <> "" Then
      cmdAdd.Visible = True
   Else  'all values not entered
      cmdAdd.Visible = False
   End If
   
End Sub

