VERSION 5.00
Begin VB.Form frmEditPart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Part"
   ClientHeight    =   8535
   ClientLeft      =   8235
   ClientTop       =   330
   ClientWidth     =   3735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   3735
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   8
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   7614
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   7
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   5358
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   6
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   4692
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   5
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   4026
      Width           =   855
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   4
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   3678
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   3
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   3330
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   2
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   2982
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   1
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   2634
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   0
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   2286
      Width           =   615
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   2640
      TabIndex        =   40
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FullOrder"
      Height          =   285
      Index           =   22
      Left            =   2040
      TabIndex        =   37
      Top             =   7296
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "D1dimension"
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   35
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ddimension"
      Height          =   285
      Index           =   20
      Left            =   2055
      TabIndex        =   33
      Top             =   6660
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1dimension"
      Height          =   285
      Index           =   19
      Left            =   2055
      TabIndex        =   31
      Top             =   6342
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Cdimension"
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   29
      Top             =   6024
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E2dimension"
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   27
      Top             =   5706
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E1dimension"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   24
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BlankLength"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   21
      Top             =   4374
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BldQnty"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1968
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Quantity"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1650
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Scheduled Ship Date"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1332
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Sequence Number"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1014
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Item"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   696
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Release"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   378
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Order Number"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Build:"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   38
      Top             =   7644
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "FullOrder:"
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   36
      Top             =   7311
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "D1dimension:"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   34
      Top             =   6993
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ddimension:"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   32
      Top             =   6675
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "C1dimension:"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   30
      Top             =   6357
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cdimension:"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   28
      Top             =   6039
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E2dimension:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   26
      Top             =   5721
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E2figure:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   25
      Top             =   5388
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E1dimension:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   23
      Top             =   5055
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E1figure:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   22
      Top             =   4722
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BlankLength:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   20
      Top             =   4389
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BarWidth:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   19
      Top             =   4056
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Material:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   18
      Top             =   3708
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BarType:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stack:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Leg:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   15
      Top             =   2664
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Phase:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2316
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Desired Qnty:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1983
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Default Qnty:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1665
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Scheduled Ship Date:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1347
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sequence Number:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1029
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   711
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Release:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   393
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Order Number:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAccept_Click()
   Dim longTemp As Long          'temp variable
   Dim longDefQty As Long        'temp default quantity
   Dim longDesQty As Long        'temp desired quantity
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
   If Combo(8).Text = "F" Then                        'can't have a full build
      MsgBox ("Invalid Build:  Must be Partial or Remake!")
      Combo(8).Text = "R"                             'reset to default
      Exit Sub
   End If
   '--- Desired Quantity
   longDefQty = Val(txtFields(5).Text)                'get default quantity
   If longDefQty > 0 Then                             'if default qty > 0
      longDesQty = Val(txtFields(6).Text)             'get desired quantity
      If longDesQty > longDefQty Then                 'desired > default
         MsgBox ("Invalid Quantity: Desired Qty must be <= Default Qty!")
         txtFields(6).Text = "0"                      'reset qty
         Exit Sub
      End If
   End If
     
   '------------------- update the part -------------------
   '--- update entry in the parts array for this part
   arrParts(intIndex).Job = txtFields(0).Text
   arrParts(intIndex).Rel = txtFields(1).Text
   arrParts(intIndex).Item = txtFields(2).Text
   arrParts(intIndex).Seq = txtFields(3).Text
   arrParts(intIndex).ShipDate = txtFields(4).Text
   arrParts(intIndex).Qnty = txtFields(5).Text
   arrParts(intIndex).BldQnty = txtFields(6).Text
   arrParts(intIndex).Phase = Combo(0).Text
   arrParts(intIndex).Leg = Combo(1).Text
   arrParts(intIndex).Stack = Combo(2).Text
   arrParts(intIndex).BarType = Combo(3).Text
   arrParts(intIndex).Material = Combo(4).Text
   arrParts(intIndex).BarWidth = Combo(5).Text
   arrParts(intIndex).BlankLength = txtFields(13).Text
   arrParts(intIndex).E1fig = Combo(6).Text
   arrParts(intIndex).E1dim = txtFields(15).Text
   arrParts(intIndex).E2fig = Combo(7).Text
   arrParts(intIndex).E2dim = txtFields(17).Text
   arrParts(intIndex).Cdim = txtFields(18).Text
   arrParts(intIndex).C1dim = txtFields(19).Text
   arrParts(intIndex).Ddim = txtFields(20).Text
   arrParts(intIndex).D1dim = txtFields(21).Text
   arrParts(intIndex).FullJobNum = txtFields(22).Text
   arrParts(intIndex).Build = Combo(8).Text
   
   '-------------------- Clean up ---------------------
   frmPartial.RefreshPartialQue                       'refresh the que
   'MsgBox ("Part added to Partial Que!")             'OP feedback
   Unload Me                                          'unload form
   
End Sub  'cmdAccept_Click()

Private Sub Form_Load()
   '--- Global Variables
   'intIndex                           'index of selected part
   
   '--- Variable Declarations
   Dim conLocal As Connection
   Dim adoRS As ADODB.Recordset
   Dim strSQL As String
   Dim strTemp As String
   Dim dblTemp As Double
   
   '--------------------- Build Combo Boxes ----------------------
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
         strTemp = Str(dblTemp)
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
                                   
                                   
   '-------------------- populate the form --------------------------
   txtFields(0).Text = arrParts(intIndex).Job
   txtFields(1).Text = arrParts(intIndex).Rel
   txtFields(2).Text = arrParts(intIndex).Item
   txtFields(3).Text = arrParts(intIndex).Seq
   txtFields(4).Text = arrParts(intIndex).ShipDate
   txtFields(5).Text = arrParts(intIndex).Qnty
   txtFields(6).Text = arrParts(intIndex).BldQnty
   Combo(0).Text = arrParts(intIndex).Phase
   Combo(1).Text = arrParts(intIndex).Leg
   Combo(2).Text = arrParts(intIndex).Stack
   Combo(3).Text = arrParts(intIndex).BarType
   Combo(4).Text = arrParts(intIndex).Material
   Combo(5).Text = arrParts(intIndex).BarWidth
   txtFields(13).Text = arrParts(intIndex).BlankLength
   Combo(6).Text = arrParts(intIndex).E1fig
   txtFields(15).Text = arrParts(intIndex).E1dim
   Combo(7).Text = arrParts(intIndex).E2fig
   txtFields(17).Text = arrParts(intIndex).E2dim
   txtFields(18).Text = arrParts(intIndex).Cdim
   txtFields(19).Text = arrParts(intIndex).C1dim
   txtFields(20).Text = arrParts(intIndex).Ddim
   txtFields(21).Text = arrParts(intIndex).D1dim
   txtFields(22).Text = arrParts(intIndex).FullJobNum
   Combo(8).Text = arrParts(intIndex).Build
                                    
End Sub

Private Sub txtFields_Change(Index As Integer)

   '--- check to see if all required values are entered
   If txtFields(0).Text <> "" And _
      txtFields(1).Text <> "" And _
      txtFields(2).Text <> "" And _
      txtFields(4).Text <> "" And _
      txtFields(6).Text <> "" And _
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
      cmdAccept.Visible = True
   Else  'all values not entered
      cmdAccept.Visible = False
   End If
   
End Sub


Private Sub cmdClose_Click()
   Unload Me
End Sub

