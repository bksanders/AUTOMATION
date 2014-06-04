VERSION 5.00
Begin VB.Form frmViewHist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Completed Part"
   ClientHeight    =   9285
   ClientLeft      =   8235
   ClientTop       =   330
   ClientWidth     =   3690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   3690
   Begin VB.TextBox txtFields 
      DataField       =   "Scheduled Ship Date"
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   2040
      TabIndex        =   53
      Top             =   8070
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Scheduled Ship Date"
      Enabled         =   0   'False
      Height          =   285
      Index           =   24
      Left            =   2040
      TabIndex        =   51
      Top             =   7740
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Scheduled Ship Date"
      Enabled         =   0   'False
      Height          =   285
      Index           =   26
      Left            =   2040
      TabIndex        =   49
      Top             =   8370
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   48
      Top             =   8760
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Build"
      Enabled         =   0   'False
      Height          =   285
      Index           =   23
      Left            =   2040
      TabIndex        =   47
      Top             =   7420
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FullOrder"
      Enabled         =   0   'False
      Height          =   285
      Index           =   22
      Left            =   2040
      TabIndex        =   45
      Top             =   7100
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "D1dimension"
      Enabled         =   0   'False
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   43
      Top             =   6780
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ddimension"
      Enabled         =   0   'False
      Height          =   285
      Index           =   20
      Left            =   2040
      TabIndex        =   41
      Top             =   6460
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1dimension"
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   2040
      TabIndex        =   39
      Top             =   6140
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Cdimension"
      Enabled         =   0   'False
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   37
      Top             =   5820
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E2dimension"
      Enabled         =   0   'False
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   35
      Top             =   5500
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E2figure"
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   2040
      TabIndex        =   33
      Top             =   5180
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E1dimension"
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   31
      Top             =   4860
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "E1figure"
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   29
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BlankLength"
      Enabled         =   0   'False
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   27
      Top             =   4220
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BarWidth"
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   25
      Top             =   3915
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Material"
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Top             =   3580
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BarType"
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   3260
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Stack"
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   2940
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Leg"
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2620
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Phase"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2300
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BldQnty"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Quantity"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Scheduled Ship Date"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Sequence Number"
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Item"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Release"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Order Number"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Machine:"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   54
      Top             =   8100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Truck NO:"
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   52
      Top             =   7770
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Date/Time Stamp:"
      Height          =   255
      Index           =   26
      Left            =   120
      TabIndex        =   50
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Build:"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   46
      Top             =   7420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "FullOrder:"
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   44
      Top             =   7100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "D1dimension:"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   42
      Top             =   6780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ddimension:"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   40
      Top             =   6460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "C1dimension:"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   38
      Top             =   6140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cdimension:"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   36
      Top             =   5820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E2dimension:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   5500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E2figure:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Top             =   5180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E1dimension:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   30
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "E1figure:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   4540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BlankLength:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BarWidth:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Material:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BarType:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stack:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Leg:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Phase:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BldQnty:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Quantity:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Scheduled Ship Date:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sequence Number:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Release:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Order Number:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmViewHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Dim strSQL As String
Dim conLocal As Connection
Dim adorsHist As ADODB.Recordset

'----------------------------- Program Control Method ------------------------
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                              
   '--- create the recordset
   Set adorsHist = New ADODB.Recordset          'init recordset
                                                'build the SQL string
   strSQL = "SELECT * FROM tblHist WHERE " & _
            "[Order Number] = '" & usrRecID.Job & "' AND " & _
            "Release = '" & usrRecID.Rel & "' AND " & _
            "Item = " & usrRecID.Item & " AND " & _
            "[Sequence Number] = " & usrRecID.Seq
                                                'open recordset
   adorsHist.Open strSQL, conLocal, adOpenStatic, adLockOptimistic

   If adorsHist.RecordCount > 0 Then            'if a Record is found
      adorsHist.MoveFirst
   Else                                         'if not found
      MsgBox ("Record Not Located in Exec Que!")
      adorsHist.Close                           'close recordset
      Set adorsHist = Nothing                   'unload recordset
      conLocal.Close                          'close connection
      Set conLocal = Nothing                  'unload connection
      Unload Me
      Exit Sub
   End If
                                    
   '--- populate the form
   txtFields(0).Text = adorsHist.Fields("Order Number")
   txtFields(1).Text = adorsHist.Fields("Release")
   txtFields(2).Text = adorsHist.Fields("Item")
   txtFields(3).Text = adorsHist.Fields("Sequence Number")
   txtFields(4).Text = adorsHist.Fields("Scheduled Ship Date")
   txtFields(5).Text = adorsHist.Fields("Quantity")
   txtFields(6).Text = adorsHist.Fields("BldQnty")
   txtFields(7).Text = adorsHist.Fields("Phase")
   txtFields(8).Text = adorsHist.Fields("Leg")
   txtFields(9).Text = adorsHist.Fields("Stack")
   txtFields(10).Text = adorsHist.Fields("BarType")
   txtFields(11).Text = adorsHist.Fields("Material")
   txtFields(12).Text = adorsHist.Fields("BarWidth")
   txtFields(13).Text = adorsHist.Fields("BlankLength")
   txtFields(14).Text = adorsHist.Fields("E1figure")
   txtFields(15).Text = adorsHist.Fields("E1dimension")
   txtFields(16).Text = adorsHist.Fields("E2figure")
   txtFields(17).Text = adorsHist.Fields("E2dimension")
   txtFields(18).Text = adorsHist.Fields("Cdimension")
   txtFields(19).Text = adorsHist.Fields("C1dimension")
   txtFields(20).Text = adorsHist.Fields("Ddimension")
   txtFields(21).Text = adorsHist.Fields("D1dimension")
   txtFields(22).Text = adorsHist.Fields("FullOrder")
   txtFields(23).Text = adorsHist.Fields("Build")
   If IsNull(adoRS.Fields("Truck")) Then
      txtFields(24).Text = "None"
   Else
      txtFields(24).Text = adorsHist.Fields("Truck")
   End If
   txtFields(25).Text = adorsHist.Fields("Machine")
   txtFields(26).Text = adorsHist.Fields("DTStamp")
                                    
   adorsHist.Close                              'close recordset
   Set adorsHist = Nothing                      'unload recordset
   conLocal.Close                             'close connection
   Set conLocal = Nothing                     'unload connection
                    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub


