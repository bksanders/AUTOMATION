VERSION 5.00
Begin VB.Form frmTruck 
   Caption         =   "Truck Screen"
   ClientHeight    =   5700
   ClientLeft      =   5310
   ClientTop       =   1695
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   4095
   Begin VB.CheckBox chkEnPIgrnds 
      Caption         =   "Check1"
      Height          =   240
      Left            =   240
      TabIndex        =   20
      Top             =   5280
      Width           =   240
   End
   Begin VB.CheckBox chkEnPlugIn 
      Caption         =   "Check1"
      Height          =   240
      Left            =   240
      TabIndex        =   18
      Top             =   4920
      Width           =   240
   End
   Begin VB.CheckBox chkEnFDR 
      Caption         =   "Check1"
      Height          =   240
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Truck# Options"
      Height          =   1575
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
      Begin VB.OptionButton optStandard 
         Caption         =   "Standard"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optHand 
         Caption         =   "Hand Carry"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optRemake 
         Caption         =   "Remakes"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintPWO 
      Caption         =   "Print PWO's"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintSum 
      Caption         =   "Print Summary"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cmbTruck 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox chkAutoPWO 
      Caption         =   "Check1"
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   240
   End
   Begin VB.CheckBox chkAutoSum 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Enable Plug-In Gnds Truck"
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   5280
      Width           =   2025
   End
   Begin VB.Label Label5 
      Caption         =   "Enable Plug-In Truck"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   4920
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "Enable Feeder Truck"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   4560
      Width           =   1545
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4560
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblType 
      Caption         =   "(Feeder)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   1080
      TabIndex        =   11
      Top             =   480
      Width           =   1530
   End
   Begin VB.Line Line1 
      X1              =   -360
      X2              =   4200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      Caption         =   "Auto Print PWO's"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   4200
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Auto Print Summary"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3840
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Current Truck:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1530
   End
End
Attribute VB_Name = "frmTruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bnlLOAD As Boolean

Private Sub chkEnFDR_Click()
   If blnload = False Then
      If chkEnFDR.Value = vbChecked Then
         EnFDRTruck = True
         frmRun.lblFDR.Visible = True
         frmRun.txtFDRTruck.Visible = True
      Else
         EnFDRTruck = False
         frmRun.lblFDR.Visible = False
         frmRun.txtFDRTruck.Visible = False
      End If
   End If
End Sub

Private Sub chkEnPIgrnds_Click()
   If blnload = False Then
      If chkEnPIgrnds.Value = vbChecked Then
         EnPIGTruck = True
         frmRun.lblPIG.Visible = True
         frmRun.txtPIGTruck.Visible = True
      Else
         EnPIGTruck = False
         frmRun.lblPIG.Visible = False
         frmRun.txtPIGTruck.Visible = False
      End If
   End If
End Sub

Private Sub chkEnPlugIn_Click()
   If blnload = False Then
      If chkEnPlugIn.Value = vbChecked Then
         EnPITruck = True
         frmRun.lblPI.Visible = True
         frmRun.txtPITruck.Visible = True
      Else
         EnPITruck = False
         frmRun.lblPI.Visible = False
         frmRun.txtPITruck.Visible = False
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   
   Select Case strTruckType                        'pass truck# back
   Case "FDR"
      frmRun.txtFDRTruck.Text = cmbTruck.Text
   Case "PI"
      frmRun.txtPITruck.Text = cmbTruck.Text
   Case "PIG"
      frmRun.txtPIGTruck.Text = cmbTruck.Text
   End Select
   
   Me.Hide                                         'hide screen
   Unload Me                                       'unload screen
   
End Sub

Private Sub cmdNew_Click()

Dim strTruck As String              'truck no
Dim strType As String               'truck type
Dim filename As String
   
   '---- call the update function
   strTruck = cmbTruck.Text                                 'get truck#
   filename = "C:\Mubea\VB\TruckTrack.exe  " & strTruck
   'filename = "D:\GE-Selmer\shell_test\shell_test.exe " & strTruck
   varAPP = Shell(filename, vbNormalFocus)
   If chkAutoSum.Value = 1 Then                             'print Truck Summary
      'genReport "P", cmbTruck.Text
   End If
   
   If chkAutoPWO.Value = 1 Then                             'print PWO's
      'cmdPrintPWO                                           'execute the command click
   End If
   
   'If strTruckType = "S" Then                               'determine truck Type
      strType = "B"                                         'only straight lengths
   'Else
   '   strType = "BF"
   'End If
   
   strTruck = newTruck(strType, optRemake.Value, optHand.Value)  'calc the new truck#
   
   cmbTruck.Text = strTruck

End Sub

Private Sub cmdPrintPWO_Click()
Dim varAPP As Variant
Dim strTruck As String
Dim intResults As Integer

On Error GoTo errorHandler

   strTruck = cmbTruck.Text                              'get truck#
   'clrLocalTBL "PRTPWO"                                 'do not clear per J.Gray
   intResult = genPWOList(strTruck)                      'generate PWO list table
   If intResult > 0 Then                                 'something to print
      varAPP = Shell("C:\MUBEA\DB\PWOPRT.BAT", vbNormalFocus)
      'MsgBox ("This will call .bat file to print PWO's")
   End If
Exit Sub

errorHandler:
   MsgBox ("Error Printing PWO's!")
End Sub

Private Sub cmdPrintSum_Click()
Dim varAPP As Variant
Dim strTruck As String
Dim filename As String

On Error GoTo errorHandler

   strTruck = cmbTruck.Text                         'get truck#
   filename = "C:\Mubea\VB\TruckTrack.exe  " & strTruck
   'filename = "C:\Mubea\VB\shell_test.exe  " & strTruck                'test app on mubea pc
   'filename = "D:\GE-Selmer\shell_test\shell_test.exe " & strTruck     'test app on JKA laptop
   varAPP = Shell(filename, vbNormalFocus)
   
Exit Sub

errorHandler:
   MsgBox ("Error Printing Truck Summary!")
End Sub

Private Sub cmdView_Click()
   genReport "V", cmbTruck.Text
End Sub

Private Sub Form_Load()
   blnload = True                                  'set loading flag
   
   cmbTruck.Clear                                  'empty truck combo
   getTruckList                                    'populate the truck# combo
   
   Select Case strTruckType                        'set up for truck type
   Case "FDR"
      cmbTruck.Text = frmRun.txtFDRTruck.Text
      lblType.Caption = "Feeder"
   Case "PI"
      cmbTruck.Text = frmRun.txtPITruck.Text
      lblType.Caption = "Plug-In"
   Case "PIG"
      cmbTruck.Text = frmRun.txtPIGTruck.Text
      lblType.Caption = "Plug-In Grnds"
   End Select

   '--- initialize checkboxes
   If EnFDRTruck = True Then                       'feeder
      chkEnFDR.Value = 1
   Else
      chkEnFDR.Value = 0
   End If
   If EnPITruck = True Then                        'plug-in
      chkEnPlugIn.Value = 1
   Else
      chkEnPlugIn.Value = 0
   End If
   If EnPIGTruck = True Then                       'plug-in grnds
      chkEnPIgrnds.Value = 1
   Else
      chkEnPIgrnds.Value = 0
   End If
   
   blnload = False
End Sub


