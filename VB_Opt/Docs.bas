Attribute VB_Name = "Docs"
Option Explicit

'------------------------------------------------------------------------------
'Name:         Mubea Bar HMI
'Created by:   GE Automation Services
'Project #:    61-7446
'Created for:  GE Busway, Selmer, TN
'Created by:   J.Keith Anderson
'----------------------------- Version History --------------------------------
'Version:      Date:          Description:
'--------      --------       -------------------------------------------------
'1.0.0         06-04-03       Initial copy from RemBar program
'1.0.1         06-07-03       Modify Select
'1.0.2         06-09-03       Optimization Routine
'1.0.3         06-11-03       Modify Run Screen
'1.1.0         06-12-03       Added DDE Handshaking
'1.1.1         06-30-03       convert to UDP Handshaking
'1.1.2         07-06-03       convert to TCP Handshaking
'1.2.0         07-06-03       basic exec functions: Modes, Estop & Messaging
'1.2.1         07-08-03       proccess req,compl,batch
'                             clean up Hist GUI...finish batch command
'1.2.2         07-9-03        clean up button/flag permissives
'1.3.0         07-10-03       general fixes from sim testing
'1.3.1         07-11-03       added select check vs hist tbl
'                             added batch & connection status displays
'                             sendback & misc fixes
'1.3.2         07-13-03       modified procBatch to handle incomplete items
'2.0.0         07-13-03       installed for startup
'2.0.1         07-14-03       change TCP comm over to Client/Server config.
'2.0.2         08-15-03       allowed "enter" at select screen
'                             fixed suspend problem
'2.0.3         08-22-03       added time delay before ack grp req.
'2.0.4         08-27-03       activated check against hist table.
'                             increased time delay for REQ ACK
'2.1.0         05-29-04       added barTracking...preinstall
'                             also includes disabled portions of bar optimization
'2.1.1         05-29-04       Post install for tracking
'2.1.2         06-08-04       Latest tracking changes
'2.1.3         06-25-04       latest tracking changes + disabled bar opt
'                             bar opt, includes: new fill algorithm, batch submit
'                             and assignment updates for manual batching.
'2.1.4         06-30-04       After successfull tracking test.
'2.1.5         07-19-04       BarWidth fix for Optimization.
'2.1.6         07-21-04       Set up Optimize paramaters in local db.
'                             Added "Find" button to Batch Screen.
'2.1.7         07-23-04       Fix to prevent Closing of app via
'                             windows close(X) button.
'                             Zero Length string fix in procComp()
'2.1.8         07-27-04       Fix's from 7/26 production run
'              07-28-04       Change Qty button fix's.
'                             Re-optimize error after accept/clear batch
'                             Place added item/seq at top of listbox que.
'------------------------------------------------------------------------------


'------------------ definitions for running app in system tray ----------------
'constants required by Shell_NotifyIcon API call:
'Public Const NIM_ADD = &H0
'Public Const NIM_MODIFY = &H1
'Public Const NIM_DELETE = &H2
'Public Const NIF_MESSAGE = &H1
'Public Const NIF_ICON = &H2
'Public Const NIF_TIP = &H4
'Public Const WM_MOUSEMOVE = &H200
'Public Const WM_LBUTTONDOWN = &H201     'Button down
'Public Const WM_LBUTTONUP = &H202       'Button up
'Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
'Public Const WM_RBUTTONDOWN = &H204     'Button down
'Public Const WM_RBUTTONUP = &H205       'Button up
'Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

'user defined type required by Shell_NotifyIcon API call
'Public Type NOTIFYICONDATA
' cbSize As Long
' hwnd As Long
' uId As Long
' uFlags As Long
' uCallBackMessage As Long
' hIcon As Long
' szTip As String * 64
'End Type

'Public Declare Function SetForegroundWindow Lib "user32" _
'(ByVal hwnd As Long) As Long
'Public Declare Function Shell_NotifyIcon Lib "shell32" _
'Alias "Shell_NotifyIconA" _
'(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Public nid As NOTIFYICONDATA

