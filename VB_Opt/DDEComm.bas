Attribute VB_Name = "DDEComm"
Option Explicit

'-----------------------------------------------------------------------------
'Name:      enDDElinks
'Accepts:   intEnable = 0 = disables DDE links
'                     = 1 = enables DDE links
'Returns:   intResult = 0 = request failed
'                     = 1 = request fullfilled
'Requires:  frmRun w/ DDE listed DDE objects.
'Discrip:   This sub enables the DDE destination links.
'Notes:
'------------------------------------------------------------------------------

Public Sub enDDElinks(intEnable As Integer, intResult As Integer)
   
'On Error GoTo errorHandler
   
   frmRun.txtMubMessage.Text = App.EXEName
   frmRun.txtMubMessage.Text = frmRun.txtMubMessage.Text & "|" & frmRun.LinkTopic
   
   '----------------------- Define DDE  Destination Links ----------------------
   frmRun.ddeGrpREQ.LinkTopic = "MBSim|frmSim"
   frmRun.ddeGrpREQ.LinkItem = "txtReqGrp"
   
   frmRun.ddeGrpCmpl.LinkTopic = "MBSim|frmSim"
   frmRun.ddeGrpCmpl.LinkItem = "txtGrpCmpl"
   
   frmRun.ddeMSG.LinkTopic = "MBSim|frmSim"
   frmRun.ddeMSG.LinkItem = "txtMSG"
       
   '------------------------- enable/disable DDE Links ------------------------
   Select Case intEnable
   Case 0
      frmRun.ddeGrpREQ.LinkMode = 0             'disable links
      frmRun.ddeGrpCmpl.LinkMode = 0
      frmRun.ddeMSG.LinkMode = 0
   Case 1
      frmRun.ddeGrpREQ.LinkMode = 1             'enable links
      frmRun.ddeGrpCmpl.LinkMode = 1
      frmRun.ddeMSG.LinkMode = 1
   End Select
   intResult = 1                                'return OK
Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
   If Err.Number = 282 Then                     'error 282
      MsgBox ("The Mubea OI is NOT running! " & Chr$(13) & _
              "Start the OI and try again.")
   Else
      MsgBox Err.Description
   End If
   intResult = 0                                'return FAILED
End Sub
