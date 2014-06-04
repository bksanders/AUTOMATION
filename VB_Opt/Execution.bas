Attribute VB_Name = "Execution"
'---------------------------------------------------------------------------
'Name:      Execution
'Accepts:
'Returns:
'Requires:
'Discrip:   The Execution Module contains the code requred for to execute the
'           sequence of operations needed to build parts.
'Notes:
'---------------------------------------------------------------------------
Option Explicit

Global AUTO As Boolean              'auto mode flag
Global MANUAL As Boolean            'manual mode flag
Global ESTOP As Boolean             'ESTOP flag, true when ESTOP Hit
Global STARTUP As Boolean           'start flag, true after STARTUP sub run
Global PAUSE As Boolean             'PAUSE flag, ture when OP wants to PAUSE system
Global SUSPEND As Boolean           'SUSPEND flag, true when SUSPENDing a piece
Global SINGLEGRP As Boolean         'single step exec mode
Global NoStopPROMPT As Boolean      'flag to skip STOP prompt
Global BATCH As Boolean             'batch in progress flag
Global OPTIMIZED As Boolean         'optimized flag
Global BARinPROG As Boolean         'bar in progress flag
Global pendMSG  As Boolean          'recv signal pending
Global intTotBlanks As Integer      'total # of blanks in batch
Global intRemBlanks As Integer      'remain blanks in batch
Global intReqCnt As Integer         'count of Req's received--debug only
Global ACKdelayDN As Boolean        'Req ACK Delay

'------------------------------------------------------------------------------
'Name:      startExec()
'Accepts:   None
'Returns:   None
'Requires:
'Discrip:   Sets up and starts execution module.
'Notes:
'------------------------------------------------------------------------------
Sub startExec()
   '--- global variables
   'ESTOP
   'AUTO
   'BATCH
   'pendMSG
   
   '--- local variables
   Dim intLength As Integer                'temp. string length variable
   
   '---------------------------------------------------------------------------
   ESTOP = False                                      'reset Estop flag
   AUTO = True                                        'set AUTO flag
   BATCH = True                                       'set the BATCH flag
   
   '-------------------------------------------------- set Button visibility
   'frmRun.cmdOptimize.Visible = False                 'hide Optimize Button
   'frmRun.cmdSendBack.Visible = False                 'hide Sendback Button
   'frmRun.cmdSUSPEND.Visible = True                   'show SUSPEND Button
   
   'With frmRun.tcpConn  '---------------------------- check/establish TCP conn
   '   If .State <> sckConnected Then
   '      If .State <> sckListening Then .Close
   '      .Connect                                     'initiate a tcp connection
   '   End If
   'End With
   
   '------------------------------------------------- check for pending message
   intLength = Len(frmRun.txtTCPRecv.Text)
   If intLength > 1 Then
      pendMSG = True                                  'set pending msg flag
      frmRun.tcpServer_DataArrival intLength          'proc the message
   End If

End Sub  'startExec()
   

'---------------------------------------------------------------------------
'Name:      stopExec()
'Accepts:
'Returns:
'Requires:
'Discrip:
'Notes:
'---------------------------------------------------------------------------
Sub stopExec()
   '--- globals used
   'AUTO
   'SINGLEGRP
   'PAUSE = False
   
   '--- variable declarations
   
   '------------------------------------------------------------------------
   AUTO = False
   SINGLEGRP = False
   PAUSE = False
   NoStopPROMPT = False
   frmRun.Shape1.FillColor = &HFF&                    'show stop ind (red)
   'frmRun.cmdSUSPEND.Visible = False                  'hide SUSPEND Button
   'frmRun.cmdOptimize.Visible = True                  'show Optimize Button
   'frmRun.cmdSendBack.Visible = True                  'show Sendback Button
   
End Sub  'stopExec()
   


'---------------------------------------------------------------------------
'Name:      procESTOP()
'Accepts:   None
'Returns:   None
'Requires:
'Discrip:   This sub processes an Estop.  It turns of the Execution program.
'           and moves any pending parts back to the exec que.
'Notes:
'---------------------------------------------------------------------------
Sub procESTOP()
   '--- globals used
   '--- variable declarations
   '------------------------------------------------------------------------
   
   'stopExec                                           'stop execution
   frmRun.optStop.Value = True                        'show stop mode
   
   '---Message to operator
   MsgBox ("ProcError: An ESTOP has been detected!" & Chr$(13) & _
          "Note status of bar currently in Mubea. ")
         
End Sub  'procESTOP()




'------------------------------------------------------------------------------
'Name:      procMSG
'Accepts:   intCode, message code integer
'Returns:
'Requires:  tblMSG in local db
'Discrip:   This sub processes a message code from the Mubea-OI.
'Notes:
'------------------------------------------------------------------------------
Sub procMSG(intCode As Integer)
   '--- globals used
   'ESTOP      ESTOP flag
   
   '--- variable declarations
   '---------------------------------------------------------------------------
   
   If intCode > 139 And intCode < 146 Then '------------ check for ESTOP
      ESTOP = True                                       'set ESTOP flag
      procESTOP                                          'process the estop
   End If

   If intCode > 0 Then                                   'if an msg exists
      dspMSG intCode                                     'display the MSG
   End If
End Sub  'procMSG

'------------------------------------------------------------------------------
'Name:      procReq
'Accepts:   none
'Returns:   none
'Requires:
'Discrip:   This sub processes a request(from the mubeaOI)for a new group.  It
'           detects if a group is already in the run Que, and prompts the OP to
'           determine if the part should be built or not.   It also detects an
'           empty batch tbl, indicating that the batch is done.
'Notes:
'------------------------------------------------------------------------------

Sub procReq()

'--- Global Variables
'intRunCnt                    '# of parts in the run Que
'PAUSE
'SUSPEND

'--- Variable Declarations
Dim conLocal As Connection    'dbase variables
Dim adoRS As ADODB.Recordset
Dim strSQL As String
Dim strMSG As String          'message box variables
Dim intStyle As Integer
Dim strTitle As String
Dim intResponse As Integer
Dim intOK As Integer          'getRunGroup result
Dim clrOK As Boolean          'result of clearing the run que
      
   '---------------------------------------------------------------------------
   On Error GoTo errorHandler
   
   intReqCnt = intReqCnt + 1                          'track req's & display, debug only
   frmRun.txtCount.Text = intReqCnt
   
   If PAUSE = True And BATCH = True Then    '-------- check for complete batch
      '--- make connection to local database
      Set conLocal = New Connection
      conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
      '--- make recordset for Batch table
      Set adoRS = New ADODB.Recordset
      strSQL = "SELECT * FROM tblBatch "
      adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
      
      If adoRS.RecordCount < 1 Then                   'batch empty
         procBatch                                    'process batch as complete
      End If
      Exit Sub
   End If
   
   If SUSPEND = True Then  '------------------------- handle SUSPEND
      procSusp                                        'process the suspend
      Exit Sub
   End If
   
   If intRunCnt > 0 Then   '------------------------- parts already in Run Que
      '--- operator prompt
      strMSG = "There's already a group in the Run Que." & Chr$(13) & _
               "Has it been successfully built?"      'define msgbox test
      intStyle = vbYesNo + vbDefaultButton2           'define msgbox buttons
      strTitle = "Part in Run Que..." 'define msgbox title
      '--- OP response
      intResponse = MsgBox(strMSG, intStyle, strTitle)
      If intResponse = vbYes Then  '--- OP chose Yes
         clrRun clrOK                                 'empty the Run Que
         getRunGroup intOK                            'get a group from the batch table
      Else  '--- OP chose NO
         intOK = 1                                    'already have a group
      End If
   Else  '------------------------------------------- RunQue empty
      getRunGroup intOK                               'get a group from the batch table
   End If   'parts in que
   
   Select Case intOK
   Case 1   '--- got group OK
      frmRun.txtTCPRecv.Text = ""                     'clear the recv text
      frmRun.txtTCPSend.Text = "GEAS_ReqAck"
      BARinPROG = True
      
      wait 2
      
      frmRun.tcpClient.Connect
           
   Case 2   '--- batch done
      procBatch                                       'process batch
   Case 3   '--- error getting group
      MsgBox ("Could NOT get a group from tblBatch!")
   End Select
Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
      
End Sub  'procReq

'---------------------------------------------------------------------------
'Name:      procCmpl
'Accepts:   none
'Returns:   none
'Requires:  local db & frmRun
'Discrip:   this sub processes a complete signal from the Mubea-OI
'Notes:
'---------------------------------------------------------------------------
Sub procCmpl()

'--- variable declarations
Dim MATCH As Boolean                   'match flag
Dim intTemp As Integer                 'temp integer
Dim clrOK As Boolean                   'clear result
Dim strMSG As String                   'TCP message string
Dim usrTemp As part                    'temp part
Dim usrRID As RecordID                 'temp record ID
Dim strTemp As String                  'temp string
Dim strDTS As String                   'date/time stamp

'On Error GoTo errorHandler
   
   strDTS = Now
      
   With frmRun.AdodcRun.Recordset
      If .RecordCount > 0 Then                           'make sure tbl NOT empty
         .MoveFirst
         Do Until .EOF  '------------------------------- loop thru recordset
            '--- build temp record id
            usrRID.Job = .Fields("Order Number")
            usrRID.Rel = .Fields("Release")
            usrRID.Item = .Fields("Item")
            usrRID.Seq = .Fields("Sequence Number")
                       
            '----------------------- Locate Seq in Exec Que ---------------------
            With frmRun.AdodcExec.Recordset   '--------- locate record in Exec Que
               If .RecordCount > 0 Then                  'skip if ExecQue empty
                  .MoveFirst
                  MATCH = False                          'reset match flag
                  Do Until .EOF Or MATCH
                     If .Fields("Order Number") = usrRID.Job And _
                        .Fields("Release") = usrRID.Rel And _
                        .Fields("Item") = usrRID.Item And _
                        .Fields("Sequence Number") = usrRID.Seq Then
                        
                        '--- increment the build quantity for the exec que
                        intTemp = .Fields("BldQnty")
                        .Fields("BldQnty") = intTemp + 1
                        .Update
                        MATCH = True
                        
                        '--- tracking --------------------------------------------
                        '--- build temp part
                        usrTemp.FullJobNum = .Fields("FullOrder")
                        usrTemp.Job = .Fields("Order Number")
                        usrTemp.Rel = .Fields("Release")
                        usrTemp.Item = .Fields("Item")
                        usrTemp.Seq = .Fields("Sequence Number")
                        usrTemp.ShipDate = .Fields("Scheduled Ship Date")
                        usrTemp.Qnty = .Fields("Quantity")
                        usrTemp.BldQnty = .Fields("BldQnty")
                        usrTemp.Phase = .Fields("Phase")
                        usrTemp.Leg = .Fields("Leg")
                        usrTemp.Stack = .Fields("Stack")
                        usrTemp.BarType = .Fields("BarType")
                        usrTemp.Material = .Fields("Material")
                        usrTemp.BarWidth = .Fields("BarWidth")
                        usrTemp.BlankLength = .Fields("BlankLength")
                        usrTemp.RunLength = .Fields("RunLength")
                        usrTemp.E1fig = .Fields("E1figure")
                        usrTemp.E1dim = .Fields("E1dimension")
                        usrTemp.E2fig = .Fields("E2figure")
                        usrTemp.E2dim = .Fields("E2dimension")
                        usrTemp.Cdim = .Fields("Cdimension")
                        usrTemp.C1dim = .Fields("C1dimension")
                        usrTemp.Ddim = .Fields("Ddimension")
                        usrTemp.D1dim = .Fields("D1dimension")
                        usrTemp.Build = .Fields("Build")
                        usrTemp.Priority = .Fields("Priority")
                        usrTemp.Status = .Fields("Status")
                        
                        '----------------------- Log tracking info -----------------
                        '--- Assign truck #
                        'this code will set truck number for each part based on bartype & phase
                        If blnEnTRACK = True Then
                           Select Case usrTemp.BarType
                           Case "F"
                              'if seq >=19 then this part is part of a riser
                              'plug-in and should be assinged to the plug-in truck.
                              If usrTemp.Seq >= 19 Then
                                 If EnPITruck = True Then
                                    strTemp = frmRun.txtPITruck.Text
                                 Else
                                    strTemp = "None"
                                 End If
                              Else
                                 If EnFDRTruck = True Then
                                    strTemp = frmRun.txtFDRTruck.Text
                                 Else
                                    strTemp = "None"
                                 End If
                              End If
                           Case "P"
                              If usrTemp.Phase = "3" Or _
                                 usrTemp.Phase = "4" Then
                                 If EnPIGTruck = True Then
                                    strTemp = frmRun.txtPIGTruck.Text
                                 Else
                                    strTemp = "None"
                                 End If
                              Else
                                 If EnPITruck = True Then
                                    strTemp = frmRun.txtPITruck.Text
                                 Else
                                    strTemp = "None"
                                 End If
                              End If
                           Case Else
                              strTemp = "None"
                           End Select
                           usrTemp.Truck = strTemp                      'assign truck #
                           
                           usrTemp.DTS = strDTS                         'set DTS
                           usrTemp.Machine = "M"                        'set Machine
                           
                           If IsNull(strTemp) Or strTemp = "" Then strTemp = "None"
                           If strTemp <> "None" Then                    'only log part if it has a truck#
                              addTracking usrTemp, intTemp              'log to tracking table
                           End If
                        End If   '------------------------------------- tracking enabled
                                    
                     End If
                     .MoveNext
                  Loop  '--- match
               Else  ' Nothing in Exec Que
                  MsgBox ("procCmpl: Could NOT locate complete parts in ExecQue!")
                  Exit Sub
               End If
            End With
            
            If MATCH = False Then                        'NO match found
               MsgBox ("procCmpl: Could NOT locate complete parts in ExecQue!")
               Exit Sub
            End If
                   
            .MoveNext
         Loop           '-------------------------------  run que
      End If   'NOT empty
   End With 'adodcrun
   
   intRemBlanks = intRemBlanks - 1                       'decr remain blank cnt
   BARinPROG = False                                     'clear bar in prog flag
   
   clrRun clrOK                                          'clr the run que
   frmRun.txtTCPRecv.Text = ""                           'clear the recv text
   frmRun.RefreshExecQue                                 'refresh the ExecQue
   
   'if the single grp has been processed then switch to pause
   If SINGLEGRP = True Then   '------------------------- check for singlegrp mode
      frmRun.optPause = True                             'switch to pause
   End If
   
   '---------------------------------------------------- ACK the complete
   frmRun.txtTCPSend.Text = "GEAS_CmplAck"
   frmRun.tcpClient.Connect

Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  
End Sub  'procCmpl


'---------------------------------------------------------------------------
'Name:      procBatch
'Accepts:   none
'Returns:   none
'Requires:  local db & frmRun
'Discrip:   This sub processes a batch
'Notes: (1) This sub handles complete or suspended batches.  If an item is
'           complete its status will be marked as " C".  If incomplete, the
'           item will be marked as " S", if Suspending, or " I" = incomplete
'           if completing the batch normally.
'---------------------------------------------------------------------------
Sub procBatch()

'--- variable declarations
Dim intItemCnt As Integer              '# of items in Batch
Dim intSusCnt As Integer               '# of suspended items
Dim intIncCnt As Integer               '# of incomplete items
Dim intCmpCnt As Integer               '# of complete items
Dim clrOK As Boolean                   'clear result
Dim strDTS As String                   'date/time stamp
Dim usrTemp As part                    'temp. part holder
Dim intAddOK As Integer                'add to HIST result

'---------------------------------------------------------------------------

'On Error GoTo errorHandler
                         
   strDTS = Now                                          'capture DTS
   
   With frmRun.AdodcExec.Recordset
      If .RecordCount > 0 Then                           'skip if ExecQue empty
         intItemCnt = .RecordCount                       'capture # of items
         .MoveFirst
         Do Until .EOF                                   'loop thru ExecQue
            '--- build temp part
            usrTemp.FullJobNum = .Fields("FullOrder")
            usrTemp.Job = .Fields("Order Number")
            usrTemp.Rel = .Fields("Release")
            usrTemp.Item = .Fields("Item")
            usrTemp.Seq = .Fields("Sequence Number")
            usrTemp.ShipDate = .Fields("Scheduled Ship Date")
            usrTemp.Qnty = .Fields("Quantity")
            usrTemp.BldQnty = .Fields("BldQnty")
            usrTemp.Phase = .Fields("Phase")
            usrTemp.Leg = .Fields("Leg")
            usrTemp.Stack = .Fields("Stack")
            usrTemp.BarType = .Fields("BarType")
            usrTemp.Material = .Fields("Material")
            usrTemp.BarWidth = .Fields("BarWidth")
            usrTemp.BlankLength = .Fields("BlankLength")
            usrTemp.RunLength = .Fields("RunLength")
            usrTemp.E1fig = .Fields("E1figure")
            usrTemp.E1dim = .Fields("E1dimension")
            usrTemp.E2fig = .Fields("E2figure")
            usrTemp.E2dim = .Fields("E2dimension")
            usrTemp.Cdim = .Fields("Cdimension")
            usrTemp.C1dim = .Fields("C1dimension")
            usrTemp.Ddim = .Fields("Ddimension")
            usrTemp.D1dim = .Fields("D1dimension")
            usrTemp.Build = .Fields("Build")
            usrTemp.Priority = .Fields("Priority")
            usrTemp.DTS = strDTS                         'set DTS
         
            usrTemp.Machine = "R"                        'set machine
              
            If usrTemp.BldQnty >= usrTemp.Qnty Then      'set status
               usrTemp.Status = " C"
               intCmpCnt = intCmpCnt + 1
            Else
               If SUSPEND = True Then
                  usrTemp.Status = " S"
                  intSusCnt = intSusCnt + 1
               Else
                  usrTemp.Status = " I"
                  intIncCnt = intIncCnt + 1
               End If
            End If
            
            intAddOK = 0                                 'reset add result
            addHist usrTemp, intAddOK                    'add to HistQue
            
            .MoveNext
         Loop  '--- match
      Else  ' Nothing in Exec Que
         MsgBox ("procBatch: ExecQue empty!")
         Exit Sub
      End If   'skip
   End With
            
            
   OPTIMIZED = False                                     'reset optimize flag
   BATCH = False                                         'reset batch(in prog) flag
   NoStopPROMPT = True                                   'skip STOP prompt
   frmRun.optStop.Value = True                           'go to Stop mode
   intTotBlanks = 0                                      'clear blank counts
   intRemBlanks = 0
   clrExec clrOK                                         'clear the ExecQue
   clrRun clrOK                                          'clear RunQue
   
Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  
End Sub  'procBatch

'---------------------------------------------------------------------------
'Name:      dspMSG
'Accepts:   intMsgCode = message code
'Returns:
'Requires:  message table in local dbase
'Discrip:   This sub accepts a msg code, locates the msg in the message
'           table, gets the dispcription and displays it in the Mubea
'           message textbox.
'Notes:
'---------------------------------------------------------------------------

Public Sub dspMSG(intMsgCode As Integer)
   '--- Variable Declarations
   Dim conLocal As Connection
   Dim adoRS As ADODB.Recordset
   Dim strSQL As String
   
   '-------------------- display the error description ----------------
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
   '--- make recordset for Message list
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblMSG " & _
            "WHERE MsgCode = " & intMsgCode
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                      'found a record
      adoRS.MoveFirst                                 'go to the record
      frmRun.txtMubMessage = adoRS.Fields("MsgDescription")
   Else
      MsgBox ("procMsg: Message = " & intMsgCode & " NOT found!")
   End If
   
   '------------------------------- cleanup --------------------------------
   adoRS.Close                                        'unload recordset
   Set adoRS = Nothing
   conLocal.Close                                     'unload connection
   Set conLocal = Nothing
   
End Sub  'dspMsg

'---------------------------------------------------------------------------
'Name:      getRunGroup()
'Accepts:   none
'Returns:   an integer, intResult = 1 = got a group
'                                 = 2 = batch done
'                                 = 3 = error
'Requires:  local db
'Discrip:   this moves the next group, from the batch table to the run Que.
'Notes:
'---------------------------------------------------------------------------
Public Sub getRunGroup(intResult As Integer)

Dim i As Integer                       'loop index
Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String
Dim usrPart As part                    'temp part holder
Dim intGroup As Integer                'group #
Dim intAddOK As Integer                'add result

'On Error GoTo errorHandler
   
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                               
   '--------------------------- locate the Next Group ----------------------
   '--- make recordset for Batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBatch " & _
            "ORDER BY Priority ASC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
    
   With adoRS
      If .RecordCount > 0 Then                  'tblBatch has a group
         .MoveFirst
         intGroup = .Fields("Priority")         'get the next group's priority#
         .Close
      Else                                      'tblBatch empty...batch done
         intResult = 2
         .Close
         Set adoRS = Nothing                    'clean up befor exit
         conLocal.Close
         Set conLocal = Nothing
         Exit Sub
      End If
   End With
         
   '---------------------------------------------------------------------------
   '--- create a recordset of the next group
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBatch " & _
            "WHERE Priority = " & intGroup & " " & _
            "ORDER BY BlankLength DESC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic

   With adoRS
      If .RecordCount > 0 Then
         .MoveFirst
         For i = 1 To .RecordCount
            '--- bld usrPart
            'usrPart.FullJobNum = .Fields("FullOrder")
            usrPart.Job = .Fields("Order Number")
            usrPart.Rel = .Fields("Release")
            usrPart.Item = .Fields("Item")
            usrPart.Seq = .Fields("Sequence Number")
            'usrPart.ShipDate = .Fields("Scheduled Ship Date")
            'usrPart.Qnty = .Fields("Quantity")
            'usrPart.BldQnty = .Fields("BldQnty")
            usrPart.Phase = .Fields("Phase")
            usrPart.Leg = .Fields("Leg")
            usrPart.Stack = .Fields("Stack")
            usrPart.BarType = .Fields("BarType")
            usrPart.Material = .Fields("Material")
            usrPart.BarWidth = .Fields("BarWidth")
            usrPart.BlankLength = .Fields("BlankLength")
            'usrPart.RunLength = .Fields("RunLength")
            usrPart.E1fig = .Fields("E1figure")
            usrPart.E1dim = .Fields("E1dimension")
            usrPart.E2fig = .Fields("E2figure")
            usrPart.E2dim = .Fields("E2dimension")
            usrPart.Cdim = .Fields("Cdimension")
            usrPart.C1dim = .Fields("C1dimension")
            usrPart.Ddim = .Fields("Ddimension")
            usrPart.D1dim = .Fields("D1dimension")
            'usrPart.Build = .Fields("Build")
            'usrPart.Status = .Fields("Status")
            usrPart.Priority = .Fields("Priority")
         
            addRun usrPart, intAddOK            'add part to run table
            .Delete
            .Update
            .MoveNext
         Next i
         intResult = 1
         frmRun.RefreshRunQue                   'refresh the run que
         frmRun.RefreshExecQue                  'refresh the run que
      Else                                      'error-should have records
         intResult = 3
      End If
   
   End With
   
   '--- clean up
   adoRS.Close                                  'unload recordset & connection
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing
   
Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  intResult = 3                                 'some kind access error
  
End Sub  'getRunGroup()

'------------------------------------------------------------------------------
'Name:      procSusp
'Accepts:   none
'Returns:   none
'Requires:
'Discrip:   This sub handles a suspend function
'Notes:
'------------------------------------------------------------------------------
Sub procSusp()
'--- globals variables
'--- local variables

   frmRun.cmdSuspend.Visible = True                'reshow button
   frmRun.Shape2.Visible = False                   'hide reminder
   frmRun.Label33.Visible = False
   
   frmRun.optPause.Value = True                    'switch to PAUSE mode
   
   procBatch                                       'process the batch
   
   SUSPEND = False                                 'clear SUSPEND flag
   
End Sub  'procSusp







'---sample header
'------------------------------------------------------------------------------
'Name:      x
'Accepts:   none
'Returns:   none
'Requires:
'Discrip:
'Notes:
'------------------------------------------------------------------------------
'--- globals variables
'--- local variables

