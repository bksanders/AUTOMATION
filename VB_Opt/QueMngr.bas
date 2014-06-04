Attribute VB_Name = "QueMngr"
'---------------------------------------------------------------------------
'Name:      QueMngr.bas
'Accepts:   none
'Returns:   none
'Requires:  none
'Discrip:   code module for managing ques
'Notes:
'---------------------------------------------------------------------------

Public Type SelectItem
   Job As String
   Rel As String
   Item As Integer
   ShipDate As Long
   Mat As String
   Width As String
   Qnty As Integer
   Build As String
End Type
   
Public Type part
   FullJobNum As String
   Job As String
   Rel As String
   Item As Integer
   Seq As Integer
   ShipDate As Long
   SchedDate As Date
   Qnty As Integer
   BldQnty As Integer
   Phase As String
   Leg As String
   Stack As String
   BarType As String
   Material As String
   BarWidth As String
   BlankLength As Double
   RunLength As Double
   E1fig As Double
   E1dim As Double
   E2fig As Double
   E2dim As Double
   Cdim As Double
   C1dim As Double
   Ddim As Double
   D1dim As Double
   Build As String
   Status As String
   Priority As Integer     'used for order of execution & batch groups
   Group As Integer        'not used in mubea bar hmi
   Truck As String
   Machine As String
   DTS As String
End Type

Public Type RecordID       'info needed to uniquely
   Job As String           'identity an item or part record
   Rel As String
   Item As Integer
   Seq As Integer
   Pri As Integer          'a group can be ID'd by its priority while in Exec Que
End Type

'-----* a peice is made up of grouped parts.
'     * there can be up to 3 parts grouped together into a peice.
'     * the piece type can = 0,1,2,3
'           0 = a standard piece, ie. 1 part per piece.
'           1 = a short piece, < 15inches
'           2 = two parts grouped together
'           3 = 3 parts grouped together
Public Type Piece
   Type As Integer      'see above note
   Qty As Integer       'desired qty
   SubQty As Integer    'submitted qty
   bldQTY As Integer    'built qty
   RunLen As Long       'runlength
   part(3) As RecordID  'part ID's
   SubCnt As Integer    '# of times piece has been submitted
End Type

'--- Global Variable Declarations
Global usrSelItem As SelectItem
Global intSelCnt As Integer         '# of items in select que
Global intExecCnt As Integer        '# of parts in Exec Que
Global usrRecID As RecordID
Global strSelMat As String          'select que material
Global strSelWidth As String        'select que width
Global strBatchMat As String        'batch material
Global strBatchWidth As String      'batch bar width
Global arrParts(100) As part        'parts array for submitting to exec que
Global intParts   As Integer        '# of parts in parts array
Global usrRUNParts(20) As part      'Run Listbox Storage
Global intRunCnt As Integer         '# of parts in RunParts array
Global intIndex   As Integer        'index into parts array
Global intHiPri As Integer          'Highest priority in the Exec Que
Global strItem As String
Global usrNULLItem As SelectItem    'empty item
Global usrNULLPiece As Piece        'empty piece
Global usrNULLRecID As RecordID     'empty rec id
Global usrNULLPart As part          'empty part

Global SawKerf As Long              '--- system parameters
Global CutAllow As Long
Global ScrapFactor As Long
Global StockLength As Long
Global arrCoinPress(4) As Long
Global arrSortOrder(6) As String
Global BlankMin As Long
Global BlankMax As Long
Global E1DimMin As Long
Global E1DimMax As Long
Global E2DimMin As Long
Global E2DimMax As Long
Global CDimMin As Long
Global CDimMax As Long
Global C1DimMin As Long
Global C1DimMax As Long
Global DDimMin As Long
Global DDimMax As Long
Global D1DimMin As Long
Global D1DimMax As Long

Global FirstPass As Boolean          'jng
Global dblBatchOpt As Double         'jng 03/31/08-added to keep optimization for history tracking
Global dblTotPartsLen As Double      'jng 03/31/08-added to keep total parts length for history tracking
Global dblBatchLngth As Double       'jng 03/31/08-added to keep total Stock length for history tracking
Global intBatchBlanks As Integer     'jng 3/31/08
Global dblPartLen As Double          'jng
Global intPartQty As Integer         'jng
Global dblLongScrap As Double        'jng
Global intOptType As Integer         'jng

Global dblBatchOpt2 As Double        'jng 03/31/08-added to keep optimization for history tracking
Global dblTotPartsLen2 As Double      'jng 03/31/08-added to keep total parts length for history tracking
Global dblBatchLngth2 As Double      'jng 03/31/08-added to keep total Stock length for history tracking
Global intBatchBlanks2 As Integer     'jng 3/31/08
Global dblPartLen2 As Double          'jng
Global dblLongScrap2 As Double        'jng
Global intOptType2 As Integer         'jng



'---------------------------------------------------------------------------
'Name:      setup()
'Accepts:   none
'Returns:   none
'Requires:  local hsg dbase w/ paramaters table
'Discrip:   This sub loads system paramaters from the paramaters table of
'           the local dbase.
'Notes: (1) paramaters are stored in inches as floating points.  Converted to
'           mills by x 1000.
'       (2) This sub runs once at STARTUP.
'---------------------------------------------------------------------------

Public Sub setup()
   Dim conLocal As Connection
   Dim adoRS As ADODB.Recordset
   Dim BLANK As Boolean                'blank flag
   
   '---------------------- Load System Paramaters --------------------------
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset of paramaters table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblParameters"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount = 1 Then                      'found 1 record
      adoRS.MoveFirst                                 'go to the record
      
      SawKerf = adoRS.Fields("SawKerf") * 1000
      CutAllow = adoRS.Fields("CutAllow") * 1000
      ScrapFactor = adoRS.Fields("ScrapFactor") * 1000
      StockLength = adoRS.Fields("StockLength") * 1000
      frmRun.txtStkLength = StockLength
      BlankMin = adoRS.Fields("BlankMin") * 1000
      BlankMax = adoRS.Fields("BlankMax") * 1000
      E1DimMin = adoRS.Fields("E1DimMin") * 1000
      E1DimMax = adoRS.Fields("E1DimMax") * 1000
      E2DimMin = adoRS.Fields("E2DimMin") * 1000
      E2DimMax = adoRS.Fields("E2DimMax") * 1000
      CDimMin = adoRS.Fields("CDimMin") * 1000
      CDimMax = adoRS.Fields("CDimMax") * 1000
      C1DimMin = adoRS.Fields("C1DimMin") * 1000
      C1DimMax = adoRS.Fields("C1DimMax") * 1000
      DDimMin = adoRS.Fields("DDimMin") * 1000
      DDimMax = adoRS.Fields("DDimMax") * 1000
      D1DimMin = adoRS.Fields("D1DimMin") * 1000
      D1DimMax = adoRS.Fields("D1DimMax") * 1000
      
      blnEnTRACK = adoRS.Fields("EnableTracking")
      'blnEnTRACK = False
      blnEnOPT = adoRS.Fields("EnableOPT")
      blnAutoPICK = adoRS.Fields("AutoPick")
      intFillDays = adoRS.Fields("FSDays")
      intMaxRel = adoRS.Fields("MaxRelease")
         
   Else                                                  'record problem
      MsgBox ("Local DB error: Paramater Table Corrupt!")
   End If
    
   adoRS.Close                                           'unload recordset
   Set adoRS = Nothing
   
   '-------------------------- Set up for Tracking -------------------------
   EnFDRTruck = True
   EnPITruck = False
   EnPIGTruck = False
   
   '--------------------------- ID Blank Counts ------------------------
   '--- check tblBatch
   BLANK = False                                         'reset blank flag
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBatch " & _
            "ORDER BY Priority ASC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   With adoRS
      If .RecordCount > 0 Then                           'batch NOT empty
         .MoveLast                                       'go to last record
         intTotBlanks = .Fields("Priority")              'Hightest prior = total blanks
         BLANK = True                                    'set blank flag
         .MoveFirst                                      'go to 1st record
         intTemp = .Fields("Priority")
         intRemBlanks = intTotBlanks - intTemp + 1       'rem blanks = HI-LO + 1
      Else                                               'batch empty
         intTotBlanks = 0
         intRemBlanks = 0
      End If
      .Close                                             'close recordset
   End With
   Set adoRS = Nothing                                   'unload recordset
   
   '--- check tblRun
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblRun"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   With adoRS
      If .RecordCount > 0 Then                           'RunQue NOT empty
         If BLANK = False Then                           'if tblBatch Empty
            .MoveFirst                                   'go to 1st record
            intTotBlanks = .Fields("Priority")           'Highest prior = total blanks
         End If
         intRemBlanks = intRemBlanks + 1                 'add the blank
      End If
      .Close                                             'close recordset
   End With
   Set adoRS = Nothing                                   'unload recordset
   
   '-------------------------- Set Up for TCP Comms -------------------------
   With frmRun.tcpClient   '------------------------- Setup TCP Client
      If frmRun.optSim.Value = True Then              'connect to simulator
         .RemoteHost = "TNSELMUBEA"
         '.RemoteHost = "ANDERSKEFLDGE"
         .RemotePort = 1001
      Else
         .RemoteHost = "3.94.104.126"                 'connect to Mubea OI
         '.RemoteHost = "ANDERSKEFLDGE"
         .RemotePort = 200
      End If
   
      'If .State <> sckListening Then
      '   .Close
      'End If
     ' .Connect                                        'initiate a tcp connection
   End With
   
   With frmRun.tcpServer   '-------------------------------- Setup TCP Server
      .LocalPort = 201
      .Listen
   End With
   
   
   '-------------------------- Set UP for Execution -------------------------
   STARTUP = True                                     'set STARTUP flag
   NoStopPROMPT = True                                'set flag to skip STOP prompt
  
   frmRun.txtCount = ""                               'reset troublshoot counter
   
   '--- Open capture file for dbugging
   'Open "C:\RemBar\Test\hsgTest.txt" For Append As #1
   'Print #1, "Open capture file" & Now

End Sub  'setup

'---------------------------------------------------------------------------
'Name:      addSel
'Accepts:   usrItem
'Returns:   result = 1 = item added
'                  = 2 = item Not added
'                  = 3 = Material Missmatch
'                  = 4 = Width Mismatch
'                  = 5 = que full
'Requires:  This sub adds an item to the Select Que.
'Discrip:
'Notes: (1) Will not add item if already in que.
'       (2) Will not add item if mat or width mismatch
'       (3) Will not add item if sel que count > 14
'---------------------------------------------------------------------------

Public Sub addSel(usrItem As SelectItem, result As Integer)

Dim conLocal As Connection
Dim adorsSelect As ADODB.Recordset

   '---check for full que
   'select que must me limited to 15 items, or ExecQue size will be to large
   If intSelCnt > 14 Then
      result = 5
      Exit Sub
   End If
   
   '---check for matching MAT & Width
   'once an item is added to the SelQue...all other items must have
   'matching materials and barwidths
   If strSelMat <> "" Then                            'skip if que empty
      If usrItem.Mat <> strSelMat Then                'check Mat
         result = 3
         Exit Sub
      End If
      If usrItem.Width <> strSelWidth Then            'check width
         result = 4
         Exit Sub
      End If
   End If
     
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset of select que
   Set adorsSelect = New ADODB.Recordset
   strSQL = "SELECT * FROM tblSelQue " & _
            "WHERE [Order Number] = '" & usrItem.Job & "'" & _
            "AND Release = '" & usrItem.Rel & "'" & _
            "AND Item = " & usrItem.Item
   
   adorsSelect.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adorsSelect.RecordCount < 1 Then                'if NO duplicate, Add to Que
      With adorsSelect
         .AddNew
         ![Order Number] = usrItem.Job
         !Release = usrItem.Rel
         !Item = usrItem.Item
         ![Scheduled Ship Date] = usrItem.ShipDate
         ![Quantity] = usrItem.Qnty
         ![Mat] = usrItem.Mat
         ![Width] = usrItem.Width
         .Update
         result = 1
      End With
   Else  ' duplicate
      result = 2                                   'signal duplicate
   End If
      
   adorsSelect.Close                               'unload recordset & connection
   Set adorsSelect = Nothing
   conLocal.Close
   Set conLocal = Nothing
   
   frmSelect.RefreshSelectQue                      'update the que display
      
End Sub 'addSel

'---------------------------------------------------------------------------
' This Sub removes a part from the Select Que
'---------------------------------------------------------------------------
Public Sub rmvSel(usrItemID As RecordID, result As Boolean)

Dim conLocal As Connection
Dim adorsSelect As ADODB.Recordset

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for History que
   Set adorsSelect = New ADODB.Recordset
   strSQL = "SELECT * FROM tblSelQue " & _
            "WHERE [Order Number] = '" & usrItemID.Job & "'" & _
            "AND Release = '" & usrItemID.Rel & "'" & _
            "AND Item = " & usrItemID.Item
   adorsSelect.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adorsSelect.RecordCount > 0 Then                'if part found in que
      With adorsSelect
         .MoveFirst
         .Delete                                      'remove it
         .Update
         result = True
      End With
   Else  ' duplicate
      MsgBox ("Item Not found in Select Que.")
   End If
         
   '--- cleanup
   adorsSelect.Close                                  'unload recordset & connection
   Set adorsSelect = Nothing
   conLocal.Close
   Set conLocal = Nothing
   
End Sub  'rmvSel


'------------------------------------------------------------------------------
'Name:      addExec
'Accepts:   none
'Returns:   results integer, intresult = 1 , added OK
'                                      = 2 , already in que
'                                      = 3 , Mat mismatch
'                                      = 4,  Width mismatch
'Requires:
'Discrip:   This Sub adds a part to the Exec Que
'Notes: (1) Will not add part if already in que, or if mat/width mismatch
'       (2) This sub does not determine the parts priority in the exec que.
'           This must be done prior to calling this sub.
'------------------------------------------------------------------------------

Public Sub addExec(usrPart As part, intResult As Integer)
   Dim conLocal As Connection
   Dim adorsExec As ADODB.Recordset
                          
   '----------------------- check mat/width agaist batch ----------------------
   If intExecCnt > 0 Then                             'if ExecQue NOT empty
      If usrPart.Material <> strBatchMat Then         'check material
         intResult = 3
         Exit Sub
      End If
      If usrPart.BarWidth <> strBatchWidth Then       'check width
         intResult = 4
         Exit Sub
      End If
   End If
                                                    
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                          
   '------------------------------- Add the Part --------------------------
   '--- make recordset for Exec que
   Set adorsExec = New ADODB.Recordset
   strSQL = "SELECT * FROM tblExecQue " & _
            "WHERE [Order Number] = '" & usrPart.Job & "' " & _
            "AND Release = '" & usrPart.Rel & "' " & _
            "AND Item = " & usrPart.Item & _
            "AND [Sequence Number] = " & usrPart.Seq
   adorsExec.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adorsExec.RecordCount < 1 Then                  'if NO duplicate, Add to Que
      With adorsExec
        .AddNew                                       'add a record to Exec
        '--- fill out the record
        !FullOrder = usrPart.FullJobNum
        ![Order Number] = usrPart.Job
        !Release = usrPart.Rel
        !Item = usrPart.Item
        ![Sequence Number] = usrPart.Seq
        ![Scheduled Ship Date] = usrPart.ShipDate
        !Quantity = usrPart.Qnty
        !BldQnty = usrPart.BldQnty
        !Phase = usrPart.Phase
        !Leg = usrPart.Leg
        !Stack = usrPart.Stack
        !BarType = usrPart.BarType
        !Material = usrPart.Material
        !BarWidth = usrPart.BarWidth
        !BlankLength = usrPart.BlankLength
        !RunLength = usrPart.RunLength
        !E1figure = usrPart.E1fig
        !E1dimension = usrPart.E1dim
        !E2figure = usrPart.E2fig
        !E2dimension = usrPart.E2dim
        !Cdimension = usrPart.Cdim
        !C1dimension = usrPart.C1dim
        !Ddimension = usrPart.Ddim
        !D1dimension = usrPart.D1dim
        !Build = usrPart.Build
        !Status = usrPart.Status
        !Priority = usrPart.Priority
        !DTStamp = usrPart.DTS
        .Update                                       'update the recordset
      End With
      intResult = 1                                   'successfull result
      OPTIMIZED = False                               'reset optimized flag
   Else  ' duplicate
      'MsgBox ("Part already in Que.")
      intResult = 2
   End If
           
   adorsExec.Close                                    'unload recordset & connection
   Set adorsExec = Nothing
   conLocal.Close
   Set conLocal = Nothing
   frmRun.RefreshExecQue                              'refresh the Que
   
End Sub 'addExec


'---------------------------------------------------------------------------
' This Sub removes a part from the Exec Que

Public Sub rmvExec(usrPartID As RecordID, result As Integer)

Dim conLocal As Connection
Dim adorsExec As ADODB.Recordset

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for History que
   Set adorsExec = New ADODB.Recordset
   strSQL = "SELECT * FROM tblExecQue " & _
            "WHERE [Order Number] = '" & usrPartID.Job & "'" & _
            "AND Release = '" & usrPartID.Rel & "'" & _
            "AND Item = " & usrPartID.Item & _
            "AND [Sequence Number] = " & usrPartID.Seq
   adorsExec.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adorsExec.RecordCount > 0 Then                  'if part found in que
      With adorsExec
         .MoveFirst
         .Delete                                      'remove it
         .Update
      End With
      result = 1
   Else  'part not in Que
      MsgBox ("Item Not found in Exec Que.")
   End If
         
   OPTIMIZED = False                                  'reset optimize flag
         
   adorsExec.Close                                    'unload recordset & connection
   Set adorsExec = Nothing
   conLocal.Close
   Set conLocal = Nothing
   
   frmRun.RefreshExecQue                              'refresh the Exec Que
   frmRun.RefreshRunQue                               'refresh the Run Que
   
End Sub  'rmvExec

'---------------------------------------------------------------------------
' This Sub adds a part to the History Que

Public Sub addHist(usrPart As part, intResult As Integer)

Dim conLocal As Connection
Dim adorsHist As ADODB.Recordset

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for History que
   Set adorsHist = New ADODB.Recordset
   strSQL = "SELECT * FROM tblHist " & _
            "WHERE [Order Number] = '" & usrPart.Job & "'" & _
            "AND Release = '" & usrPart.Rel & "'" & _
            "AND Item = " & usrPart.Item & " " & _
            "AND [Sequence Number] = " & usrPart.Seq
   adorsHist.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adorsHist
      .AddNew                                   'add a record to HIST
      '--- fill out the record
      !FullOrder = usrPart.FullJobNum
      ![Order Number] = usrPart.Job
      !Release = usrPart.Rel
      !Item = usrPart.Item
      ![Sequence Number] = usrPart.Seq
      ![Scheduled Ship Date] = usrPart.ShipDate
      ![Quantity] = usrPart.Qnty
      !BldQnty = usrPart.BldQnty
      !Phase = usrPart.Phase
      !Leg = usrPart.Leg
      !Stack = usrPart.Stack
      !BarType = usrPart.BarType
      !Material = usrPart.Material
      !BarWidth = usrPart.BarWidth
      !BlankLength = usrPart.BlankLength
      !RunLength = usrPart.RunLength
      !E1figure = usrPart.E1fig
      !E1dimension = usrPart.E1dim
      !E2figure = usrPart.E2fig
      !E2dimension = usrPart.E2dim
      !Cdimension = usrPart.Cdim
      !C1dimension = usrPart.C1dim
      !Ddimension = usrPart.Ddim
      !D1dimension = usrPart.D1dim
      !Build = usrPart.Build
      !Status = usrPart.Status
      !Priority = usrPart.Priority
      !Truck = "Test"
      !Machine = usrPart.Machine
      !DTStamp = usrPart.DTS
      !BatchOpt = dblBatchOpt   'jng 03/31/08-added to keep optimization for history tracking
      !BatchTotPartLen = dblTotPartsLen   'jng 03/31/08-added to keep total parts length for history tracking
      !BatchTotLen = dblBatchLngth     'jng 03/31/08-added to keep total Stock length for history tracking
      !BatchTotBlanks = intBatchBlanks     'jng 3/31/08 -Total blanks in batch
      !BatchStockLen = StockLength       'jng 3/31/08 -Blanks Stocklength
      !LongScrap = dblLongScrap         'jng
      !BatchType = intOptType           'jng
      .Update                                   'update the recordset
   End With
         
   '--- clean up
   adorsHist.Close                              'unload recordset & connection
   Set adorsHist = Nothing
   conLocal.Close
   Set conLocal = Nothing
   frmHist.RefreshHistQue                       'refresh the Hist Que
   intResult = 1                                'return result
   
End Sub  'addHist

'--- this sub builds a parts array for an item

Public Sub bldPartsArray(usrItem As SelectItem, intOption As Integer)
   Dim conCamdata As Connection
   Dim adorsCamdata As ADODB.Recordset
   Dim strSQL As String
   Dim strTemp As String
   Dim i As Integer

   '---------------------------- Locate Parts in Camdata ----------------------
   '--- make connection to camdata database
   Set conCamdata = New Connection
   conCamdata.Open "PROVIDER=MSDASQL;dsn=dsnMBCamdata;uid=;pwd=;"
                                 
   '--- make recordset of all parts in this item
   Set adorsCamdata = New ADODB.Recordset
   If usrItem.Rel = "000" Then
      strSQL = "SELECT * FROM mubbarff " & _
               "WHERE ([Order Number] = '" & usrItem.Job & "' " & _
               "AND Release Is Null " & _
               "AND Item = " & usrItem.Item & ") "
   Else
      strSQL = "SELECT * FROM mubbarff " & _
               "WHERE ([Order Number] = '" & usrItem.Job & "' " & _
               "AND Release = '" & usrItem.Rel & "' " & _
               "AND Item = " & usrItem.Item & ") "
   End If
   
   Select Case intOption   'get only those parts designated by the build option
   Case 1      'full build w/o grounds
      strSQL = strSQL & "AND Phase <> '3' AND Phase <> '4' "
   Case 2      'full build grounds only
      strSQL = strSQL & "AND (Phase = '3' OR Phase = '4') "
   End Select
   
   strSQL = strSQL & "ORDER BY Stack ASC"
   
   adorsCamdata.Open strSQL, conCamdata, adOpenStatic, adLockOptimistic


   '----------------------------- Create Parts Array --------------------------
   intParts = adorsCamdata.RecordCount             'get the number of parts in item
   
   If intParts > 0 Then                            'if parts, generate the array
      adorsCamdata.MoveFirst                       'start loop at 1st part
      i = 1
      Do Until adorsCamdata.EOF                    'loop thru parts in item
         arrParts(i).FullJobNum = adorsCamdata.Fields("FullOrder")
         arrParts(i).Job = adorsCamdata.Fields("Order Number")
         If IsNull(adorsCamdata.Fields("Release")) Then
            arrParts(i).Rel = "000"
         Else
            arrParts(i).Rel = adorsCamdata.Fields("Release")
         End If
         arrParts(i).Item = adorsCamdata.Fields("Item")
         arrParts(i).Seq = adorsCamdata.Fields("Sequence Number")
         arrParts(i).ShipDate = adorsCamdata.Fields("Scheduled Ship Date")
         arrParts(i).Qnty = adorsCamdata.Fields("Quantity")
         arrParts(i).BldQnty = 0
         arrParts(i).Phase = adorsCamdata.Fields("Phase")
         arrParts(i).Leg = adorsCamdata.Fields("Leg")
         If IsNull(adorsCamdata.Fields("Stack")) Then
            arrParts(i).Stack = "n"
         Else
            arrParts(i).Stack = adorsCamdata.Fields("Stack")
         End If
         arrParts(i).BarType = adorsCamdata.Fields("BarType")
         arrParts(i).Material = adorsCamdata.Fields("Material")
         strTemp = adorsCamdata.Fields("BarWidth")
         arrParts(i).BarWidth = getWidth(strTemp)
         arrParts(i).BlankLength = adorsCamdata.Fields("BlankLength")
         arrParts(i).RunLength = -1
         arrParts(i).E1fig = adorsCamdata.Fields("E1figure")
         arrParts(i).E1dim = adorsCamdata.Fields("E1dimension")
         arrParts(i).E2fig = adorsCamdata.Fields("E2figure")
         arrParts(i).E2dim = adorsCamdata.Fields("E2dimension")
         arrParts(i).Cdim = adorsCamdata.Fields("Cdimension")
         arrParts(i).C1dim = adorsCamdata.Fields("C1dimension")
         arrParts(i).Ddim = adorsCamdata.Fields("Ddimension")
         arrParts(i).D1dim = adorsCamdata.Fields("D1dimension")
         arrParts(i).Build = usrItem.Build
         arrParts(i).Status = "  "
         arrParts(i).Priority = -1
         arrParts(i).Group = -1
         arrParts(i).DTS = "               "
         adorsCamdata.MoveNext
         i = i + 1
      Loop ' end of parts loop
   End If
   
   '--- close conn to camdata
   adorsCamdata.Close                              'close recordset
   Set adorsCamdata = Nothing                      'unload recordset
   conCamdata.Close                                'close connection
   Set conCamdata = Nothing                        'unload connection
    
End Sub  'bldPartsArray()
  
'---------------------------------------------------------------------------
'Name:      purgePartsArray()
'Accepts:   none
'Returns:   none
'Requires:  arrParts(i) = parts array
'Discrip:   This sub purges the parts array of any parts w/ Qty = 0
'Notes:
'---------------------------------------------------------------------------
Public Sub purgePartsArray()
   Dim ZERO As Boolean        'zero qty flag
   Dim i, j As Integer        'loop index
   
   '--- remove zero qty's from array
   Do
      ZERO = False
      For i = 1 To intParts                           'loop thru parts
         If arrParts(i).Qnty < 1 Then                 'Zero Qty found
            If i < intParts Then                      'NOT last part
               For j = i To intParts - 1              'shift parts down
                  arrParts(j) = arrParts(j + 1)
               Next j
               ZERO = True                            'zero part found
            End If
            intParts = intParts - 1                   'decr part count
            Exit For
         End If  'zero qty
      Next i
   Loop Until ZERO = False

End Sub  'purgePartsArray()
 
'---------------------------------------------------------------------------
'Name:      subItemF
'Accepts:   usrItem = item to be submited
'Returns:   result = 1 if item submitted OK
'Requires:  Listed subs
'Discrip:   This sub submits an item to the Exec Que as a FULL build.
'Notes:
'---------------------------------------------------------------------------
Public Sub subItemF(usrItem As SelectItem, intBldOpt As Integer, result As Integer)
 
   bldPartsArray usrItem, intBldOpt                'build the parts array
   prioritize                                      'call prioritize routine
   InsertParts                                     'Insert Parts into Exec Que

   'possible location for updating machine table assignments

   result = 1
  
End Sub  'subItemF

'---------------------------------------------------------------------------
'Name:      subItemPR
'Accepts:   none
'Returns:   result = 1 if item submitted OK
'Requires:  Listed subs, valid parts array
'Discrip:   This sub submits an item to the Exec Que as a PARTIAL or REMAKE.
'Notes:     In this case, the parts array has already been manually built
'           Therefore, No need to send an Item...just use the global parts
'           array.
'---------------------------------------------------------------------------
Public Sub subItemPR(result As Integer)
      
   purgePartsArray                                 'remove parts w/ .qty=0
   prioritize                                      'call prioritize routine
   InsertParts                                     'Insert Parts into Exec Que

   'possible loca for
   
   result = 1
  
End Sub  'subItemPR


'----------------------------------------------------------------------
'This sub inserts all the parts in the parts array into the Exec Que
'It 1st determines the highest priority# in the exec que.
'Then presets each parts priority so it will takes its proper place
'(per priority) in the Exec que.  The correct priority will be the
'Exec que's base priority + the parts current priority w/in the item.
'Note:  The priority for parts in the item must already be set when this
'sub is called.
'----------------------------------------------------------------------
Public Sub InsertParts()
   Dim conLocal As Connection
   Dim adorsExec As ADODB.Recordset
   Dim strSQL As String
   Dim intExecPrior As Integer                  'Exec Que's priority
   Dim i As Integer
   Dim intAddOK As Integer
   
   '------------------------- Determine Priority -----------------------
   
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                             
   '--- make recordset of Exec Que
   Set adorsExec = New ADODB.Recordset
   strSQL = "SELECT * FROM tblExecQue "
   adorsExec.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
 
   '--- determine the highest priority number in the Que
   intExecPrior = 0
   If adorsExec.RecordCount > 0 Then               'if parts in que
      adorsExec.MoveFirst
      Do While Not adorsExec.EOF                   'loop thru Exec Que
         If adorsExec.Fields("Priority") > intExecPrior Then
            intExecPrior = adorsExec.Fields("Priority")
         End If
         adorsExec.MoveNext
      Loop
   End If
                
   adorsExec.Close                                 'unload recordset & connection
   Set adorsExec = Nothing
   conLocal.Close
   Set conLocal = Nothing
                         
   '--------------------------- Insert the Parts ---------------------------
   For i = 1 To intParts                           'loop thru parts array
      If arrParts(i).Qnty > 0 Then                 'only submit if qty >0
         arrParts(i).Priority = intExecPrior + arrParts(i).Priority     'set the priority
         arrParts(i).Status = "EQ"                 'set part status
         addExec arrParts(i), intAddOK
         If intAddOK = 1 Then                      'part added to exec
            'If the part was successfully added to the Exec Que and
            'optimization is enabled...then we must capture this part
            'to update its assignment in the machine table. NO need to
            'assign remakes.
            If blnEnOPT And arrParts(i).Build <> "R" Then
               intAddOK = 0
               addHoldAssign arrParts(i), intAddOK
               intAsgnCnt = intAsgnCnt + 1
            End If
         Else                                      'part NOT added to exec
            'MsgBox ("Part #" & i & "NOT added to Exec Que!")
         End If
      End If
   Next i  'end of parts loop
   
   frmRun.RefreshExecQue
   
End Sub  'InsertParts()

'-------------------------------------------------------------------------
'This subroutine works with the the global parts array.
'The runlengths of all parts in the array should have been init'd to -1
'After this sub processes the parts array, the parts array will be grouped.
'Grouped parts will have the same priority.
'Runlengths for each part/group will be calculated.  The runlengths for
'grouped parts will be the same.
'-------------------------------------------------------------------------

Public Sub Group()
   Dim i As Integer              'master part index
   Dim j As Integer              'compare index
   Dim intPartsDone As Integer   'parts done counter
   Dim intGrpCntr As Integer     'group counter, uniquely #'s groups
   Dim blnMaster As Boolean      'true if master part located
   Dim NoGroup As Boolean        'flag = true when unable to group current master
   Dim shortPart As Boolean      'true if at least 1 of the parts < 15"
   Dim intMatch As Integer       '# of parts that could group w/ the current master
   Dim intLastMatch As Integer   'index # of the last matching found
   Dim tmpRunLength As Long      'temporary run legth holder
   Dim X As Boolean              'x leg flag
   Dim Y As Boolean              'y leg flag
   Dim Z As Boolean              'z leg flag
   Dim idx As Integer            'x leg index
   Dim idY As Integer            'y leg index
   Dim idZ As Integer            'z leg index
   Dim strSTK As String          'special case string flag
   Dim blnTemp As Boolean        'temporary boolean variable
   
   '------------------------ Check for Bar Type T or S ------------------------
   'Only group Types T and S
   If arrParts(1).BarType = "T" Or arrParts(1).BarType = "S" Then
   
      '------------------------- Process Parts -----------------------------
      intPartsDone = 0                                'init the parts done counter
      intGrpCntr = 0                                  'init the priority counter
      shortPart = False                               'init short part flag
      Do    'loop until all parts in array are processed
                  
         '--- locate a master part
         'The master part is the 1st unprocessed part, ie runlength = -1
         blnMaster = False
         For i = 1 To intParts                        'loop thru the array
            If arrParts(i).RunLength = -1 Then        'NOT processed
            blnMaster = True
            Exit For
            End If
         Next i   'upon exiting this loop i will point to the master part
         
         If arrParts(i).BlankLength < 15000 Then      'check for short part
            shortPart = True
         End If
         
         If Not blnMaster Then                        'NO more to process
            MsgBox ("Group:Could Not Locate a master part!")
            Exit Sub    'no need to continue everything processed
         End If
                      
         '--- increment the priority counter
         intGrpCntr = intGrpCntr + 1                  'incr the priority
            
         '------------------------- find match count --------------------------
         'determine the number of parts that could group w/ the current master part
         intMatch = 1                                 'account for the master
         strSTK = ""                                  'reset special stack flag
         For j = 1 To intParts
            If arrParts(j).RunLength = -1 And _
               arrParts(j).Phase = arrParts(i).Phase And _
               arrParts(j).Material = arrParts(i).Material And _
               arrParts(j).BarWidth = arrParts(i).BarWidth And _
               arrParts(j).BarType = arrParts(i).BarType And _
               arrParts(j).Stack = arrParts(i).Stack And _
               arrParts(j).Leg <> arrParts(i).Leg Then
               intMatch = intMatch + 1                'incr the match count
               intLastMatch = j                       'hang on the match index
               If arrParts(j).BlankLength < 15000 Then      'check for short part
                  shortPart = True
               End If
            End If
         Next j   'end of match loop
           
         '---------------- handle grouping based on the Bar Type  -------------
         NoGroup = False                              'reset grouping flag
         Select Case arrParts(i).BarType
         
         Case "T"    '--------------- Grouping for BarType T ------------------
            'only group a barType T when at least 1 part is short(< 15")
            If shortPart = True Then
               Select Case intMatch                      'distribut based on # of matches
               
               Case 1   '--- only 1 part no grouping
                  NoGroup = True
                  
               Case 2   '--- 2 parts to group
                  j = intLastMatch                       '--- locate the matched part
                  
                  '--- determine leg type of master part
                  X = False                              'init leg flags
                  Y = False
                  Z = False
                  Select Case arrParts(i).Leg
                  Case "X"
                     X = True
                  Case "Y"
                     Y = True
                  Case "Z"
                     Z = True
                  End Select
                  
                  '--- determine leg type of the match part
                  Select Case arrParts(j).Leg
                  Case "X"
                     X = True
                  Case "Y"
                     Y = True
                  Case "Z"
                     Z = True
                  End Select
                  
                  If (X And Y) Or (Y And Z) Then         'can group an XY or YZ
                     '---calc the combined run length
                     tmpRunLength = arrParts(i).BlankLength + _
                                    arrParts(j).BlankLength + _
                                    SawKerf
                  
                     If tmpRunLength < 15000 Then        'check for min runlength
                        tmpRunLength = 15000
                     End If
                     
                     If tmpRunLength < StockLength Then        'combined length OK, so group
                        arrParts(i).RunLength = tmpRunLength   'do master part
                        arrParts(i).Group = intGrpCntr
                        arrParts(j).RunLength = tmpRunLength   'do matched part
                        arrParts(j).Group = intGrpCntr
                        intPartsDone = intPartsDone + 2        'processed 2 parts
                     Else     'length to long to group
                        NoGroup = True
                     End If
                  Else  '--- can't group these 2
                     NoGroup = True
                  End If
                  
               Case 3   '--- 3 parts to group
                  '--- if the current master is NOT an X leg find the X leg and make it the master
                  If arrParts(i).Leg = "X" Then                'master is an X leg
                     idx = i
                  Else                                         'master NOT an X
                     For j = 1 To intParts                     'find the X
                        If arrParts(j).Phase = arrParts(i).Phase And _
                           arrParts(j).Stack = arrParts(i).Stack And _
                           arrParts(j).Material = arrParts(i).Material And _
                           arrParts(j).BarWidth = arrParts(i).BarWidth And _
                           arrParts(j).BarType = arrParts(i).BarType And _
                           arrParts(j).Leg = "X" Then
                           idx = j                             'make the X leg the master
                           Exit For
                        End If
                     Next j   'end of X search
                  End If
               
                  '--- locate the y leg
                  For j = 1 To intParts                        'find the Y
                     If arrParts(j).Phase = arrParts(idx).Phase And _
                        arrParts(j).Stack = arrParts(idx).Stack And _
                        arrParts(j).Material = arrParts(idx).Material And _
                        arrParts(j).BarWidth = arrParts(idx).BarWidth And _
                        arrParts(j).BarType = arrParts(idx).BarType And _
                        arrParts(j).Leg = "Y" Then
                        idY = j                                'locate Y
                        Exit For
                     End If
                  Next j   'end of y search loop
                             
                  '--- check the combined xy
                  '---calc the combined run length
                  tmpRunLength = arrParts(idx).BlankLength + _
                                 arrParts(idY).BlankLength + _
                                 SawKerf
                  
                  If tmpRunLength < StockLength Then           'combined length OK, so group
                     arrParts(idx).RunLength = tmpRunLength    'do master part = X
                     arrParts(idx).Group = intGrpCntr
                     arrParts(idY).RunLength = tmpRunLength    'do matched part = Y
                     arrParts(idY).Group = intGrpCntr
                     intPartsDone = intPartsDone + 2           'processed 2 parts
                  Else     'length to long to group, can't group the X & Y
                     NoGroup = True
                  End If
                  
                  If NoGroup = False Then    'grouped XY, now check the Z
                     '--- locate the z leg
                     For j = 1 To intParts
                        If arrParts(j).Phase = arrParts(i).Phase And _
                           arrParts(j).Stack = arrParts(i).Stack And _
                           arrParts(j).Material = arrParts(i).Material And _
                           arrParts(j).BarWidth = arrParts(i).BarWidth And _
                           arrParts(j).BarType = arrParts(i).BarType And _
                           arrParts(j).Leg = "Z" Then
                           idZ = j                                'locate Z
                           Exit For
                        End If
                     Next j   'end of z search loop
                                
                     '--- check the combined xy-z length
                     '---calc the combined run length
                     tmpRunLength = arrParts(idx).RunLength + _
                                    arrParts(idZ).BlankLength + _
                                    SawKerf
                     If tmpRunLength < StockLength Then           'combined length OK, so group
                        arrParts(idx).RunLength = tmpRunLength    'update the X
                        arrParts(idx).Group = intGrpCntr
                        arrParts(idY).RunLength = tmpRunLength    'update the Y
                        arrParts(idY).Group = intGrpCntr
                        arrParts(idZ).RunLength = tmpRunLength    'update the Z
                        arrParts(idZ).Group = intGrpCntr
                        intPartsDone = intPartsDone + 1           'processed the 3rd part
                     End If   'length to long to group, can't group XY w/ Z
                             
                  End If   'check Z
                     
               Case Else
                  MsgBox ("group:To many parts matched for item!")
               
               End Select  'Type T match select
               
            Else                                            'no short part
               NoGroup = True                               'so don't group
            End If   'short part
            
         Case "S" '----------------- Grouping for BarType S -------------------
            If intMatch > 1 Then
               idx = 0                                      'reset index pointers
               idY = 0
               
               '--- if the current master is NOT an X leg find the X leg and
               '    make it the master
               If arrParts(i).Leg = "X" Then                'master is an X leg
                  idx = i
               Else                                         'master NOT an X
                  For j = 1 To intParts                     'find the X
                     If arrParts(j).RunLength = -1 And _
                        arrParts(j).Phase = arrParts(i).Phase And _
                        arrParts(j).Stack = arrParts(i).Stack And _
                        arrParts(j).Material = arrParts(i).Material And _
                        arrParts(j).BarWidth = arrParts(i).BarWidth And _
                        arrParts(j).BarType = arrParts(i).BarType And _
                        arrParts(j).Leg = "X" Then
                        idx = j                             'make the X leg the master
                        Exit For
                     End If
                  Next j   'end of X search
               End If
            
               '--- locate the y leg
               For j = 1 To intParts                        'find the Y
                  If arrParts(j).RunLength = -1 And _
                     arrParts(j).Phase = arrParts(idx).Phase And _
                     arrParts(j).Stack = arrParts(idx).Stack And _
                     arrParts(j).Material = arrParts(idx).Material And _
                     arrParts(j).BarWidth = arrParts(idx).BarWidth And _
                     arrParts(j).BarType = arrParts(idx).BarType And _
                     arrParts(j).Leg = "Y" Then
                     idY = j                                'locate Y
                     Exit For
                  End If
               Next j   'end of y search loop
                          
               '--- check the combined xy
               If idx * idY > 0 Then                        'x & y found
                  '---calc the combined run length
                  tmpRunLength = arrParts(idx).BlankLength + _
                                 arrParts(idY).BlankLength
                  If tmpRunLength < StockLength Then           'combined length OK, so group
                     arrParts(idx).RunLength = tmpRunLength    'do master part = X
                     arrParts(idx).Group = intGrpCntr
                     arrParts(idY).RunLength = tmpRunLength    'do matched part = Y
                     arrParts(idY).Group = intGrpCntr
                     intPartsDone = intPartsDone + 2           'processed 2 parts
                  Else     'length to long to group, can't group the X & Y
                     NoGroup = True
                  End If 'runlength chec
               Else
                   NoGroup = True
               End If   'x & y found
            Else                                      'only 1 part no grouping
               NoGroup = True
            End If
         
         End Select  'Select on Bartype
                             
         '-------------------------- master not grouped -----------------------
         If NoGroup Then                                 'nothing grouped w/ Master
            If arrParts(i).BlankLength > 15000 Then      'if blank length is > 15" us it
               arrParts(i).RunLength = arrParts(i).BlankLength
            Else
               arrParts(i).RunLength = 15000
            End If
            intPartsDone = intPartsDone + 1
            arrParts(i).Group = intGrpCntr
         End If
            
      Loop Until intPartsDone = intParts     'all parts processed
   
   Else  '----------------------- NOT a Type T or S ---------------------------
      '--- loop thru all the parts
      For i = 1 To intParts
         If arrParts(i).BlankLength > 15000 Then      'if blank length is > 15" us it
            arrParts(i).RunLength = arrParts(i).BlankLength
         Else
            arrParts(i).RunLength = 15000             'if <15" then use 15"
         End If
         arrParts(i).Group = i                        'set the group# for the parts
      Next i
      
   End If  'end of type IF
   
End Sub  'group

'----------------------------------------------------------------------------
'This sub works w/ the parts array.  This sub pre-sorts the parts array by phase
'per the designator sort order contained in the local dbase. Once the array has
'been ordered by phase... this routine then assigns a priority to the
'parts/groups.
'-----------------------------------------------------------------------------
Public Sub sort()
Dim i, j As Integer                             'array indexes
Dim intGrpPrior As Integer                      'group priority counter
Dim usrTempPart As part
Dim thisGroup As Integer                        'group# of current part
Dim intPartsDone As Integer                     '# of prioritized parts
Dim idx As Integer                              'pointer into parts array

'--- setup
intPartsDone = 0                                'init parts done counter
intGrpPrior = 0                                 'init group priority counter
idx = 1                                         'init array index

'------------------------- sort by custom phase order -------------------------
'This is accomplished using a bubble sort routine.  After this routine executes
'the parts array will be arranged by phase per the sort order.  This routine
'maintains groups established in the group sub.  These groups will
'also remain together in the array due to their equivelent phase.  This sub
'will NOT maintain the original order(from DB) of the parts.
For i = 1 To 6                                  'loop thru each possible phase
   'If (idx + 1) <= intParts Then                'check if idx is out of range
      For j = idx To intParts                   'loop thru parts
         If arrParts(j).Phase = arrSortOrder(i) Then
            If idx <> j Then                    'skip swap if idx = j
               usrTempPart = arrParts(j)        'swap part j & idx
               arrParts(j) = arrParts(idx)
               arrParts(idx) = usrTempPart
            End If   'idx = j
            idx = idx + 1                       'incr the index pointer
         End If   'phase match
      Next j   'part loop
   'End If   'idx OOR
Next i   'phase loop

'-------------------------- Prioritize the Array -----------------------------
'loop thru the unprioritized parts, prioritize them by groups, end loop when
'all parts in array are prioritized.
Do                                                    'loop thru unprioritized parts
   '--- locate 1st unprioritized part
   For i = 1 To intParts
      If arrParts(i).Priority = -1 Then
         Exit For
      End If
   Next i   'i will index the 1st unprioritized part in the array
 
   '--- increment the priority #
   intGrpPrior = intGrpPrior + 1                   'incr priority
   
   '--- loop thru all parts in array, if the part is in the current group
   '    then set its prior to the current priority count
   thisGroup = arrParts(i).Group                   'get current group number
   For j = 1 To intParts                           'loop thru all parts in array
      If arrParts(j).Group = thisGroup Then        'locate parts in this group
         arrParts(j).Priority = intGrpPrior        'set priority
         intPartsDone = intPartsDone + 1
      End If
   Next j
   
Loop Until intPartsDone = intParts                    'all parts prioritized

End Sub  'sort

'---------------------------------------------------------------------------
'Name:      getWidth
'Accepts:   strCode = bar width code string
'Returns:   the barwidth string = barwidth * 1000
'Requires:  width table
'Discrip:   The function looks up the actual barwidth from the barwidth table
'           and returns the barwidth as a string.
'Notes:
'---------------------------------------------------------------------------
Public Function getWidth(strCode As String) As String
   '--- globals used
   '--- variable declarations
   Dim conLocal As Connection
   Dim adorsWidth As ADODB.Recordset
   Dim strSQL As String
   Dim dblTemp As Double               'temp float for calcs
   Dim strWidth As String              'width string, len = 6
   
   '------------------------------------------------------------------------
   
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                             
   '--- make recordset of Width Table
   Set adorsWidth = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBarWidth " & _
            "WHERE WidthCode = '" & strCode & "'"
   adorsWidth.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
  
   '--- get the width
   If adorsWidth.RecordCount > 0 Then                 'if width located
      adorsWidth.MoveFirst                            'go to 1st record
      dblTemp = adorsWidth.Fields("BarWidth")         'get barwidth
      strWidth = Str(dblTemp * 1000)
   Else                                               'NO width match found
      MsgBox ("bldPartStr: Unable to locate BarWidth info!")
      strWidth = ""
   End If   'width if
                
   '--- unload recordset & connection
   adorsWidth.Close
   Set adorsWidth = Nothing
   conLocal.Close
   Set conLocal = Nothing

   getWidth = strWidth                                'return the width string
      
End Function   'getWidth
'---------------------------------------------------------------------------
'Name:      getWdthCode
'Accepts:   Width = bar width (double)
'Returns:   the bar width code
'Requires:  width table
'Discrip:   The function looks up the barwidth code from the barwidth table
'           and returns the barwidth code as a string.
'Notes:
'---------------------------------------------------------------------------
Public Function getWdthCode(Width As Double) As String
   '--- globals used
   '--- variable declarations
   Dim conLocal As Connection
   Dim adorsWidth As ADODB.Recordset
   Dim strSQL As String
   Dim dblTemp As Double               'temp float for calcs
   Dim strCode As String               'width code
   
   '------------------------------------------------------------------------
   
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                             
   '--- make recordset of Width Table
   Set adorsWidth = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBarWidth " & _
            "WHERE BarWidth = " & Width
   adorsWidth.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
  
   '--- get the width
   If adorsWidth.RecordCount > 0 Then                 'if width located
      adorsWidth.MoveFirst                            'go to 1st record
      strCode = adorsWidth.Fields("WidthCode")        'get barwidth code
   Else                                               'NO width match found
      MsgBox ("bldPartStr: Unable to locate BarWidth info!")
      strCode = ""
   End If   'width if
                
   '--- unload recordset & connection
   adorsWidth.Close
   Set adorsWidth = Nothing
   conLocal.Close
   Set conLocal = Nothing

   getWdthCode = strCode                             'return the width string
      
End Function   'getWdthCode




'---------------------------------------------------------------------------
'Name:      prioritize()
'Accepts:   none
'Returns:   none
'Requires:  global parts array
'Discrip:   This sub prioritized the parts array.  This sub assigns a parts
'           priority = default priority = to its position in the array.
'Notes: (1) the build parts array sets the priority of all parts to -1, which
'           was utilized by the Remmele grouping subroutine.  The default
'           priority could have been done in the build parts array...however,
'           when parts are submitted as partial/remake...the build parts
'           array is not used.
'---------------------------------------------------------------------------

Public Sub prioritize()

Dim i As Integer           'loop counter

   For i = 1 To intParts                        'loop thru all parts in array
      arrParts(i).Priority = i                  'set priority
   Next i

End Sub

'---------------------------------------------------------------------------
'Name:      bldBatch()
'Accepts:   none
'Returns:   none
'Requires:  Run Screen
'Discrip:   This sub generates a raw batch table from the exec que.
'Notes:
'---------------------------------------------------------------------------
Public Sub bldBatch()

Dim i As Integer           'loop index
Dim usrBatch As part       'batch entry variable
Dim addOK As Boolean       'add result
Dim intQnty As Integer     'quantity to add to batch

With frmRun.AdodcExec.Recordset
   .MoveFirst
   If Not .EOF Then       'do nothing if ExecQue empty
      '--- loop thru recordset for Exec table
      Do Until .EOF
         '--- build batch entry data
         'usrBatch.FullJobNum = .Fields("FullOrder")
         usrBatch.Job = .Fields("Order Number")
         usrBatch.Rel = .Fields("Release")
         usrBatch.Item = .Fields("Item")
         usrBatch.Seq = .Fields("Sequence Number")
         'usrBatch.ShipDate = .Fields("Scheduled Ship Date")
         usrBatch.Qnty = .Fields("Quantity")
         usrBatch.BldQnty = .Fields("BldQnty")
         usrBatch.Phase = .Fields("Phase")
         usrBatch.Leg = .Fields("Leg")
         usrBatch.Stack = .Fields("Stack")
         usrBatch.BarType = .Fields("BarType")
         usrBatch.Material = .Fields("Material")
         usrBatch.BarWidth = .Fields("BarWidth")
         usrBatch.BlankLength = .Fields("BlankLength")
         'usrBatch.RunLength = .Fields("RunLength")
         usrBatch.E1fig = .Fields("E1figure")
         usrBatch.E1dim = .Fields("E1dimension")
         usrBatch.E2fig = .Fields("E2figure")
         usrBatch.E2dim = .Fields("E2dimension")
         usrBatch.Cdim = .Fields("Cdimension")
         usrBatch.C1dim = .Fields("C1dimension")
         usrBatch.Ddim = .Fields("Ddimension")
         usrBatch.D1dim = .Fields("D1dimension")
         'usrBatch.Build = .Fields("Build")
         'usrBatch.Status = .Fields("Status")
         usrBatch.Priority = 0
         
         '--- generate batch entry's
         intQnty = usrBatch.Qnty - usrBatch.BldQnty
         For i = 1 To intQnty
            addOK = False
            addBatch usrBatch, addOK
            If addOK = False Then
               MsgBox ("bldBatch: An error occured when Adding a record " & _
                      Chr$(13) & "to the Batch Table. Optimization aborted!")
               Exit Sub
            End If
         Next i
         .MoveNext                              'incr for next part
      Loop 'end of recordset loop
   End If 'end of empty que if
End With

End Sub  'bldBatch()

'---------------------------------------------------------------------------
'Name:      addBatch()
'Accepts:
'Returns:   a boolean, result = true if added ok
'Requires:  local db
'Discrip:   this sub adds an entry into the batch table.
'Notes:
'---------------------------------------------------------------------------
Public Sub addBatch(usrPart As part, result As Boolean)

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset

On Error GoTo errorHandler

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBatch "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      .AddNew                                   'add a record to tblBatch
      '--- fill out the record
      '!FullOrder = usrPart.FullJobNum
      ![Order Number] = usrPart.Job
      !Release = usrPart.Rel
      !Item = usrPart.Item
      ![Sequence Number] = usrPart.Seq
      '![Scheduled Ship Date] = usrPart.ShipDate
      '![Quantity] = usrPart.Qnty
      '!BldQnty = usrPart.BldQnty
      !Phase = usrPart.Phase
      !Leg = usrPart.Leg
      !Stack = usrPart.Stack
      !BarType = usrPart.BarType
      !Material = usrPart.Material
      !BarWidth = usrPart.BarWidth
      !BlankLength = usrPart.BlankLength
      '!RunLength = usrPart.RunLength
      !E1figure = usrPart.E1fig
      !E1dimension = usrPart.E1dim
      !E2figure = usrPart.E2fig
      !E2dimension = usrPart.E2dim
      !Cdimension = usrPart.Cdim
      !C1dimension = usrPart.C1dim
      !Ddimension = usrPart.Ddim
      !D1dimension = usrPart.D1dim
      '!Build = usrPart.Build
      '!Status = usrPart.Status
      !Priority = usrPart.Priority
      '!DTStamp = usrPart.DTS
      .Update                                   'update the recordset
   End With
         
   '--- clean up
   adoRS.Close                                  'unload recordset & connection
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing
   
   result = True                                'return result

Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  result = False
  
End Sub  'addBatch()

'---------------------------------------------------------------------------
'Name:      clrBatch()
'Accepts:
'Returns:   a boolean, result = true if added ok
'Requires:  local db
'Discrip:   this sub emties the batch table.
'Notes:
'---------------------------------------------------------------------------
Public Sub clrBatch(result As Boolean)

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset

'On Error GoTo errorHandler

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for the Batch Table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBatch "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then                  'make sure tbl NOT empty
         .MoveFirst
         For i = 1 To .RecordCount              'loop thru recordset
            .Delete
            .Update
            .MoveNext
         Next i
         '.Update                                'update the recordset
      End If
   End With
               
   '--- clean up
   result = True
   adoRS.Close                                  'unload recordset & connection
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing

Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  result = False
  
End Sub  'clrBatch()

'---------------------------------------------------------------------------
'Name:      optBatch()
'Accepts:   none
'Returns:   none
'Requires:  local db
'Discrip:   this sub optimizes the batch table
'Notes:
'---------------------------------------------------------------------------
Public Sub optBatch()

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim adoRS1 As ADODB.Recordset
Dim lngBlnkLen As Long
Dim lngGrpLen As Long
Dim lngRemLen As Long
Dim lngGrpCnt As Long
Dim flgNoMatch As Boolean

'On Error GoTo errorHandler
   
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for Batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBatch " & _
            "ORDER BY BlankLength DESC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   dblLongScrap = 0                              'jng
   dblTotPartsLen = 0                            'JNG add on 03/31/08-reset batch parts total lenght
   lngGrpCnt = 0                                'reset group count
   flgNoMatch = False                           'reset NO Match flag
   
   With adoRS
      .MoveFirst
      Do Until .EOF                             'loop thru batch table
         If .Fields("Priority") = 0 Then        'locate a master part
            flgNoMatch = False
            lngGrpCnt = lngGrpCnt + 1
            lngRemLen = StockLength             'start out w/ a full bar
            lngGrpLen = 0                       'reset group length
            .Fields("Priority") = lngGrpCnt     'set master's priority
            .Update
            lngBlnkLen = .Fields("BlankLength")
            lngGrpLen = lngGrpLen + lngBlnkLen
            lngRemLen = lngRemLen - lngBlnkLen - 750
            dblTotPartsLen = dblTotPartsLen + lngBlnkLen 'jng 3/31/08
            Do Until lngRemLen < 20000 Or flgNoMatch
               '--- make another recordset
               Set adoRS1 = New ADODB.Recordset
               strSQL = "SELECT * FROM tblBatch " & _
                        "WHERE Priority = 0 " & _
                        "AND Blanklength < " & lngRemLen & " " & _
                        "ORDER BY BlankLength DESC"
               adoRS1.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
               
               If adoRS1.RecordCount > 0 Then
                  adoRS1.MoveFirst
                  lngBlnkLen = adoRS1.Fields("BlankLength")
                  lngGrpLen = lngGrpLen + lngBlnkLen
                  lngRemLen = lngRemLen - lngBlnkLen - 750
                  dblTotPartsLen = dblTotPartsLen + lngBlnkLen 'jng 3/31/08
                  adoRS1.Fields("Priority") = lngGrpCnt
                  adoRS1.Update
               Else                             'NO Match
                  flgNoMatch = True
               End If
               
               adoRS1.Close                     'close recordset
               Set adoRS1 = Nothing
            Loop  'end of lngRemLen loop
            
         End If   'master
         
            If lngRemLen > dblLongScrap Then       'jng
               dblLongScrap = lngRemLen            'jng
            End If                                 'jng
        
         .MoveNext
      Loop  'end of batch tbl loop
   End With
         
   intTotBlanks = lngGrpCnt                     'capture total blanks in batch
   intRemBlanks = intTotBlanks                  'set remaining blanks
   
    intBatchBlanks = intTotBlanks                    'jng 03/31/08
    dblBatchLngth = (intBatchBlanks * StockLength)   'jng 03/31/08
    dblBatchOpt = dblTotPartsLen / dblBatchLngth    'jng3/31/08
    dblBatchOpt = dblBatchOpt * 100                 'jng 3/31/08
    intOptType = 1
   
   '--- clean up
   adoRS.Close                                  'unload recordset & connection
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing
   
Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  result = False
  
End Sub  'optBatch()

'---------------------------------------------------------------------------
'Name:      addRun()
'Accepts:   usrPart = the part to add
'Returns:   intResult = 1 added OK
'Requires:  Run Screen w/ AdodcRun
'Discrip:   This sub generates a raw batch table from the exec que.
'Notes:
'---------------------------------------------------------------------------
Public Sub addRun(usrPart As part, intResult As Integer)

With frmRun.AdodcRun.Recordset
   .AddNew
   '--- fill out the record
   '!FullOrder = usrPart.FullJobNum
   ![Order Number] = usrPart.Job
   !Release = usrPart.Rel
   !Item = usrPart.Item
   ![Sequence Number] = usrPart.Seq
   '![Scheduled Ship Date] = usrPart.ShipDate
   '![Quantity] = usrPart.Qnty
   '!BldQnty = usrPart.BldQnty
   !Phase = usrPart.Phase
   !Leg = usrPart.Leg
   !Stack = usrPart.Stack
   !BarType = usrPart.BarType
   !Material = usrPart.Material
   !BarWidth = usrPart.BarWidth
   !BlankLength = usrPart.BlankLength
   '!RunLength = usrPart.RunLength
   !E1figure = usrPart.E1fig
   !E1dimension = usrPart.E1dim
   !E2figure = usrPart.E2fig
   !E2dimension = usrPart.E2dim
   !Cdimension = usrPart.Cdim
   !C1dimension = usrPart.C1dim
   !Ddimension = usrPart.Ddim
   !D1dimension = usrPart.D1dim
   '!Build = usrPart.Build
   '!Status = usrPart.Status
   !Priority = usrPart.Priority
   '!DTStamp = usrPart.DTS
   .Update                                            'update the recordset
End With

intResult = 1

End Sub  'addRun()

'---------------------------------------------------------------------------
'Name:      clrRun
'Accepts:
'Returns:   a boolean, result = true if cleared ok
'Requires:  local db
'Discrip:   this sub empties tblRun
'Notes:
'---------------------------------------------------------------------------
Public Sub clrRun(result As Boolean)

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String                   'sql string
Dim i, j  As Integer                   'loop variables

'On Error GoTo errorHandler
      
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                                                         
   '--- create the recordset
   strSQL = "SELECT * FROM tblRun"                 'set SQL string
   Set adoRS = New ADODB.Recordset                 'init record set
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then                     'make sure tbl NOT empty
         .MoveFirst
         Do Until .EOF                             'loop thru recordset
            .Delete
            .Update
            .MoveNext
         Loop
         '.Update                                  'update the recordset
      End If
   End With
   
   adoRS.Close                                     'unload recordset
   Set adoRS = Nothing
   conLocal.Close                                  'unload connection
   Set conLocal = Nothing
   
   result = True                                   'table cleared OK
   frmRun.RefreshRunQue                            'refresh the RunQue
Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  result = False
  
End Sub  'clrRun

'---------------------------------------------------------------------------
'Name:      clrExec
'Accepts:
'Returns:   a boolean, result = true if cleared ok
'Requires:  local db
'Discrip:   this sub empties the tblExecQue
'Notes:
'---------------------------------------------------------------------------
Public Sub clrExec(result As Boolean)

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String                   'sql string
Dim i, j  As Integer                   'loop variables

'On Error GoTo errorHandler
      
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                                                         
   '--- create the recordset
   strSQL = "SELECT * FROM tblExecQue "            'set SQL string
   Set adoRS = New ADODB.Recordset                 'init record set
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then                     'make sure tbl NOT empty
         .MoveFirst
         Do Until .EOF                             'loop thru recordset
            .Delete
            .Update
            .MoveNext
         Loop
      End If
   End With
   
   adoRS.Close                                     'unload recordset
   Set adoRS = Nothing
   conLocal.Close                                  'unload connection
   Set conLocal = Nothing
   
   result = True                                   'table cleared OK
   frmRun.RefreshExecQue                           'refresh the ExecQue
Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  result = False
  
End Sub  'clrExec











'---sample header
'------------------------------------------------------------------------------
'Name:      x
'Accepts:   none
'Returns:   none
'Requires:
'Discrip:
'Notes:
'------------------------------------------------------------------------------
