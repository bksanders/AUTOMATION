Attribute VB_Name = "Tracking"
Option Explicit

Public Type PWOItem
   Job As String
   Rel As String
   Item As Integer
   Printer As String
   PWOCode As String
   Filler As String
End Type

Global blnEnTRACK As Boolean
Global EnFDRTruck As Boolean
Global EnPITruck As Boolean
Global EnPIGTruck As Boolean
Global strTruckType As String

'---------------------------------------------------------------------------
'Name:      newTruck
'Accepts:   strType = H,HF,B or BF (Housing/Bar,Fitting or straight)
'Returns:   string = truck #
'Requires:
'Discrip:   This function returns a new truck number.
'Notes:
'---------------------------------------------------------------------------
Function newTruck(strType As String, REMAKE As Boolean, HANDCARRY As Boolean) As String

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String
Dim strTruck As String
Dim intTruck As Integer
   
'===========================================================================

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                             
   '--- make recordset of Width Table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM [Bar Envlope #] " & _
            "ORDER BY barnum DESC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then                        'if a record is found
         .MoveFirst
         intTruck = !barnum                           'get last truck#
         
         .AddNew                                      'add record to truck# table
         
         '--- construct the new truck#
         intTruck = intTruck + 1                      'incr tracking #
         strTruck = strType
         If HANDCARRY Then strTruck = strTruck & "D"
         If REMAKE Then strTruck = strTruck & "R"
         strTruck = strTruck & intTruck
         
         ![Bar Envelope #] = strTruck             'store new truck# in table
         !barnum = intTruck
         .Update
      End If
      .Close
   End With

   '--- unload recordset & connection
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing
   
   newTruck = strTruck
   
End Function   'newTruck

'---------------------------------------------------------------------------
'Name:      getTruckList
'Accepts:
'Returns:
'Requires:
'Discrip:   This Sub generates a truck list from the history table and
'           populates the combobox, cmbTruck.
'Notes:
'---------------------------------------------------------------------------
Sub getTruckList()

Dim conDB As Connection
Dim adoRS As ADODB.Recordset
Dim strConn As String
Dim strTable As String
Dim strSQL As String
Dim i As Integer                    'list count
Dim j As Integer                    'list index
Dim blnFOUND As Boolean             'truck number found in list flag
Dim intCDBok As Integer             'central DB status, 1= avail

   '--- make connection to local db
   Set conDB = New Connection
   conDB.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                               
   '--- make recordset for History que
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT Truck FROM tblTracking " & _
            "WHERE Truck IS NOT NULL " & _
            "AND Truck <> 'NONE'" & _
            "ORDER BY DTStamp DESC"
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then
         .MoveFirst                                         'go to 1st #
         frmTruck.cmbTruck.AddItem !Truck                   'add 1st truck# to list
         .MoveNext                                          'go to next #
         i = 1
         Do Until .EOF Or i = 20                            'loop thru truck #'s
            blnFOUND = False
            For j = 0 To i - 1                              'loop thur list
               If frmTruck.cmbTruck.List(j) = !Truck Then   'check truck# against list
                  blnFOUND = True
                  Exit For
               End If
            Next j
            If blnFOUND = False Then                        'if # not in list
               frmTruck.cmbTruck.AddItem !Truck             'add to list
               i = i + 1                                    'incr list count
            End If
            .MoveNext
         Loop
      End If
      .Close
   End With
   
   Set adoRS = Nothing
   conDB.Close
   Set conDB = Nothing

End Sub

'---------------------------------------------------------------------------
'Name:      genPWOList()
'Accepts:
'Returns:   integer = # of JRI's in PWO list
'Requires:  dsn to central db
'Discrip:   this sub generates the table w/ the PWO list for printing.
'Notes:
'---------------------------------------------------------------------------
Public Function genPWOList(strTruckNO As String) As Integer

Dim conDB As Connection
Dim strConn As String
Dim strTable As String
Dim adoRS As ADODB.Recordset           'recordset
Dim strSQL As String
Dim conLocal As Connection             'connection to local database
Dim adoRS1 As ADODB.Recordset          'recordset
Dim pwoList(200) As PWOItem
Dim pwoPTR As Integer                  'index into pwolist
Dim j As Integer                       'loop index
Dim blnFOUND As Boolean
Dim intCDBok As Integer             'central DB status, 1= avail

   '--- make connection to central db
   strConn = "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
   Set conDB = New Connection
   conDB.Open strConn
   
   '--- make recordset for history table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT FullOrder,Release,Item " & _
            "FROM tblTracking " & _
            "WHERE Truck = '" & strTruckNO & "'"
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then
         pwoPTR = 1
         .MoveFirst                                         'go to 1st record
         pwoList(pwoPTR).Job = !FullOrder              'add 1st record to list
         pwoList(pwoPTR).Rel = !Release
         pwoList(pwoPTR).Item = !Item
         pwoList(pwoPTR).Printer = "A8"
         pwoList(pwoPTR).PWOCode = "B"
         '.MoveNext                                          'go to next #
         
         Do Until .EOF                                      'loop thru recordset
            blnFOUND = False
            For j = 1 To pwoPTR                             'loop thur list
               If !FullOrder = pwoList(j).Job And _
                  !Release = pwoList(j).Rel And _
                  !Item = pwoList(pwoPTR).Item Then         'check JRI against list
                  blnFOUND = True
                  Exit For
               End If
            Next j   'pwo list loop
            If blnFOUND = False Then                        'if # not in list
               pwoPTR = pwoPTR + 1                          'incr list count
               pwoList(pwoPTR).Job = !FullOrder        'add record to list
               pwoList(pwoPTR).Rel = !Release
               pwoList(pwoPTR).Item = !Item
               pwoList(pwoPTR).Printer = "A8"
               pwoList(pwoPTR).PWOCode = "B"
            End If
            .MoveNext
         Loop  'recordset loop
         
         '------------------------ store PWO list in table -----------------
         
         '--- make recordset for machine table
         Set adoRS1 = New ADODB.Recordset
         strSQL = "SELECT * FROM PRTPWO"
         adoRS1.Open strSQL, conDB, adOpenStatic, adLockOptimistic
         
         With adoRS1
            For j = 1 To pwoPTR                             'loop thru list
               .AddNew                                      'add each PWO item
               ![PRSHP#] = pwoList(j).Job             'to table
               ![PRREL#] = pwoList(j).Rel
               ![PRITM#] = pwoList(j).Item
               !PRPRNT = pwoList(j).Printer
               !PRPAGE = pwoList(j).PWOCode
               .Update
            Next j   'list
            .Close                                          'close recordset
         End With
         genPWOList = pwoPTR
         
         Set adoRS1 = Nothing                               'unload recordset
      
      Else
         genPWOList = 0
      End If
      .Close                                                'close recordset
   End With
   
   Set adoRS = Nothing                                      'unload recordset
   conDB.Close                                       'close connection
   Set conDB = Nothing                               'unload connection
   
End Function 'genPWOList()

'---------------------------------------------------------------------------
'Name:      genReport()
'Accepts:   strMode, P = print; V = show
'           strTruckNO = desired truck #
'Returns:
'Requires:
'Discrip:   This Sub prints/shows the truck summary report.
'Notes:
'---------------------------------------------------------------------------

Public Sub genReport(strMode As String, strTruckNO As String)

Dim conDB As Connection                'connection to local database
Dim adoRS As ADODB.Recordset           'recordset
Dim strSQL As String
   
   '--- make connection to db
   Set conDB = New Connection
   conDB.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
   
   '--- make recordset
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblTracking WHERE TRUCK = '" & strTruckNO & "'"
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic
   
   With rptTruckSum
      .Title = "Truck Summary: " & strTruckNO
      Set .DataSource = adoRS
      .DataMember = adoRS.DataMember
      .Sections("Section4").Controls("Label14").Caption = strTruckNO
      .Sections("Section4").Controls("Label15").Caption = strTruckNO
      DoEvents
      If strMode = "P" Then
         .PrintReport
      Else
         .Show
      End If
   End With
   
   wait 5                                    'wait for report to display
   
   adoRS.Close
   Set adoRS = Nothing
   conDB.Close
   Set conDB = Nothing

End Sub

'---------------------------------------------------------------------------
'Name:      inHist()
'Accepts:   a record ID for the part/item to be checked
'Returns:   an result integer, indicating the results of the check
'              0 = NOT in hist; 1 = in hist; 2= in hist& complete
'Requires:  dsn's to local history tables
'Discrip:   this sub checks to see if a part/item is in the history table
'Notes: (1) if the .seq of the recID is left blank...then the sub will treat
'           the the recID as an Item instead of a part.
'Last Mod:  5/28/04, GEAS, JKA
'---------------------------------------------------------------------------
Public Function inHist(recID As RecordID) As Integer

Dim conDB As Connection             'camdata connection
Dim adoRS As ADODB.Recordset        'recordset of this part in hist table
Dim strSQL As String                'SQL query string
Dim strConn As String               'connection string
Dim strTable As String              'history table string
Dim intDesiredQty As Integer        'desired qty for this part
Dim intPartCount As Integer         'total qty of this part in hist table
Dim intCDBok As Integer             'central DB avail.
      
   '------------------------- build connection string ----------------------
   strTable = "tblHist"
   strConn = "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
   
   '------------------------ build query string ---------------------------
   If recID.Seq = 0 Then
      strSQL = "SELECT * FROM " & strTable & " " & _
               "WHERE ([Order Number] = '" & recID.Job & "' " & _
               "AND Release = '" & recID.Rel & "' " & _
               "AND Item = " & recID.Item & ") "
   Else
      strSQL = "SELECT * FROM " & strTable & " " & _
               "WHERE ([Order Number] = '" & recID.Job & "' " & _
               "AND Release = '" & recID.Rel & "' " & _
               "AND Item = " & recID.Item & " " & _
               "AND [Sequence Number] = " & recID.Seq & ") "
   End If
      
   '--- make connection to db
   Set conDB = New Connection
   conDB.Open strConn

   '--- make recordset
   Set adoRS = New ADODB.Recordset
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic
   
   
   'Note: The desired quantity of the current part being checked is passed in
   '      the .pri field record ID.
   intDesiredQty = recID.Pri
     
   '------------------------------ check part ---------------------------------
   With adoRS
      If .RecordCount > 0 Then                              'if a Record is found
         inHist = 1
         .MoveFirst
         
         '----- loop thru recordset
         Do Until .EOF Or inHist = 2
            Select Case .Fields("Build")
            Case "F"
               If .Fields("BldQnty") = intDesiredQty Then
                  inHist = 2                             'part complete
               End If
            Case "P"
               intPartCount = intPartCount + .Fields("BldQnty")
               If intPartCount = intDesiredQty Then
                  inHist = 2                             'part complete
               End If
            Case "R"
               'Don't count remakes
            End Select
         .MoveNext
         Loop  'recordset
         
      Else
         inHist = 0                                      'part NOT in Hist table
      End If   'recordcount
         
      .Close                                                'close recordset
   End With
      
   '--------------------------------- clean up --------------------------------
   Set adoRS = Nothing                                      'unload recordset
   conDB.Close                                              'close connection
   Set conDB = Nothing                                      'unload connection
   
End Function 'inHist

'---------------------------------------------------------------------------
'Name:      addTracking()
'Accepts:   usrPart, the part to be added
'Returns:   intResult, 0 = not added, 1=added OK
'Requires:
'Discrip:   This Sub adds a part to the tracking table
'Notes:
'---------------------------------------------------------------------------
Public Sub addTracking(usrPart As part, intResult As Integer)

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim intBuildQty As Integer
Dim strSQL As String

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for History que
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblTracking " & _
            "WHERE [Order Number] = '" & usrPart.Job & "' " & _
            "AND Release = '" & usrPart.Rel & "' " & _
            "AND Item = " & usrPart.Item & " " & _
            "AND [Sequence Number] = " & usrPart.Seq & " " & _
            "AND Truck = '" & usrPart.Truck & "'"
   
   
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then
         intBuildQty = !BldQnty                 'get existing  build qty
         intBuildQty = intBuildQty + 1          'incr build qty
         !BldQnty = intBuildQty                 'store new build qty
         .Update                                'update record
      Else
      .AddNew                                   'add a record
         '--- fill out the record
         !FullOrder = usrPart.FullJobNum
         ![Order Number] = usrPart.Job
         !Release = usrPart.Rel
         !Item = usrPart.Item
         ![Sequence Number] = usrPart.Seq
         ![Scheduled Ship Date] = usrPart.ShipDate
         ![Quantity] = usrPart.Qnty
         !BldQnty = 1
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
         !Truck = usrPart.Truck
         !Machine = usrPart.Machine
         !DTStamp = usrPart.DTS
         .Update                                   'update the recordset
      End If
   End With
         
   '--- clean up
   adoRS.Close                              'unload recordset & connection
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing
   frmHist.RefreshHistQue                       'refresh the Hist Que
   intResult = 1                                'return result
   
End Sub  'addTracking

