Attribute VB_Name = "Optimize"
Option Explicit

Public Type BatchInfo
   InitBatch As Boolean       'flag for initial batch
   InitPick As Boolean        'flag for initial pick
   Fill As Boolean            'flag for batch filled
   Lock As Boolean            'flag for batch locked
   Job As String              'Init Pick -- Job
   Rel As String              'Init Pick -- Rel
   SDate As Date              'Init Pick -- schedule date
   Mat As String              'batch material
   Width As String            'batch width
   Blanks As Integer          'total # of blanks in Batch
   TotOFF As Long             'total off-fall for batch
   TotLen As Long             'Total utilized length
   LongOFF As Long            'longest off-fall for batch
   Opt As Double              '%Optimization for batch
End Type

Public Type Release
   Job As String              'Job #
   Rel As String              'Rel #
End Type

Public Type MatchParams
   Index As Integer           'pointer to item in machine list
   Stacks As Integer          '# of stacks matched
   Complete As Boolean        'complete item matched flag
   Length As Long             'total recovered length
   MaxSeqLen As Long          'length of longest seq
   SeqCnt As Integer          '# of seq in item
End Type

'----- Optimization paramaters
Public blnEnOPT As Boolean                'enable optimization
Public blnAutoPICK As Boolean             'flag for AutoPICKing initial batch
Public intFillDays As Integer             '# of days for fill search
Public intMaxRel As Integer               'max # of rel's allowed in batch

'----- global variables
Public usrBatch As BatchInfo              'batch info
Public usrNullBatch As BatchInfo          'null batch info
Public arrOffFall(20) As Long             'off fall array
Public usrMACHlist(7000) As part          'listbox Storage array
Public intMachCnt As Integer              '# of parts in machine array
Public intMachIDX As Integer              'index for selected part
Public arrRelease(10) As Release          'release array
Public intRelCnt As Integer               '# of release's in release array
Public intAsgnCnt As Integer              '# of assinged parts(for machine list assignment update)

'---------------------------------------------------------------------------
'Name:      clrLocalTBL()
'Accepts:
'Returns:
'Requires:  dsn to central db
'Discrip:   this sub clears a table in the local db.
'Notes:
'---------------------------------------------------------------------------
Public Sub clrLocalTBL(strTable As String)

Dim conLocalDB As Connection           'connection to local database
Dim adoRS As ADODB.Recordset           'recordset
Dim strSQL As String

   '--- make connection to central db
   Set conLocalDB = New Connection
   conLocalDB.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
   
   '--- make recordset for machine table
   Set adoRS = New ADODB.Recordset
   strSQL = "DELETE * FROM " & strTable
   adoRS.Open strSQL, conLocalDB, adOpenStatic, adLockOptimistic
   
   Set adoRS = Nothing                       'unload recordset
   
   conLocalDB.Close                          'close connection
   Set conLocalDB = Nothing                  'unload connection

End Sub  'clrLocalTBL()

'---------------------------------------------------------------------------
'Name:      addBatch2()
'Accepts:   usrPart, the part
'           strTable, the name of the batch table
'Returns:   a boolean, result = true if added ok
'Requires:  local db
'Discrip:   this sub adds an entry into the batch table.
'Notes:
'---------------------------------------------------------------------------
Public Sub addBatch2(usrPart As part, strTable As String, result As Boolean)

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String

On Error GoTo errorHandler

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM " & strTable
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
  
End Sub  'addBatch2()


'---------------------------------------------------------------------------
'Name:      bldBatch2()
'Accepts:   none
'Returns:   none
'Requires:  Run Screen
'Discrip:   This sub generates a raw batch table from the batch que.
'Notes:
'---------------------------------------------------------------------------
Public Sub bldBatch2()

Dim i As Integer           'loop index
Dim j As Integer
Dim usrBatch As part       'batch entry variable
Dim addOK As Boolean       'add result
Dim intQnty As Integer     'quantity to add to batch
         
For i = 1 To intMachCnt
   DoEvents
   '--- generate batch entry's
   'intQnty = usrMACHlist(i).Qnty - usrMACHlist(i).BldQnty
   intQnty = usrMACHlist(i).BldQnty
   If intQnty > 0 Then
      For j = 1 To intQnty
         addOK = False
         addBatch2 usrMACHlist(i), "tblBatch2", addOK
         If addOK = False Then
            MsgBox ("bldBatch2: An error occured when Adding a record " & _
                   Chr$(13) & "to the Batch Table. Optimization aborted!")
            Exit Sub
         End If
      Next j
   End If
Next i

End Sub  'bldBatch2()


'---------------------------------------------------------------------------
'Name:      optBatch2()
'Accepts:   none
'Returns:   none
'Requires:  local db
'Discrip:   this sub optimizes the batch table
'Notes:
'---------------------------------------------------------------------------
Public Sub optBatch2()

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim adoRS1 As ADODB.Recordset
Dim lngBlnkLen As Long
Dim lngGrpLen As Long
Dim lngRemLen As Long
Dim lngGrpCnt As Long
Dim flgNoMatch As Boolean
Dim strSQL As String
   
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for Batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBatch2 " & _
            "ORDER BY BlankLength DESC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   lngGrpCnt = 0                                'reset group count
   flgNoMatch = False                           'reset NO Match flag
   clrOffFall                                   'reset the off-fall array
   usrBatch.TotLen = 0                          'reset total blank length
   dblTotPartsLen2 = 0                           'jng- reset batch parts total length
   
   With adoRS
      If .RecordCount > 0 Then                     'no need to opt is tbl empty
         .MoveFirst
         Do Until .EOF                             'loop thru batch table
            DoEvents
            If .Fields("Priority") = 0 Then        'locate a master part
               flgNoMatch = False
               lngGrpCnt = lngGrpCnt + 1
               lngRemLen = StockLength             'start out w/ a full bar
               lngGrpLen = 0                       'reset group length
               .Fields("Priority") = lngGrpCnt     'set master's priority
               .Update
               lngBlnkLen = .Fields("BlankLength")
               
               dblTotPartsLen2 = dblTotPartsLen2 + lngBlnkLen   'jng
               
               lngGrpLen = lngGrpLen + lngBlnkLen + 750
               lngRemLen = lngRemLen - lngBlnkLen - 750
               Do Until lngRemLen < 24000 Or flgNoMatch
                  '--- make another recordset
                  Set adoRS1 = New ADODB.Recordset
                  strSQL = "SELECT * FROM tblBatch2 " & _
                           "WHERE Priority = 0 " & _
                           "AND Blanklength < " & lngRemLen & " " & _
                           "ORDER BY BlankLength DESC"
                  adoRS1.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
                  
                  If adoRS1.RecordCount > 0 Then
                     adoRS1.MoveFirst
                     lngBlnkLen = adoRS1.Fields("BlankLength")
                     
                     dblTotPartsLen2 = dblTotPartsLen2 + lngBlnkLen      'jng
                     
                     lngGrpLen = lngGrpLen + lngBlnkLen + 750
                     lngRemLen = lngRemLen - lngBlnkLen - 750
                     adoRS1.Fields("Priority") = lngGrpCnt
                     adoRS1.Update
                  Else                             'NO Match
                     flgNoMatch = True
                  End If
                  
                  adoRS1.Close                     'close recordset
                  Set adoRS1 = Nothing
               Loop  'end of lngRemLen loop
               
               logOffFall lngRemLen, lngGrpCnt     'capture off-fall
               usrBatch.TotLen = usrBatch.TotLen + lngGrpLen
               
            End If   'master
            .MoveNext
         Loop  'end of batch tbl loop
         
         '--- calc opt stats
         usrBatch.Blanks = lngGrpCnt                     'capture total blanks in batch
        
        intBatchBlanks2 = lngGrpCnt      'jng
        
        If usrBatch.Blanks > 0 Then
            usrBatch.Opt = usrBatch.TotLen / (usrBatch.Blanks * StockLength)
            usrBatch.Opt = usrBatch.Opt * 100
            
            dblBatchLngth2 = (usrBatch.Blanks * StockLength)   'jng 03/31/08
            dblBatchOpt2 = dblTotPartsLen2 / dblBatchLngth2     'jng 03/31/08 added for optimization tracking
            dblBatchOpt2 = dblBatchOpt2 * 100                 'jng
            intOptType2 = 2                                   'jng
          
         Else
            usrBatch.Opt = 0
         End If
         usrBatch.LongOFF = arrOffFall(19)
         usrBatch.TotOFF = arrOffFall(20)
         
         dblLongScrap2 = usrBatch.LongOFF                       'jng
      
      Else                                         'batch table empty
         usrBatch.Blanks = 0                       'clear batch stats
         usrBatch.Opt = 0
         usrBatch.LongOFF = 0
         usrBatch.TotOFF = 0
      End If   'tbl empty
   End With
            
   '--- clean up
   adoRS.Close                                  'unload recordset & connection
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing

End Sub  'optBatch2()

'---------------------------------------------------------------------------
'Name:      clrOffFall()
'Accepts:
'Returns:
'Requires:  off fall array
'Discrip:   this sub clears the off fall array
'Notes:
'---------------------------------------------------------------------------
Public Sub clrOffFall()
Dim i As Integer              'loop index
   For i = 0 To 20                              'clear the off-fall array
      arrOffFall(i) = 0
   Next i
   clrLocalTBL "tblOffFall"                     'clear the off-fall table
End Sub  'clrOffFall

'---------------------------------------------------------------------------
'Name:      logOffFall()
'Accepts:
'Returns:
'Requires:  off fall array
'Discrip:   this sub logs the current off-fall into the array
'Notes:
'---------------------------------------------------------------------------
Public Sub logOffFall(curOFFfall As Long, groupID As Long)

Dim i As Integer              'loop index
Dim addOK As Boolean          'add result

   '--- catagories off-fall
   If curOFFfall >= 100000 Then i = 18
   If curOFFfall < 100000 Then i = 17
   If curOFFfall < 95000 Then i = 16
   If curOFFfall < 90000 Then i = 15
   If curOFFfall < 85000 Then i = 14
   If curOFFfall < 80000 Then i = 13
   If curOFFfall < 75000 Then i = 12
   If curOFFfall < 70000 Then i = 11
   If curOFFfall < 65000 Then i = 10
   If curOFFfall < 60000 Then i = 9
   If curOFFfall < 55000 Then i = 8
   If curOFFfall < 50000 Then i = 7
   If curOFFfall < 45000 Then i = 6
   If curOFFfall < 40000 Then i = 5
   If curOFFfall < 35000 Then i = 4
   If curOFFfall < 30000 Then i = 3
   If curOFFfall < 25000 Then i = 2
   If curOFFfall < 20000 Then i = 1
   If curOFFfall < 15000 Then i = 0
      
   arrOffFall(i) = arrOffFall(i) + 1                     'log off-fall into array
      
   If curOFFfall > 0 Then
      If curOFFfall > arrOffFall(19) Then                'check for longest off-fall
         arrOffFall(19) = curOFFfall
      End If
      arrOffFall(20) = arrOffFall(20) + curOFFfall       'add to total off-fall
   End If
   
   If curOFFfall > 24000 Then
      addOffFall curOFFfall, groupID, addOK              'add to off-fall table
   End If
End Sub  'clrOffFall
'---------------------------------------------------------------------------
'Name:      locOffFall()
'Accepts:
'Returns:
'Requires:  off fall array
'Discrip:   this function locates an unused piece of off-fall of a
'           particular size'
'Notes:
'---------------------------------------------------------------------------
Function locOffFall(Length1 As Long) As Integer

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblOffFall " & _
            "WHERE Status = 'N' AND Length > " & Length1 & " " & _
            "ORDER BY Length DESC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then
         .MoveFirst                          'this will locate longest unused off-fall
         locOffFall = !Group                 'get its group
         !Status = "U"
         .Update
      Else
         locOffFall = 0
      End If
      .Close
   End With
   
   Set adoRS = Nothing                       'unload recordset
   conLocal.Close                            'close conn
   Set conLocal = Nothing                    'unload conn
   
End Function   'locOffFall
'---------------------------------------------------------------------------
'Name:      chkOffFall()
'Accepts:
'Returns:
'Requires:  off fall table
'Discrip:   this function locates an unused piece of off-fall of a
'           particular size for an item check.
'Notes:
'---------------------------------------------------------------------------
Function chkOffFall(Length1 As Double) As Integer

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblOffFall " & _
            "WHERE Status = 'N' AND Check = 'N' AND Length > " & Length1 & " " & _
            "ORDER BY Length ASC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then
         .MoveFirst                          'this will locate shortest unused off-fall
         chkOffFall = !Group                 'get its group
         !Check = "U"                        'mark check field as used
         .Update
      Else
         chkOffFall = 0
      End If
      .Close
   End With
   
   Set adoRS = Nothing                       'unload recordset
   conLocal.Close                            'close conn
   Set conLocal = Nothing                    'unload conn
   
End Function   'chkOffFall
'---------------------------------------------------------------------------
'Name:      fillOffFall()
'Accepts:
'Returns:
'Requires:  off fall array
'Discrip:   this function locates an unused piece of off-fall of a
'           particular size and marks it as used for the filling process.
'Notes:
'---------------------------------------------------------------------------
Function fillOffFall(Length1 As Double) As Integer

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblOffFall " & _
            "WHERE Status = 'N' AND Length > " & Length1 & " " & _
            "ORDER BY Length ASC"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then
         .MoveFirst                          'this will locate longest unused off-fall
         fillOffFall = !Group                 'get its group
         !Status = "U"
         .Update
      Else
         fillOffFall = 0
      End If
      .Close
   End With
   
   Set adoRS = Nothing                       'unload recordset
   conLocal.Close                            'close conn
   Set conLocal = Nothing                    'unload conn
   
End Function   'fillOffFall

'---------------------------------------------------------------------------
'Name:      compareMatch()
'Accepts:   two sets of match paramaters
'Returns:   the # of the better match: 1 = 1st match; 2 = 2nd
'Requires:
'Discrip:   this function compares the match paramaters for two matches and
'           determines which match is better.
'Notes:
'---------------------------------------------------------------------------
Public Function compareMatch(match1 As MatchParams, match2 As MatchParams) As Integer
   
   '---------------------------- Complete ----------------------------
   If match2.Complete = True And match1.Complete = False Then
      compareMatch = 2
      Exit Function
   End If
   If match2.Complete = False And match1.Complete = True Then
      compareMatch = 1
      Exit Function
   End If
   
   '------------------------ recovered length -----------------------
   If match2.Length > match1.Length Then
      compareMatch = 2
      Exit Function
   End If
   If match1.Length > match2.Length Then
      compareMatch = 1
      Exit Function
   End If
   
   '------------------------ Longest sequence -------------------------
   If match2.MaxSeqLen >= match1.MaxSeqLen Then
      compareMatch = 2
   Else
      compareMatch = 1
   End If
   
End Function   'compareMatch
'---------------------------------------------------------------------------
'Name:      addOffFall()
'Accepts:   lngLength = length of the off-fall
'           intGroup = group from which off-fall cam
'Returns:   a boolean, blnResult = true if added ok
'Requires:  local db
'Discrip:   this sub adds an entry into the Off-Fall table.
'Notes:
'---------------------------------------------------------------------------
Public Sub addOffFall(lngLength As Long, lngGroup As Long, blnResult As Boolean)

Dim conLocal As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String

On Error GoTo errorHandler

   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for batch table
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblOffFall "
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   With adoRS
      .AddNew                                   'add a record to tblBatch
      '--- fill out the record
      !Length = lngLength
      !Group = lngGroup
      !Status = "N"
      !Check = "N"
      .Update                                   'update the recordset
   End With
         
   '--- clean up
   adoRS.Close                                  'unload recordset & connection
   Set adoRS = Nothing
   conLocal.Close
   Set conLocal = Nothing
   
   blnResult = True                             'return result

Exit Sub

errorHandler:     '------------------ Error Handler ---------------------
  MsgBox Err.Description
  blnResult = False
  
End Sub  'addOffFall()
'---------------------------------------------------------------------------
'Name:      chkItemFit()
'Accepts:   index to item in machine list
'Returns:   match paramaters
'Requires:
'Discrip:   this sub checks an item to see if it will fit into the existing
'           off-fall
'Notes:
'---------------------------------------------------------------------------
Public Sub chkItemFit(thisMatch As MatchParams)
Dim idx As Integer            'index of current item
Dim strJRI As String
Dim thisJRI As String
Dim longestSeqLen As Long
Dim intSeqCnt As Integer      '# of seq in curr item
Dim totRecLength As Long      'total recovered length
Dim i, j As Integer           'loop variable
Dim intDesrdSTK As Integer    'desired stack quantity
Dim NOMATCH As Boolean        'no match flag
Dim intGroup                  'check result

   idx = thisMatch.Index
   strJRI = usrMACHlist(idx).Job & usrMACHlist(idx).Rel & usrMACHlist(idx).Item
   thisJRI = strJRI
   
   '-------------------------------- Pre-check --------------------------------
   'the precheck performs a quick check to determine if the item has a seq longer
   'than the longest off-fall piece.  If so, no need to check further...item
   'won't fit.
   'the precheck also identifies the # of seq's in the item & the length of the
   'longest seq in the item.
   i = idx
   Do Until thisJRI <> strJRI
      If usrMACHlist(i).BlankLength > usrBatch.LongOFF Then
         Exit Sub
      End If
      If usrMACHlist(i).BlankLength > longestSeqLen Then    'id longest seq
         longestSeqLen = usrMACHlist(i).BlankLength
      End If
      intSeqCnt = intSeqCnt + 1                             'tally up seq's in the item
      i = i + 1                                             'incr loop pointer
      thisJRI = usrMACHlist(i).Job & usrMACHlist(i).Rel & usrMACHlist(i).Item
   Loop
   
   '------------------------------- Item-check ---------------------------------
   resItemCHK                                               'clear check field for item check
   intDesrdSTK = usrMACHlist(idx).Qnty
   
   For i = 1 To intDesrdSTK
      NOMATCH = False
      For j = idx To (idx + intSeqCnt - 1)
         intGroup = chkOffFall(usrMACHlist(j).BlankLength)
         If intGroup > 0 Then
            totRecLength = totRecLength + usrMACHlist(j).BlankLength
         Else
            NOMATCH = True
            Exit For
         End If
      Next j
      
      If NOMATCH = True Then
         Exit For
      Else
         thisMatch.Stacks = thisMatch.Stacks + 1            'incr fill'd stack count
      End If
   Next i
   
   If NOMATCH = True Then
      thisMatch.Complete = False
   Else
      thisMatch.Complete = True
   End If
   
   '--- return match params
   thisMatch.SeqCnt = intSeqCnt
   thisMatch.MaxSeqLen = longestSeqLen
   thisMatch.Length = totRecLength
   
End Sub  'chkItemFit
'---------------------------------------------------------------------------
'Name:      resItemCHK()
'Accepts:
'Returns:
'Requires:  dsn to central db
'Discrip:   this sub reset the check field in the off-fall table, to prepare
'           for an item check.
'Notes:
'---------------------------------------------------------------------------
Public Sub resItemCHK()

Dim conLocalDB As Connection           'connection to local database
Dim adoRS As ADODB.Recordset           'recordset
Dim strSQL As String

   '--- make connection to central db
   Set conLocalDB = New Connection
   conLocalDB.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
   
   '--- make recordset for machine table
   Set adoRS = New ADODB.Recordset
   
   strSQL = "UPDATE tblOffFall SET Check = 'N'"
   adoRS.Open strSQL, conLocalDB, adOpenStatic, adLockOptimistic
   
   Set adoRS = Nothing                       'unload recordset
   
   conLocalDB.Close                          'close connection
   Set conLocalDB = Nothing                  'unload connection

End Sub  'resItemCHK

'---------------------------------------------------------------------------
'Name:      clrRelArr()
'Returns:
'Requires:  , array
'Discrip:   this sub clears the release array
'Notes:
'---------------------------------------------------------------------------
Public Sub clrRelArr()
Dim i As Integer              'loop index
   For i = 1 To 10                              'clear the off-fall array
      arrRelease(i).Job = ""
      arrRelease(i).Rel = ""
      intRelCnt = 0
   Next i
End Sub  'clrRelArr
'---------------------------------------------------------------------------
'Name:      inRelease()
'Accepts:   curRel, type release
'Returns:   the location of the current release in the array, 0 if not in
'           array.
'Requires:  arrRelease, array
'Discrip:   this sub checks to see if the given release is already in the
'           release array
'Notes:
'---------------------------------------------------------------------------
Public Function inRelease(curRel As Release) As Integer
Dim i As Integer              'loop index
   For i = 1 To intRelCnt
      If curRel.Job = arrRelease(i).Job And _
         curRel.Rel = arrRelease(i).Rel Then
         inRelease = i
         Exit Function
      End If
   Next i
   inRelease = 0
End Function 'inRelease

'---------------------------------------------------------------------------
'Name:      addRelease()
'Accepts:   a job, string
'           a release, string
'Returns:   the location of the current release in the array, 0 if not in
'           array.
'Requires:  arrRelease, array
'Discrip:   this sub checks to see if the given release is already in the
'           release array
'Notes:
'---------------------------------------------------------------------------
Public Function addRelease(strJob As String, strRel As String) As Integer
Dim i As Integer              'loop index
Dim relIDX As Integer         'release index
Dim myRel As Release          'current release
   
   myRel.Job = strJob
   myRel.Rel = strRel
   
   relIDX = inRelease(myRel)                       'check if new release is
   If relIDX > 0 Then                              'already in rel array
      addRelease = relIDX                          'return its position
   Else
      If intRelCnt < intMaxRel Then                'check if OK to add
         intRelCnt = intRelCnt + 1                 'incr rel count
         arrRelease(intRelCnt).Job = myRel.Job     'add rel to array
         arrRelease(intRelCnt).Rel = myRel.Rel
         addRelease = intRelCnt                    'return new position
      Else
         addRelease = 0                            'item not added
      End If
   End If

End Function 'addRelease

'---------------------------------------------------------------------------
'Name:      autoPick()
'Accepts:
'Returns:
'Requires:  global, usrBatch --- batch info variable
'Discrip:   this sub determines the initial pick for the batch.
'Notes:
'---------------------------------------------------------------------------
Public Sub autoPick()

Dim conDB As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String
   
   '--- Construct Search Criteria
   strSQL = "SELECT [Order Number], Release, " & _
            "Material, BarWidth, ShipDate  " & _
            "FROM tblMachine " & _
            "WHERE (((Machine) = 'M' Or (Machine) = 'E') And ((OpenQty) > 0)) " & _
            "ORDER BY ShipDate"
    
   '--- make connection to central database"
   Set conDB = New Connection
   conDB.Open "PROVIDER=MSDASQL;dsn=dsnCentralDB;uid=;pwd=;"
                              
   '--- create the recordset
   Set adoRS = New ADODB.Recordset                          'record set
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then                              'if a Record is found
         .MoveFirst
         usrBatch.Job = ![Order Number]                     'get the init batch info
         usrBatch.Rel = !Release
         usrBatch.SDate = !ShipDate
         usrBatch.Mat = !Material
         usrBatch.Width = !BarWidth
         usrBatch.InitBatch = True                          'mark init batch as complete
         usrBatch.InitPick = True                           'mark INIT pick as complete
         
         intRelCnt = 1                                      'log the initial release
         arrRelease(intRelCnt).Job = usrBatch.Job
         arrRelease(intRelCnt).Rel = usrBatch.Rel
      Else
         MsgBox ("autoPick:  No Records found in Machine table!")
      End If
      .Close                                                'close recordset
   End With
   
   '------------------------------- Clean UP -------------------------------
   Set adoRS = Nothing                                      'unload recordset
   conDB.Close                                              'close connection
   Set conDB = Nothing                                      'unload connection
   
End Sub  'autoPick

'---------------------------------------------------------------------------
'Name:      fillSearch()
'Accepts:   release index, integer
'           search type, string
'Returns:   index to match, 0 if unmatched.
'Requires:
'Discrip:   this function performs the fill search.
'Notes:
'---------------------------------------------------------------------------
Function fillsearch(rID As Integer, sType As String, umQty As Integer, Length1 As Long) As Integer
Dim fillDate As Date          'date range for search
Dim Length2 As Long           'upper length range
Dim i As Integer              'loop index
Dim strJob As String          'Job #
Dim strRel As String          'Rel #
Dim mIDX As Integer           'match index
Dim mQTY As Integer           'match quantity

   fillsearch = 0                                     'init return value
   fillDate = usrBatch.SDate + intFillDays                'calc the fill date
   Length2 = Length1 + 6000                           'calc upper length range
   
   If rID > 0 Then   '-------------------------------- search w/in rel list
      strJob = arrRelease(rID).Job                    'get release info
      strRel = arrRelease(rID).Rel
      i = 1                                           'init loop index
      Select Case sType
      Case "="                                        'search for qty = & length =
         'this search returns the 1st qty & length match found
         Do Until usrMACHlist(i).SchedDate > fillDate
            If usrMACHlist(i).Job = strJob And _
               usrMACHlist(i).Rel = strRel And _
               usrMACHlist(i).BldQnty <> usrMACHlist(i).Qnty And _
               (usrMACHlist(i).BlankLength > Length1 And _
                usrMACHlist(i).BlankLength <= Length2) And _
               usrMACHlist(i).Qnty = umQty Then
               fillsearch = i
               Exit Do
            Else
               i = i + 1                              'incr loop index
            End If
         Loop  'machine list loop
      Case "<"
         'this search returns the largest quantity < unmatched quantity
         mIDX = 0
         mQTY = 0
         Do Until usrMACHlist(i).SchedDate > fillDate
            If usrMACHlist(i).Job = strJob And _
               usrMACHlist(i).Rel = strRel And _
               usrMACHlist(i).BldQnty <> usrMACHlist(i).Qnty And _
               (usrMACHlist(i).BlankLength > Length1 And _
                usrMACHlist(i).BlankLength <= Length2) And _
               usrMACHlist(i).Qnty < umQty Then
               If usrMACHlist(i).Qnty > mQTY Then
                  mQTY = usrMACHlist(i).Qnty
                  mIDX = i
               End If
            End If
            i = i + 1                                 'incr loop index
         Loop  'machine list loop
         fillsearch = mIDX
      Case ">"
         'this search returns the smallest quantity > unmatched quantity
         mIDX = 0
         mQTY = 30000
         Do Until usrMACHlist(i).SchedDate > fillDate
            If usrMACHlist(i).Job = strJob And _
               usrMACHlist(i).Rel = strRel And _
               usrMACHlist(i).BldQnty <> usrMACHlist(i).Qnty And _
               (usrMACHlist(i).BlankLength > Length1 And _
                usrMACHlist(i).BlankLength <= Length2) And _
               usrMACHlist(i).Qnty > umQty Then
               If usrMACHlist(i).Qnty < mQTY Then
                  mQTY = usrMACHlist(i).Qnty
                  mIDX = i
               End If
            End If
            i = i + 1                                 'incr loop index
         Loop  'machine list loop
         fillsearch = mIDX
      End Select
   Else    '------------------------------------------ search outside rel list
      i = 1                                           'init loop index
      Select Case sType
      Case "="                                        'search for qty = & length =
         'this search returns the 1st qty & length match found
         Do Until usrMACHlist(i).SchedDate > fillDate
            If usrMACHlist(i).BldQnty <> usrMACHlist(i).Qnty And _
               (usrMACHlist(i).BlankLength > Length1 And _
                usrMACHlist(i).BlankLength <= Length2) And _
               usrMACHlist(i).Qnty = umQty Then
               fillsearch = i
               Exit Do
            Else
               i = i + 1                              'incr loop index
            End If
         Loop  'machine list loop
      Case "<"
         'this search returns the largest quantity < unmatched quantity
         mIDX = 0
         mQTY = 0
         Do Until usrMACHlist(i).SchedDate > fillDate
            If usrMACHlist(i).BldQnty <> usrMACHlist(i).Qnty And _
               (usrMACHlist(i).BlankLength > Length1 And _
                usrMACHlist(i).BlankLength <= Length2) And _
               usrMACHlist(i).Qnty < umQty Then
               If usrMACHlist(i).Qnty > mQTY Then
                  mQTY = usrMACHlist(i).Qnty
                  mIDX = i
               End If
            End If
            i = i + 1                                 'incr loop index
         Loop  'machine list loop
         fillsearch = mIDX
      Case ">"
         'this search returns the smallest quantity > unmatched quantity
         mIDX = 0
         mQTY = 30000
         Do Until usrMACHlist(i).SchedDate > fillDate
            If usrMACHlist(i).BldQnty <> usrMACHlist(i).Qnty And _
               (usrMACHlist(i).BlankLength > Length1 And _
                usrMACHlist(i).BlankLength <= Length2) And _
               usrMACHlist(i).Qnty > umQty Then
               If usrMACHlist(i).Qnty < mQTY Then
                  mQTY = usrMACHlist(i).Qnty
                  mIDX = i
               End If
            End If
            i = i + 1                                 'incr loop index
         Loop  'machine list loop
         fillsearch = mIDX
      End Select
   End If
End Function

'---------------------------------------------------------------------------
'Name:      autoOPT()
'Returns:
'Requires:
'Discrip:   this sub automates the optimization process
'Notes:
'---------------------------------------------------------------------------
Public Sub autoOPT()
   frmBatch.lblStatus.Caption = "Optimizing..."
   frmBatch.lblStatus.Visible = True
   DoEvents
   clrLocalTBL "tblBatch2"
   bldBatch2
   optBatch2
   displaySTATS
   frmBatch.lblStatus.Visible = False
End Sub  'autoOPT
'---------------------------------------------------------------------------
'Name:      updateSTATS()
'Accepts:   dblBLANKlen, = length of the bar being added to batch
'Returns:
'Requires:
'Discrip:   this sub displays the batch stats
'Notes:
'---------------------------------------------------------------------------
Public Sub updateSTATS(dblBLANKlen As Double)
   
   arrOffFall(20) = usrBatch.TotOFF - dblBLANKlen           'recalc total off-fall
   usrBatch.TotLen = usrBatch.TotLen + dblBLANKlen          'recalc total used length
   
   dblTotPartsLen2 = dblTotPartsLen2 + dblBLANKlen          'jng
   
   
   'calc % optimization
   If usrBatch.Blanks > 0 Then
      usrBatch.Opt = usrBatch.TotLen / (usrBatch.Blanks * StockLength)
      usrBatch.Opt = usrBatch.Opt * 100
      
      intBatchBlanks2 = usrBatch.Blanks      'jng
      dblBatchLngth2 = (usrBatch.Blanks * StockLength)   'jng 03/31/08
      dblBatchOpt2 = dblTotPartsLen2 / dblBatchLngth2     'jng 03/31/08 added for optimization tracking
      dblBatchOpt2 = dblBatchOpt2 * 100                 'jng
      'jng - program should be re-calculating longest piece of offall for batch screen
      'jng - after sucessful fill-search on batch form, but it does not.
      'jng - bug in original program
      
   Else
      usrBatch.Opt = 0
   End If
     
   usrBatch.TotOFF = arrOffFall(20)

End Sub  'updateSTATS

'---------------------------------------------------------------------------
'Name:      displaySTATS()
'Returns:
'Requires:
'Discrip:   this sub displays the batch stats
'Notes:
'---------------------------------------------------------------------------
Public Sub displaySTATS()
   frmBatch.txtMaterial.Text = usrBatch.Mat
   frmBatch.txtWidth.Text = usrBatch.Width
   frmBatch.txtTotBlnks.Text = usrBatch.Blanks
   frmBatch.txtLongOff.Text = usrBatch.LongOFF
   frmBatch.txtTotOff.Text = usrBatch.TotOFF
   frmBatch.txtOPT.Text = usrBatch.Opt
End Sub  'displaySTATS

'------------------------------------------------------------------------------
'Name:      addBatchQue
'Accepts:   usrMachList index of seq to be added.
'Returns:   results integer, intresult = 0, part NOT added
'                                        1 , added OK
'Requires:
'Discrip:   This Sub adds a part to the Batch Que
'Notes:
'------------------------------------------------------------------------------

Public Sub addBatchQue(idx As Integer, intResult As Integer)
   Dim conLocal As Connection
   Dim adoRS As ADODB.Recordset
   Dim strSQL As String
   
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                          
   '------------------------------- Add the Part --------------------------
   '--- make recordset
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblBatchQue"
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
         
   With adoRS
      .AddNew                                         'add a record
      '--- fill out the record
      !FullOrder = usrMACHlist(idx).FullJobNum
      ![Order Number] = usrMACHlist(idx).Job
      !Release = usrMACHlist(idx).Rel
      !Item = usrMACHlist(idx).Item
      ![Sequence Number] = usrMACHlist(idx).Seq
      ![Scheduled Ship Date] = usrMACHlist(idx).ShipDate
      !ShipDate = usrMACHlist(idx).SchedDate
      !Quantity = usrMACHlist(idx).BldQnty            'set qty = desired qty = bldqty
      !BldQnty = 0                                    'reset build qty
      !Phase = usrMACHlist(idx).Phase
      !Leg = usrMACHlist(idx).Leg
      !Stack = usrMACHlist(idx).Stack
      !BarType = usrMACHlist(idx).BarType
      !Material = usrMACHlist(idx).Material
      !BarWidth = usrMACHlist(idx).BarWidth
      !BlankLength = usrMACHlist(idx).BlankLength
      !RunLength = usrMACHlist(idx).RunLength
      !E1figure = usrMACHlist(idx).E1fig
      !E1dimension = usrMACHlist(idx).E1dim
      !E2figure = usrMACHlist(idx).E2fig
      !E2dimension = usrMACHlist(idx).E2dim
      !Cdimension = usrMACHlist(idx).Cdim
      !C1dimension = usrMACHlist(idx).C1dim
      !Ddimension = usrMACHlist(idx).Ddim
      !D1dimension = usrMACHlist(idx).D1dim
      If usrMACHlist(idx).BldQnty = usrMACHlist(idx).Qnty Then
         !Build = "F"
      Else
         !Build = "P"
      End If
      !Status = "BQ"
      !Priority = usrMACHlist(idx).Priority
      !DTStamp = Now
      .Update                                         'update the recordset
      .Close
   End With
      
   intResult = 1                                      'successfull result
                                                
   Set adoRS = Nothing                                'unload recordset & connection
   conLocal.Close
   Set conLocal = Nothing
   
End Sub 'addBatchQue

'---------------------------------------------------------------------------
'Name:      runQuery0()
'Accepts:
'Returns:
'Requires:  conDB, conn to central db
'Discrip:   this sub executes a query with no paramaters.
'Notes:
'---------------------------------------------------------------------------
Public Sub runQuery0(strQry As String)

Dim conDB As Connection
Dim adoRS As ADODB.Recordset           'recordset
Dim strSQL As String
Dim cmd As ADODB.Command

   'make connection to local database
   Set conDB = New Connection
   conDB.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"

   Set adoRS = New ADODB.Recordset                    'init recordset object
   
   Set cmd = New ADODB.Command                        'init command object
   cmd.ActiveConnection = conDB
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = strQry
   
   Set adoRS = cmd.Execute                            'execute the command
          
   Set adoRS = Nothing                                'unload recordset
   Set cmd = Nothing
   conDB.Close
   Set conDB = Nothing
   
End Sub  'runQuery0

'---------------------------------------------------------------------------
'Name:      copyBatchQue()
'Accepts:
'Requires:  local db
'Discrip:   this sub appends the records in batch que to the records in exec que.
'Notes:     test only...this sub not used in code!
'---------------------------------------------------------------------------
Public Sub copyBatchQue()

Dim conDB As Connection                'db connection
Dim adoRS As ADODB.Recordset           'recordset
Dim strSQL As String                   'SQL string
   
   'make connection to local database
   Set conDB = New Connection
   conDB.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
   
   '--- make recordset
   Set adoRS = New ADODB.Recordset

   strSQL = "INSERT INTO tblExecQue " & _
            "([Order Number], Release, Item, [Sequence Number], " & _
            "[Scheduled Ship Date], ShipDate, Quantity, BldQnty, " & _
            "Phase, Leg, Stack, BarType, Material, BarWidth, " & _
            "BlankLength, RunLength, E1figure, E1dimension, E2figure, " & _
            "E2dimension, Cdimension, C1dimension, Ddimension, D1dimension, " & _
            "FullOrder, Build, Status, Priority, DTStamp )" & _
            "SELECT [Order Number], Release, Item, [Sequence Number], " & _
            "[Scheduled Ship Date], ShipDate, Quantity, BldQnty, " & _
            "Phase, Leg, Stack, BarType, Material, BarWidth, BlankLength, " & _
            "RunLength, E1figure, E1dimension, E2figure, E2dimension, " & _
            "Cdimension, C1dimension, Ddimension, D1dimension, " & _
            "FullOrder, Build, Status, Priority, DTStamp " & _
            "FROM tblBatchQue"
                    
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic
         
   Set adoRS = Nothing                          'unload recordset
   conDB.Close
   Set conDB = Nothing
      
End Sub  'copyBatchQue

'---------------------------------------------------------------------------
'Name:      displayButtons()
'Returns:
'Requires:
'Discrip:   this sub updates the visibility of the command buttons for the
'           batch screen.
'Notes:
'---------------------------------------------------------------------------
Public Sub displayButtons()
   If usrBatch.Lock = True Then   '------------------ batch locked
      frmBatch.cmdView.Visible = False
      frmBatch.cmdAdd.Visible = False
      frmBatch.cmdRemove.Visible = False
      frmBatch.cmdChange.Visible = False
      
      frmBatch.cmdClear.Visible = True
      frmBatch.cmdInitBatch.Visible = False
      frmBatch.cmdFill.Visible = False
      frmBatch.cmdAccept.Visible = False
      frmBatch.cmdSubmit.Visible = True
      Exit Sub
   End If
   
   If usrBatch.Fill = True Then   '------------------ batch filled
      frmBatch.cmdView.Visible = False
      frmBatch.cmdAdd.Visible = False
      frmBatch.cmdRemove.Visible = False
      frmBatch.cmdChange.Visible = False
      
      frmBatch.cmdClear.Visible = True
      frmBatch.cmdInitBatch.Visible = False
      frmBatch.cmdFill.Visible = False
      frmBatch.cmdAccept.Visible = True
      frmBatch.cmdSubmit.Visible = False
      Exit Sub
   End If
   
   '--- initial batch selection made
   If usrBatch.InitBatch = True Then    '------------- init batch made
      frmBatch.cmdView.Visible = True
      frmBatch.cmdAdd.Visible = True
      frmBatch.cmdRemove.Visible = True
      frmBatch.cmdChange.Visible = True
      
      frmBatch.cmdClear.Visible = True
      frmBatch.cmdInitBatch.Visible = False
      frmBatch.cmdFill.Visible = True
      frmBatch.cmdAccept.Visible = True
      frmBatch.cmdSubmit.Visible = False
      Exit Sub
   End If
   
   '--- initial pick made
   If usrBatch.InitPick = True Then    '------------- init pick made
      frmBatch.cmdView.Visible = True
      frmBatch.cmdAdd.Visible = True
      frmBatch.cmdRemove.Visible = True
      frmBatch.cmdChange.Visible = True
      
      frmBatch.cmdClear.Visible = True
      frmBatch.cmdInitBatch.Visible = True
      frmBatch.cmdFill.Visible = True
      frmBatch.cmdAccept.Visible = True
      frmBatch.cmdSubmit.Visible = False
      Exit Sub
   Else                           '------------------ init pick NOT made
      frmBatch.cmdView.Visible = False
      frmBatch.cmdAdd.Visible = True
      frmBatch.cmdRemove.Visible = True
      frmBatch.cmdChange.Visible = True
      
      frmBatch.cmdClear.Visible = True
      frmBatch.cmdInitBatch.Visible = True
      frmBatch.cmdFill.Visible = False
      frmBatch.cmdAccept.Visible = False
      frmBatch.cmdSubmit.Visible = False
      Exit Sub
   End If
      
End Sub  'displayButtons

'---------------------------------------------------------------------------
'Name:      addHoldAssign()
'Accepts:   usrPart, the part to be added
'Returns:   intResult, 0 = not added, 1=added OK
'Requires:
'Discrip:   This Sub adds a part to the hold assignemnt table
'Notes:
'---------------------------------------------------------------------------
Public Sub addHoldAssign(usrPart As part, intResult As Integer)

Dim conDB As Connection
Dim adoRS As ADODB.Recordset
Dim strSQL As String

   '--- make connection to local database
   Set conDB = New Connection
   conDB.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for History que
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblHoldAssign " & _
            "WHERE [Order Number] = '" & usrPart.Job & "' " & _
            "AND Release = '" & usrPart.Rel & "' " & _
            "AND Item = " & usrPart.Item & " " & _
            "AND [Sequence Number] = " & usrPart.Seq
   
   adoRS.Open strSQL, conDB, adOpenStatic, adLockOptimistic
   
   With adoRS
      If .RecordCount > 0 Then
         !BldQnty = usrPart.Qnty                'store latest qty
         .Update                                'update record
      Else
      .AddNew                                   'add a record
         '--- fill out the record ---- --------- only need the JRIS & Qty
         '!FullOrder = usrPart.FullJobNum
         ![Order Number] = usrPart.Job
         !Release = usrPart.Rel
         !Item = usrPart.Item
         ![Sequence Number] = usrPart.Seq
         '![Scheduled Ship Date] = usrPart.ShipDate
         ![Quantity] = usrPart.Qnty
         '!BldQnty = 1
         '!Phase = usrPart.Phase
         '!Leg = usrPart.Leg
         '!Stack = usrPart.Stack
         '!BarType = usrPart.BarType
         '!Material = usrPart.Material
         '!BarWidth = usrPart.BarWidth
         '!BlankLength = usrPart.BlankLength
         '!RunLength = usrPart.RunLength
         '!E1figure = usrPart.E1fig
         '!E1dimension = usrPart.E1dim
         '!E2figure = usrPart.E2fig
         '!E2dimension = usrPart.E2dim
         '!Cdimension = usrPart.Cdim
         '!C1dimension = usrPart.C1dim
         '!Ddimension = usrPart.Ddim
         '!D1dimension = usrPart.D1dim
         '!Build = usrPart.Build
         '!Status = usrPart.Status
         '!Priority = usrPart.Priority
         '!Truck = usrPart.Truck
         '!Machine = usrPart.Machine
         '!DTStamp = usrPart.DTS
         .Update                                   'update the recordset
      End If
   End With
         
   '--- clean up
   adoRS.Close                                     'unload recordset & connection
   Set adoRS = Nothing
   conDB.Close
   Set conDB = Nothing
   intResult = 1                                   'return result
   
End Sub  'addHoldAssign

