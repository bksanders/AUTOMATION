Attribute VB_Name = "S6comm"
'----------------------------------------------------------------------------
'Name:      S6comm.bas
'Descrip:   This file contains all the neccessary subroutines and functions
'           to implement series communications between a PC and a CCM module
'           in a Series6 PLC.
'Notes: (1) This implements the series Master-Slave communications protocal
'           as set forth in Ch#4 of the CCM manual.
'       (2) The PC acts as the Master, the PLC as the Slave
'       (3) The PLC's ID = 01
'       (4) The PC's ID = 05
'       (5) The PC uses comm 1
'       (6) The PLC-CCM uses port J2(9pin D-shell)
'       (7) Comm Paramaters = 4800Baud, 8-bit, No Parity, 1 Stop bit
'       (8) Communication is implemented using the mscomm control,MSComm1,
'           which must be located on the main form.
'       (9) communication timeouts are implemented, using a timer,tmrTOUT,
'           which must be located on the main form.
'       (10) All timeout times are taken from table 4.5 of the CCM manual.
'       (11) S6comm referenced 3 TextBox Controls for Comm Feedback:
'            - txtStat: serves as a status window
'            - txtTX:   displays the TX buffer
'            - txtRX:   displays the RX buffer
'       (11) The above requirments are NOT repeated inside the individual
'            subs/functs of this file.  They are simply refered to as requiring
'            MSComm1.
'----------------------------------------------------------------------------
' USAGE:  To use S6comm.bas with your VB app.
'        - If your main form is not frmMain....then search & replace "frmMain"
'          w/ your forms name.
'        - create the following controls on the main form
'           - MSComm1 as an mscomm control
'           - tmrTOUT as a timer control
'           - txtStat: as a textbox control
'           - txtTX:  as a textbox control
'           - txtRX: as a textbox control
'        - If you do not wish to use the feedback windows...you can make them
'          invisible on your form...or comment out any reference to them in
'          S6comm.bas
'----------------------------------------------------------------------------
'Descrip:   This specific application has modified to work with the GE Selmer
'           Remmele Bar Machine.  Modifications are listed below:
'       (1) setmode() has been modified to send the proper "Mode" string.
'----------------------------------------------------------------------------

Option Explicit

Public NUL As String    '--- comm protocol ASCII Control Codes
Public SOH As String
Public STX As String
Public ETX As String
Public EOT As String
Public ENQ As String
Public ACK As String
Public NAK As String
Public ETB As String

Public timeout As Boolean           'comm timeout flag

'----------------------------------------------------------------------------
' Name:     TxPLC
' Accepts:  strToSend = string to send
' Returns:  None
' Requires: MSComm1
' Descrip:  This sub receives a string, and sends it out the open comm port.
'----------------------------------------------------------------------------

Public Sub TxPLC(strToSend As String)
   
   frmRun.MSComm1.Output = strToSend               'send the string
   'frmRun.txtTX.Text = strToSend                   'OP Feedback
   'Print #1, "TX:" & strToSend                        'dbugging
   
End Sub

'---------------------------------------------------------------------------
'Name:      TxENQ
'Accepts:   None
'Returns:   result = true if Enquiry Tx'd and ACK Rx'd OK
'Requires:  mscomm1
'Discrip:   This sub xmits the Normal Enquiry String.  It will send the
'           string 32 times before quiting.
'           If it sends the string 32 times w/o getting an ACK, then it
'           will term. the comm session by sending the EOT char.
'---------------------------------------------------------------------------
Public Sub TxENQ(result As Boolean)
   '--- Variable Declarations
   Dim strOutput As String
   Dim AckOK As Boolean                            'ACK flag
   Dim loopCNT As Integer                          'send count
   
   '---build the enquiry string
   strOutput = "N" & Chr$(33) & ENQ
   
   loopCNT = 1                                     'init loop
   Do   '--- Enq Xmit loop
      TxPLC strOutput                              'xmit the string
        
      RxENQ AckOK                                  'check for response
      If AckOK Then
         result = True
      End If
      
      loopCNT = loopCNT + 1                        'incr loop count
   Loop Until result = True Or loopCNT > 32
   
   If result = False Then                          'if NO ACk
      TxPLC EOT                                    'send the EOT to Term Comm session
   End If
   
End Sub  'TxENQ

'----------------------------------------------------------------------------
' Name:     RxENQ
' Accepts:  result, should be set to FALSE prior to calling this sub.
' Returns:  result = true if ACK received, False if NOT
' Requires: timeout, global timeout flag
'           tmrTOUT, a timeout timer located on the current form.
'           mscomm1
' Descrip:  This sub receives the response to a Normal Enquiry String.
'           If an ACK is received, before the timeout,result is set TRUE.
'           Otherwise it is false.  The timeout timer interval is set for the
'           time specified by the CMM protocal(See cond#1, Table 4.5)'
'----------------------------------------------------------------------------

Public Sub RxENQ(result As Boolean)
   '----- global variables
   'timeout is a global boolean used as a timeout flag
   
   '----- Variable Declarations
   Dim strInput As String
   
   '--- setup
   frmRun.tmrTOUT.Interval = 800                  'set timeout
   frmRun.tmrTOUT.Enabled = True                  'enable TimeOut Timer
   timeout = False                                 'reset timeout flag
   
   Do                                              '---look for ACK
      DoEvents                                     'check for other events
      strInput = strInput & frmRun.MSComm1.Input  'get the input
      'frmRun.txtRX.Text = strInput                'OP feedback
      If InStr(1, strInput, ACK, 1) > 0 Then       'ACK Recv'd
         result = True
      End If
   Loop Until result Or timeout                    'loop until ack or timeout
   
   'Print #1, "RX:" & strInput                      'debug
   
   '--- cleanup
   frmRun.tmrTOUT.Enabled = False                 'disable TimeOut Timer
End Sub  'RxENQ

'---------------------------------------------------------------------------
'Name:      TxHDR
'Accepts:   strHDR, header string to be sent
'Returns:   result = true if ENQ and HDR ACK'd
'Requires:  mscomm1
'Discrip:   This sub xmits a Header string.
'           It will send the header3 times before quiting.
'           If the PLC ACK's the header everything is OK.
'           If no response, bad char recv'd or NAK'd 3 times then send the
'           EOT to end the comm session.
'---------------------------------------------------------------------------

Public Sub TxHDR(strHDR As String, result As Boolean)
   '--- Variable Declarations
   Dim EnqOK As Boolean
   Dim strACK As String
   Dim intNAK As Integer
   Dim badACK As Boolean
   
   EnqOK = False
   TxENQ EnqOK                                        'xmit the enquiry
   
   If EnqOK = True Then                               'PLC ACK'd Enq
      intNAK = 0                                      'reset NAK count
      badACK = False                                  'reset badACK flag
      Do '--- send the header
         TxPLC strHDR                                 'xmit the header
         
         'frmRun.txtStat.Text = "TxHdr: Waiting for ACK..."
         strACK = ""                                  'reset ack result
         RxACK 2000, strACK                           'check for response
         Select Case strACK                           'handle response
         Case ""                                      'timeout
            'frmRun.txtStat.Text = "TxHdr: HDR Timeout!"
            badACK = True
         Case ACK                                     'ACK Recv'd
            'frmRun.txtStat.Text = "TxHdr: ACK recv'd!"
            result = True
         Case NAK                                     'NAK Recv'd
            intNAK = intNAK + 1
            'frmRun.txtStat.Text = "TxHdr: NAK recv'd!"
         Case Else                                    'bad char Recv'd
            'frmRun.txtStat.Text = "TxHdr: Bad Char recv'd!"
            badACK = True
         End Select
      Loop Until result = True Or intNAK > 2 Or badACK = True

   Else
      'frmRun.txtStat.Text = "TxHdr: No Response to ENQ!"
   End If
   
   If result = False Then                          'NO ACK recv'd
      TxPLC EOT                                    'xmit the EOT to end comm session
   End If
   
End Sub  'TxHDR

'----------------------------------------------------------------------------
' Name:     RxACK
' Accepts:  intTimeOut, desired timeout
' Returns:  ACKchar, the response character read, returns "" on timeout
' Requires: timeout, global timeout flag
'           tmrTOUT, a timeout timer located on the current form.
'           MSComm1
' Descrip:  This sub gets a single char response.
'----------------------------------------------------------------------------

Public Sub RxACK(intTOUT As Integer, ACKchar As String)
   '--- global variables
   'timeout                            'timeout flag
      
   '--- variable declarations
   Dim strTemp As String               'temp string
   
   '--- setup
   frmRun.MSComm1.InputLen = 1                       'set read for 1 char
   frmRun.tmrTOUT.Interval = intTOUT                 'set timeout
   timeout = False                                   'reset timeout flag
   frmRun.tmrTOUT.Enabled = True                     'enable timer
   
   Do '--- check for response
      DoEvents                                        'handle events
   Loop Until frmRun.MSComm1.InBufferCount > 0 Or timeout
   ACKchar = frmRun.MSComm1.Input                    'get the response
   
   Select Case ACKchar                                '---debug
   Case ACK
   strTemp = "ACK"
   Case NAK
      strTemp = "NAK"
   Case EOT
      strTemp = "EOT"
   Case Else
      strTemp = ACKchar
   End Select
   'Print #1, "RX:" & strTemp
   
   '--- cleanup
   frmRun.tmrTOUT.Enabled = False                    'disable timer
   frmRun.MSComm1.InputLen = 0                       'set read for entire buffer

End Sub  'RxACK

Public Sub TxBLK(strBlock As String, result As Boolean)
   '--- Variable Declarations
   Dim gotACK As Boolean               'ACK flag
   Dim intNAK As Integer               'NAK counter
   Dim badACK As Boolean               'bad response flag
   Dim strChar As String               'input char

     
   Do    '--- send the block
      'frmRun.txtStat.Text = "TxBLK:Sending Clear String..." 'OP feedback
      TxPLC strBlock                                  'send the string
      'frmRun.txtTX.Text = strBlock                    'display TX string
       
      '--- get response, expecting an ACK
      gotACK = False                                  'setup for loop
      badACK = False
      intNAK = 0
      strChar = ""
      'set timeout = 20s per cond#6, table 4.5
      RxACK 20000, strChar                            'get the response
      
      '--- handle response
      Select Case strChar
      Case ""                                         'timeout
         'frmRun.txtStat.Text = "TxBLK: ACK Timeout!"  'OP feedback
         badACK = True
      Case ACK                                        'ACK Recv'd
         'frmRun.txtStat.Text = "TxBLK: ACK recv'd!"   'OP feedback
         gotACK = True
      Case NAK                                        'NAK Recv'd
         intNAK = intNAK + 1
         'frmRun.txtStat.Text = "TxBLK: NAK recv'd!"   'OP feedback
      Case Else                                       'bad char Recv'd
         'frmRun.txtStat.Text = "TxBLK: Bad Char recv'd!"   'OP feedback
         strChar = frmRun.MSComm1.Input               'empty buffer
         'frmRun.txtRX.Text = strChar
         badACK = True
      End Select
   Loop Until gotACK = True Or intNAK > 2 Or badACK = True  'send block

   TxPLC EOT                                          'send EOT
   
   result = gotACK                                    'set the result
   
   frmRun.MSComm1.RThreshold = 3                      'Set the Rx Threshold
  
End Sub  'TxBLK

'----------------------------------------------------------------------------
' Name:     RxBLK
' Accepts:  intLength, # of expected characters
' Returns:  strBlock, the desired string, returns NULL if no string recv'd
' Requires: timeout, global timeout flag
'           tmrTOUT, a timeout timer located on the current form.
'           MSComm1
' Descrip:  This sub recv's a string from the PLC, via the comm port.
' Note:     Calling program should verify that the string recv'd is ok by
'           checking its length against what was expected.
'----------------------------------------------------------------------------

Public Sub RxBLK(intLength, strBlock)
   '--- Variable Declarations
   Dim gotACK As Boolean               'ACK flag
   Dim intNAK As Integer               'NAK counter
   Dim badACK As Boolean               'bad response flag
   Dim gotBLK As Boolean               'good Block flag
   Dim strChar As String               'input char
   
   '--- setup
   intNAK = 0
   gotBLK = False
   'set timeout = 8340ms per cond#7, table 4.5
   frmRun.tmrTOUT.Interval = 8340                    'set timeout
   frmRun.tmrTOUT.Enabled = True                     'enable timer
   timeout = False                                    'reset timeout flag
   
   Do    '--- get the block
      DoEvents                                        'handle events
      If frmRun.MSComm1.InBufferCount >= intLength Then
         frmRun.tmrTOUT.Enabled = False              'disable timer
         strBlock = frmRun.MSComm1.Input             'get the block string
         
         'Print #1, "RX:" & strBlock                   'debug
         
         '--- verify the block is OK
         If Len(strBlock) = intLength And _
            Mid(strBlock, 1, 1) = STX And _
            Mid(strBlock, Len(strBlock) - 1, 1) = ETX Then
            gotBLK = True                             'block is OK
         Else                                         'bad block
            intNAK = intNAK + 1                       'incr NAK counter
            If intNAK < 3 Then
               TxPLC NAK                              'send a NAK
               frmRun.tmrTOUT.Enabled = True         'enable timer
            End If
         End If
      End If
   Loop Until gotBLK Or intNAK = 3 Or timeout
      
   If gotBLK = True Then                              'block is OK
      TxPLC ACK                                       'send an ACK
      
      '--- check for EOT
      strChar = ""                                    'reset input char
      'set timeout = 800ms per cond#8, table 4.5
      RxACK 800, strChar                              'get the response
      If strChar <> EOT Then                          'NOT EOT
         strChar = frmRun.MSComm1.Input               'empty buffer
         TxPLC EOT                                    'send EOT
      Else                                            'GOT the EOT
         TxPLC EOT                                    'send EOT
      End If
   Else                                               'bad block
      TxPLC EOT                                       'send the EOT
   End If   'blk ok
         
   '--- cleanup
   
End Sub  'RxBLK

'----------------------------------------------------------------------------
' Name:     calcLRC
' Accepts:  strData, the string to be sent
' Returns:  the LRC char
' Requires: None
' Descrip:  This function recv's a string from the calling program and
'           calculates its corresponding LRC error checking byte.
' Note:     LRC = Longitudinal Redundancy Checking
'----------------------------------------------------------------------------

Public Function calcLRC(strData As String) As String
   '--- Variable Declarations
   Dim i As Integer                                'loop index
   Dim tmpByte As Byte                             'byte for indv. chars
   Dim char As String
   Dim xorByte As Byte                             'byte for accumulation of LRC
   
   tmpByte = 0                                     'reset the bytes
   xorByte = 0
   
   For i = 1 To Len(strData)                       'loop thru string
      char = Mid(strData, i, 1)                    'get the next char
      tmpByte = Asc(char)                          'conv the char to a byte
      xorByte = xorByte Xor tmpByte                'xor with Accumulator
   Next i   'end of string loop
   
   calcLRC = Chr$(xorByte)                         'return the LRC
   
End Function   'calcLRC

'---------------------------------------------------------------------------
'Name:      bldHdr
'Accepts:   HdrType, 1 char string representing the Header Type to be built.
'Returns:   The header string
'Requires:  None
'Discrip:   This function accepts the header type char and generates the
'           entire header string to be sent, including the STX, ETB and LRC.
'Notes:     None
'---------------------------------------------------------------------------

Public Function bldHdr(HdrType As String) As String
   '--- Variable Declarations
   Dim LRCData As String
   Dim LRC As String
   Dim strHDR As String
   
   strHDR = " "                                 'init header string
   Select Case HdrType                          'Evaluate HdrType
      Case "N"
         LRCData = "0101179F000205"             '6047-new part status
      Case "I"
         LRCData = "01811771005E05"             '6001-dim/stamp info
      Case "E"
         LRCData = "01011A47000405"             '6727-error status
      Case "S"
         LRCData = "01811A47000205"             '6727-service plc error
      Case "P"
         LRCData = "01011A3F000C05"             '6719-part complete check
      Case "C"
         LRCData = "01811A44000205"             '6724-clear complete
      Case "F"
         LRCData = "01811A44000205"             '6724-batch finished
      Case "R"
         LRCData = "01010119000205"             '281-check mode change ready
      Case "Z"
         LRCData = "01810101003205"             '257-change mode/constants
      Case Else   ' Other values.
         MsgBox ("Illegal Header Type!")
         bldHdr = " "
         Exit Function
   End Select

   LRC = calcLRC(LRCData)                       'get the LRC for the header
   
   strHDR = SOH & LRCData & ETB & LRC           'build the header string
   bldHdr = strHDR                              'return as function result
   
End Function   'bldHdr



'---------------------------------------------------------------------------
'Name:      chkMode
'Accepts:   None
'Returns:   ready = true if PLC ready for mode change
'Requires:  None
'Discrip:   This sub sends the R-HDR, and gets the PLC's reply to the
'           the R-HDR.
'Notes: (1) The PLC should respond w/ a 5 char string
'           STX + ? + ? + ETB + LRC.  If 2nd & 3rd char are ?'s then
'           The PLC is ready for a mode change.
'---------------------------------------------------------------------------

Public Sub chkMode(ready As Boolean)
   '--- Variable Declarations
   Dim HdrOK As Boolean                            'header result
   Dim strHeader As String                         'header string
   Dim strReply As String                          'reply string
   Dim char2, char3 As String                      'status characters
    
   strHeader = bldHdr("R")                         'build the check mode header
    
   '--- send the header
   'frmRun.txtStat.Text = "chkMode:Sending R-HDR..."   'OP feedback
   HdrOK = False
   TxHDR strHeader, HdrOK                          'xmit the header

   If HdrOK = True Then                            'if header Ack'd
      '--- get the reply
      'frmRun.txtStat.Text = "chkMode:Waiting for R-HDR block..."
      
      strReply = ""
      'timout set to 8340 per cond#7 of Table 4-5.
      RxBLK 5, strReply                            'get the reply string
      
      '--- handle reply
      If Len(strReply) = 5 Then                    'if proper reply received
         char2 = Mid(strReply, 2, 1)               'get the 2nd char
         char3 = Mid(strReply, 3, 1)               'get the 3rd char
         If char2 = "?" And char3 = "?" Then       'if two ?? then PLC ready for Mode change
            ready = True
         End If   '??
      Else
         'frmRun.txtStat.Text = "chkMode: Bad/No R-HDR block!"   'OP feedback
      End If 'reply OK
   Else                                               'if header xmit failed
      'frmRun.txtStat.Text = "chkMode:R-HDR Failed!"    'OP feedback
   End If 'hdrok
   
End Sub  'chkMode

'---------------------------------------------------------------------------
'Name:      setMode
'Accepts:   charMode, a 1 char string: A=auto, M=Manual
'Returns:   result = true if the PLC accepted mode change
'Requires:  None
'Discrip:   This sub sets the Auto/Manual mode in the PLC.  It also sets
'           various system parameters.
'Notes: (1) Sending the mode string writes spaces into PLC register 281, if the
'           PLC accepted the mode change it will write ?? back to the register.
'           This is how we confirm acceptance of the mode change.
'       (2) .The sub 1st checks if the PLC is ready for a mode change.
'           .If so, it sends the mode change, "Z", header.
'           .If the PLC ACK receipt of the "Z" header the sub constructs
'            the mode change string and sends it to the PLC.
'           .Once the PLC ACK's the mode change string, the sub sends an EOT
'            to end the comm session.
'           .If no/bad response is received then the sub sends and EOT to
'            end the session anyway.
'---------------------------------------------------------------------------
'Revision Notes:
'           - 3/10/03   modified for use w/ RemBar Mode String
'---------------------------------------------------------------------------

Public Sub setMode(charMode As String, result As Boolean)
   '--- Variable Declarations
   Dim PLCrdy As Boolean               'PLC ready flag
   Dim HdrOK As Boolean                'header result
   Dim BlkOK As Boolean                'block result
   Dim strHeader As String             'header string
   Dim strACK As Integer               'ACK result
   Dim intNAK As Integer               'NAK counter
   Dim badACK As Boolean               'Bad ACK flag
   Dim dblTemp As Double               'temp float for calcs
   Dim strTemp As String               'temp string for calcs
   Dim strSawKerf As String            '6 char sawkerf string
   Dim strCP1 As String                '6 char coin press string
   Dim strCP2 As String                '6 char coin press string
   Dim strCP3 As String                '6 char coin press string
   Dim strCP4 As String                '6 char coin press string
   Dim strManBendPress As String       '6 char manual bend pressure string
   Dim strTx As String                 'complete Tx string
   Dim LRC As String                   'LRC byte
   Dim strMode As String               'mode string
   Dim decpos As Long                  'location of decimal place in string
   
   PLCrdy = False
   chkMode PLCrdy                                     'check if PLC ready for mode change
   
   If PLCrdy = True Then                              'PLC is ready
      strHeader = bldHdr("Z")                         'build the mode change header
      'frmRun.txtStat.Text = "Sending Z-HDR..."         'OP feedback
      
      HdrOK = False
      TxHDR strHeader, HdrOK                          'xmit the header
      
      If HdrOK = True Then    '--- if header xmitted ok
         '---------- construct the mode string
         'frmRun.txtStat.Text = "Constructing Mode String..."    'OP feedback
         
         '--- construct 6char SawKerf
         strSawKerf = Val(SawKerf)
         strSawKerf = Format(strSawKerf, "000000")
         
         '--- construct 6char Coin Pressures
         strCP1 = Val(arrCoinPress(1))
         strCP1 = Format(strCP1, "000000")
        
         strCP2 = Val(arrCoinPress(2))
         strCP2 = Format(strCP2, "000000")
        
         strCP3 = Val(arrCoinPress(3))
         strCP3 = Format(strCP3, "000000")
         
         strCP4 = Val(arrCoinPress(4))
         strCP4 = Format(strCP4, "000000")
         
         '--- construct 6char Manual Bend Pressure
         strManBendPress = Val(ManBendPress)
         strManBendPress = Format(strManBendPress, "000000")
         
         '--- construct the mode string
         strMode = charMode & " " & _
                   strSawKerf & "  " & _
                   strManBendPress & "00" & _
                   strCP1 & "  " & _
                   strCP2 & "  " & _
                   strCP3 & "  " & _
                   strCP4 & "  "
                   
         LRC = calcLRC(strMode)                       'calc the LRC
         strTx = STX & strMode & ETX & LRC            'build the Tx string
         
         TxBLK strTx, BlkOK                           'send the string/block
         
         'Sending the mode string writes spaces into PLC register 281, if the
         'PLC accepted the mode change it will write ?? back to the register
         PLCrdy = False
         chkMode PLCrdy                               'check if PLC ready
         result = PLCrdy                              'return mode change result
           
      Else                                            'if header failed
         'frmRun.txtStat.Text = "setMode: Z-HDR Failed!"
      End If 'hdrok
   
   Else                                               'PLC NOT ready
      'frmRun.txtStat.Text = "PLC NOT ready for MODE change!"
      MsgBox ("PLC NOT responding to MODE change!" & Chr$(13) & _
              "Verify Control Power is ON and Retry!" & Chr$(13) & _
              "If problem persists...contact Engineering. ")
   End If  'PLCrdy
   
End Sub  'setMode

'---------------------------------------------------------------------------
'Name:      chkNew
'Accepts:   none
'Returns:   ready = true if PLC ready for a NEW part.
'Requires:  none
'Discrip:   This sub determines if the PLC is ready for a new part.
'Notes: (1) The PLC will write "??" into register 6047 when ready for a
'           new part.
'---------------------------------------------------------------------------

Public Sub chkNew(ready As Boolean)
   '--- Variable Declarations
   Dim HdrOK As Boolean                'header result
   Dim strHeader As String             'header string
   Dim strReply As String              'reply string
   Dim char2, char3 As String          'status characters
     
   strHeader = bldHdr("N")                            'build the check New header
   
   '--- send the check NEW header
   HdrOK = False
   TxHDR strHeader, HdrOK                             'xmit the header
   
   '--- get the response block
   If HdrOK = True Then                               'if header Ack'd
      '--- get the reply
      'frmRun.txtStat.Text = "chkNew:Waiting for N-HDR block..."
      strReply = ""
      'timout set to 8340 per cond#7 of Table 4-5.
      RxBLK 5, strReply                               'get the reply string
      
      '--- handle reply
      If Len(strReply) = 5 Then                       'if proper reply received
         char2 = Mid(strReply, 2, 1)                  'get the 2nd char
         char3 = Mid(strReply, 3, 1)                  'get the 3rd char
         If char2 = "?" And char3 = "?" Then          'if two ?? then PLC ready for new part
            ready = True
         End If   '??
      Else
         'frmRun.txtStat.Text = "chkNew:Bad/No N-HDR block!"   'OP feedback
      End If 'reply OK
   Else                                               'if header xmit failed
      'frmRun.txtStat.Text = "chkNew:N-HDR Failed!"     'OP feedback
   End If 'hdrok
     
End Sub 'chkNew

'---------------------------------------------------------------------------
'Name:      TxPart
'Accepts:   strPart = part string, must be 94 bytes/chars in length
'Returns:   result = true if part string sent successfully
'Requires:  none
'Discrip:   This sub sends the part string to the PLC.
'Notes: (1) The PLC will write the part string into PLC registers starting
'           at register 6001.
'---------------------------------------------------------------------------

Public Sub TxPart(strPart As String, result As Boolean)
   '--- Variable Declarations
   Dim HdrOK As Boolean                'header result
   Dim strHeader As String             'header string
   Dim strACK As Integer               'ACK response
   Dim intNAK As Integer               'NAK counter
   Dim badACK As Boolean               'Bad ACK flag
   Dim LRC As String                   'LRC byte
   Dim strTx As String                 'block string to Tx
   'Dim intTemp As Integer              'temp integer for debugging
   
   strHeader = bldHdr("I")                            'build the clear part complete header
   
   HdrOK = False
   TxHDR strHeader, HdrOK                             'xmit the header
        
   If HdrOK = True Then    '--- if header Tx'd & ACK'd
      '--- construct the Tx string
      LRC = calcLRC(strPart)                          'calc the LRC
      strTx = STX & strPart & ETX & LRC               'build the TX string
      'intTemp = Len(strTx)                            'length of string for debugging
                  
      '--- capture part string for troubleshooting
      If frmRun.chkStampLog.Value = 1 Then
         Open "C:\RemBar\Temp\partstring.txt" For Append As #1
         Print #1, strTx
         Close #1
      End If
      
      TxBLK strTx, result                             'send the string/block
      
   Else                                               'if header failed
      'frmRun.txtStat.Text = "TxPart: I-HDR Failed!"
   End If 'hdrok
          
End Sub  'TxPart

'---------------------------------------------------------------------------
'Name:      chkComp
'Accepts:   none
'Returns:   result = true if PLC has a complete part
'           strStamp = 12 char stamp string of the completed part
'Requires:  none
'Discrip:   This sub determines if the PLC has a complete part.
'Notes: (1) The PLC will write the stamp string of the completed part
'           into register 6719-6724 when the part is complete.
'       (2) The PLC will write "00" into char's 12 & 13 of the reply string
'           if the part was advanced but NOT complete.---NOT UTILITZED
'       (3) If char's 12 & 13 = "??" then NO part complete or advance.
'---------------------------------------------------------------------------

Public Sub chkComp(strStamp As String, result As Boolean)
   '--- Variable Declarations
   Dim HdrOK As Boolean                'header result
   Dim strHeader As String             'header string
   Dim strReply As String              'reply string
   Dim char12, char13 As String        'status characters
    
   strHeader = bldHdr("P")                            'build the check mode header
   
   HdrOK = False
   TxHDR strHeader, HdrOK                             'xmit the header
   
   '--- get the response block
   If HdrOK = True Then                               'if header Ack'd
      '--- get the reply
      'frmRun.txtStat.Text = "chkComp:Waiting for P-HDR block..."
      strReply = ""
      'timout set to 8340 per cond#7 of Table 4-5.
      RxBLK 15, strReply                              'get the reply string
      
      '--- handle reply
      If Len(strReply) = 15 Then                      'if proper reply received
         char12 = Mid(strReply, 12, 1)                'get the 2nd char
         char13 = Mid(strReply, 13, 1)                'get the 3rd char
         If char12 = "?" And char13 = "?" Then        'if two ?? then NO Advance/Complete
            result = False
            strStamp = "FFFFFFFFFFFF"                 'impossible stamp#
         Else                                         'Not ??, thus adv/compl
            result = True
            strStamp = Mid(strReply, 2, 12)           'get the stamp string
         End If   '??
      Else
         'frmRun.txtStat.Text = "chkComp:Bad/No P-HDR block!"   'OP feedback
      End If 'reply OK
   Else                                               'if header xmit failed
      'frmRun.txtStat.Text = "chkComp:P-HDR Failed!"    'OP feedback
   End If 'hdrok
      
End Sub 'chkComp

'---------------------------------------------------------------------------
'Name:      clrComp
'Accepts:   none
'Returns:   result = true if complete successfully cleared.
'Requires:  none
'Discrip:   This sub lets the PLC know that it has processed the part complete.
'Notes: (1) The PLC will write the stamp string of the completed part
'           into register 6719-6724 when the part is complete.
'       (2) Writing "??" into 6724 ACK's processing of the completed part.
'---------------------------------------------------------------------------

Public Sub clrComp(result As Boolean)
   '--- Variable Declarations
   Dim HdrOK As Boolean                'header result
   Dim strHeader As String             'header string
   Dim strACK As Integer               'ACK response
   Dim intNAK As Integer               'NAK counter
   Dim badACK As Boolean               'Bad ACK flag
   Dim strClear As String              'clear string
   Dim LRC As String                   'LRC byte
   Dim strTx As String                 'Tx string
      
   strHeader = bldHdr("C")                            'build the clear part complete header
   
   HdrOK = False
   TxHDR strHeader, HdrOK                             'xmit the header

   If HdrOK = True Then    '--- if header Tx'd & ACK'd
      '--- construct the clear string
      'frmRun.txtStat.Text = "clrComp:Constructing Clear String..."    'OP feedback
      strClear = "??"                                 'build clear string
      LRC = calcLRC(strClear)                         'calc the LRC
      strTx = STX & strClear & ETX & LRC              'build the Tx string
                 
      TxBLK strTx, result                             'send the block
      
   Else                                               'if header failed
      'frmRun.txtStat.Text = "clrComp: C-HDR Failed!"
   End If 'hdrok
    
End Sub  'clrComp

'---------------------------------------------------------------------------
'Name:      chkErr
'Accepts:   None
'Returns:   intError = the error# that occured
'Requires:  None
'Discrip:   This subs checks the PLC to see if an Error has occured.
'Notes: (1) When an error occurs, the PLC places the error code in register
'           6727.  The CCM transmits the error code in a Tx block as follows:
'           STX + LObyte + HIbyte = ETX = LRC
'---------------------------------------------------------------------------

Public Sub chkErr(intError As Integer, result As Boolean)
   '--- Variable Declarations
   Dim HdrOK As Boolean                'header result
   Dim strHeader As String             'header string
   Dim strReply As String              'reply string
   Dim intTemp As Integer              'temp integer for calcs
   Dim char2, char3 As String          'status characters
   Dim intLO, intHI As Integer         'ascii codes of LO & HI bytes
   
   strHeader = bldHdr("E")                            'build the check intError header
   
   HdrOK = False
   TxHDR strHeader, HdrOK                             'xmit the header
   
   '--- get the response block
   If HdrOK = True Then                               'if header Ack'd
      '--- get the reply
      'frmRun.txtStat.Text = "chkErr:Waiting for E-HDR block..."
      strReply = ""
      'timout set to 8340 per cond#7 of Table 4-5.
      RxBLK 7, strReply                               'get the reply string
      
      '--- handle reply
      If Len(strReply) = 7 Then                       'if proper reply received
         If Mid(strReply, 2, 2) = NUL & NUL Then      'Nul's found, NO PLC error
            intError = 0
            result = True
         Else                                         'No ?? must be an error
            '--- reconstruct error#
            char2 = Mid(strReply, 2, 1)               'get the 2nd char
            char3 = Mid(strReply, 3, 1)               'get the 3rd char
            intLO = Asc(char2)                        'get ascii code for LO byte
            intHI = Asc(char3)                        'get ascii code for HI byte
            intError = intHI * 256 + intLO            'reconstruct error
            result = True
         End If '??
      Else
         'frmRun.txtStat.Text = "chkErr:Bad/No N-HDR block!"   'OP feedback
      End If 'reply OK
   Else                                               'if header failed
      'frmRun.txtStat.Text = "chkErr:E-HDR Failed!"     'OP feedback
   End If 'hdrok

'--- capture error string for troubleshooting
   'Open "C:\RemBar\Test\Error.txt" For Append As #1
   'Print #1, "ErrString Len = " & Len(strReply) & " ErrString:" & strReply
   'Print #1, "Error = " & intError
   'Close #1

End Sub 'chkErr

'---------------------------------------------------------------------------
'Name:      clrErr
'Accepts:   none
'Returns:   result = true if error successfully cleared.
'Requires:  none
'Discrip:   This sub lets the PLC know that it has processed the error.
'Notes: (1) When an error occurs, the PLC places the error code in register
'           6727.
'       (2) Writing "??" into 6727 ACK's processing of the error.
'---------------------------------------------------------------------------

Public Sub clrErr(result As Boolean)
   '--- Variable Declarations
   Dim HdrOK As Boolean                'header result
   Dim strHeader As String             'header string
   Dim strACK As Integer               'ACK response
   Dim intNAK As Integer               'NAK counter
   Dim badACK As Boolean               'Bad ACK flag
   Dim strClear As String              'clear string
   Dim LRC As String                   'LRC byte
   Dim strTx As String                 'Tx string
      
   strHeader = bldHdr("S")                            'build the clear error header
   
   HdrOK = False
   TxHDR strHeader, HdrOK                             'xmit the header

   If HdrOK = True Then    '--- if header Tx'd & ACK'd
      '--- construct the clear string
      'frmRun.txtStat.Text = "clrErr:Constructing Clear String..."    'OP feedback
      strClear = NUL & NUL                            'build clear string
      LRC = calcLRC(strClear)                         'calc the LRC
      strTx = STX & strClear & ETX & LRC              'build the Tx string
                 
      TxBLK strTx, result                             'send the block
      
   Else                                               'if header failed
      'frmRun.txtStat.Text = "clrErr: S-HDR Failed!"
   End If 'hdrok
    
End Sub  'clrErr

'---------------------------------------------------------------------------
'Name:      dspErr
'Accepts:   intError = error code
'Returns:
'Requires:  error table in remmele housing local dbase
'Discrip:   This sub accepts an error code, locates the error in the error
'           table, gets the erro dispcriptions and displays it in the PLC
'           message textbox.
'Notes:
'---------------------------------------------------------------------------

Public Sub dspErr(intError As Integer)
   '--- Variable Declarations
   Dim conLocal As Connection
   Dim adoRS As ADODB.Recordset
   Dim strSQL As String
   
   '-------------------- display the error description ----------------
   '--- make connection to local database
   Set conLocal = New Connection
   conLocal.Open "PROVIDER=MSDASQL;dsn=dsnMBLocal;uid=;pwd=;"
                
   '--- make recordset for error list
   Set adoRS = New ADODB.Recordset
   strSQL = "SELECT * FROM tblError " & _
            "WHERE ErrCode = " & intError
   adoRS.Open strSQL, conLocal, adOpenStatic, adLockOptimistic
   
   If adoRS.RecordCount > 0 Then                      'found a record
      adoRS.MoveFirst                                 'go to the record
      frmRun.txtPLCMessage = adoRS.Fields("ErrDescription")
   Else
      MsgBox ("procError: Error = " & intError & " NOT found!")
   End If
   
   '------------------------------- cleanup --------------------------------
   adoRS.Close                                        'unload recordset
   Set adoRS = Nothing
   conLocal.Close                                     'unload connection
   Set conLocal = Nothing
   
End Sub  'dspErr









'--- func/sub comment header
'---------------------------------------------------------------------------
'Name:      x
'Accepts:
'Returns:
'Requires:  MScomm1
'Discrip:
'Notes:
'---------------------------------------------------------------------------












