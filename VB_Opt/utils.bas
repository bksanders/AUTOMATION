Attribute VB_Name = "Utils"
Option Explicit

'------------------------------------------------------------------------------
' Name:     utils.bas
' By:       J.Keith Anderson, GEAS
' Type:     module
' Accepts:  none
' Returns:  none
' Reqrd:    none
' Descr:    This module contains general utility subs &  functions, written
'           to enhance VB.
' Notes:
'------------------------------------------------------------------------------
' Rev Hist:
' Date      Description
' 8-22-03   added incr & decr
' 8-8-03    added wait sub
' 8-7-03    bitVal, verifyZone
'------------------------------------------------------------------------------



'------------------------------------------------------------------------------
' Name:     bitVal
' Type:     function
' Accepts:  intWord, the integer word
'           intBit, the bit your checking
' Returns:  0 if bit off, 1 if bit on
' Requires: none
' Descr:    This function checks the status of a desired bit, in a given integer
'           word.
' Notes: (1) the bits are numbered: 0 to 15, left to right, lsb to msb
'------------------------------------------------------------------------------
Public Function bitVal(intWord As Integer, intBit As Integer) As Integer
   Dim intMask As Integer
   Dim intResult As Integer
   Dim intTemp As Integer
   
   intMask = 2 ^ intBit
   intResult = intWord And intMask
   If intResult > 0 Then bitVal = 1
End Function

'------------------------------------------------------------------------------
' Name:     wait
' Type:     subr
' Accepts:  intSecs, the # of secs to wait
' Returns:  none
' Requires: none
' Descr:    This function kills some time.  It will wait for the # of secs
'           passed in intSecs.
' Notes: (1)wont work across midnight barier
'------------------------------------------------------------------------------
Public Sub wait(intSecs As Integer)
   Dim i As Long
   Dim curTime As String
   Dim baseTime As String
   
   For i = 1 To intSecs
      baseTime = Format(Now, "HH:mm:ss")
      Do
         DoEvents
         curTime = Format(Now, "HH:mm:ss")
      Loop Until curTime > baseTime
   Next i
End Sub

'------------------------------------------------------------------------------
' Name:     incr
' Type:     function
' Accepts:  a varient
' Returns:  a varient
' Requires: none
' Descr:    This function increments the number passed to it.
' Notes:
'------------------------------------------------------------------------------
Public Function incr(varNumber As Variant) As Variant
   incr = varNumber + 1
End Function 'incr

'------------------------------------------------------------------------------
' Name:     decr
' Type:     function
' Accepts:  a varient
' Returns:  a varient
' Requires: none
' Descr:    This function decrements the number passed to it.
' Notes:
'------------------------------------------------------------------------------
Public Function decr(varNumber As Variant) As Variant
   decr = varNumber - 1
End Function 'decr
