Attribute VB_Name = "MODULE_HiddenSurfaceRemoval"
'000000000000000000000000000000000000000000000000000000000000000000000
'000000000000000000011000000000011000000000000000000000000000000000000
'000000011111100000101000110000111000011000000000001111110000000000000
'000000100001100001010000110001110000110000000000001000111100000000000
'000011000001100010100001110011010000110000000000001100001100000000000
'000011000011000111000000100010100001110000001110001100111000010000000
'000110000011001110100000100011000000100000110100001111000001111000000
'000110000000001101100111100110000101100000110001111100000010011011000
'000011000000111110111001111001111001100011111110011111001100010100000
'000011111111001100000001100001100001111110001000010001110000111000000
'000000011000000000000000000000000001111000000000010000011100111000000
'000000000000000000000000000000000000000000000000000000000001011000000
'000000000000000000000000000000000000000000000000000000000010110000000
'(c) 2002 by Riley T. Perry - Chillers of Entropy

'-> If the comments below look garbled then change font to COURIER NEW

'                                                 ,  ,
'                                                / \/ \
'                                              (/ //_ \_
'     .-._                                      \||  .  \
'      \  '-._                            _,:__.-"/---\_ \
' ______/___  '.    .--------------------'~-'--.)__( , )\ \
'`'--.___  _\  /    | HiddenSurface (Module)  ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 22/5/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Hidden surface removal methods.                        *
' *--------------------------------------------------------*
'
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   QSort_Numeric_Ascending()              |
' |  /   \  -------------------------              |
' | |\_.  | Sorts the array of triangles. Original |
' |\|  | /| by Kenneth Ives (converted by me)     /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Array of triangles
'   `---'   2.>> Lowest array value
'           3.>> Highest array value
'
Public Sub QSort_Numeric_Ascending(arNumeric() As Triangle, _
                                   lngLow As Long, lngHi As Long)

' ***************************************************************************
' Routine:       QSort_Numeric_Ascending
'
' Description:   This routine will accept and sort in ascending order a
'                numeric array of data.  This routine is used when the data
'                to be sorted is known to be ALL numeric.
'
'                This routine can also be used for sorting dates since they
'                are stored in a double precision format.  the system date
'                is retrieved as a variant and then can be stored as a
'                double.
'
' Parameters:    arNumeric() - Array to be sorted
'                lngLow      - Minimum number of elements in the array
'                lngHi       - Maximum numer of elements in the array
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 30-APR-2000  Kenneth Ives     Routine created kenaso@home.com
' ***************************************************************************

' ----------------------------------------------------------------------
' Define local variables
' ----------------------------------------------------------------------
  Dim dblMidPoint  As Double    ' midpoint of the array to be sorted
  Dim dblHold      As Triangle  ' Temp hold area for swapping values
  Dim lngTmpLow    As Long      ' Index pointer
  Dim lngTmpHi     As Long      ' Index pointer

' ----------------------------------------------------------------------
' See if this is an empty array by checking to see if there is data in
' the first element.  If not, then leave.
' ----------------------------------------------------------------------
  If Len(Trim(arNumeric(0).Middle_z)) = 0 Then
      Exit Sub
  End If
  
' ----------------------------------------------------------------------
' Leave if there is nothing to sort
' ----------------------------------------------------------------------
  If lngLow >= lngHi Then
      Exit Sub
  End If

' ----------------------------------------------------------------------
' Save the count of the minimum and maximum number of elements in the
' array to be sorted.
' ----------------------------------------------------------------------
  lngTmpLow = lngLow
  lngTmpHi = lngHi
   
' ----------------------------------------------------------------------
' Calculate the midpoint of the array
' ----------------------------------------------------------------------
  dblMidPoint = arNumeric((lngLow + lngHi) / 2).Middle_z

' ----------------------------------------------------------------------
' Start the sorting process
' ----------------------------------------------------------------------
  While (lngTmpLow <= lngTmpHi)
       
      ' Always process the low end first.  Loop as long the array data
      ' element is LESS than the data in the temporary holding area
      ' and the temporary low value is LESS than the maximum number of
      ' array elements.
      While (arNumeric(lngTmpLow).Middle_z < dblMidPoint And lngTmpLow < lngHi)
          lngTmpLow = lngTmpLow + 1  ' Increment the temp low counter
      Wend
   
      ' Now, we will process the high end.  Loop as long the data in the
      ' temporary holding area is LESS than the array data element
      ' and the temporary high value is GREATER than the minimum number
      ' of array elements.
      While (dblMidPoint < arNumeric(lngTmpHi).Middle_z And lngTmpHi > lngLow)
          lngTmpHi = lngTmpHi - 1    ' Decrement the temp high counter
      Wend

      ' if the temp low end is LESS than or equal to the temp high end,
      ' then swap places
      If (lngTmpLow <= lngTmpHi) Then
          dblHold = arNumeric(lngTmpLow)                                    ' Move the Low value to Temp Hold
          arNumeric(lngTmpLow) = arNumeric(lngTmpHi)                        ' Move the high value to the low
          arNumeric(lngTmpHi) = dblHold                                     ' move the Temp Hold to the High
          lngTmpLow = lngTmpLow + 1                                         ' Increment the temp low counter
          lngTmpHi = lngTmpHi - 1                                           ' Decrement the temp high counter
      End If
 Wend
    
' ----------------------------------------------------------------------
' If the minimum number of elements in the array is LESS than the temp
' high end, then make a recursive call to this routine.  Always sort
' the low end of the array first.  This gives you a solid base.
' ----------------------------------------------------------------------
  If (lngLow < lngTmpHi) Then
      QSort_Numeric_Ascending arNumeric(), lngLow, lngTmpHi
  End If
   
' ----------------------------------------------------------------------
' If the temp low end is LESS than the maximum number of elements in
' the array, then make a recursive call to this routine.  The high end
' is always sorted last.
' ----------------------------------------------------------------------
  If (lngTmpLow < lngHi) Then
      QSort_Numeric_Ascending arNumeric(), lngTmpLow, lngHi
  End If

End Sub
