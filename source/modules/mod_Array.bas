Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Array
' VERSION:      1.01
' Description:  Array related functions & subroutines
'
' Source/date:  Bonnie Campbell, 7/24/2014
' Revisions:    BLC, 7/24/2014 - 1.00 - initial version
'               BLC, 8/19/2014 - 1.01 - removed unused functions (initArray, IsArrayEmpty, NumberOfDimensions),
'                                       added versioning
' =================================

' ---------------------------------
' FUNCTION:     addArrayValueSuffix
' Description:  Add a suffix to array values
' Parameters:   array - array to modify
'               suffix - string to add to the value
'               position - optional position of value to modify for 2D arrays
'               dimension - optional array dimension of value to modify for 2D arrays
' Returns:      array with values altered to (ValueSuffix, etc.)
' Throws:       -
' References:   -
' Note:
' Source/date:  Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Adapted:      -
' Revisions:    BLC, 7/25/2014 - initial version
' ---------------------------------
Public Function addArrayValueSuffix(ary As Variant, suffix As String, Optional position As Integer, Optional dimension As Integer) As Variant
On Error GoTo Err_Handler:
    Dim i As Integer
    
    '1 D
    For i = 0 To UBound(ary)
        If Not IsNull(ary(i)) Then
            'modify value
            ary(i) = ary(i) & suffix
        End If
    Next i
    
    '2 D
    If dimension > 0 Then
        For i = 0 To UBound(ary, dimension)
            If Not IsNull(ary(position, i)) Then
                'modify value
                ary(position, i) = ary(position, i) & suffix
            End If
        Next i
    End If
    
    addArrayValueSuffix = ary
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - renameArrayLifeforms[mod_Array])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     mergeArrays
' Description:  Merges 1D or 2D arrays (appends together), arrays must both be either 1D or 2D
' Parameters:   arr1, arr2 - arrays to merge
' Returns:      merged array
' Throws:       -
' References:   -
' Source/date:
'   Johannes - June 6, 2009
'   http://stackoverflow.com/questions/1588913/how-do-i-merge-two-arrays-in-vba
' Adapted:      Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Revisions:    BLC, 7/25/2014 - initial version
'               BLC, 8/12/2014 - fixed bug where first value of second array
'                                was skipped in merged array
' ---------------------------------
Public Function mergeArrays(ByRef arr1 As Variant, arr2 As Variant, Optional dimension As Integer, Optional columns As Integer) As Variant
On Error GoTo Err_Handler:
    Dim aryMerged() As Variant
    Dim len1 As Integer, len2 As Integer, lenRe As Integer, counter As Integer, i As Integer
        
    '2D arrays
    If dimension > 0 Then
        len1 = UBound(arr1, dimension)
        len2 = UBound(arr2, dimension)
    
    '1D arrays
    Else
        len1 = UBound(arr1)
        len2 = UBound(arr2)
    End If
    
    lenRe = len1 + len2
    counter = 0

    '2D arrays
    If dimension > 0 Then
        ReDim aryMerged(columns, 0 To lenRe + 1)
        Do While counter <= len1 'add first array
            For i = 0 To columns
                aryMerged(i, counter) = arr1(i, counter)
            Next
            counter = counter + 1
        Loop
        Do While counter <= lenRe + 1 'add the second array
            For i = 0 To columns
                aryMerged(i, counter) = arr2(i, counter - len1 - 1)
            Next
            counter = counter + 1
        Loop
    
    '1D arrays
    Else
        'assume 0 based so add 1
        ReDim aryMerged(0 To lenRe + 1)
        Do While counter <= len1 'add first array
            aryMerged(counter) = arr1(counter)
            counter = counter + 1
        Loop
        Do While counter <= lenRe + 1 'add the second array
            aryMerged(counter) = arr2(counter - len1 - 1)
            counter = counter + 1
        Loop
    End If
    
    mergeArrays = aryMerged
  
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - mergeLifeformArrays[mod_Array])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:     DeallocateArrays
' Description:  Releases memory allocated to arrays
' Parameters:   ary - array of arrays to deallocate
' Returns:      -
' Throws:       -
' References:   -
' Source/date:
'   Lance Roberts, March 6, 2009
'   http://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
' Adapted:      Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Revisions:    BLC, 7/8/2014 - XX
' ---------------------------------
Public Sub DeallocateArrays(ary As Variant)
On Error Resume Next
Dim i As Integer

    'release memory allocated to arrays
    For i = 0 To UBound(ary, 1)
        If IsArray(ary) Then
            Erase Array(i)
        End If
    Next

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DeallocateArrays[mod_Array])"
    End Select
    Resume Exit_Sub
End Sub