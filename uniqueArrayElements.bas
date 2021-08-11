'''''''''''''''''''''''''''''''''''''''''''''''
' uniqueArrayElements                         '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Reference "Microsoft Scripting Library" ***

'recieves input of
'       arr as variant array

'outputs the only the unique elements of the input array as variant array

''' From The Author '''
'@Description: Computes the unique elements of an array.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0

Function uniqueArrayElements(arr As Variant) As Variant
    Dim length As Long: length = UBound(arr)
    Dim dict As New Scripting.Dictionary
    Dim vals() As String
    Dim val As String
    
    For i = 1 To length
        val = arr(i)
        If Not dict.Exists(val) Then dict.Add val, True
    Next i
    
    i = 1
    For Each Key In dict.Keys
        ReDim Preserve vals(1 To i)
        vals(i) = Key
        i = i + 1
    Next Key
    
    uniqueArrayElements = vals
End Function