'''''''''''''''''''''''''''''''''''''''''''''''
' Count Non-Blank Array Items                 '
'''''''''''''''''''''''''''''''''''''''''''''''
' recieves an array as input
' outputs a integer count of the number of non-blank items

Function countNonBlankArrayItems(arr As Variant) As Integer
        'dimension variables
        Dim i As Integer: i = 0
        Dim count As Integer: count = 0
        
        'loop through each item in the array
        For i = LBound(arr) To UBound(arr)

                'if the array item contains a value increment count
                If Not arr(i, 1) = 0 Then
                        count = count + 1
                End If
        Next
        
        'return
        CountNonBlankArrayElements = count
End Function