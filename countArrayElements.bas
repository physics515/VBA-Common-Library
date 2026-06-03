'''''''''''''''''''''''''''''''''''''''''''''''
' Count Array Elements                        '
'''''''''''''''''''''''''''''''''''''''''''''''
' receives an array as input along with boolean flags indicating which counts to return
' outputs a long array containing the requested counts in order: [nonBlanks, blanks, total]

''' From The Author '''
'@Description: Counts the elements of an array. Returns a long array containing the requested counts in order: [nonBlanks, blanks, total].
'@Author: Justin Icenhour
'@Version: 2.0.0
'@License: GPL-3.0

Function countArrayElements(arr As Variant, nonBlanks As Boolean, blanks As Boolean, total As Boolean) As Long()
        'dimension variables
        Dim i As Long: i = 0
        Dim nonBlankCount As Long: nonBlankCount = 0
        Dim blankCount As Long: blankCount = 0
        Dim totalCount As Long: totalCount = 0
        Dim resultSize As Long: resultSize = 0
        Dim result() As Long
        Dim isBlank As Boolean
        
        'loop through each item in the array
        For i = LBound(arr) To UBound(arr)
                totalCount = totalCount + 1
                
                'determine if the array item is blank
                isBlank = False
                If IsEmpty(arr(i)) Then
                        isBlank = True
                ElseIf IsError(arr(i)) Then
                        isBlank = False
                ElseIf IsObject(arr(i)) Then
                        isBlank = False
                ElseIf CStr(arr(i)) = "" Then
                        isBlank = True
                End If
                
                'increment the appropriate counter
                If isBlank Then
                        blankCount = blankCount + 1
                Else
                        nonBlankCount = nonBlankCount + 1
                End If
        Next i
        
        'determine size of result array based on boolean flags
        If nonBlanks Then resultSize = resultSize + 1
        If blanks Then resultSize = resultSize + 1
        If total Then resultSize = resultSize + 1
        
        'build result array; if no flags are set, return an empty array
        If resultSize > 0 Then
                ReDim result(1 To resultSize)
                
                Dim idx As Long: idx = 1
                If nonBlanks Then
                        result(idx) = nonBlankCount
                        idx = idx + 1
                End If
                If blanks Then
                        result(idx) = blankCount
                        idx = idx + 1
                End If
                If total Then
                        result(idx) = totalCount
                End If
        Else
                ReDim result(0 To -1)
        End If
        
        'return
        countArrayElements = result
End Function
