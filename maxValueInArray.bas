'''''''''''''''''''''''''''''''''''''''''''''''
' max                                         '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of
'       numbers array (ex. Array(1,3,5,9)) as single number, multiple numbers, or multiple arrays of numbers

'outputs the maximum number contained within an array

''' From The Author '''
'@Description: This function takes multiple numbers or multiple arrays of numbers and returns the max number. This function also accounts for numbers that are formatted as strings by converting them into numbers
'@Author: Anthony Mancini
'@Version: 1.0.0
'@License: MIT
'@Example: =Max(1, 2, 3) -> 3
'@Example: =Max(4.4, 5, "6") -> 6
'@Example: =Max(x) -> 3; Where x is an array with these values [1, 2.2, "3"]
'@Example: =Max(x, y, 10) -> 15; Where x = [1, 2.2, "3"] and y = [5, 15, -100]

Public Function max(ParamArray numbers() As Variant) As Double
    Dim individualParamArrayValue As Variant
    Dim individualValue As Variant
    Dim maxValue As Variant
    
    maxValue = Empty
    
    For Each individualParamArrayValue In numbers
        If IsArray(individualParamArrayValue) Then
            For Each individualValue In individualParamArrayValue
                If TypeName(individualValue) = "String" Then
                    individualValue = CDbl(individualValue)
                End If
            
                If IsEmpty(maxValue) Then
                    maxValue = individualValue
                ElseIf individualValue > maxValue Then
                    maxValue = individualValue
                End If
            Next
        Else
            If TypeName(individualParamArrayValue) = "String" Then
                individualParamArrayValue = CDbl(individualParamArrayValue)
            End If
        
            If IsEmpty(maxValue) Then
                maxValue = individualParamArrayValue
            ElseIf individualParamArrayValue > maxValue Then
                maxValue = individualParamArrayValue
            End If
        End If
    Next
    
    max = maxValue
End Function