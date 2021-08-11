'''''''''''''''''''''''''''''''''''''''''''''''
' min                                         '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of
'       numbers array (ex. "foo") as single number, multiple numbers, or multiple arrays of numbers

'outputs the minimum number contained within an array

''' From The Author '''
'@Description: This function takes multiple numbers or multiple arrays of numbers and returns the min number. This function also accounts for numbers that are formatted as strings by converting them into numbers
'@Author: Anthony Mancini
'@Version: 1.0.0
'@License: MIT
'@Example: =Min(1, 2, 3) -> 1
'@Example: =Min(4.4, 5, "6") -> 4.4
'@Example: =Min(-1, -2, -3) -> -3
'@Example: =Min(x) -> 1; Where x is an array with these values [1, 2.2, "3"]
'@Example: =Min(x, y, 10) -> -100; Where x = [1, 2.2, "3"] and y = [5, 15, -100]

Public Function min(ParamArray numbers() As Variant) As Double
    Dim individualParamArrayValue As Variant
    Dim individualValue As Variant
    Dim minValue As Variant
    
    minValue = Empty
    
    For Each individualParamArrayValue In numbers
        If IsArray(individualParamArrayValue) Then
            For Each individualValue In individualParamArrayValue
                If TypeName(individualValue) = "String" Then
                    individualValue = CDbl(individualValue)
                End If
            
                If IsEmpty(minValue) Then
                    minValue = individualValue
                ElseIf individualValue < minValue Then
                    minValue = individualValue
                End If
            Next
        Else
            If TypeName(individualParamArrayValue) = "String" Then
                individualParamArrayValue = CDbl(individualParamArrayValue)
            End If
        
            If IsEmpty(minValue) Then
                minValue = individualParamArrayValue
            ElseIf individualParamArrayValue < minValue Then
                minValue = individualParamArrayValue
            End If
        End If
    Next
    
    min = minValue

End Function