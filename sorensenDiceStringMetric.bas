'''''''''''''''''''''''''''''''''''''''''''''''
' sorensenDice                                '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***
' *** Requires Function "ngrams" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string

'       optional caseSensitive (ex. True) as CaseSensitivity: default = CaseSensitivity.NotSensitive

'outputs the metric as double

''' From The Author '''
'@Description: Get the edit-distance according to Dice between two values.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0


''' Inspired By '''
'@url: https://github.com/words/dice-coefficient
'@language: JavaScript
'@description: Get the edit-distance according to Dice between two values.
'@author: @words
'@version: 2.0
'@license: MIT

Function sorensenDice(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity = CaseSensitivity.NotSensitive) As Double
    
    'if not case sensitive then convert to upper case
    Select Case caseSensitive
        Case CaseSensitivity.NotSensitive
            string1 = UCase(string1)
            string2 = UCase(string2)
        Case CaseSensitivity.DefaultSensitivity
            string1 = UCase(string1)
            string2 = UCase(string2)
    End Select
    
    'build bi-grams
    Dim left As Variant: left = nGrams(string1, 2)
    Dim right As Variant: right = nGrams(string2, 2)
    
    Dim index As Integer: index = 1
    Dim intersections As Integer: intersection = 0
    Dim leftPair As String
    Dim rightPair As String
    Dim offset As Integer: offset = 1
    
    'record intersections and offsets
    While index <= UBound(left)
        leftPair = left(index)
        offset = 1
        
        Do While offset < UBound(right)
            rightPair = right(offset)
            
            If leftPair = rightPair Then
                intersections = intersections + 1
                
                right(offset) = ""
                Exit Do
            End If
            
            offset = offset + 1
        Loop
    
        index = index + 1
    Wend
    
    sorensenDice = (2 * intersections) / (UBound(left) + UBound(right))
End Function