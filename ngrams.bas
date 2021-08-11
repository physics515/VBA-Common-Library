'''''''''''''''''''''''''''''''''''''''''''''''
' ngrams                                      '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of
'       text (ex. "foo") as string
'       n (ex. 2) as integer

'outputs grams of n length as string array

''' From The Author '''
'@Description: Determine the grams of a given length. (ex. nGrams("Hello World", 2) = ("He", "el", "ll", "lo", "o ", " W", "Wo", "or", "rl", "ld")
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0

Function nGrams(text As String, n As Integer) As Variant
    Dim grams() As String
    Dim index As Integer
    Dim i As Integer
    Dim source() As String
    
    'split text in a character array
    For i = 1 To Len(text)
        ReDim Preserve source(1 To i)
        source(i) = Mid(text, i, 1)
    Next i
    
    index = UBound(source) - n + 1
    
    'exit if the number of characters is less than n - 1 length
    If index < 1 Then Exit Function
    
    'create the grams
    ReDim grams(1 To index)
    While index > 0
        grams(index) = source(index) & source(index + 1)
        index = index - 1
    Wend
    
    nGrams = grams
End Function