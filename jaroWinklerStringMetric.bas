'''''''''''''''''''''''''''''''''''''''''''''''
' jaroWinkler                                 '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string
'       optional caseSensitive (ex. True) as CaseSensitivity: default = CaseSensitivity.Sensitive

'outputs the metric as double

''' From The Author '''
'@Description: calculate the Jaro-Winkler distance.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0


''' Inspired By '''
'@url: https://github.com/jordanthomas/jaro-winkler
'@language: Javascript
'@description: The Jaro-Winkler distance metric for node and browser.
'@author: @jordanthomas
'@version: 0.2.7
'@license: MIT

Function jaroWinkler(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity = CaseSensitivity.Sensitive) As Double
    Dim i As Double: i = 0
    Dim j As Double: j = 0
    Dim k As Double: k = 0
    Dim l As Double: l = 0
    Dim p As Double: p = 0.1
    Dim low As Double: low = 0
    Dim high As Double: high = 0
    Dim numTrans As Double: numTrans = 0
    Dim weight As Double: weight = 0
    Dim string1Arr() As String
    Dim string2Arr() As String
    Dim stringUBound As Double: stringUBound = 0
    
    'exit early if either string is empty
    If string1 = "" Or string2 = "" Then
        jaroWinkler = 0
        Exit Function
    End If
    
    'if not case sensitive then convert to upper case
    Select Case caseSensitive
        Case CaseSensitivity.NotSensitive
            string1 = UCase(string1)
            string2 = UCase(string2)
    End Select
    
    'exit early if the strings are the same
    If string1 = string2 Then
        jaroWinkler = 1
        Exit Function
    End If
    
    Dim range1 As Double: range1 = (Application.WorksheetFunction.Floor(Application.WorksheetFunction.max(Len(string1), Len(string2)) / 2, 1)) - 1
    ReDim string1Matches(0 To Len(string1)) As Boolean
    ReDim string2Matches(0 To Len(string2)) As Boolean
    
    For i = 1 To UBound(string1Matches)
        string1Matches(i) = False
    Next i
    
    For i = 1 To UBound(string2Matches)
        string2Matches(i) = False
    Next i
    
    'split string1 into an array of characters
    For i = 1 To Len(string1)
        ReDim Preserve string1Arr(0 To i)
        string1Arr(i) = Mid(string1, i, 1)
    Next i
    
    'split string2 into an array of characters
    For i = 1 To Len(string2)
        ReDim Preserve string2Arr(0 To i)
        string2Arr(i) = Mid(string2, i, 1)
    Next i
    
    'find matching characters
    For i = 0 To Len(string1) - 1
        If i > range1 Then low = i - range1 Else low = 0
        If i + range1 <= (Len(string2) - 1) Then high = (i + range1) Else high = (Len(string2) - 1)
        
        For j = low To high
            If Not string1Matches(i) = True And Not string2Matches(j) = True And string1Arr(i) = string2Arr(j) Then
                m = m + 1
                string1Matches(i) = True
                string2Matches(j) = True
                Exit For
            End If
        Next j
    Next i
    
    'exit early if no matches were found
    If m = 0 Then
        jaroWinkler = 0
        Exit Function
    End If
    
    'count the transpositions
    For i = 0 To Len(string1)
        If string1Matches(i) = True Then
            For j = k To Len(string2)
                If string2Matches(j) = True Then
                    k = j + 1
                    Exit For
                End If
            Next j
            
            If Not string1Arr(i) = string2Arr(j) Then
                numTrans = numTrans + 1
            End If
        End If
    Next i
    
    weight = (m / Len(string1) + m / Len(string2) + (m - (numTrans / 2)) / m) / 3
    
    If UBound(string1Arr) > UBound(string2Arr) Then stringUBound = UBound(string2Arr) Else stringUBound = UBound(string1Arr)
    
    If weight > 0.7 Then
        While string1Arr(l) = string2Arr(l) And l < 4 And l < stringUBound
            l = l + 1
        Wend
        weight = weight + l * p * (1 - weight)
    End If
    
    jaroWinkler = weight
End Function