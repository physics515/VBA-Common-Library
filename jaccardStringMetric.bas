'''''''''''''''''''''''''''''''''''''''''''''''
' jaccard                                     '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string
'       optional caseSensitive (ex. True) as CaseSensitivity: default = CaseSensitivity.Sensitive

'outputs the metric as double

''' From The Author '''
'@Description: calculate the Jaccard Similarity Coefficient.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0


''' Inspired By '''
'@url: https://github.com/DigitecGalaxus/Jaccard
'@language: C#
'@description: Small tool to calculate the Jaccard Similarity Coefficient.
'@author: @DigitecGalaxus
'@version: 0.0.0
'@license: MIT

Function jaccard(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity = CaseSensitivity.Sensitive) As Double
    Dim string1Arr() As String
    Dim string2Arr() As String
    Dim intersectionCount As Long
    Dim unionCoun As Long
    Dim i As Integer: i = 1
    Dim j As Integer: j = 1
    Dim union As New Scripting.Dictionary
    
    'if not case sensitive then convert to upper case
    Select Case caseSensitive
        Case CaseSensitivity.NotSensitive
            string1 = UCase(string1)
            string2 = UCase(string2)
    End Select
    
    'split string1 into an array of characters
    For i = 1 To Len(string1)
        ReDim Preserve string1Arr(1 To i)
        string1Arr(i) = Mid(string1, i, 1)
    Next i
    
    'split string2 into an array of characters
    For i = 1 To Len(string2)
        ReDim Preserve string2Arr(1 To i)
        string2Arr(i) = Mid(string2, i, 1)
    Next i
    
    'check for nonzero values
    If (UBound(string1Arr) >= 0 And UBound(string2Arr) = 0) Or (UBound(string1Arr) = 0 And UBound(string2Arr) > 0) Then
        jaccard = 0
        Exit Function
    End If
    
    If UBound(string1Arr) = 0 And UBound(string2Arr) = 0 Then
        jaccard = 0
        Exit Function
    End If
    
    'determine the intersection of the two arrays
    i = 1
    For Each str1 In string1Arr
        j = 1
        For Each str2 In string2Arr
            If str1 = str2 Then intersectionCount = intersectionCount + 1
        Next str2
    Next str1
    
    'create union
    For i = 1 To UBound(string1Arr)
        If Not union.Exists(string1Arr(i)) Then union.Add string1Arr(i), True
    Next i
    
    For i = 1 To UBound(string2Arr)
        If Not union.Exists(string2Arr(i)) Then union.Add string2Arr(i), True
    Next i
    
    unioncount = union.Count
    
    jaccard = (intersectionCount / unioncount)
End Function