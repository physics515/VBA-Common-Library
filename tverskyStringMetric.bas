'''''''''''''''''''''''''''''''''''''''''''''''
' tversky                                     '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***
' *** Requires Function "uniqueArrayElements" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string

'       optional caseSensitive (ex. True) as CaseSensitivity: default = CaseSensitivity.Sensitive
'       optional symetric (ex True) as boolean: default = false
'       optional string1Weight (ex. .5) as double: default = 1
'       optional string2Weight (ex. 2) as double: default = 1

'outputs the metric as double

''' From The Author '''
'@Description: Computes the Tversky index between two sequences. For alpha = beta = 0.5, the index is equal to Dice's coefficient. For alpha = beta = 1, the index is equal to the Tanimoto coefficient.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0


''' Inspired By '''
'@url: https://github.com/compute-io/tversky-index
'@language: JavaScript
'@description: Computes the Tversky index between two sequences. For alpha = beta = 0.5, the index is equal to Dice's coefficient. For alpha = beta = 1, the index is equal to the Tanimoto coefficient.
'@author: @compute-io
'@version: 0.0.0
'@license: MIT

Function tversky(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity = CaseSensitivity.Sensitive, Optional symmetric As Boolean = False, Optional string1Weight As Double = 1, Optional string2Weight As Double = 1) As Double
    Dim i As Integer: i = 1
    Dim string1Arr() As String
    Dim string2Arr() As String
    Dim uniqueString1ArrLength As Long
    Dim uniqueString2ArrLength As Long
    Dim dict1n2 As New Scripting.Dictionary
    Dim aCompl As Long: aCompl = 0
    Dim bCompl As Long: bCompl = 0
    Dim min As Long: min = 1
    Dim max As Long: max = 1
    
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
    
    'determine unique characters in each array
    uniqueString1ArrLength = UBound(common.uniqueArrayElements(string1Arr))
    uniqueString2ArrLength = UBound(common.uniqueArrayElements(string2Arr))
    string1Arr = common.uniqueArrayElements(string1Arr)
    string2Arr = common.uniqueArrayElements(string2Arr)
    ReDim Preserve string1Arr(1 To uniqueString1ArrLength)
    ReDim Preserve string2Arr(1 To uniqueString2ArrLength)
    
    'determine the intersection between the two arrays
    For i = 1 To UBound(string1Arr)
        If Not dict1n2.Exists(string1Arr(i)) Then dict1n2.Add string1Arr(i), True
    Next i
    
    For i = 1 To UBound(string2Arr)
        If Not dict1n2.Exists(string2Arr(i)) Then dict1n2.Add string2Arr(i), True
    Next i
    
    length = dict1n2.Count
    
    'compute the relative complements
    For i = 1 To UBound(string1Arr)
        If Not string1Arr(i) = dict1n2.Keys(i) Then aCompl = aCompl + 1
    Next i
    
    For i = 1 To UBound(string2Arr)
        If Not i > UBound(dict1n2.Keys) Then
            If Not string2Arr(i) = dict1n2.Keys(i) Then bCompl = bCompl + 1
        End If
    Next i
    
    If symmetric Then
        If aCompl > bCompl Then
            min = bCompl
            max = aCompl
        Else
            min = aCompl
            max = bCompl
        End If
        tversky = length / (length + string2Weight * (string1Weight * min + max * (1 - string1Weight)))
        Exit Function
    End If
    
    tversky = length / (length + (string1Weight * aCompl) + (string2Weight * bCompl))
End Function