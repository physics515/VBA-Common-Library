'''''''''''''''''''''''''''''''''''''''''''''''
' simpleMatching                              '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string
'       optional caseSensitive (ex. True) as CaseSensitivity: default = CaseSensitivity.Sensitive

'outputs the metric as double

''' From The Author '''
'@Description: calculate the simple matching metric.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0

Function simpleMatching(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity = CaseSensitivity.Sensitive) As Double
    Dim string1Arr() As String
    Dim string2Arr() As String
    Dim nAttributes As Long
    Dim matches As Long: matches = 0
    
    'if not case sensitive then convert to upper case
    Select Case caseSensitive
        Case CaseSensitivity.NotSensitive
            string1 = UCase(string1)
            string2 = UCase(string2)
    End Select
    
    If Len(string1) > Len(string2) Then nAttributes = Len(string1) Else nAttributes = Len(string2)
    
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
    
    'determine matching attributes
    For i = 1 To nAttributes
        If i <= UBound(string1Arr) And i <= UBound(string2Arr) Then
            If string1Arr(i) = string2Arr(i) Then matches = matches + 1
        End If
    Next i
    
    simpleMatching = matches / nAttributes
End Function