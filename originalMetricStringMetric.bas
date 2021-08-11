'''''''''''''''''''''''''''''''''''''''''''''''
' originalMetric                              '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string

'       optional caseSensitive (ex. True) as CaseSensitivity: default = Sensitive

'outputs the metric as double

'requirements
'       common.CaseSensitivity Enum

''' From the Author '''
'@Description: String metric.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0

Function originalMetric(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity) As Double
    Dim i As Integer
    Dim j As Integer
    Dim foundPosition As Integer: foundPosition = 0
    Dim distance As Double
    Dim currentScore As Double
    
    'if not case sensitive then convert to upper case
    Select Case caseSensitive
        Case CaseSensitivity.NotSensitive
            string1 = UCase(string1)
            string2 = UCase(string2)
    End Select
    
    'loop through each letter in query
    For i = 1 To Len(string1)
    
            'loop through each letter in the current cell value
            For j = 1 To Len(string2)
    
                    'if the current query letter apeers in the the cell value then update the score
                    If InStr(1, Mid(string2, j, 1), Mid(string1, i, 1), vbTextCompare) > 0 Then
    
                            'if the last letter that was found came before the current letter (the letters are found in the correct order)
                            If foundpostition < j Then
    
                                    'allways return a positive distance
                                    If i < j Then
                                            distance = j - i
                                    Else
                                            distance = i - j
                                    End If
    
                                    'record the current found position as the previouly found position
                                    foundPosition = j
    
                                    'add distance to socre as a percentage matched, and round to 4 decimal places
                                    If distance > 0 Then
                                            currentScore = Round(Abs(currentScore + Round((1 / distance), 4)), 4)
                                    End If
                            End If
                    End If
            Next j
    Next i
    
    originalMetric = currentScore
End Function