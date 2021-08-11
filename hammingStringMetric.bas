'''''''''''''''''''''''''''''''''''''''''''''''
' hamming                                     '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string

'       optional caseSensitive (ex. True) as CaseSensitivity: default = CaseSensitivity.

'outputs the metric as integer

''' From The Author '''
'@Description: This function takes two strings of the same length and calculates the Hamming Distance between them. Hamming Distance measures how close two strings are by checking how many Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers.
'@Author: Anthony Mancini
'@Version: 1.0.0
'@License: MIT
'@Example: =Hamming("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
'@Example: =Hamming("Cat", "Bag") -> 2; 2 changes are needed, changing the "B" to "C" and the "g" to "t"
'@Example: =Hamming("Cat", "Dog") -> 3; Every single character needs to be substituted in this case

Public Function hamming(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity = CaseSensitivity.NotSensitive) As Integer

    'if not case sensitive then convert to upper case
    Select Case caseSensitive
        Case CaseSensitivity.NotSensitive
            string1 = UCase(string1)
            string2 = UCase(string2)
        Case CaseSensitivity.DefaultSensitivity
            string1 = UCase(string1)
            string2 = UCase(string2)
    End Select

    If Len(string1) <> Len(string2) Then
        If Len(string1) > Len(string2) Then
            string1 = left(string1, Len(string2))
        Else
            string2 = left(string2, Len(string1))
        End If
    End If
    
    Dim totalDistance As Integer
    totalDistance = 0
    
    Dim i As Integer
    
    For i = 1 To Len(string1)
        If Mid(string1, i, 1) <> Mid(string2, i, 1) Then
            totalDistance = totalDistance + 1
        End If
    Next
    
    hamming = totalDistance
    
End Function