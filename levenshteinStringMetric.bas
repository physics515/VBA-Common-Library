'''''''''''''''''''''''''''''''''''''''''''''''
' levenshtein                                 '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string

'       optional caseSensitive (ex. True) as CaseSensitivity: default = CaseSensitivity.Sensitive

'outputs the metric as long

''' From The Author '''
'@Description:This function takes two strings of any length and calculates the Levenshtein Distance between them. Levenshtein Distance measures how close two strings are by checking how many Insertions, Deletions, or Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers. Unlike Hamming Distance, Levenshtein Distance works for strings of any length and includes 2 more operations. However, calculation time will be slower than Hamming Distance for same length strings, so if you know the two strings are the same length, its preferred to use Hamming Distance.
'@Author: Anthony Mancini
'@Version: 1.1.0
'@License: MIT
'@Example: =Levenshtein("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
'@Example: =Levenshtein("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
'@Example: =Levenshtein("Cat", "Cta") -> 2; Since the "t" in "Cta" needs to be substituted into an "a", and the final character "a" needs to be substituted into a "t"

Public Function levenshtein(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity = CaseSensitivity.Sensitive) As Long

    'if not case sensitive then convert to upper case
    Select Case caseSensitive
        Case CaseSensitivity.NotSensitive
            string1 = UCase(string1)
            string2 = UCase(string2)
    End Select

    ' **Error Checking**
    'quick returns for common errors
    If string1 = string2 Then
        levenshtein = 0
        Exit Function
    ElseIf string1 = Empty Then
        levenshtein = Len(string2)
        Exit Function
    ElseIf string2 = Empty Then
        levenshtein = Len(string1)
        Exit Function
    End If
    

    ' **Algorithm Code**
    'creating the distance metrix and filling it with values
    Dim numberOfRows As Integer
    Dim numberOfColumns As Integer
    
    numberOfRows = Len(string1)
    numberOfColumns = Len(string2)
    
    Dim distanceArray() As Integer
    ReDim distanceArray(numberOfRows, numberOfColumns)
    
    Dim r As Integer
    Dim c As Integer
    
    For r = 0 To numberOfRows
        For c = 0 To numberOfColumns
            distanceArray(r, c) = 0
        Next c
    Next r
    
    For r = 1 To numberOfRows
        distanceArray(r, 0) = r
    Next r
    
    For c = 1 To numberOfColumns
        distanceArray(0, c) = c
    Next c
    
    'non-recursive Levenshtein Distance matrix walk
    Dim operationCost As Integer
    
    For c = 1 To numberOfColumns
        For r = 1 To numberOfRows
            If Mid(string1, r, 1) = Mid(string2, c, 1) Then
                operationCost = 0
            Else
                operationCost = 1
            End If
                                                           
            distanceArray(r, c) = min(distanceArray(r - 1, c) + 1, distanceArray(r, c - 1) + 1, distanceArray(r - 1, c - 1) + operationCost)
        Next r
    Next c
    
    levenshtein = distanceArray(numberOfRows, numberOfColumns)
End Function