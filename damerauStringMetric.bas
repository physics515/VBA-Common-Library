'''''''''''''''''''''''''''''''''''''''''''''''
' damerau                                     '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***
' *** Requires Reference "Microsoft Scripting Library" ***

'recieves input of
'       string1 (ex. "foo") as string
'       string2 (ex. "bar") as string

'       optional caseSensitive (ex. NotSensitive) as CaseSensitivity: default = Sensitive

'outputs the metric as integer

''' From The Author '''
'@Description: This function takes two strings of any length and calculates the Damerau-Levenshtein Distance between them. Damerau-Levenshtein Distance differs from Levenshtein Distance in that it includes an additional operation, called Transpositions, which occurs when two adjacent characters are swapped. Thus, Damerau-Levenshtein Distance calculates the number of Insertions, Deletions, Substitutions, and Transpositons needed to convert string1 into string2. As a result, this function is good when it is likely that spelling errors have occured between two string where the error is simply a transposition of 2 adjacent characters.
'@Author: Anthony Mancini
'@Version: 1.1.0
'@License: MIT
'@Example: =Damerau("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
'@Example: =Damerau("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
'@Example: =Damerau("Cat", "Cta") -> 1; Since the "t" and "a" can be transposed as they are adjacent to each other. Notice how LEVENSHTEIN("Cat","Cta")=2 but DAMERAU("Cat","Cta")=1

''' Modified By '''
'@Description: Add ability to change case sensitivity
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3

Public Function damerau(string1 As String, string2 As String, Optional caseSensitive As CaseSensitivity) As Integer

    'if not case sensitive then convert to upper case
    Select Case caseSensitive
        Case CaseSensitivity.NotSensitive
            string1 = UCase(string1)
            string2 = UCase(string2)
    End Select

    ' **Error Checking**
    'quick returns for common errors
    If string1 = string2 Then
        damerau = 0
    ElseIf string1 = Empty Then
        damerau = Len(string2)
    ElseIf string2 = Empty Then
        damerau = Len(string1)
    End If
    
    Dim inf As Long
    Dim da As Object
    inf = Len(string1) + Len(string2)
    Set da = CreateObject("Scripting.Dictionary")
    
    'filling the dictionary
    Dim i As Integer
    For i = 1 To Len(string1)
        If da.Exists(Mid(string1, i, 1)) = False Then
            da.Add Mid(string1, i, 1), "0"
        End If
    Next
    
    For i = 1 To Len(string2)
        If da.Exists(Mid(string2, i, 1)) = False Then
            da.Add Mid(string2, i, 1), "0"
        End If
    Next
    
    'creating h matrix
    Dim H() As Long
    ReDim H(Len(string1) + 1, Len(string2) + 1)
    
    Dim k As Integer
    For i = 0 To (Len(string1) + 1)
        For k = 0 To (Len(string2) + 1)
            H(i, k) = 0
        Next
    Next
    
    'updating the matrix
    For i = 0 To Len(string1)
        H(i + 1, 0) = inf
        H(i + 1, 1) = i
    Next
    For k = 0 To Len(string2)
        H(0, k + 1) = inf
        H(1, k + 1) = k
    Next
    
    'running the array
    Dim db As Long
    Dim i1 As Long
    Dim k1 As Long
    Dim cost As Long
    
    For i = 1 To Len(string1)
        db = 0
        For k = 1 To Len(string2)
            i1 = CInt(da(Mid(string2, k, 1)))
            k1 = db
            cost = 1
            
            If Mid(string1, i, 1) = Mid(string2, k, 1) Then
                cost = 0
                db = k
            End If
            
            H(i + 1, k + 1) = min(H(i, k) + cost, H(i + 1, k) + 1, H(i, k + 1) + 1, H(i1, k1) + (i - i1 - 1) + 1 + (k - k1 - 1))
        Next k
        
        If da.Exists(Mid(string1, i, 1)) Then
            da.Remove Mid(string1, i, 1)
            da.Add Mid(string1, i, 1), CStr(i)
        Else
            da.Add Mid(string1, i, 1), CStr(i)
        End If
    Next

    damerau = H(Len(string1) + 1, Len(string2) + 1)
End Function