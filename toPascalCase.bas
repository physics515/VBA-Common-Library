'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Converts String To Pascal Case                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'recieves an input of text (ex. "Hello World") as string
'outputs string converted to pascal case (ex. "HelloWorld")

Function toPascalCase(ByVal text As String) As String

        'dimension variables
        Dim words() As String
        Dim i As Integer

        'conver text to proper case
        text = WorksheetFunction.Proper(text)

        'if there are less that 2 characters, just return the string as uppercase
        If Len(text) < 2 Then
                ToPascalCase = UCase$(text)
                Exit Function
        End If

        'split the string into words
        words = Split(text)

        'capitalize each word
        For i = LBound(words) To UBound(words)
                If (Len(words(i)) > 0) Then
                        Mid$(words(i), 1, 1) = UCase$(Mid$(words(i), 1, 1))
                End If
        Next i

        'return - combine the words
        toPascalCase = Join(words, "")
End Function