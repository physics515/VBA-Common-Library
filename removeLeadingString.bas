'''''''''''''''''''''''''''''''''''''''''''''''
' Remove String From Begining Of String       '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of lead (ex. "foo ") as string, and whole (ex. "foo bar")
'outputs the whole with the lead removed (ex. "bar")

Function removeLeadingString(lead As String, whole as string) As String

        'take the first length of lead characters of whole and check if they match lead
        If Left(whole, Len(lead)) = lead Then

                'if the characters match remove length of lead many characters from the beginning of whole and return
                removeLeadingString = Right(whole, Len(whole) - Len(lead))
        Else

                'else if the character do not match return whole
                removeLeadingString = whole
        End If
End Function