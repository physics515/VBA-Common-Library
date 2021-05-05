'''''''''''''''''''''''''''''''''''''''''''''''
' Remove String From Begining Of String       '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of lead (ex. "foo ") as string, whole (ex. "foo bar"), and an Option asFormula (ex. True) as bool
'outputs the whole with the lead removed (ex. "bar")

'asFormula=True outputs an excel formula that performs the same function

Function removeLeadingString(lead As String, whole as String, Optional asFormula as Boolean = False) As String

        'if user did Not request an excel formula
        If Not asFormula Then

                'take the first length of lead characters of whole and check if they match lead
                If Left(whole, Len(lead)) = lead Then

                        'if the characters match remove length of lead many characters from the beginning of whole and return
                        removeLeadingString = Right(whole, Len(whole) - Len(lead))
                Else

                        'else if the character do not match return whole
                        removeLeadingString = whole
                End If

        'if the user requested an excel formula
        Else

                'return an excel formula that will perform the same function
                ' *** Note: the formula is missing the "=" at the beginning so that it can be easily concatinated with othe formulas ***
                removeLeadingString = "IF(LEFT(" & whole & ",LEN("& lead &"))=""" & lead & """,RIGHT(" & whole & ",LEN(" & whole & ")-LEN(" & lead & ")), " & whole & ")"
        End If
End Function