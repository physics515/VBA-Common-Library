'''''''''''''''''''''''''''''''''''''''''''''''
' One Digit Number To Text                    '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves any one digit number as digit (ex. "3") as string
'outputs a string as text (ex. "Three")

Function oneDigitNumberToText(digit As String) As String
    Select Case val(digit)
        Case 1: oneDigitNumberToText = "One"
        Case 2: oneDigitNumberToText = "Two"
        Case 3: oneDigitNumberToText = "Three"
        Case 4: oneDigitNumberToText = "Four"
        Case 5: oneDigitNumberToText = "Five"
        Case 6: oneDigitNumberToText = "Six"
        Case 7: oneDigitNumberToText = "Seven"
        Case 8: oneDigitNumberToText = "Eight"
        Case 9: oneDigitNumberToText = "Nine"
        Case Else: oneDigitNumberToText = ""
    End Select
End Function