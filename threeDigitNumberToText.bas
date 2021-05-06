'''''''''''''''''''''''''''''''''''''''''''''''
' Three Digit Number To Text                  '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves any three digit number as threeDigits (ex. "125") as string
'outputs a string as text (ex. "One Hundred Twenty Five")

' *** Requires Function "oneDigitNumberToText" ***
' *** Requires Function "twoDigitNumberToText" ***

Function threeDigitNumberToText(ByVal threeDigits As String) As String

        'dimension variables
        Dim result As String

        'if threeDigits = 0 then exit
        If val(threeDigits) = 0 Then Exit Function

        'ensure three threeDigits is three digits by appending zeros to the left and then trimming the length to 3 digits starting at the right
        threeDigits = Right("000" & threeDigits, 3)

        'convert the hundreds place
        'if the first digit is > 0
        If Mid(threeDigits, 1, 1) <> "0" Then

                'convert the first digit to text and append " Hundred "
                result = common.oneDigitNumberToText(Mid(threeDigits, 1, 1)) & " Hundred "
        End If

        'convert the tens and ones place
        'if the second digit > 0 then convert the last two digits to text else only convert the last digit to text
        If Mid(threeDigits, 2, 1) <> "0" Then
                result = result & common.twoDigitNumberToText(Mid(threeDigits, 2))
        Else
                result = result & common.oneDigitNumberToText(Mid(threeDigits, 3))
        End If

        'return
        threeDigitNumberToText = result
End Function