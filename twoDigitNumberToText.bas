'''''''''''''''''''''''''''''''''''''''''''''''
' Two Digit number To Text                    '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves any two digit number as twoDigits (ex. "12") as string
'outputs a string as text (ex. "Twelve")

' *** Requires Function "oneDigitNumberToText" ***

Function twoDigitNumberToText(twoDigits As String) as String

        'dimension variables
        Dim result As String: result = ""
        
        ' If value between 10-19
        If val(Left(twoDigits, 1)) = 1 Then
                Select Case val(twoDigits)
                        Case 10: result = "Ten"
                        Case 11: result = "Eleven"
                        Case 12: result = "Twelve"
                        Case 13: result = "Thirteen"
                        Case 14: result = "Fourteen"
                        Case 15: result = "Fifteen"
                        Case 16: result = "Sixteen"
                        Case 17: result = "Seventeen"
                        Case 18: result = "Eighteen"
                        Case 19: result = "Nineteen"
                        Case Else
                End Select

        ' If value between 20-99
        Else
                Select Case val(Left(twoDigits, 1))
                        Case 2: result = "Twenty "
                        Case 3: result = "Thirty "
                        Case 4: result = "Forty "
                        Case 5: result = "Fifty "
                        Case 6: result = "Sixty "
                        Case 7: result = "Seventy "
                        Case 8: result = "Eighty "
                        Case 9: result = "Ninety "
                        Case Else
                End Select

                ' Retrieve ones place
                result = result & common.oneDigitNumberToText(Right(twoDigits, 1))
        End If

        'return
        twoDigitNumberToText = result
End Function