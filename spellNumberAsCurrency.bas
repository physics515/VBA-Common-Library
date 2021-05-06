'''''''''''''''''''''''''''''''''''''''''''''''
' Spell Numbers As Currency                   '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves numberToSpell (ex. 123) as variant, and currencyName (ex. "Dollars") as string
'outputs number written as string (ex. "One Hundred Twenty Three Dollars And No Cents")

' *** Requires Function "oneDigitNumberToText" ***
' *** Requires Function "twoDigitNumberToText" ***
' *** Requires Function "threeDigitNumberToText" ***

Function spellNumberAsCurrency(ByVal numberToSpell As Variant, currencyName As String) As String

        'dimension variables
        Dim Dollars As String
        Dim Cents As String
        Dim temp As String
        Dim DecimalPlace As Integer
        Dim count As Integer
        Dim Place(1 To 9) As String

        'define decimal place terms
        Place(2) = " Thousand "
        Place(3) = " Million "
        Place(4) = " Billion "
        Place(5) = " Trillion "
        
        'convert number to spell to a string and trim leading and trailing spaces
        numberToSpell = Trim(str(numberToSpell))

        'find the position of the decimal point
        DecimalPlace = InStr(numberToSpell, ".")

        'if number to spell contains a decimal
        If DecimalPlace > 0 Then

                'send the first two digits to the right of the decimal place to the getTens function
                Cents = common.twoDigitNumberToText(Left(Mid(numberToSpell, DecimalPlace + 1) & "00", 2))

                'remove the decimal place and all digits to the right of it
                numberToSpell = Trim(Left(numberToSpell, DecimalPlace - 1))
        End If

        'loop until numberToSpell = ""
        count = 1
        Do While numberToSpell <> ""

                'convert the rightmost three digits to text
                temp = common.threeDigitNumberToText(Right(numberToSpell, 3))

                'if temp returns value append templ and place text to Dollars
                If temp <> "" Then Dollars = temp & Place(count) & Dollars

                'if numberToSpell is greater that 3 digits long
                If Len(numberToSpell) > 3 Then

                        'trim the rightmost 3 digits
                        numberToSpell = Left(numberToSpell, Len(numberToSpell) - 3)
                Else
                        'else set numberToSpell to ""
                        numberToSpell = ""
                End If
                
                'itterate count
                count = count + 1
        Loop

        'append currencyName to Dollars
        Select Case Dollars
                Case ""
                Dollars = "No " & currencyName

                Case "One"
                Dollars = "One" & currencyName
                
                Case Else
                Dollars = Dollars & " " & currencyName
        End Select

        'append cent to Cents
        Select Case Cents
                Case ""
                Cents = " and No Cents"

                Case "One"
                Cents = " and One Cent"

                Case Else
                Cents = " and " & Cents & " Cents"
        End Select

        'return
        SpellNumber = Dollars & Cents
End Function