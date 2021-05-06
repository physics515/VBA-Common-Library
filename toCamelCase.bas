'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Converts String To Camel Case                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of text (ex. "Hello World")
'outputs the string converted to camel case (ex. "helloWorld")

' *** Requires Function "toPascalCase" ***

Function toCamelCase(ByVal text As String) As String

        'dimension variables
        Dim result As String
        
        'convert string to pascal case
        result = common.toPascalCase(text)
        
        'if the string length is > 0
        If Len(result) > 0 Then

                'change the first character in the string to lower case
                Mid$(result, 1, 1) = LCase$(Mid$(result, 1, 1))
        End If

        'return result
        toCamelCase = result
End Function