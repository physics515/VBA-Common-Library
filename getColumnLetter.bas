'''''''''''''''''''''''''''''''''''''''''''''''
' Get Column Letter From Number               '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves column number as input
'returns the corresponding column letter

Function getColumnLetter(columnNumber As Long) As String
        'dimension variables
        Dim ColumnLetter As String
        
        'get the column letter from the specified address
        ColumnLetter = Split(Cells(1, columnNumber).Address, "$")(1)

        'return
        Col_Letter = ColumnLetter
End Function