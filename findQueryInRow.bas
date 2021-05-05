'''''''''''''''''''''''''''''''''''''''''''''''
' Find A Query In Row                         '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of worksheet (ex. "Sheet 1"), search term (ex. "foo"), and range (ex. "1:1") as string
'outputs column number as integer

Function findQueryInRow(searchWorksheet As String, searchTerm As Variant, searchRow As String) As Integer

        '''method 1
        'dimension variables
        Dim wb As Workbook: Set wb = ThisWorkbook
        Dim ws As Worksheet: Set ws = wb.Sheets(searchWorksheet)
        Dim foundCol As Integer
        
        'find the search term within the search range
        Dim foundSearchTerm As range: Set foundSearchTerm = ws.range(searchRow).Find(what:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)

        'if search term is found return row number
        'else try method 2
        If Not foundSearchTerm Is Nothing Then
                foundCol = foundSearchTerm.Column

                ' return function
                findQueryInRow = foundCol
                Exit Function
        Else
                '''method 2
                'dimension variables

                'find the column letter of the last column in the search row
                Dim searchRowLastColumnLetter As String: searchRowLastColumnLetter = common.Col_Letter(ws.range(common.Col_Letter(Columns.count) & ws.range(searchRow).row).End(xlToLeft).Column)

                'find the last row in the search search row
                Dim searchRowLastRow As String: searchRowLastRow = ws.range(searchRow).row

                'find the last cell in the search row
                Dim searchRowLastCell As range: Set searchRowLastCell = ws.range(searchRowLastColumnLetter & searchRowLastRow)
        
                Dim row As Integer: row = ws.range(searchRow).row
                Dim i As Long
                Dim foundMatch As Boolean: foundMatch = False
                Dim compare As String
                
                'loop through each cell in search row
                For i = 1 To searchRowLastCell.Column

                        'convert to cell value to uppercase
                        compare = UCase(CStr(ws.range(common.Col_Letter(CLng(i)) & row).Value))

                        'convert search term to upper case
                        searchTerm = UCase(searchTerm)

                        'if current cell value is equal to the search term return current column
                        If compare = searchTerm Then
                                findQueryInRow = ws.range(common.Col_Letter(CLng(i)) & row).Column
                                foundMatch = True
                                Exit For
                        End If
                Next i

                ' if no match is found return 0
                If foundMatch = False Then
                        findQueryInRow = 0
                End If
        End If
End Function