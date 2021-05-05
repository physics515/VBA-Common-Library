'''''''''''''''''''''''''''''''''''''''''''''''
' Find A Query In Column                      '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of worksheet (ex. "Sheet 1"), search term (ex. "foo"), and range (ex. "A:A") as string
'outputs row number as integer

Function findQueryInColumn(searchWroksheet As String, searchTerm As Variant, searchColumn As String) As Integer

        '''method 1
        'dimension variables
        Dim wb As Workbook: Set wb = ThisWorkbook
        Dim ws As Worksheet: Set ws = wb.Sheets(searchWroksheet)
        Dim foundRow As Integer

        'find the search term within the search range
        Dim foundSearchTerm As range: Set foundSearchTerm = ws.range(searchColumn).Find(what:=searchTerm, after:=searchColumnLastCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)

        'if search term is found return row number
        'else try method 2
        If Not foundSearchTerm Is Nothing Then
                foundRow = foundSearchTerm.row
                
                ' return function
                findQueryInColumn = foundRow
                Exit Function
        Else
                '''method 2
                'dimension variables

                'find the column letter of the search column
                Dim searchColumnLetter As String: searchColumnLetter = common.Col_Letter(ws.range(searchColumn).Column)

                'find the last row in the search column
                Dim searchColumnLastRow As String: searchColumnLastRow = ws.range(common.Col_Letter(ws.range(searchColumn).Column) & Rows.count).End(xlUp).row

                'find the last cell in the search column
                Dim searchColumnLastCell As range: Set searchColumnLastCell = ws.range(searchColumnLetter & searchColumnLastRow)

                Dim Column As Integer: Column = ws.range(searchColumn).Column
                Dim i As Long
                Dim foundMatch As Boolean: foundMatch = False
                Dim compare As String
                
                'loop through each cell in search column
                For i = 1 To searchColumnLastCell.row

                        'convert to cell value to uppercase
                        compare = UCase(CStr(ws.range(common.Col_Letter(CLng(Column)) & i).Value))

                        'convert search term to upper case
                        searchTerm = UCase(searchTerm)

                        'if current cell value is equal to the search term return current row
                        If compare = searchTerm Then
                                findQueryInColumn = ws.range(common.Col_Letter(CLng(Column)) & i).row
                                foundMatch = True
                                Exit For
                        End If
                Next i
                
                ' if no match is found return 0
                If foundMatch = False Then
                        findQueryInColumn = 0
                End If
        End If
End Function