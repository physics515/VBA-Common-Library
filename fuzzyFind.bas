'''''''''''''''''''''''''''''''''''''''''''''''
' Fuzzy Find                                  '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of query (ex. "foo"), searchRange (ex. "A1:B5") and searchSheet (ex. "Sheet 1") as string
'outputs the closest match to the query text as a string

Function fuzzyFind(query As String, searchRange As String, searchSheet As String) As String

        'dimension variables
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(searchSheet)
        Dim lookupRange As range: Set lookupRange = ws.range(searchRange)
        Dim cell As range
        Dim i As Integer
        Dim j As Integer
        Dim topScore As Double
        Dim currentScore As Double
        Dim topScoringCell As range
        Dim foundPosition As Integer: foundPosition = 0
        Dim distance As Double
        
        'loop through each cell in search range
        For Each cell In lookupRange

                'loop through each letter in query
                For i = 1 To Len(query)

                        'loop through each letter in the current cell value
                        For j = 1 To Len(cell.Value)

                                'if the current query letter apeers in the the cell value then update the score
                                If InStr(1, Mid(cell.Value, j, 1), Mid(query, i, 1), vbTextCompare) > 0 Then

                                        'if the last letter that was found came before the current letter (the letters are found in the correct order)
                                        If foundpostition < j Then

                                                'allways return a positive distance
                                                If i < j Then
                                                        distance = j - i
                                                Else
                                                        distance = i - j
                                                End If

                                                'allways return a non-zero distance
                                                ' If distance = 0 Then
                                                '         distance = distance + 0.5
                                                ' End If

                                                'record the current found position as the previouly found position
                                                foundPosition = j

                                                'add distance to socre as a percentage matched, and round to 4 decimal places
                                                If distance > 0 Then
                                                        currentScore = Round(Abs(currentScore + Round((1 / distance), 4)), 4)
                                                End If
                                        End If
                                End If
                        Next j
                Next i
                
                'if the current cells is the highest scoring cell then record it as the top score and record its range
                If currentScore > topScore Then
                        Set topScoringCell = cell
                        topScore = currentScore
                End If
                
                'reset current scores and positions
                foundPosition = 0
                currentScore = 0
        Next cell
        
        'return the value of the top scoring cell
        fuzzyFind = topScoringCell.Value
End Function