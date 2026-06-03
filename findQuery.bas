'''''''''''''''''''''''''''''''''''''''''''''''
' Find Query                                  '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Function "fuzzyFind" (for search algorithms 4 and 5) ***
'       *** Requires Enum Type "CaseSensitivity" ***
'       *** Requires Function "originalMetric" ***
'       *** Requires Function "damerau" ***
'             *** Requires Reference "Microsoft Scripting Library" ***
'       *** Requires Function "hamming" ***
'       *** Requires Function "levenshtein" ***
'       *** Requires Function "sorensenDice" ***
'             *** Requires Function "ngrams" ***
'       *** Requires Function "tversky" ***
'             *** Requires Function "uniqueArrayElements" ***
'       *** Requires Function "jaccard" ***
'       *** Requires Function "jaroWinkler" ***
'       *** Requires Function "simpleMatching" ***
'       *** Requires Function "min" ***
'       *** Requires Function "max" ***

'receives input of:
'       searchWorksheet (ex. ThisWorkbook.Sheets("Sheet 1")) as Worksheet
'       searchRange (ex. ThisWorkbook.Sheets("Sheet 1").Range("A:A")) as Range
'       searchTerm (ex. "foo") as String
'       optional searchAlgorithm (ex. 2) as Integer: default = 2
'           1 - Quick Search: Uses Excel's built-in Find function. Works quickly but does not always return the correct value.
'           2 - Accurate Search: Uses Excel's built-in Find function. If no result is found, falls back to Brute Force Search. A middle ground for accuracy and speed. (Default)
'           3 - Brute Force Search: Loops through every cell in the range and manually compares the values. Returns the first matched cell. Slow but always gets the job done.
'           4 - Return Something Search: Uses Accurate Search. If nothing is returned, falls back to Fuzzy Search.
'           5 - Fuzzy Search: Uses common.fuzzyFind to return the closest match cell.

'outputs an Integer Array [column, row]. If nothing is found the function returns [0, 0].

''' From the Author '''
'@Description: Finds a queried value in a specified range and returns the column and row where the query is found.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0

Function findQuery(searchWorksheet As Worksheet, searchRange As Range, searchTerm As String, Optional searchAlgorithm As Integer = 2) As Variant

        'dimension variables
        Dim result(1) As Integer
        result(0) = 0
        result(1) = 0

        Dim foundCell As Range
        Dim cell As Range
        Dim fuzzyMatchValue As String
        Dim fallbackResult As Variant

        Select Case searchAlgorithm

                Case 1
                        'Quick Search: Use Excel's built-in Find
                        Set foundCell = searchRange.Find(what:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                        If Not foundCell Is Nothing Then
                                result(0) = foundCell.Column
                                result(1) = foundCell.Row
                        End If
                        findQuery = result

                Case 2
                        'Accurate Search: Try Quick Search, fall back to Brute Force
                        Set foundCell = searchRange.Find(what:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                        If Not foundCell Is Nothing Then
                                result(0) = foundCell.Column
                                result(1) = foundCell.Row
                                findQuery = result
                                Exit Function
                        End If

                        'fall back to Brute Force
                        For Each cell In searchRange
                                If UCase(CStr(cell.Value)) = UCase(searchTerm) Then
                                        result(0) = cell.Column
                                        result(1) = cell.Row
                                        findQuery = result
                                        Exit Function
                                End If
                        Next cell
                        findQuery = result

                Case 3
                        'Brute Force Search: Loop through every cell in the range
                        For Each cell In searchRange
                                If UCase(CStr(cell.Value)) = UCase(searchTerm) Then
                                        result(0) = cell.Column
                                        result(1) = cell.Row
                                        findQuery = result
                                        Exit Function
                                End If
                        Next cell
                        findQuery = result

                Case 4
                        'Return Something Search: Accurate Search, fall back to Fuzzy Search
                        fallbackResult = findQuery(searchWorksheet, searchRange, searchTerm, 2)
                        If fallbackResult(0) <> 0 Or fallbackResult(1) <> 0 Then
                                findQuery = fallbackResult
                                Exit Function
                        End If
                        fallbackResult = findQuery(searchWorksheet, searchRange, searchTerm, 5)
                        findQuery = fallbackResult

                Case 5
                        'Fuzzy Search: Use common.fuzzyFind to find the closest matching cell
                        fuzzyMatchValue = common.fuzzyFind(searchTerm, searchRange, searchWorksheet, CaseSensitivity.NotSensitive)
                        Set foundCell = searchRange.Find(what:=fuzzyMatchValue, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                        If Not foundCell Is Nothing Then
                                result(0) = foundCell.Column
                                result(1) = foundCell.Row
                        End If
                        findQuery = result

                Case Else
                        'Default to Accurate Search for unrecognized algorithm values
                        findQuery = findQuery(searchWorksheet, searchRange, searchTerm, 2)

        End Select

End Function
