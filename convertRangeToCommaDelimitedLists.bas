'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Convert (Named)Range to Delimited List      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'receives input as workSheetName (ex. "Sheet 1") as string, rangeName (ex. "A1:B5" or "clientNames") as string, delimitor (ex. ";" or ", ") as string, and optional removeFinalDelimiter as boolean
'outputs a string of all values in the range separated by a delimitor, with optional removal of the trailing delimitor

Function convertRangeToDelimitedLists(workSheetName As String, rangeName As String, delimitor As String, Optional removeFinalDelimiter As Boolean = False) As String

        'dimension variables
        Dim rng As range
        Dim cell As range
        Dim lst As String
        
        'find the range
        Set rng = range(ThisWorkbook.Sheets(workSheetName).Range(rangeName).Address)
        
        'loop through reach cell in the range
        For Each cell In rng

                'add value to list
                lst = lst & cell.Value & delimitor
        Next cell
        
        If removeFinalDelimiter And Len(delimitor) > 0 And Len(lst) >= Len(delimitor) Then
                If Right$(lst, Len(delimitor)) = delimitor Then
                        lst = Left$(lst, Len(lst) - Len(delimitor))
                End If
        End If

        'return list
        convertRangeToDelimitedLists = lst
End Function