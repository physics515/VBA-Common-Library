'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Convert (Named)Range to Delimited List      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'recieves input as workSheetName (ex. "Sheet 1") as string, rangeName (ex. "A1:B5" or "clientNames") as string, and delimitor (ex. ";" or ", ") as string
'outputs a string of all values in the range separated by a delimitor

Function convertRangeToDelimitedLists(workSheetName As String, rangeName As String, delimitor as String) As String

        'diminsion variables
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
        
        'return list
        convertRangeToCommaDelimitedLists = lst
End Function