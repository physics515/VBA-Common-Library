'''''''''''''''''''''''''''''''''''''''''''''''
' Get Last Row In Column                      '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Function "getColumnLetter" ***

'recieves input of worksheet (ex. "Sheet 1") and column range (ex. "A:A") as string
'outputs last used row number as long

Function lastRow(searchWorksheet As String, searchColumn As String) As Long
        'dimension variables
        Dim wb As Workbook: Set wb = ThisWorkbook
        Dim ws As Worksheet: Set ws = wb.Sheets(searchWorksheet)

        'find the column letter of the search column
        Dim searchColumnLetter As String: searchColumnLetter = common.getColumnLetter(ws.Range(searchColumn).Column)

        'find and return the last row in the search column
        lastRow = ws.Range(searchColumnLetter & Rows.Count).End(xlUp).Row
End Function
