'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generate Range Of Available Printers              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Function "findQueryInRow" ***
' *** Requires Function "getColumnLetter" ***

'recieves an input of where the list should be saved in the form of desinationSheet (ex. "Sheet 1") as string, and destinationColumnHeader (ex. "Printer List") as string.
'outputs a list of currently connected printers to an excel column based on the sheet name an the column header

Sub generateRangeOfAvailablePrinters(destinationSheet As String, destinationColumnHeader as String)

        'dimension variables
        Dim wb As Workbook: Set wb = ThisWorkbook
        Dim ws As Worksheet: Set ws = wb.Sheets(destinationSheet)
        Dim printers() As String
        Dim i As Integer: i = 1

        'add optional selections to the top of the printer list
        ReDim printers(1 To 2)
        printers(1) = "-- SELECT PRINTER --"
        printers(2) = "*** Print to PDF ***"
        
        'create a nextwork object
        With CreateObject("WScript.Network")

                'loop through each printer on the network by incrementing by 2 each loop
                For i = 1 To .EnumPrinterConnections.count Step 2

                        'redimension the printers array
                        ReDim Preserve printers(1 To UBound(printers) + 1)

                        'add the printer to the to the printers array
                        printers(UBound(printers)) = .EnumPrinterConnections(i)
                Next i
        End With
        
        'loop through each item in the printers array
        For i = 1 To UBound(printers)
        
                'copy current of printer array to destination
                ws.range(common.getColumnLetter(common.findQueryInRow(destinationSheet, destinationColumnHeader, "1:1")) & i + 1).Value = printers(i)
        Next i
End Function