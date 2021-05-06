'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Remove Duplicate Values From Range                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'recieves input as originSheet (ex. ThisWorkbook.Sheets("Sheet 1")) as worksheet, and originRange (ex. ThisWorkbook.Sheets("Sheet 1").Range("A1:B5")) as range
'outputs a Scripting.Dictionary of which the keys are the values of the range with duplicates removed


' *** Note: Issue: This function only gets values from the fist column in the origin range ***

Function removeDuplicates(originSheet As Worksheet, originRange As range) As Scripting.dictionary

        'diminsion variables
        Dim i As Long
        Dim dictionary As New Scripting.dictionary
        Dim dictAdd As String

        'find the first row of the origin range
        Dim FirstRow As Long: FirstRow = originSheet.range(originRange.Address).row

        'find the last row of the origin range
        Dim LastRow As Long: LastRow = originRange.row + originRange.Rows.count - 1
        
        'ignore errors
        On Error Resume Next

                'loop through each row in range
                For i = FirstRow To LastRow

                        'find the current cell's value
                        dictAdd = originSheet.range(common.Col_Letter(originRange.Column) & i).Value

                        'if the current cell's value is not blank
                        If Len(dictAdd) <> 0 Then

                                'add current cell's value to the dictionary
                                dictionary.Add dictAdd, 1
                        End If
                Next i

        On Error Goto 0
        
        'return dictionary
        Set removeDuplicates = dictionary
End Function