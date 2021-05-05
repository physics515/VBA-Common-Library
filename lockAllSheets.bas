'''''''''''''''''''''''''''''''''''''''''''''''
' Lock or Unlock All Worksheets               '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves an input of optional locked (ex. True) and an optional input of singleSheet (ex. ThisWorkbook.Sheets("Sheet 1"))
'locks or unlocks sheets based on locked input

'this sub will only lock the user interface features of the worksheets so that they can continue to be manipulated by
'VBA without the need to unlock

'this sub will allow the insertion of hyperlinks in unlocked cells

'this sub will allow the selection of unlocked cells

'''locked
'if locked is not supplied the sheets will be locked by default

'''singleSheet
'if singleSheet is not supplied all sheets on the workbook will be locked

Sub lockAllSheets(Optional locked As Boolean = True, Optional singleSheet As Worksheet)
        Dim wks As Worksheet
        Dim password as String: password = ""
        
        'if singleSheet was supplied
        If Not singleSheet Is Nothing Then

                'if unlock was requested then ulock the single sheet else lock the single sheet
                If Not locked Then
                        singleSheet.Unprotect Password:=password
                Else
                        With singleSheet
                                .Protect Password:=password, UserInterfaceOnly:=True, AllowInsertingHyperlinks:=True
                                .EnableSelection = xlUnlockedCells
                        End With
                End If

        'if singleSheet was NOT supplied
        Else

                'if unlock was requested
                If Not locked Then

                        'loop through each sheet in the workbook
                        For Each wks In ActiveWorkbook.Worksheets

                                'only unlock sheets that are locked
                                If wks.ProtectContents = True Then

                                        'unlock sheet
                                        With wks
                                                .Unprotect Password:=password
                                                .EnableSelection = xlNoRestrictions
                                        End With
                                End If
                        Next wks

                'if lock was requested
                Else

                        'loop through each sheet in the workbook
                        For Each wks In ActiveWorkbook.Worksheets

                                'only lock sheets that are unlocked
                                If wks.ProtectContents = False Then

                                        'lock sheet
                                        With wks
                                                .Protect Password:=password, UserInterfaceOnly:=True, AllowInsertingHyperlinks:=True
                                                .EnableSelection = xlUnlockedCells
                                        End With
                                End If
                        Next wks
                End If
        End If
End Sub