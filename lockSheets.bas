'''''''''''''''''''''''''''''''''''''''''''''''
' Lock or Unlock Worksheets                   '
'''''''''''''''''''''''''''''''''''''''''''''''
'receives an input of optional locked (ex. True) and an optional input of sheets (ex. Array(ThisWorkbook.Sheets("Sheet 1")))
'locks or unlocks sheets based on locked input

'this sub will only lock the user interface features of the worksheets so that they can continue to be manipulated by
'VBA without the need to unlock

'this sub will allow the insertion of hyperlinks in unlocked cells

'this sub will allow the selection of unlocked cells

'''locked
'if locked is not supplied the sheets will be locked by default

'''sheets
'if sheets is not supplied all sheets on the workbook will be locked

Sub lockSheets(Optional locked As Boolean = True, Optional sheets As Variant)
        Dim wks As Worksheet
        Dim targetSheet As Variant
        Dim password as String: password = ""
        
        'if sheets was supplied
        If Not IsMissing(sheets) Then

                'loop through each supplied sheet
                If IsArray(sheets) Then
                        For Each targetSheet In sheets
                                If Not targetSheet Is Nothing Then
                                        setSheetLockState targetSheet, locked, password
                                End If
                        Next targetSheet

                'support a single supplied sheet
                ElseIf IsObject(sheets) Then
                        setSheetLockState sheets, locked, password
                End If

        'if sheets was NOT supplied
        Else

                'if unlock was requested
                If Not locked Then

                        'loop through each sheet in the workbook
                        For Each wks In ActiveWorkbook.Worksheets

                                'only unlock sheets that are locked
                                If wks.ProtectContents = True Then

                                        setSheetLockState wks, locked, password
                                End If
                        Next wks

                'if lock was requested
                Else

                        'loop through each sheet in the workbook
                        For Each wks In ActiveWorkbook.Worksheets

                                'only lock sheets that are unlocked
                                If wks.ProtectContents = False Then

                                        setSheetLockState wks, locked, password
                                End If
                        Next wks
                End If
        End If
End Sub

Private Sub setSheetLockState(ByVal targetSheet As Worksheet, ByVal locked As Boolean, ByVal password As String)
        If Not locked Then
                With targetSheet
                        .Unprotect Password:=password
                        .EnableSelection = xlNoRestrictions
                End With
        Else
                With targetSheet
                        .Protect Password:=password, UserInterfaceOnly:=True, AllowInsertingHyperlinks:=True
                        .EnableSelection = xlUnlockedCells
                End With
        End If
End Sub