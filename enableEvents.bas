'''''''''''''''''''''''''''''''''''''''''''''''
' enable application events                   '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves a boolean value
'enable or disable events depending on bool value

Sub enableEvents(Optional enable As Boolean = True)

        'if input is fales disable events else enable events
        If Not enable Then
                Application.Calculation = xlCalculationManual
                Application.ScreenUpdating = False
                Application.enableEvents = False
        Else
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Application.enableEvents = True
        End If
End Sub