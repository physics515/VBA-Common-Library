'''''''''''''''''''''''''''''''''''''''''''''''
' Enumerated Case Sensitivity Type            '
'''''''''''''''''''''''''''''''''''''''''''''''
'Options
'       Sensitive: Should consider case.
'       NotSensitive: Should not consider case.
'       DefaultSensitivity: Should retain the default case consideration.

''' From the Author '''
'@Description: Case sensitivity user-defined type for use in defining the case sensitivity of a function.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0

Public Enum CaseSensitivity
     Sensitive = 0
     NotSensitive = 1
     DefaultSensitivity = 2
End Enum