'''''''''''''''''''''''''''''''''''''''''''''''
' Convert String to Base64 encoding           '
'''''''''''''''''''''''''''''''''''''''''''''''
' recieves input of type String
' outputs same string with base64 encoding applied

Function base64Encode(input as String) As String

        'dimension tools
        Dim oXML, oNode
        Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
        Set oNode = oXML.createElement("base64")
        
        'set datatype
        oNode.DataType = "bin.base64"

        'encode to base64
        oNode.nodeTypedValue = Stream_StringToBinary(sText)

        'return
        Base64Encode = oNode.Text
        
        'garbage collection
        Set oNode = Nothing
        Set oXML = Nothing

End Function