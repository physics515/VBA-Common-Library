'''''''''''''''''''''''''''''''''''''''''''''''
' Convert Binary To String                    '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Note: 2003 Antonin Foller, http://www.motobit.com ***
'recieves bianary data (ex. VT_UI1 | VT_ARRAY) as variant
'outputs a string

Function binaryToString(Binary As Variant) As String

        'dimension constants
        Const adTypeText = 2
        Const adTypeBinary = 1

        'create stream object
        Dim BinaryStream 'As New Stream
        Set BinaryStream = CreateObject("ADODB.Stream")

        'specify stream type - we want to save text/string data
        BinaryStream.Type = adTypeBinary

        'open the stream and write text/string data to the object
        BinaryStream.Open
        BinaryStream.Write Binary

        'change stream type To binary
        BinaryStream.Position = 0
        BinaryStream.Type = adTypeText

        'specify charset for the source text (unicode) data
        BinaryStream.Charset = "us-ascii"

        'return - open the stream and get binary data from the object
        binaryToString = BinaryStream.ReadText

        'garbage collection
        Set BinaryStream = Nothing
End Function