'''''''''''''''''''''''''''''''''''''''''''''''
' Convert String To Binary                    '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Note: 2003 Antonin Foller, http://www.motobit.com ***
'recieves text (ex. "Hello World") as string
'outputs binary data

Function stringToBinary(Text As String)

        'dimension constants
        Const adTypeText = 2
        Const adTypeBinary = 1

        'create stream object
        Dim BinaryStream
        Set BinaryStream = CreateObject("ADODB.Stream")

        'specify stream type - we want to save text/string data
        BinaryStream.Type = adTypeText

        'specify charset for the source text (unicode) data
        BinaryStream.Charset = "us-ascii"

        'open the stream and write text/string data to the object
        BinaryStream.Open
        BinaryStream.WriteText Text

        'change stream type to binary
        BinaryStream.Position = 0
        BinaryStream.Type = adTypeBinary

        'ignore first two bytes - sign of
        BinaryStream.Position = 0

        'return - open the stream and get binary data from the object
        stringToBinary = BinaryStream.Read
        
        'garbage collection
        Set BinaryStream = Nothing
End Function