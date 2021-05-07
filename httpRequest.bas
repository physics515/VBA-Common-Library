'''''''''''''''''''''''''''''''''''''''''''''''
' Http Request                                '
'''''''''''''''''''''''''''''''''''''''''''''''
'recieves url as string, and post (ex. True or False) as boolean
'outputs the http resonse

'this function defaults to a GET request unless post boolean is true

Function httpRequest(url As String, Optional post as Boolean) As String

        'dimension variables
        Dim hReq As Object
        Dim apiKeys as String: apiKeys = ""
        Dim httpType as String: type = "GET"

        'is the request to POST or GET
        If post Then
                httpType = "POST"
        End If
        
        'on error return ""
        On Error GoTo httpRequestError
                
                'create XML HTTP object
                Set hReq = CreateObject("MSXML2.XMLHTTP")

                'open request and assign headers
                With hReq
                        .Open httpType, url, False
                        .SetRequestHeader "Authorization", "Basic " & common.Base64Encode(apiKeys)
                        .Send
                End With
                
                'return response
                httpRequest = hReq.ResponseText
                
                'garbage collection
                Set hReq = Nothing  
        Exit Function

httpRequestError:
        httpRequest = ""
End Function