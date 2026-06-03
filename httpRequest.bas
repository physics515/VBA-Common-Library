'''''''''''''''''''''''''''''''''''''''''''''''
' Http Request                                '
'''''''''''''''''''''''''''''''''''''''''''''''
'receives url as string, optional request method (or legacy post boolean), and optional request body
'outputs the http response

'this function defaults to a GET request unless a supported request method is provided

Function httpRequest(url As String, Optional post As Variant, Optional requestBody As Variant) As String

        'dimension variables
        Dim hReq As Object
        Dim apiKeys as String: apiKeys = ""
        Dim httpType as String: httpType = "GET"

        'allow legacy boolean post flag while supporting explicit request methods
        If Not IsMissing(post) Then
                If VarType(post) = vbBoolean Then
                        If CBool(post) Then
                                httpType = "POST"
                        End If
                ElseIf Trim$(CStr(post)) <> "" Then
                        httpType = UCase$(Trim$(CStr(post)))
                End If
        End If

        Select Case httpType
                Case "GET", "HEAD", "POST", "PUT", "DELETE", "CONNECT", "OPTIONS", "TRACE", "PATCH"
                Case Else
                        httpType = "GET"
        End Select
        
        'on error return ""
        On Error GoTo httpRequestError
                
                'create XML HTTP object
                Set hReq = CreateObject("MSXML2.XMLHTTP")

                'open request and assign headers
                With hReq
                        .Open httpType, url, False
                        .SetRequestHeader "Authorization", "Basic " & common.Base64Encode(apiKeys)
                        If IsMissing(requestBody) Or httpType = "GET" Or httpType = "HEAD" Then
                                .Send
                        Else
                                .Send requestBody
                        End If
                End With
                
                'return response
                If httpType = "HEAD" Then
                        httpRequest = hReq.GetAllResponseHeaders
                Else
                        httpRequest = hReq.ResponseText
                End If
                
                'garbage collection
                Set hReq = Nothing  
        Exit Function

httpRequestError:
        httpRequest = ""
End Function