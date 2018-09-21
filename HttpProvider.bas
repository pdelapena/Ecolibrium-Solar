Attribute VB_Name = "HttpProvider"
Public Function HttpGet(ByVal Url As String) As Object

    Dim hReq As Object
    Set hReq = CreateObject("MSXML2.XMLHTTP")
        With hReq
            .Open "GET", Url, False
            .Send
        End With
        
    Dim response As Object
    Set response = hReq
    
    Set HttpGet = response
End Function

Public Function HttpPost(ByVal Url As String, ByVal JsonBody As String) As Object

    Dim hReq As Object
    Set hReq = CreateObject("MSXML2.XMLHTTP")
        With hReq
            .Open "POST", Url, False
            .SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
            .Send (JsonBody)
        End With
        
    Dim response As Object
    Set response = hReq
    
    Set HttpPost = response
End Function
