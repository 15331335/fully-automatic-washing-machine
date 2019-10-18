Attribute VB_Name = "Ä£¿é1"
Function lbs(mcc, mnc, lac, ci) As String
    target_lac = lac  '55254
    target_ci = ci  '162643337

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Url = "http://api.cellocation.com:81/cell/?coord=gcj02&output=csv&mcc=" & mcc & "&mnc=" & mnc & "&lac=" & lac & "&ci=" & ci
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    lbs = objHTTP.ResponseText
    
End Function

Function ipapi(ip) As String
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Url = "http://ip-api.com/json/" & ip & "?lang=zh-CN&fields=17"
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    ipapi = objHTTP.ResponseText
End Function
