Attribute VB_Name = "Ä£¿é1"
Function lbs(infos) As String
    arr = Split(infos, ",")
    mcc = arr(0)
    mnc = arr(1)
    If mcc = 460 And mnc = 3 Then
        mnc = 11
    End If
    lac = arr(2)
    ci = arr(3)

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Url = "http://api.cellocation.com:81/cell/?coord=gcj02&output=csv&mcc=" & mcc & "&mnc=" & mnc & "&lac=" & lac & "&ci=" & ci
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    lbs = Split(objHTTP.ResponseText, """")(1)
    
End Function

Function ipapi(ip) As String
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    'Url = "http://ip-api.com/json/" & ip & "?lang=zh-CN&fields=17"
    Url = "http://ip.cz88.net/data.php?ip=" & ip
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    ipapi = Split(objHTTP.ResponseText, "'")(3)
End Function
