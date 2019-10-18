Attribute VB_Name = "模块1"
Function a(lac, ci) As String
    target_lac = lac  '55254
    target_ci = ci  '162643337
    Dim URLStr As String        'API
    Dim originalStr As String   '网络请求后的原始字符串数据
    
    '随意使用的一个免费API用于下面的练习
    'URLStr = "http://api.cellocation.com:81/cell/?coord=gcj02&output=csv&mcc=460&mnc=1&lac=" & target_lac & "&ci=" & target_ci
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Url = "http://api.cellocation.com:81/cell/?coord=gcj02&output=csv&mcc=460&mnc=1&lac=" & target_lac & "&ci=" & target_ci
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    'MsgBox (objHTTP.ResponseText)
    a = objHTTP.ResponseText
    
End Function
