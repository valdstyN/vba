Option Explicit
Public xmlHttpRequest As MSXML2.XMLHTTP60

Function UpdateData(strURL, ppayload, strUser, strPass, error) As String

   Dim uri As String, payload As String, sendUrl As String
   Dim env As String, user As String, pwd As String, encB64 As String, token as String
        
   env = "https://..."
   user = "API_user"
   pwd = "API_password"
   encB64 = Replace(EncodeBase64(user & ":" & pwd), Chr(10), "")
  
   Set xmlHttpRequest = New MSXML2.XMLHTTP60
    
   ' get a token (authenticated GET request with header x-csrf-token = Fetch here returns a token)
   xmlHttpRequest.Open "GET", env & "/cpd/SC_PROJ_ENGMT_CREATE_UPD_SRV/A_CustProjSlsOrd?$top=1", False
   xmlHttpRequest.SetRequestHeader "X-CSRF-Token", "Fetch"
   xmlHttpRequest.SetRequestHeader "Content-Type", "application/json"
   xmlHttpRequest.SetRequestHeader "Authorization", "Basic " & encB64
   xmlHttpRequest.send ""
   token = xmlHttpRequest.GetResponseHeader("X-CSRF-Token")

   ' now send a POST request
   method  = "POST
   uri = "/cpd./endpoint"
   sendUrl = env & uri
   xmlHttpRequest.Open method, sendUrl, False
   xmlHttpRequest.SetRequestHeader "X-CSRF-Token", token
   xmlHttpRequest.SetRequestHeader "Content-Type", "application/json"
   xmlHttpRequest.SetRequestHeader "Authorization", "Basic " & encB64
   myPayload = "here your payload"
   xmlHttpRequest.send myPayload
    
End Function
   
' encode a string to base 64
Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)
  Dim objXML As MSXML2.DOMDocument60
  Dim objNode As MSXML2.IXMLDOMElement
  Set objXML = New MSXML2.DOMDocument60
  Set objNode = objXML.createElement("b64")
  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text
  Set objNode = Nothing
  Set objXML = Nothing
End Function
