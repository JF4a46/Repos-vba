Sub request()

'Dim xmlhttp As New MSXML2.XMLHTTP60, myurl As String


xmlhttp.Open "GET", myurl, False
xmlhttp.send
Debug.Print xmlhttp.responseText

'inac = Add_Item("testinglist", "ur", "test1", "Title")


End Sub

Sub RESTCall()
    Const sUrl As String = "https://synergi.ssc-spc.gc.ca/IS/SMO-OGS/SMTPS/Lists/logs/AllItems.aspx"
    Dim oRequest As WinHttp.WinHttpRequest
    Dim sResult As String
    
    Set oRequest = New WinHttp.WinHttpRequest
With oRequest
    .Open "GET", sUrl, True
    .setRequestHeader "Content-Type", "application/json"
    .SetCredentials "domain\administrator", "pw", HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
    .send
    .waitForResponse
    sResult = .responseText
    Debug.Print sResult
    sResult = oRequest.Status
    Debug.Print sResult
End With
End Sub


Function Add_ItemArray(ListName As String, SharepointUrl As String, ValueVar As Variant, FieldNameVar As Variant)

Dim objXMLHTTP As MSXML2.xmlhttp

Dim strListNameOrGuid As String
Dim strBatchXml As String
Dim strSoapBody As String

Set objXMLHTTP = New MSXML2.xmlhttp

strListNameOrGuid = ListName


'Add New Item'
strBatchXml = "<Batch OnError='Continue'><Method ID='3' Cmd='New'><Field Name='" + FieldNameVar + "'>" + ValueVar + "</Field></Method></Batch>"


objXMLHTTP.Open "POST", SharepointUrl & "_vti_bin/Lists.asmx", False
objXMLHTTP.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
objXMLHTTP.setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"

strSoapBody = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " _
 & "xmlns:xsd='http://www.w3.org/2001/XMLSchema' " _
 & "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><UpdateListItems " _
 & "xmlns='http://schemas.microsoft.com/sharepoint/soap/'><listName>" & strListNameOrGuid _
 & "</listName><updates>" & strBatchXml & "</updates></UpdateListItems></soap:Body></soap:Envelope>"

objXMLHTTP.send strSoapBody

If objXMLHTTP.Status = 200 Then
    MsgBox "reussit"
End If

Set objXMLHTTP = Nothing

End Function






Function Add_Item(ListName As String, SharepointUrl As String, ValueVar As String, FieldNameVar As String)

Dim objXMLHTTP As MSXML2.xmlhttp

Dim strListNameOrGuid As String
Dim strBatchXml As String
Dim strSoapBody As String

Set objXMLHTTP = New MSXML2.xmlhttp

strListNameOrGuid = ListName


'Add New Item'
strBatchXml = "<Batch OnError='Continue'><Method ID='3' Cmd='New'><Field Name='ID'>New</Field><Field Name='" + FieldNameVar + "'>" + ValueVar + "</Field></Method></Batch>"


objXMLHTTP.Open "POST", SharepointUrl & "_vti_bin/Lists.asmx", False
objXMLHTTP.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
objXMLHTTP.setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"

strSoapBody = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " _
 & "xmlns:xsd='http://www.w3.org/2001/XMLSchema' " _
 & "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><UpdateListItems " _
 & "xmlns='http://schemas.microsoft.com/sharepoint/soap/'><listName>" & strListNameOrGuid _
 & "</listName><updates>" & strBatchXml & "</updates></UpdateListItems></soap:Body></soap:Envelope>"

objXMLHTTP.send strSoapBody

If objXMLHTTP.Status = 200 Then
    MsgBox "reussit"
End If

Set objXMLHTTP = Nothing

End Function



