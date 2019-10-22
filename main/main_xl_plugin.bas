Attribute VB_Name = "Module2"
Function api()
    Dim url As String
    url = "https://stats.pacificdata.org/data-nsi/Rest/dataflow/SPC/all?detail=allstubs"
        
    Dim hReq As New MSXML2.XMLHTTP60
    hReq.Open "GET", url, False
    ' hReq.send
    
    hReq.send CreateObject("ADODB.Stream")
    
    Dim resp As String
    ' Get XML
    resp = hReq.responseText
    
    Dim objXML As New MSXML2.DOMDocument60
    'MsgBox TypeName(resp)
    If objXML.LoadXML(resp) Then
        objXML.setProperty "SelectionNamespaces", "xmlns:message=""http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message"" xmlns:structure=""http://www.sdmx.org/resources/sdmxml/schemas/v2_1/structure"" xmlns:common=""http://www.sdmx.org/resources/sdmxml/schemas/v2_1/common"""
        'MsgBox objXML.SelectSingleNode("message:Structure/message:Header/message:ID").Text
    Else
        MsgBox "Error"
    End If
    
    Dim xPath As String
    xPath = "message:Structure/message:Structures/structure:Dataflows/structure:Dataflow"
    'xPath = "Structure/Header/ID"
    
    Dim nodeList As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    
    Set nodeList = objXML.SelectNodes(xPath)
    
    ' Get length of nodelist
    Dim nodeLen As Integer
    nodeLen = nodeList.Length
    ' Make arrays for IDs and Names of DFs
    Dim dfIds() As String
    ReDim dfIds(nodeLen)
    Dim dfNames() As String
    ReDim dfNames(nodeLen)
    
    Dim X As Integer
    X = 0
    Dim dfId As String
    Dim dfName As String
    Dim dict As Collection
    Set dict = New Collection
    For Each node In nodeList
        dfIds(X) = node.Attributes.getNamedItem("id").Text
        dfNames(X) = node.SelectSingleNode("common:Name").Text
        X = X + 1
        'MsgBox "Dataflow: " & node.SelectSingleNode("common:Name").Text & "ID: " & node.Attributes.getNamedItem("id").Text
    Next
    'MsgBox Join(dfNames, vbCrLf)
    'MsgBox Join(dfIds, vbCrLf)
    api = dfIds
End Function















