Attribute VB_Name = "Module3"
'/*
'* Bits and pieces assembled by Conor (Github: conorg000, roly97)
'* Big thanks to Mark from "Excel Off The Grid" (https://exceloffthegrid.com/about/) for his tutorial on creating drop down menus for Excel add-ins.
'* The code in this add-in is mostly from Mark's tutorial (https://exceloffthegrid.com/inserting-a-dynamic-drop-down-in-ribbon/)
'* The getTimeSeries() function in Module "SDMX" is credited to the Bank Of Italy, as it is an adaptation of their SDMX plug-in (https://github.com/amattioc/SDMX/tree/master/EXCEL)
'*/
'testRibbon is a variable which contains the Ribbon
Option Explicit
Public testRibbon As IRibbonUI
Public listLen As Integer
Public dfIds() As String

Sub testRibbon_onLoad(ByVal ribbon As Office.IRibbonUI)
'This is the Callback for the whole Ribbon.
'When the workbook is opened it sets the testRibbon variable
Dim url As String
url = "https://stats.pacificdata.org/data-nsi/Rest/dataflow/SPC/all?detail=allstubs"

'Connect to API
Dim hReq As New MSXML2.XMLHTTP60
hReq.Open "GET", url, False
hReq.send CreateObject("ADODB.Stream")

Dim resp As String
' Get XML from API
resp = hReq.responseText

'Set namespace for XML
Dim objXML As New MSXML2.DOMDocument60
'MsgBox TypeName(resp)
If objXML.LoadXML(resp) Then
    objXML.setProperty "SelectionNamespaces", "xmlns:message=""http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message"" xmlns:structure=""http://www.sdmx.org/resources/sdmxml/schemas/v2_1/structure"" xmlns:common=""http://www.sdmx.org/resources/sdmxml/schemas/v2_1/common"""
    'MsgBox objXML.SelectSingleNode("message:Structure/message:Header/message:ID").Text
Else
    MsgBox "Error"
End If

' Use this to find dataflows in XML
Dim xPath As String
xPath = "message:Structure/message:Structures/structure:Dataflows/structure:Dataflow"

' nodeList contains all instances of xPath (dataflows)
Dim nodeList As IXMLDOMNodeList
Dim node As IXMLDOMNode
Set nodeList = objXML.SelectNodes(xPath)

' Get length of nodelist
Dim nodeLen As Integer
nodeLen = nodeList.Length
listLen = nodeList.Length
'MsgBox nodeLen
'MsgBox listLen
' Make arrays for IDs and Names of DFs
'Dim dfIds() As String
ReDim dfIds(nodeLen)
Dim dfNames() As String
ReDim dfNames(nodeLen)

Dim X As Integer
X = 0
Dim dfId As String
Dim dfName As String
Dim dict As Collection
Set dict = New Collection
' Assign each dataflow id and name to arrays
For Each node In nodeList
    dfIds(X) = node.Attributes.getNamedItem("id").Text
    dfNames(X) = node.SelectSingleNode("common:Name").Text
    X = X + 1
    'MsgBox "Dataflow: " & node.SelectSingleNode("common:Name").Text & "ID: " & node.Attributes.getNamedItem("id").Text
Next
Set testRibbon = ribbon

End Sub


Public Sub DropDown_getItemCount(control As IRibbonControl, ByRef returnedVal)
'This Callback will create the number of drop-down items as determined
'by the returnedVal value

'returnedVal is set to be equal to public variable listLen
returnedVal = listLen
'MsgBox listLen

End Sub

Public Sub DropDown_getItemID(control As IRibbonControl, index As Integer, ByRef id)
'This Callback will set the id for each item created.
'It provides the index value within the Callback.
'The index is the position within the drop-down list.
'The index can be used to create the id.

'In this example, the id is based on the index number of the item
'first item will be '0', second item will be '1'
id = index

End Sub


Public Sub DropDown_getItemLabel(control As IRibbonControl, index As Integer, _
ByRef returnedVal)
'This Callback will set the displayed label for each item created.
'It provides the index value within the Callback.
'The index is the position within the drop-down list.
'The index can be used to create the id.

'In this example, the label is based on the name of the worksheet
'The index number is used to obtain the sheet name
'Index numbers start at 0, sheet numbers start 1, so +1 to the index.
'returnedVal = ActiveWorkbook.Sheets(index + 1).name

'returnedVal is the name of the dataflow defined by index of the global variable
returnedVal = dfIds(index)


End Sub


Public Sub DropDown_getSelectedItemID(control As IRibbonControl, ByRef id)
'This Callback will change the drop-down to be set to a specific the id.
'This could be used to set a default value or reset the first item in the list

'This example will set the selected item to the id "0"
id = "0"

End Sub

Public Sub DropDown_onAction(control As IRibbonControl, id As String, _
index As Integer)
'This Callback will return the id or index of the item selected.

'This example returns the id in a message box
Call getTimeSeries(dfIds(index))

End Sub

Sub updateRibbon()
'This is a standard procedure, not a Callback. It is triggered by the button.
'It invalidates the Ribbon, which causes it to re-load.

testRibbon.Invalidate

End Sub
