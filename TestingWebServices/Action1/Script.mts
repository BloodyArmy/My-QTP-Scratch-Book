url = "http://maps.googleapis.com/maps/api/distancematrix/xml?" & _
		"origins=Atlanta+GA+USA&" & _
		"destinations=Dallas+TX+USA&" & _
		"units=imperial&" & _
		"sensor=false"

Set oReq = CreateObject("Microsoft.XMLHTTP")

oReq.Open "POST", url, FALSE
oReq.Send

'Response
responseText = oReq.responseText

Print responseText


sNodeDistance = "row/element/distance/text"
sNodeDuration = "row/element/duration/text"

Set oParser = XMLUtil.CreateXML()

'Note: responseText comes from XMLHTTP Request in the previous section
oParser.Load responseText 

'Using the XPath above to retrieve distance
nodeDistance = oParser.GetRootElement.GetValueByXPath(sNodeDistance)
 
 
nodeDuration = oParser.GetRootElement.GetValueByXPath(sNodeDuration)
 
Print nodeDistance
Print nodeDuration







sNodeDistance = "row/element/distance/text"
sNodeDuration = "row/element/duration/text"
 
Set oParser = CreateObject("Microsoft.XMLDOM")
 
'Note: responseText comes from XMLHTTPRequest in the previous section
oParser.loadXML(responseText)
 
'Distance node: comes from XPath above
Set nodeDistance = oParser.documentElement.selectSingleNode(sNodeDistance)
 
'Duration node: comes from XPath above
Set nodeDuration = oParser.documentElement.selectSingleNode(sNodeDuration)
 
Print nodeDistance.text
Print nodeDuration.text