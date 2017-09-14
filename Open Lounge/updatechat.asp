<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Chat Log</title>
</head>

<body>
<%

Dim objFS, objXML, objRoot, objEntries, objEntry, objLine, objChatLine, strMsg, strUsername, strFilePath, Counter

Set objXML = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
strFilePath = Server.MapPath(".")
objXML.Async = False
objXML.ValidateonParse = False
objXML.resolveExternals = False
objXML.preserveWhiteSpace = False
objXML.setProperty "SelectionLanguage","XPath"

objXML.Load strFilePath & "\chatlog.xml"

Set objRoot = objXML.documentElement

Set objEntries = objXML.selectNodes("/CHATLOG/ENTRY")


If objEntries.Length > 50 Then
	For Counter=0 To objEntries.Length - 50
		objXML.documentElement.removeChild objEntries(Counter)
	Next
End If

strUsername = Request.Form("username")
strMsg = Request.Form("mytwopennies")

Set objEntry = objXML.createElement("ENTRY")
Set objLine = objXML.createElement("LINE")

objEntry.SetAttribute "name", strUsername 
objEntry.SetAttribute "date", Now

Set objChatLine = objXML.createCDATASection (strMsg)
'objEntry.SetAttribute "message", strMsg 

objLine.appendChild objChatLine 
objEntry.appendChild objLine 
objXML.documentElement.appendChild objEntry

objXML.Save strFilePath & "\chatlog.xml"

Set objXML = Nothing
'Set objFS = Nothing
Set objEntry = Nothing
Set objLine = Nothing
Set objChatLine = Nothing
Set objEntries = Nothing

Response.Write strMsg 

%>
</body>
</html>
