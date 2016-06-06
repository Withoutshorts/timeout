<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>XML</title>
</head>

<body>
<%
set objXML = Server.CreateObject("msxml.domdocument")
objXML.async = false
objXML.load("confoutz.xml")


Response.write objXML.documentElement.firstChild.nodeName


%>


</body>
</html>
