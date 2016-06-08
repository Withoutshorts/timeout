
<% 
Dim objXMLHTTP_calender, objXMLDOM_calender, i_calender, strHTML_calender

Set objXMLDom_calender = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_calender = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_calender.open "GET", "http://localhost/inc/xml/calender_sprog.xml", False
'objXmlHttp_calender.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/calender_sprog.xml", False
objXmlHttp_calender.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/calender_sprog.xml", False
'objXmlHttp_calender.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/calender_sprog.xml", False

objXmlHttp_calender.send


Set objXmlDom_calender = objXmlHttp_calender.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_calender = Nothing



Dim Address_calender, Latitude_calender, Longitude_calender
Dim oNode_calender, oNodes_calender
Dim sXPathQuery_calender

sprog = 1 'DK
if len(trim(session("mid"))) <> 0 then
strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

oRec.open strSQL, oConn, 3
if not oRec.EOF then
sprog = oRec("sprog")
end if
oRec.close
end if

select case sprog
case 1
sXPathQuery_calender = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_calender = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_calender = "//sprog/se"
Session.LCID = 1053
case else
sXPathQuery_calender = "//sprog/dk"
Session.LCID = 1030
end select

Set oNode_calender = objXmlDom_calender.documentElement.selectSingleNode(sXPathQuery_calender)
Address_calender = oNode_calender.Text

Set oNodes_calender = objXmlDom_calender.documentElement.selectNodes(sXPathQuery_calender)

    For Each oNode_calender in oNodes_calender

          calender_txt_107 = oNode_calender.selectSingleNode("txt_107").Text
          
         
          calender_txt_108 = oNode_calender.selectSingleNode("txt_108").Text
          calender_txt_109 = oNode_calender.selectSingleNode("txt_109").Text
          calender_txt_110 = oNode_calender.selectSingleNode("txt_110").Text
          calender_txt_111 = oNode_calender.selectSingleNode("txt_111").Text
          calender_txt_112 = oNode_calender.selectSingleNode("txt_112").Text
          calender_txt_113 = oNode_calender.selectSingleNode("txt_113").Text
          calender_txt_114 = oNode_calender.selectSingleNode("txt_114").Text
        
            calender_txt_115 = oNode_calender.selectSingleNode("txt_115").Text
            calender_txt_116 = oNode_calender.selectSingleNode("txt_116").Text
  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>