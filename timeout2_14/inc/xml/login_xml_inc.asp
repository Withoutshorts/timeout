
<% 

Dim objXMLHTTP_login, objXMLDOM_login, i_login, strHTML_login

Set objXMLDom = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp.open "GET", "http://localhost/inc/xml/login_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/xml/login_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/login_sprog.xml", False
objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/login_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver4_22/inc/xml/login_sprog.xml", False

objXmlHttp.send


Set objXmlDom = objXmlHttp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp = Nothing



Dim Address_login, Latitude_login, Longitude_login
Dim oNode_login, oNodes_login
Dim sXPathQuery_login


sprog = 1 'DK
select case lto
case "epi_uk", "intranet - local"
sprog = 2
end select





select case sprog
case 1
sXPathQuery_login = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_login = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_login = "//sprog/se"
Session.LCID = 1053
case else
sXPathQuery_login = "//sprog/dk"
Session.LCID = 1030
end select

'Response.Write "Session.LCID" &  Session.LCID

Set oNode_login = objXmlDom.documentElement.selectSingleNode(sXPathQuery_login)
Address_login = oNode_login.Text

Set oNodes_login = objXmlDom.documentElement.selectNodes(sXPathQuery_login)

For Each oNode_login in oNodes_login

     login_txt_001 = oNode_login.selectSingleNode("txt_1").Text
     login_txt_002 = oNode_login.selectSingleNode("txt_2").Text
     login_txt_003 = oNode_login.selectSingleNode("txt_3").Text
     login_txt_003 = oNode_login.selectSingleNode("txt_3").Text
    login_txt_004 = oNode_login.selectSingleNode("txt_4").Text
    login_txt_005 = oNode_login.selectSingleNode("txt_5").Text

    login_txt_006 = oNode_login.selectSingleNode("txt_6").Text
    login_txt_007 = oNode_login.selectSingleNode("txt_7").Text
    login_txt_008 = oNode_login.selectSingleNode("txt_8").Text
    login_txt_009 = oNode_login.selectSingleNode("txt_9").Text

    login_txt_010 = oNode_login.selectSingleNode("txt_10").Text
    login_txt_011 = oNode_login.selectSingleNode("txt_11").Text
    login_txt_012 = oNode_login.selectSingleNode("txt_12").Text
    login_txt_013 = oNode_login.selectSingleNode("txt_13").Text
    login_txt_014 = oNode_login.selectSingleNode("txt_14").Text
    login_txt_015 = oNode_login.selectSingleNode("txt_15").Text
    login_txt_016 = oNode_login.selectSingleNode("txt_16").Text
    login_txt_017 = oNode_login.selectSingleNode("txt_17").Text
    login_txt_018 = oNode_login.selectSingleNode("txt_18").Text
    login_txt_019 = oNode_login.selectSingleNode("txt_19").Text

    login_txt_020 = oNode_login.selectSingleNode("txt_20").Text
    login_txt_021 = oNode_login.selectSingleNode("txt_21").Text
    login_txt_022 = oNode_login.selectSingleNode("txt_22").Text
    login_txt_023 = oNode_login.selectSingleNode("txt_23").Text
    login_txt_024 = oNode_login.selectSingleNode("txt_24").Text
    login_txt_025 = oNode_login.selectSingleNode("txt_25").Text
    login_txt_026 = oNode_login.selectSingleNode("txt_26").Text
    login_txt_027 = oNode_login.selectSingleNode("txt_27").Text
    login_txt_028 = oNode_login.selectSingleNode("txt_28").Text
    login_txt_029 = oNode_login.selectSingleNode("txt_29").Text

    login_txt_030 = oNode_login.selectSingleNode("txt_30").Text
    login_txt_031 = oNode_login.selectSingleNode("txt_31").Text
    login_txt_032 = oNode_login.selectSingleNode("txt_32").Text
    login_txt_033 = oNode_login.selectSingleNode("txt_33").Text
    login_txt_034 = oNode_login.selectSingleNode("txt_34").Text
    
next




%>