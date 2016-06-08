<%
Response.buffer = True '** bruges også i timereg til flush
if len(request("key")) <> "0" then
'** Indsætter cookie **
Response.Cookies("licenskey") = request("key")
Response.Cookies("licenskey").Expires = date + 360
end if


'* finder load tid ***
timeA = now
	

'ODBC
'Dim objXMLHTTP, objXMLDOM, i
'Opretter en instans af Microsoft.XMLHTTP, så det er muligt at få fat på dokumentet
Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
'Opretter en instans af Microsofts XML-parser, XMLDOM
Set objXMLDOM = Server.CreateObject("Microsoft.XMLDOM")



if len(request("key")) = 15 then 
strLicenskey = request("key")
else
'Henter Cookie
strLicenskey = Request.Cookies("licenskey")
end if



if len(strLicenskey) <> 0 then
	select case strLicenskey
	case "2.052-xxxx-B000" 'demo
		Call objXMLHTTP.Open("GET", "http://outzource.dk/demo/conf_TOser_2_demo.xml", False)
		addnewclient = 1
		lto = "demo"
	case "2.152-0416-B001" 'outzource
		Call objXMLHTTP.Open("GET", "http://outzource.dk/outz/conf_TOser_2_outz.xml", False)
		'Call objXMLHTTP.Open("GET", "http://www.outzource.dk/outz/conf2_outz.xml", False)
		lto = "outz"
	case "2.152-0416-B002" 'sthaus
		Call objXMLHTTP.Open("GET", "http://outzource.dk/sthaus/conf_TOser_2_sthaus.xml", False)
		lto = "sthaus"
	case "2.052-0423-A005" 'margin
		Call objXMLHTTP.Open("GET", "http://outzource.dk/margin/conf_TOser_2_margin.xml", False)
		lto = "margin"
	case "2.182-0416-B004" 'Buying
		Call objXMLHTTP.Open("GET", "http://outzource.dk/buy/conf_TOser_2_buy.xml", False)
		lto = "buying"
	case "2.052-2111-B010" 'Netstrategen
		Call objXMLHTTP.Open("GET", "http://outzource.dk/net/conf_TOser_2_net.xml", False)
		lto = "net"
	case "2.052-0705-B008" 'Revisor
		Call objXMLHTTP.Open("GET", "http://outzource.dk/revisor/conf_TOser_2_rev.xml", False)
		lto = "rev"
	case "2.152-2711-B012" 'Comone
		Call objXMLHTTP.Open("GET", "http://outzource.dk/comone/conf_TOser_2_comone.xml", False)
		lto = "comone" 
	case "2.052-2711-B013" 'Userneeds
		Call objXMLHTTP.Open("GET", "http://outzource.dk/usrneeds/conf_TOser_2_usrneeds.xml", False)
		lto = "usrneeds"
	case "2.052-0416-B006" 'WPS
		Call objXMLHTTP.Open("GET", "http://outzource.dk/wps/conf_TOser_2_wps.xml", False)
		lto = "wps"
	case "2.993-2701-B014" 'Voice
		Call objXMLHTTP.Open("GET", "http://outzource.dk/voice/conf_TOser_2_voice.xml", False)
		lto = "voice"
	case "xxx" 'ktv udvikling
		if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
		Call objXMLHTTP.Open("GET", "http://outzource.dk/outz/conf2_ktvudv.xml", False)
		else
		Call objXMLHTTP.Open("GET", "http://localhost/timeout_xp/inc/xml/conf3_ktvudv.xml", False)
		end if
		lto = "intranet"
	case "2.013-1803-B015" 'ravnit
		lto = "ravnit"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/ravnit/conf_TOser_3_ravnit.xml", False)
	case "2.053-0705-B016" 'inlead
		lto = "inlead"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/inlead/conf_TOser_2_inlead.xml", False)
	case "2.053-2705-B017" 'itsupportpartner
		lto = "itsuppart"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/itsuppart/conf_TOser_2_itsuppart.xml", False)
	case "2.253-1906-B018" 'titoonic
		lto = "titoonic"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/titoonic/conf_TOser_2_titoonic.xml", False)
	case "2.153-1909-B019" 'revitax	
		lto = "revitax"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/revitax/conf_TOser_2_revitax.xml", False)
	case "2.053-1011-B020" 'Mezzo	
		lto = "mezzo"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/mezzo/conf_TOser_2_mezzo.xml", False)
	case "2.153-2411-B021" 'Skybrud	
		lto = "skybrud"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/skybrud/conf_TOser_2_skybrud.xml", False)	
	case "2.154-2301-B022" 'Gramtech
		lto = "gramtech"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/gramtech/conf_TOser_2_gramtech.xml", False)	
	case "2.054-0403-B023" 'cybervision
		lto = "cybervision"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/cybervision/conf_TOser_2_cybervision.xml", False)	
	case "2.154-1203-B024" 'ferro
		lto = "ferro"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/ferro/conf_TOser_2_ferro.xml", False)	
	case "2.054-2405-B025" 'Henton
		lto = "henton"
		Call objXMLHTTP.Open("GET", "http://outzource.dk/henton/conf_TOser_2_henton.xml", False)	
	
	
	case else 'ktv udvikling (hvis der ikke er en cookie)
		Call objXMLHTTP.Open("GET", "http://localhost/timeout_xp/inc/xml/conf3_ktvudv.xml", False)
		lto = "intranet"
	end select 
else
	if instr(request.servervariables("HTTP_HOST"), "localhost") <> 0 then
	Call objXMLHTTP.Open("GET", "http://localhost/timeout_xp/inc/xml/conf3_ktvudv.xml", False)
	lto = "intranet"
	else
	Response.redirect "../login_fejlet.asp"
	end if
end if

objXMLHTTP.Send

'Lægger indholdet af dokumentet over i vores XML-parser objekt
Set objXMLDOM = objXMLHTTP.ResponseXML

'Henter indholdet af alle tags med navnet 'titel'
Set objModuler = objXMLDOM.getElementsByTagName("modul")

'Henter indholdet af alle tags med navnet 'url'
Set objDBConn = objXMLDOM.getElementsByTagName("dbconn")

strConn = objDBConn(0).text  'Connection
strA1 = objModuler(0).text   'Timeregistrering
strA2 = objModuler(1).text   'Medarbejder log
strA3 = objModuler(2).text   'Skift medarbejder 
strA4 = objModuler(3).text   'Se timereg for andre medarb.

strB1 = objModuler(4).text   'Joboversigt og Aktiviter. Mulighed for at lukke akt. 
strB2 = objModuler(5).text   'Projektgrupper
strB3 = objModuler(6).text   'Stamaktiviteter
strB4 = objModuler(7).text   'Jobplanner
strB5 = objModuler(8).text   'Km regnskab
strB6 = objModuler(9).text   'Udgifter paa job
strB7 = objModuler(10).text   'Ikke planlagt
strB8 = objModuler(11).text   'Upload af dokumenter
strB9 = objModuler(12).text   'Incident tool
strB10 = objModuler(13).text	'Aktivitetsskabeloner

strC1 = objModuler(14).text  'Kunder + filter og sog paa kunde
strC2 = objModuler(15).text  'Kunde login
strC3 = objModuler(16).text	 'Upload logoer paa kunder

strD1 = objModuler(17).text  'Medarbejdere
strD2 = objModuler(18).text  'Brugergrupper
strD3 = objModuler(19).text  'Medarbejdertyper

strE1 = objModuler(20).text  'Statistik og Fakturering + filter
strE2 = objModuler(21).text  'Joblog_z_b, joblog, top 5 job
strE3 = objModuler(22).text	 'Aars omsaetning
strE4 = objModuler(23).text  'Export
strE5 = objModuler(24).text	 'Faktura oversigt

strH1 = objModuler(25).text  'Kalender
strH2 = objModuler(26).text  'Firmaer
strH3 = objModuler(27).text	 'Aktions historik + sogefunktion
strH4 = objModuler(28).text  'Status
strH5 = objModuler(29).text  'Emner
strH6 = objModuler(30).text  'Kontaktform
strH7 = objModuler(31).text  'Outlook integration

Set objXMLHTTP = Nothing
Set objXMLDOM = Nothing


if len(strConn) <> 0 then
	strConnect = strConn
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRec = Server.CreateObject ("ADODB.Recordset")
	Set oRec2 = Server.CreateObject ("ADODB.Recordset")
	Set oRec3 = Server.CreateObject ("ADODB.Recordset")
	Set oCmd = Server.CreateObject ("ADODB.Command")
	
	oConn.open strConnect
	
	
	sub closeDB
	'now close and clean up
			oRec.Close
			oConn.close
			
			Set oRec = Nothing
			set oConn = Nothing
	
	end sub

end if

%>
