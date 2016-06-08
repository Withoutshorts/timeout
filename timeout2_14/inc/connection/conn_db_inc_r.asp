
<%
'Response.write session.sessionId

Response.buffer = True 


 
'** Bruges også i 
'** timereg.asp
'** ressourcer.asp 
'** til response.flush

if thisfile <> "printversion.asp" then

if len(request("key")) <> "0" then
session("lto") = request("key")
strLicenskey = request("key")

else
strLicenskey = session("lto") 'Request.Cookies("licenskey")
end if

else

strLicenskey = session("lto") 'Request.Cookies("licenskey")
 
end if


'* finder load tid ***
timeA = now
	

'ODBC
'Dim objXMLHTTP, objXMLDOM, i
'Opretter en instans af Microsoft.XMLHTTP, så det er muligt at få fat på dokumentet
'Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
'Opretter en instans af Microsofts XML-parser, XMLDOM
'Set objXMLDOM = Server.CreateObject("Microsoft.XMLDOM")





'Response.Write " session(lto) :" & session("lto")
'Response.Write " strLicenskey" & strLicenskey



if len(strLicenskey) <> 0 then
	select case strLicenskey
	case "2.052-xxxx-B000" 'demo
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_demo;"
		'strConnThis = "driver={MySQL ODBC 5.00 Driver};server=192.168.23.45; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
		addnewclient = 1
		lto = "demo"
	case "2.152-0416-B001", "2.152-1604-B001" 'outzource
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
		lto = "outz"
        strConnThis = "timeout_intranet64"
        'response.redirect "http://timeout2.outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.152-0416-B001&lto=outz"


	case "2.152-0416-B002" 'sthaus
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sthaus;"
		'lto = "sthaus"
		'*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.052-0423-A005" 'margin
	    'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_margin;"
	    'lto = "margin"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.182-0416-B004" 'Buying
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_buying;"
		'lto = "buying"
		  '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.052-2111-B010" 'Netstrategen
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netstrategen;"
		'lto = "net"
		'*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.052-2711-B013" 'Userneeds
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_userneeds;"
		'lto = "usrneeds"
		'*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "xxx" 'ktv udvikling
    	strConnThis = "mySQL_timeOut_intranet"
		lto = "intranet - local"
	case "2.013-1803-B015" 'ravnit
		'lto = "ravnit"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ravnit;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.053-0705-B016" 'inlead
		'lto = "inlead"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_inlead;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.253-1906-B018" 'titoonic
		lto = "titoonic"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_titoonic;"
	case "2.053-1011-B020" 'Mezzo	
		'lto = "mezzo"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mezzo;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.153-2411-B021" 'Skybrud	
		'lto = "skybrud"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skybrud;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.154-2301-B022" 'Gramtech
		'lto = "gramtech"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gramtech;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.054-0403-B023" 'cybervision
		'lto = "cybervision"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cybervision;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.154-1203-B024" 'ferro
		'lto = "ferro"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ferro;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.054-2405-B025" 'Henton
		'lto = "henton"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_henton;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.994-0206-B026" 'worldiq
		'lto = "worldiq"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_worldiq;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.054-2806-B027" 'Lysta
		lto = "lysta"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lysta;"
	case "2.054-2806-B028", "665f644e43731ff9db3d341da5c827e1" 'Execon
		lto = "execon"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_execon;"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=tid.execon.dk; Port=3306; uid=root;pwd=;database=timeout_execon;"
	case "2.054-2806-B029" 'Inclusive
		'lto = "inclusive"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_inclusive;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.154-2607-B030" 'KringIT
		lto = "kringit"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kringit;"
	case "2.054-2807-B031" 'Kasters
		'lto = "kasters"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kasters;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.154-0710-B032" 'Skousen
		'lto = "skousen"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skousen;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.084-0311-B033" 'Dencker
		lto = "dencker"
        strConnThis = "timeout_dencker64"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"
	case "2.054-0311-B034" 'Firmaservice
		'lto = "firmaservice"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_firmaservice;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.155-0804-B035" 'Webmasters
		'lto = "webmasters"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_webmasters;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.005-1005-B036" 'Proveno
		'lto = "proveno"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_proveno;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.155-2209-B037" 'Øberg Partners
	    'lto = "oberg"
	    'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_oberg;"
	    '*** 29-03-2007 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.155-0411-B038" 'Bj. bro Ingeniørkontor
		lto = "bika"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bika;"
	case "2.455-2811-B039" 'Norma
		'lto = "norma"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_norma;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.155-2012-B040" 'Juul og Stejle
		'lto = "stejle"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_stejle;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-2001-B041" 'Novo Quality Documentation Center
		'lto = "novo_qdc"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_novo_qdc;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-0302-B042" 'Simi
		'lto = "simi"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_simi;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-1302-B043" 'External
		lto = "external"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_external;"
	case "2.606-1502-B044" 'Netkoncept
		'lto = "netkoncept"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netkoncept;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-1303-B045" 'PerspektivaIT
		'lto = "perspektivait"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_perspektivait;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-1307-SB046" 'Mansoft
		lto = "mansoft"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mansoft;"
	case "2.156-1308-SB047" 'Zonerne
		'lto = "zonerne"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_zonerne;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-2408-SB048" 'Böwe
		lto = "bowe"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bowe;"
	case "2.506-1109-LB049" 'Bminds
		'lto = "bminds"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bminds;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-2509-SB050" 'Workpartners
		'lto = "workpartners"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_workpartners;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-2509-SB051" 'GP
		lto = "gp"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gp;"
	case "2.056-0210-SB052" 'Optimizer4u
		'lto = "optimizer4u"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_optimizer4u;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.156-0310-SB053" 'Bolsjebutikken
		lto = "bolsjebutikken"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bolsjebutikken;"
	case "2.156-0512-SB054" 'Maintain
		'lto = "maintain"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_maintain;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.157-0901-SB055" 'Fyn-bo
		lto = "fyn-bo"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_fyn-bo;"
        '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
    case "2.157-0602-SB056" 'Herbo
		'lto = "herbo"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_herbo;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.307-1503-SB057" 'Radius Kommunikation A/S
		'lto = "radius"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_radius;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.307-2005-SB058" 'Mansfield.tv
		'lto = "mansfield"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mansfield;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.057-1008-SB059" 'Morten Hald Mortensen
		lto = "mhm"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mhm;"
	case "2.157-1008-MB060" 'SponsorCar
		'lto = "sponsorcar"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sponsorcar;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.157-1008-MB061" 'A-BETTER-TRAFFIC
		'lto = "abt"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_abt;"
       
	case "2.157-2108-MB062" 'InformationWorker ApS
		lto = "infow"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_infow;"
	case "2.157-2108-MB063" 'InformationWorker ApS
		lto = "infow_demo"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_infowdemo;"
	case "2.307-2709-MB064" 'Accounting ApS
		lto = "acc"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_acc;"
	case "2.157-0410-SB065" 'Elkær
		lto = "elkar"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_elkar;"
	case "2.157-0410-SB066" 'Netfabik
		lto = "netfabrik"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_netfabrik;"
	case "2.307-0810-MB067" 'Ingelise
		lto = "ingelise"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_ingelise;"
	case "2.157-1110-SB068" 'PCM ApS    
		lto = "pcm"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_pcm;"
	case "2.1507-2310-SB069" 'Inger Billund
	    'lto = "ib"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_ib;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.1508-0102-SB070" '2.1507-0811-SB070 UserMinds
	    'lto = "userminds"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_userminds;"
        response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
        '** 1.10.2011 **'
	case "2.1507-0911-SB071" 'Assurator
	    lto = "assurator"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_assurator;"
	case "2.1508-0802-MB072" 'Böwe Systec Nordic SE
	    lto = "bsn_se"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_bsn_se;"
	case "2.1508-0802-MB073" 'Syncronic
	    lto = "syncronic"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_syncronic;"
	case "2.1508-1202-MB074" 'Viatech
	    lto = "viatech"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_viatech;"
	case "2.1508-1902-MB075" 'Kits
	    lto = "kits"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_kits;"
	case "2.1508-2502-MB076" 'JohannesTorpe
	    lto = "jt"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_jt;"
	case "2.1508-1303-MB077" 'HH gruppen
	    lto = "hh"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_hh;"
	case "2.1508-3004-SB078" 'PC-Manden
	    'lto = "pcmanden"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_pcmanden;"
	      '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.1508-1405-SB079" 'Accendo
	    'lto = "accendo"
        'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_accendo;"
	     response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.1508-2505-SB080" 'Dreist Administration
	    'lto = "dreist"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_dreist;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.2008-0306-MB081" 'Abusiness ApS
	    lto = "abu"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_abu;"
	case "2.2008-0306-MB082" 'LN Automatik
	    'lto = "lna"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_lna;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.2008-0306-MB083" 'Upsite
	    'lto = "upsite"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_upsite;"
	    '*** 14-01-2008 ***'
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.2008-1007-MB086" 'Q2con
	    'lto = "Q2con"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_Q2con;"
        response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
        '** 1.1.2012 **'
    case "2.2008-1508-TO087" 'Fyns Energiteknik A/S
	    'lto = "fe"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_fe;"
        response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.0408-2008-TO088" 'TimeOut Management
	    lto = "timeout"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_timeout;"
	case "2.2008-0809-TO089" 'Rosenloeve
	    'lto = "rosenloeve"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_rosenloeve;"
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.2008-1609-TO090" 'JM net
	    lto = "jm"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_jm;"
	case "2.2008-1609-TO091" 'Seepoint / Printcon
	    lto = "printcon"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_printcon;"
	case "2.2008-2511-TO092" 'Weboteket
	    'lto = "weboteket"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_weboteket;"
	    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
	case "2.2008-0212-TO093" 'Madsen + Co
	    lto = "madsen"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_madsen;"
	case "2.2008-0712-TO094" 'Råhedestål 
	    lto = "raahede"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_raahede;"
	case "2.2008-1012-TO095" 'OptionOne 
	    lto = "optionone"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_optionone;"
	case "2.3508-1812-TO096" 'FK ejendomme 
	    lto = "fk"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_fk;"
	case "2.2008-2112-TO097" 'Immenso 
	    lto = "immenso"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_immenso;"
	case "2.2009-0105-TO098" 'Essens 
	    lto = "essens"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_essens;"
	case "2.2009-1201-TO099" 'Svendborg Architects
	    lto = "svea"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_svea;"
	case "2.2009-0105-TO100" 'C.S.T
	    lto = "cst"
		'**CST server***'
        'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=93.161.131.214; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cst;"
	    strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_cst;"
        '93.161.131.214
    case "2.2009-0105-TO10x" 'C.S.T TEST
	    lto = "cst_test"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=93.161.131.214; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cst;"
	    strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_cst_BAK_20130301;"
        '93.161.131.214
	case "2.7509-1202-TO101" 'ETS-Track
	    lto = "ets-track"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_ets-track;"
	case "2.0409-1202-TO102" 'Berens
	    lto = "berens"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_berens;"
	case "2.2009-1003-TO103" 'Wowern Reklame
	    lto = "wowern"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_wowern;"
	case "2.2009-1003-TO104" 'Flere-Klik-Flere-Kunder
	    lto = "fkfk"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_fkfk;"
	case "2.2009-1805-TO105" 'skabelon Design
	    lto = "sd"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_sd;"
	case "2.2009-2906-TO106" 'Business Data
	    lto = "bd"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_bd;"
    case "2.2009-2907-TO107" '2L developemnt AS
	    lto = "2l"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_2l;"
    case "2.2009-1608-TO108" 'Symbion
	    lto = "symbion"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_symbion;"
    case "2.2009-2308-TO109" 'Tool Test ApS
	    lto = "tooltest"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_tooltest;"
    case "2.2009-2708-TO110" 'Arti Farti
	    lto = "af"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_af;"
    case "2.2010-2001-TO111" 'CasperMartens ApS
	    lto = "cma"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_cma;"
    case "2.2010-2101-TO112" 'Revisor Lars Jeppesen
	    lto = "jep"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_jep;"
    case "2.2010-1202-TO113" 'Rage Technologies
	    lto = "rage"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_rage;"
    case "2.2010-1402-TO114" 'Energi Horsens
	    lto = "enho"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_enho;"
    case "2.2010-1802-TO115" 'Internet Service
	    lto = "is"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_is;"
    case "2.2010-2004-TO116" 'Epinion
	    lto = "epi"
        strConnThis = "timeout_epi64"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi;"
    case "2.2010-2605-TO117" 'JT-Teknik
	    lto = "jttek"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_jttek;"
        strConnThis = "timeout_jttek64"
    case "2.2010-1004-TO118" 'WWF Verdensnaturfonden
	    lto = "wwf"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_wwf;"
        strConnThis = "timeout_wwf64"
        'response.redirect "http://timeout2.outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2010-1004-TO118&lto=wwf"
    case "2.2010-0311-TO119" 'Mærsk Insttuttet
	    lto = "mmmi"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_mmmi;"
     case "2.2010-0311-TO119BAK" 'Mærsk Insttuttet
	    lto = "mmmi"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_mmmi_BAK;"
    case "2.2010-2004-TO116_c" 'Epinion CATI
	    lto = "epi_cati"
        strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi_catitest;"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi_catitest_dw;"
    case "2.2011-0125-TO120" 'CommuteMedia
	    lto = "commutemedia"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_commutemedia;"
	case "2.2011-0201-TO121" 'Testhuset
	    lto = "testhuset"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_testhuset;"
	case "2.2011-0210-TO122" 'Berlingske Media
	    lto = "bm"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_bm;"
	case "2.2011-0215-TO123" 'QWERT
	    lto = "qwert"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_qwert;"
	case "2.2011-0308-TO124" 'Magnus Informatik A/S
	    lto = "mi"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_mi;"
	case "2.2011-0314-TO125" 'Frederiksberg Kommune By og Miljø
	    'lto = "fk_bpm"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_fk_bpm;"

        '** midlertidig til backup tjk af FK
        lto = "fk"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_fk_BAK;"


	case "2.2011-0315-TO126" 'Grundfos
	    lto = "grundfos"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_grundfos;"
	case "2.2011-0410-TO127" 'SDU Unik
	    lto = "unik"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_unik;"
    case "2.2011-0512-TO128" 'Synergi1
	    lto = "synergi1"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_synergi1;"
        strConnThis = "timeout_synergi164"
    case "2.2011-1024-TO129" 'BST-Nord
	    lto = "bst"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_bst;"
    case "2.2011-1115-TO130" 'Leonhard Weiss
	    lto = "lw"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_lw;"
    case "2.2012-0214-TO131" 'Verdensnaturfonden SANDKASSE
	    lto = "wwf2"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_wwf2;"

    'case "2.2012-0214-TO131_ORG" 'Verdensnaturfonden ORG
	'    lto = "wwf"
    'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_wwf_ORG_20121112;"

    case "2.2012-0214-TO132" 'Københavns Kommune Projekt & Bygherre
	    lto = "kejd_pb"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_kejd_pb;"
    case "2.2012-0214-TO133" 'Københavns Kommune Projekt & Bygherre
	    lto = "kejd_pb2"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_kejd_pb2;"
    case "2.2012-0625-TO134" 'Hvidovre Kommune 
	    lto = "hvk_bbb"
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_hvk_bbb;"
    

    case "2.2012-1010-TO135" 'Epinion Baromater STABÆK
	'lto = "epi_sta"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi_sta;"
    lto = "epi_no"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi_no;"
    strConnThis = "timeout_epi_no64"

     case "2.2012-1010-TO135b" 'Epinion Baromater STAVANGER BAK
	lto = "epi_sta"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi_sta;"'_bak
        
    case "2.2012-1012-TO136" 'Epinion AS OSLO / Epinion Baromater STAVANGER = EPI_NO
	'lto = "epi_no"
    lto = "epi_no"
    strConnThis = "timeout_epi_no64"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi_no;"

    case "2.2012-1013-TO137" 'Kestner
	lto = "kestner"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_kestner;"
        
    case "2.2012-1014-TO138" 'Ordkontoret 
	lto = "ord"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_ord;"

    case "2.2012-1015-TO139" 'Rekomenderet 
	lto = "rek"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_rek;"

    case "2.2012-1016-TO140" 'Århus Universits Hospital 
	lto = "auh"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_auh;"

    case "2.2013-0108-TO141" 'Naturgas Fyn 
	lto = "ngf"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_ngf;"

    case "2.2013-0109-TO142" 'Epi AB 
	lto = "epi_ab"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi_ab;"

    case "2.2013-0125-TO143" 'LAWABA 
	lto = "lawaba"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_lawaba;"

    case "2.2013-0625-TO144" 'IntegrationsNet 
	lto = "intg"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_intg;"


    case "2.2013-0812-TO145" 'Hidalgo 
	lto = "hidalgo"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_hidalgo;"


    case "2.2013-0912-TO146" 'Biofac Orthena
	lto = "biofac"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_biofac;"

    case "2.2013-1002-TO147" 'Økologisk landsforening
	lto = "oko"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_oko;"
    strConnThis = "timeout_oko64"
    'response.redirect "http://timeout2.outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2013-1002-TO147&lto=oko"


    case "2.2014-0130-TO147" 'Hestia Ejendomme
	lto = "hestia"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_hestia;"


     case "2.2014-0331-TO148" 'Epinion UK
	lto = "epi_uk"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi_uk;"


    case "2.2014-0501-TO149" 'NorthTex
	lto = "nt"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_nt;"
    strConnThis = "timeout_nt64"

    case "2.2014-0505-TO150" 'Øens Tag og VVS entreprise
	lto = "otvvs"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_otvvs;"

    case "2.2014-0702-TO151" 'TEC
	lto = "tec"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_tec;"
    strConnThis = "timeout_tec64"

    case "2.2014-0911-TO152" 'Zylinc
	lto = "zylinc"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_zylinc;"


    case "2.2014-1007-TO155" 'SDU Tek
	lto = "sdutek"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_sdutek;"


    case "2.2014-0930-TO153" 'Glad
	lto = "glad"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_glad;"


    case "2.2014-1001-TO154" 'Lyng
	lto = "lyng"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_lyng;"


    case "2.2014-1113-TO156" 'ESN
	lto = "esn"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_esn;"
    strConnThis = "timeout_esn64"


    case "2.2014-1219-TO157" 'Nonstop
	lto = "nonstop"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_nonstop;"

     case "2.2015-0203-TO158" 'DanWatch
	lto = "danwatch"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_danwatch;"


    case "2.2015-0413-TO159" 'Micmatic
	lto = "micmatic"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_micmatic;"


     case "2.2015-0414-TO160" 'Aalund
	lto = "aalund"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_aalund;"


    case "2.2015-0806-TO161" 'Akelius
	lto = "akelius"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_akelius;"

    case "892015-0829-TO162" 'Ascendis
	lto = "ascendis"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_ascendis;"

    case "3W2015-1102-TO163" 'Wilke
	lto = "wilke"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_wilke;"
    strConnThis = "timeout_wilke64"

    case "2Z2015-1125-TO164" 'Adra
	lto = "adra"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_adra;"

    case "9K2015-1202-TO165" 'CC Accounting
	lto = "cc"
	strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_cc;"


    case "9K2016-1203-TO166" 'KRJ Consulting
	lto = "krj"
	strConnThis = "timeout_krj64"

    case "9K2016-1403-TO167" 'Cisu
	lto = "cisu"
	strConnThis = "timeout_cisu64"

        
    
    'ODBC 3.51 Driver
	
	case else 'ktv udvikling (hvis der ikke er en cookie)
		
         'strConnThis = "timeout_intranet_351_30_64"
        'strConnThis = "timeout_intranet32" 
        lto = "intranet - local"

        'strConnThis = "mySQL_timeOut_intranet"
        'strConnThis = "timeout_wwf"
		'lto = "intranet - local"
        'lto = "tec"
        'strConnThis = "timeout_intranet64"
        strConnThis = "timeout_nt"
        'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet;"
	    'response.write strConnThis
        'response.flush
	
       'lto = "outz"
	   'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=192.168.1.35; Port=3306; uid=root;pwd=;database=timeout_intranet;"
	
    
    	
        'lto = "kejd_pb"
        'Ny IP Pr 5.4.2014
        '195.189.130.210
		'NY IP: 62.182.173.226
		''lto = "epi"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"
	    
	    'lto = "kejd_pb"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kejd_pb;"
	
		
		
	end select 



   


else
	if instr(request.servervariables("HTTP_HOST"), "localhost") <> 0 then
	
	'Zepto
	'strConnThis = "mySQL_timeOut_intranet"
    'strConnThis = "timeout_intranet_351_30_64"
    'strConnThis = "timeout_intranet32"
    lto = "intranet - local"
    'lto = "tec"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet;"
	' strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet; OPTION=32"
	   'response.write strConnThis
       ' response.flush

    'strConnThis = "timeout_intranet64"
    strConnThis = "timeout_nt"


     
    ' lto = "dencker"
    ' strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"
	

    'lto = "kejd_pb"

    'NY IP: 62.182.173.226
    'lto = "cst"
    'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=81.19.249.35; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cst;"
	 
    'lto = "wwf"
    'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wwf_BAK;"

     'lto = "dencker"
     'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"

     'lto = "tec"
	 'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_tec;"
	   

	else
	
    call showError(5)
	response.end
	
	'Response.redirect "../login_fejlet.asp"
	
	'Response.write "Sessionen er udløbet"
	'Response.end
	end if
end if

'objXMLHTTP.Send

'Lægger indholdet af dokumentet over i vores XML-parser objekt
'Set objXMLDOM = objXMLHTTP.ResponseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("modul")

'Henter indholdet af alle tags med navnet 'url'
'Set objDBConn = objXMLDOM.getElementsByTagName("dbconn")

strConn = strConnThis	'objDBConn(0).text  'Connection


strA1 = "-" 	'objModuler(0).text   'Timeregistrering
strA2 = "-"		'objModuler(1).text   'Medarbejder log



'*** ObjModuler(2).text   'Skift medarbejder 
select case lto
case "kringit", "outz", "worldiq"', "intranet"
strA3 = "on"
case else
strA3 = "off"
end select 		


strA4 = "on"		'objModuler(3).text   'Se timereg for andre medarb.
strB1 = "-"   	'Joboversigt og Aktiviter. Mulighed for at lukke akt. 
strB2 = "-"		'objModuler(5).text   'Projektgrupper
strB3 = "-"		'objModuler(6).text   'Stamaktiviteter
strB4 = "-"		'objModuler(7).text   'Jobplanner
strB5 = "-"		'objModuler(8).text   'Km regnskab
strB6 = "-"		'objModuler(9).text   'Udgifter paa job
strB7 = "-"		'objModuler(10).text   'Ikke planlagt
strB8 = "-"		'objModuler(11).text   'Upload af dokumenter
strB9 = "-"		'objModuler(12).text   'Incident tool
strB10 = "-"	'objModuler(13).text	'Aktivitetsskabeloner

strC1 = "-" 	'objModuler(14).text  'Kunder + filter og sog paa kunde
strC2 = "on" 	'objModuler(15).text  'Kunde login
strC3 = "-"		'objModuler(16).text   'Upload logoer paa kunder

strD1 = "-" 	'objModuler(17).text  'Medarbejdere
strD2 = "-" 	'objModuler(18).text  'Brugergrupper
strD3 = "-"		'objModuler(19).text  'Medarbejdertyper

strE1 = "-"		'objModuler(20).text  'Statistik og Fakturering + filter
strE2 = "-"		'objModuler(21).text  'Joblog_z_b, joblog, top 5 job
strE3 = "-"		'objModuler(22).text  'Aars omsaetning
strE4 = "-"		'objModuler(23).text  'Export
strE5 = "on"		'objModuler(24).text	 'Faktura oversigt

strH1 = "-"		'objModuler(25).text  'Kalender
strH2 = "-"		'objModuler(26).text  'Firmaer
strH3 = "-"		'objModuler(27).text	 'Aktions historik + sogefunktion
strH4 = "-"		'objModuler(28).text  'Status
strH5 = "-"		'objModuler(29).text  'Emner
strH6 = "-"		'objModuler(30).text  'Kontaktform
strH7 = "-"		'objModuler(31).text  'Outlook integration

'Set objXMLHTTP = Nothing
'Set objXMLDOM = Nothing

'if lto = "outz" then
'Response.Write "lto" & lto & "<br>"
'Response.Write "conn:" & strConn
'Response.flush
'end if

if len(strConn) <> 0 then
	strConnect_DBConn = strConn
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRec = Server.CreateObject ("ADODB.Recordset")
	Set oRec2 = Server.CreateObject ("ADODB.Recordset")
	Set oRec3 = Server.CreateObject ("ADODB.Recordset")
	Set oRec4 = Server.CreateObject ("ADODB.Recordset")
	Set oRec5 = Server.CreateObject ("ADODB.Recordset")
	Set oRec6 = Server.CreateObject ("ADODB.Recordset")
	Set oRec7 = Server.CreateObject ("ADODB.Recordset")
    Set oRec8 = Server.CreateObject ("ADODB.Recordset")
    Set oCmd = Server.CreateObject ("ADODB.Command")
	
	oConn.open strConnect_DBConn
	
	
	sub closeDB
	'now close and clean up
			oRec.Close
			oConn.close
			
			Set oRec = Nothing
			set oConn = Nothing
	
	end sub




                if thisfile <> "printversion.asp" then

                 '** Hvis der er skiftet version ****
                '** Finder bruger i ny version ****'
                if request("eksterntlnk") = "aaQWEIOC345345DFNEfjsdf7890sdfv" then
   

                        if len(trim(request("email"))) <> 0 then
                        emailThis = request("email")
                        else
                        emailThis = "99999999"
                        end if


                        'session.abandon
                        'response.flush

                        session("user") = ""
			            session("login") = ""
			            session("mid") = ""
			            session("rettigheder") = ""


                        newUSerId = 0
                        strSQLnewuser = "SELECT mid, mnavn, m.brugergruppe, b.rettigheder FROM medarbejdere AS m "_
                        &" LEFT JOIN brugergrupper AS b ON (b.id = m.brugergruppe) WHERE email = '"& emailThis & "'"
    
                        'Response.write strSQLnewuser
                        'Response.flush

                        oRec4.open strSQLnewuser, oConn, 3
                        if not oRec4.EOF then


        


                        newUSerId = oRec4("mid")

                        session("user") = trim(oRec4("mnavn"))
			            session("login") = newUSerId
			            session("mid") = newUSerId
			            session("rettigheder") = oRec4("rettigheder")
            

                        end if 
                        oRec4.close

                usemrn = newUSerId
   

                if newUSerId <> 0 then


                Response.write "<div style=""top:40px; left:40px,"">Du er godkendt til at skifte version, <a href=""timereg_akt_2006.asp"">klik her</a> for at komme videre...</div>"

                Response.end


                else

                'Response.write "her"
                'Response.end
                'Response.redirect "../login_fejlet.asp"
    
                call showError(5)
                response.end


                end if
          	
    

                end if
                end if 'printversion
   

end if

%>
