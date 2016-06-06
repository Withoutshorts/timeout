
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<% 
session.lcid = 1030

function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
end function

if len(Request.Form("Timerkom")) < 255 then 
	varTimerkomma = SQLBless(request.form("Timer"))
	strDato = request.form("aar")&"/"& request.form("mrd")&"/"& request.form("dag")
	dagsdato = year(now) &"/"& month(now) &"/"& day(now)
	str_dagsdato = formatdatetime(dagsdato, 1)
	intOff = request("FM_off")
	timerKom = replace(Request("Timerkom"), "'","''")
	
	oRec2.Open "SELECT * FROM timer WHERE Tmnavn = '"& session("user") &"' AND Tjobnr = '"&request("FM_jobnr")&"' AND TAktivitetId = "&request("FM_AktId")&" AND Tdato = '" & strDato &"' AND Tid <> " & Request.Form("id") &"", oConn, 3
	
	if oRec2.EOF then
	oRec.Open "UPDATE timer" &" SET editor = '"& session("user") &"', Tdato = '" & strDato &"', Timer = "& varTimerkomma &", Timerkom = '" & timerKom &"', offentlig = "& intOff &" WHERE Tid = " & Request.Form("id") & "", oConn, 3
	
		intJoblog = request("joblog")
		menu = request("menu")
		if (intJoblog = "1" OR intJoblog = "2") AND menu = "stat" then
		
		'strReqMrd = Request("mrd")
		medarb = request("medarb")
		intjobnr = request("Jobnr")
		eks = request("eks")
		'strReqAar = Request("year")
		lastFakdag = request("lastFakdag")
		selmedarb = request("selmedarb")
		selaktid = request("selaktid")
		strMrd = request("FM_start_mrd")
		strDag = request("FM_start_dag")
		strAar = right(request("FM_start_aar"),2) 
		strDag_slut = request("FM_slut_dag")
		strMrd_slut = request("FM_slut_mrd")
		strAar_slut = right(request("FM_slut_aar"),2)
		
			if intJoblog = "1" then
			response.redirect "joblog.asp?menu=stat&jobnr="&intJobnr&"&eks="&request("eks")&"&lastFakdag="&lastFakdag&"&selmedarb="&selmedarb&"&selaktid="&selaktid&"&FM_job="&request("FM_job")&"&FM_medarb="&request("FM_medarb")&"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""
			else
			response.redirect "joblog_korsel.asp?menu=stat&jobnr="&intJobnr&"&eks="&request("eks")&"&lastFakdag="&lastFakdag&"&selmedarb="&selmedarb&"&selaktid="&selaktid&"&FM_job="&request("FM_job")&"&FM_medarb="&request("FM_medarb")&"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""
			end if
		else
			if intJoblog = "1" then
			response.redirect "joblog.asp?menu=timereg&FM_medarb="&session("mid")&"&FM_job=0&selmedarb="&session("mid")&"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""
			else
			response.redirect "joblog_korsel.asp?menu=timereg&FM_medarb="&session("mid")&"&FM_job=0&selmedarb="&session("mid")&"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""
			end if
		end if
	
	else
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->	
	<%
	errortype = 19
	call showError(errortype)
	%>
	<!--#include file="../inc/regular/footer_inc.asp"-->
	<%
	end if
else
%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->	
	<%
	errortype = 34
	call showError(errortype)
	%>
	<!--#include file="../inc/regular/footer_inc.asp"-->
	<%
end if
%>

