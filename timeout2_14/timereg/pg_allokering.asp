<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	func = request("func")
	
	
	
	'*** Opdaterer job liste ****
    if func = "opdaterjobliste" then


    ujid = split(request("FM_jobid"), ",")
    urisiko = split(request("FM_risiko"), ",")

    for u = 0 to UBOUND(ujid)
	    'Response.write uudgifter(u) & "<br>"
    	
	    'uforvslutdato = replace(uforvslutaar(u) & "/" & uforvslutmd(u) & "/" & uforvslutdag(u), ",", "")
    	
	    strSQLjobupd = "UPDATE job SET risiko = "& urisiko(u) &" WHERE id = " & ujid(u)
	    oConn.execute(strSQLjobupd)
    	
	    'Response.write strSQLjobupd & "<br>"
    	
    next
	
	end if
	
	
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	orderkri = " jobslutdato "
	
	if len(request("FM_risiko")) <> 0 then
	risk = request("FM_risiko")
	else
	risk = 0
	end if
	
	
	if request("print") = "j" then%>	
		<html>
		<head>
		<title>timeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
		</head>
		<body topmargin="0" leftmargin="0" class="regular">
	<%end if
	
	if request("print") = "j" then
	leftPos = 10
	topPos = 60
	bgthis = "#5582d2"
	else
	leftPos = 20
	topPos = 122
	bgthis = "#5582d2"
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(2)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call webbliktopmenu()
	%>
	</div>
	
	

		
	<%
	end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<%if request("print") = "j" then%>	
	<table cellspacing="0" cellpadding="0" border="0" width="880">
		<tr>
			<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
			<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
		</tr>
		</table>
	<%end if%>
	
	
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	<h3><img src="../ill/ac0047-24.gif" width="24" height="24" alt="" border="0">&nbsp;Jobplanner (Gantt årsoversigt)</h3>
	
	
	
	
	<%
	if len(request("FM_kunde")) <> 0 then
			
			if request("FM_kunde") = 0 then
			valgtKunde = 0
			sqlKundeKri = "jobknr <> 0"
			else
			valgtKunde = request("FM_kunde")
			sqlKundeKri = "jobknr = "& valgtKunde &""
			end if
			
	else
		
			if len(request.cookies("pga")("kon")) <> 0 AND request.cookies("pga")("kon") <> 0 then
			valgtKunde = request.cookies("pga")("kon")
			sqlKundeKri = "jobknr = "& valgtKunde &""
			else
			valgtKunde = 0
			sqlKundeKri = "jobknr <> 0"
			end if
			
	end if
	
	
	response.cookies("pga")("kon") = valgtKunde
	%>
	
	<%if request("print") <> "j" then
	fDsp = ""
	fWzb = "visible"
	else
	fDsp = "none"
	fWzb = "hidden"
	end if%>
	<div id="filter" style="position:relative; visibility:<%=fWzb %>; display:<%=fDsp%>; padding:10px; border-left:1px #5582d2 solid; border-top:1px #5582d2 solid; border-right:2px #5582d2 solid; border-bottom:2px #5582d2 solid; background-color:#EFf3FF;">
	<table width=80% cellspacing=2 cellpadding=2 border=0>
	<form action="pg_allokering.asp?menu=job" method=post>
	<tr>
	    <td>
	            
	            
	
	<b>Kontakter:</b><br> <select name="FM_kunde" size="1" style="font-size : 9px; width:285px;">
		<option value="0">Alle</option>
		<%
		ketypeKri = " ketype <> 'e'"
		strKundeKri = " AND kid <> 0 "
		vlgtKunde = " Alle "
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(valgtKunde) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				vlgtKunde = oRec("Kkundenavn") & " ("& oRec("Kkundenr") &")"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select><br>
		
		<%
		if len(request("FM_start_mrd")) <> 0 then	
		strMrd = request("FM_start_mrd")
		else
			if len(request.cookies("pga")("stm")) <> 0 then
			strMrd = request.cookies("pga")("stm")
			else
			strMrd = month(now)
			end if
		end if
		
		response.cookies("pga")("stm") = strMrd
		strMrdNavn = monthname(strMrd)
		
		if len(request("FM_start_aar")) <> 0 then	
		strAar = request("FM_start_aar")
		else
			if len(request.cookies("pga")("staar")) <> 0 then
			strAar = request.cookies("pga")("staar")
			else
			strAar = year(now)
			end if
		end if
		
		useyear = strAar
		response.cookies("pga")("staar") = useyear
		
		
		
		useLowerdatoKri = useyear&"/"&strMrd&"/1"
		uDat = dateadd("m", 6, "1/"&strMrd&"/"&useyear)
		useUpperdatoKri = year(uDat)&"/"&month(uDat)&"/1"
		
		
		
		if request("viskunjobmeddeadline") = 1 then
			'Response.write "+1"
			vkjmd = 1
			tjkjmd = "CHECKED"
			sqlDatoKri = "(jobslutdato BETWEEN '"& useLowerdatoKri &"' AND '"& useUpperdatoKri &"')"
		else
			if request.cookies("pga")("vkjmd") = "1" AND len(request("FM_kunde")) = 0 then
			'Response.write "+2"
			vkjmd = 1
			tjkjmd = "CHECKED"
			sqlDatoKri = "(jobslutdato BETWEEN '"& useLowerdatoKri &"' AND '"& useUpperdatoKri &"')"
			else
			'Response.write "+3"
			vkjmd = 0
			tjkjmd = ""
			sqlDatoKri = "(jobstartdato <= '"& useUpperdatoKri &"' AND jobslutdato >= '"& useLowerdatoKri &"')"
			end if
		end if
		
		response.cookies("pga")("vkjmd") = vkjmd
		
		%>
		
		<br><b>Periode:</b><br>
		Vis job der er aktive i perioden fra <select name="FM_start_mrd" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>&nbsp;&nbsp;
		<select name="FM_start_aar" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<option value="<%=strAar%>"><%=strAar%></option>
		<option value="2002">02</option>
		<option value="2003">03</option>
	   	<option value="2004">04</option>
	   	<option value="2005">05</option>
		<option value="2006">06</option>
		<option value="2007">07</option>
		<option value="2008">08</option>
		<option value="2009">09</option>
		<option value="2010">10</option>
		<option value="2011">11</option>
		<option value="2012">12</option></select> og 6 måneder frem.
		
		<br><input type="checkbox" name="viskunjobmeddeadline" value="1" <%=tjkjmd%>>Vis kun job med slutdato i det valgte interval.
	    
	    <br>
	    <%
	    if len(request("FM_start_aar")) <> 0 then
	        if len(request("FM_status")) <> 0 then
	        stat = 1
	        stCHK = "CHECKED"
	        else
	        stat = 0
	        stCHK = ""
	        end if
	    response.cookies("pga")("status") = stat
	    else
	        if request.cookies("pga")("status") <> "" then
	        stat = 1
	        stCHK = "CHECKED"
	        else
	        stat = 0
	        stCHK = ""
	        end if
	    end if
	    %>
	    <input type="checkbox" name="FM_status" value="1" <%=stCHK%>>Vis også passive job.
	   
	    <%
	    response.cookies("pga").expires = date + 32
	    %>
	    
	    </td>
	    
	    
	    
	    <td valign=top>
		
		<%
		if len(request("FM_valgrisiko")) <> 0 AND request("FM_valgrisiko") <> 0 then
		risk = request("FM_valgrisiko")
		riskSQlKri = " risiko = " & risk &" AND "
		else
		risk = 0
		riskSQlKri = " risiko <> 4 AND " 
		end if 
		
		
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		
		
		select case risk
		case 1
		stCHK0 = ""
		stCHK1 = "SELECTED"
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		stName = "Lav"
		case 2
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = "SELECTED"
		stCHK3 = ""
		stCHK4 = ""
		stName = "Mellem"
		case 3
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = "SELECTED"
		stCHK4 = ""
		stName = "Høj"
		case 4
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = "SELECTED"
		stName = "Skjulte"
		case else
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		stName = "Alle (undt. Skjulte)"
		end select
		
		
		
		%>
		
		<b>Prioitet:</b><br /> <select name="FM_valgrisiko" id="FM_valgrisiko" style="font-size:9px; font-family:arial; width:190px;">
		<option value="0" <%=stCHK0%>>Alle (undt. Skjulte)</option>
		<option value="1" <%=stCHK1%>>Lav</option>
		<option value="2" <%=stCHK2%>>Mellem</option>
		<option value="3" <%=stCHK3%>>Høj</option>
		<option value="4" <%=stCHK4%>>Skjulte</option>
		</select>
		
		
		
		
		
		</td>
	    
	    
	    <td align=right valign=bottom style="padding:10px;">
	    <br><input type="submit" value="Vis valgte Job">
	    </td>
	</tr></form>
	</table>
	
	</div>
	
	
	<%if request("print") <> "j" then%>
	<br /><a href='pg_allokering.asp?menu=job&print=j&FM_kunde=<%=valgtKunde%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&viskunjobmeddeadline=<%=vkjmd%>' class='vmenu' target="_blank">Print venlig version</a>&nbsp;&nbsp;|&nbsp;&nbsp; <a href='jobs.asp?menu=job&func=opret&id=0&int=1&rdir=pg' class='vmenu'>Opret nyt job..</a>
    <%end if%>
	
	<%
	if request("FM_ignorer_pg") <> "1" then
		call hentbgrppamedarb(session("mid"))
	else
		strPgrpSQLkri = ""
	end if
	%>
	
	<br />
	<br>Periode:&nbsp;<b><%=useyear%>, <%=strMrdNavn%> og 6 md. frem. </b>&nbsp;(kun aktive job),<br /> 
	Valgt kontakt: <b><%=vlgtKunde %></b><br />
	Prioitet: <b><%=stName%></b>
	<br /><br />
	<table cellspacing="1" cellpadding="0" border="0" width=90% bgcolor="#8caae6">
	<form method="post" action="pg_allokering.asp?func=opdaterjobliste&FM_valgrisiko=<%=risk %>">
	<tr bgcolor="#ffffff">
	<td height=20 style="padding-left:5px;"><b>Kunde</b></td>
	<td style="padding-left:5px;"><b>Jobnavn</b></td>
	<td width=70 style="padding-left:5px;"><b>Prioitet</b></td>
	<td width=70 style="padding-left:5px;"><b>Periode</b></td>
	<%
	
	for m = 0 to 26
	beregn_thisMth = dateadd("ww", m, "1/"&strMrd&"/"&strAar)
	thisMth = datepart("m", beregn_thisMth)
	if lastmname <> left(monthname(thisMth), 3) then%>
	<td width=40 colspan=4 style="padding-left:5px;">&nbsp;<b><%=left(monthname(thisMth), 3)%>&nbsp;<%=right(datepart("yyyy", beregn_thisMth), 2)%></b></td>
	<%
	end if
	lastmname = left(monthname(thisMth), 3)
	next%>
	
	
	</tr>
	
	
	<%
	
	if cint(stat) = 1 then
	statKri = "jobstatus = 1 OR jobstatus = 2"
	else
	statKri = "jobstatus = 1"
	end if
	
	strSQL = "SELECT job.id AS jobid, jobnavn, jobstartdato, jobnr, jobslutdato, "_
	&" jobstatus, k.kkundenavn, k.kkundenr, risiko FROM job, kunder k "_
	&" WHERE "& riskSQlKri &" fakturerbart = 1 AND "& sqlDatoKri &" AND "_
	&" ("& statKri &") AND "& sqlKundeKri &" AND k.kid = jobknr ORDER BY " & orderkri
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	select case oRec("jobstatus")
	case 0
	useColorgif = "<img src='../ill/jobpl0.gif' width='8' height='8' alt='' border='0'>"
	case 1
	useColorgif = "<img src='../ill/jobpl1.gif' width='8' height='8' alt='' border='0'>"
	case 2
	useColorgif = "<img src='../ill/jobpl2.gif' width='8' height='8' alt='' border='0'>"
	end select
	
		%>
		
		<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=oRec("jobid")%>">
		
		
		<tr>
			<td style="padding-left:5px; padding-right:5px;" bgcolor="#FFFFFF" class=lille><%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>)</td>
			<td bgcolor="#FFFFFF" style="padding-left:5px; padding-right:5px;" class=lille>
			<%if request("print") = "j" then%>
			<b><%=oRec("jobnavn")%></b>
			<%else%>
			<a href="jobs.asp?menu=job&func=red&int=1&id=<%=oRec("jobid")%>&rdir=pg" class="rmenu"><%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>)</a>
			&nbsp;<a href="javascript:NewWin_large('jobplanner.asp?menu=job&id=<%=oRec("jobid")%>')"><img src="../ill/popupcal_small.gif" width="11" height="10" alt="Rediger jobperiode og aktiviteter" border="0"></a>
			<%end if%>
			
			</td>
			
			<td bgcolor="#ffffff" valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed; border-right:1px #cccccc solid" class=lille>
		<%
		
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		selbgcol = ""
		
		select case oRec("risiko")
		case 1
		stCHK0 = ""
		stCHK1 = "SELECTED"
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		selbgcol = "#ffffe1"
		stName = "Lav"
		case 2
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = "SELECTED"
		stCHK3 = ""
		stCHK4 = ""
		selbgcol = "Goldenrod"
		stName = "Mellem"
		case 3
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = "SELECTED"
		stCHK4 = ""
		selbgcol = "Red"
		stName = "Høj"
		case 4
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = "SELECTED"
		selbgcol = "silver"
		stName = "Skjult"
		case else
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		selbgcol = ""
		stName = "Ej ang."
		end select
		
		
		if request("print") <> "j" then%>
		<select name="FM_risiko" id="FM_risiko" style="background-color:<%=selbgcol%>; font-size:9px; font-family:arial;">
		<option value="0" <%=stCHK1%>>Ej angivet</option>
		<option value="1" <%=stCHK1%>>Lav</option>
		<option value="2" <%=stCHK2%>>Mellem</option>
		<option value="3" <%=stCHK3%>>Høj</option>
		<option value="4" <%=stCHK4%>>Skjul</option>
		</select>
		<%else %>
		
		<%=stName%>
		
		<%end if %>
		</td>
			
			<td bgcolor="#ffffff" style="padding-left:5px; padding-right:5px;"><font color=green size=1><%=formatdatetime(oRec("jobstartdato"), 2)%> <br />  <font color=crimson size=1><%=formatdatetime(oRec("jobslutdato"), 2)%> </td>
				
		
			<%
			for m = 0 to 26
			beregn_thisMth = dateadd("ww", m, "1/"&strMrd&"/"&strAar)
			thisMth = datepart("ww", beregn_thisMth,2,2)
			thisYear = datepart("yyyy", beregn_thisMth)
				
				%>
				<td bgcolor="#FFFFFF" style="padding:5px;">
				<%
				
				if (datepart("ww", oRec("jobstartdato"),2,2) <= thisMth AND datepart("yyyy", oRec("jobstartdato")) = thisYear)_
				OR (datepart("yyyy", oRec("jobstartdato")) < thisYear) then
				
				if  (datepart("ww", oRec("jobslutdato"),2,2) => thisMth AND datepart("yyyy", oRec("jobslutdato")) = thisYear)_
				OR (datepart("yyyy", oRec("jobslutdato")) > thisYear) then
					
				
				Response.Write useColorgif
				%>
				
				<%else%>
				<img src='../ill/blank.gif' width='8' height='8' alt='' border='0'>
				<%end if
				
				else%>
				<img src='../ill/blank.gif' width='8' height='8' alt='' border='0'>
				
				
				<%
				end if
			    %>
			    </td>
			    <%
			next%>
			
			
		
		</tr>
		<%
	oRec.movenext
	wend
	oRec.close
	%>	
	
	
	</table>
	<%if request("print") <> "j" then %>
	<br />
    <input id="Submit1" type="submit" value="Opdater liste" />
    <%end if %>
	</form>
	
	
	<br><br>
	<br>
	<%if request("print") <> "j" then%>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<%end if%>
	<br>
	<br>
	</div>
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
