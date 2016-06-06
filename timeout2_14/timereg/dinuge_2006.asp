	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%
	'*********** Sætter sidens global variable ***********************************
	
	'* Hvis der vælges en anden bruger
	'* den den der er logget på.
	 
	 '*** Mid !***
	if len(Request("FM_use_me")) <> 0 then
	usemrn = Request("FM_use_me")
	else
	usemrn = session("Mid")
	end if
	
	thisfile = "dinuge"
	func = request("func")
	
	level = session("rettigheder")
	
	session.lcid = 1030
	
	varTjDatoUS_man = request("stdato")
	varTjDatoUS_son = request("sldato")
	

	'************************************************************************************************
	'*** Din uge *******
	'************************************************************************************************
	%>
	<div id="sindhold" style="position:absolute; left:20; top:20; visibility:visible;">
	<table cellspacing="2" cellpadding="0" border="0" width=95%>
	<tr>
		<td colspan="7" valign=top>
		<h4>Din uge / Opgaver tildelt:</h4>
		Følgende timer er projekteret på din profil i denne uge:</td></tr>
	<tr><td colspan=2 valign="top">
	
	<%
	'*** tildelte timer ***
	strSQL3 = "SELECT ressourcer.id, ressourcer.dato AS rdato, timer AS sumtimer, job.jobnavn, jobid, job.jobnr AS jnr, "_
	&" job.id AS jid, k.kkundenavn, k.kkundenr FROM ressourcer "_
	&" LEFT JOIN job ON (job.id = ressourcer.jobid) "_
	&" LEFT JOIN kunder k ON (k.kid = job.jobknr) "_
	&" WHERE (ressourcer.dato >= '"&varTjDatoUS_man&"' AND ressourcer.dato <= '"&varTjDatoUS_son&"') AND mid=" & usemrn &""_
	&" GROUP BY ressourcer.dato, jobid, jobid ORDER BY ressourcer.dato, jobnavn"
	
	'Response.write strSQL3
	'Response.flush
	
	oRec3.open strSQL3, oConn, 3 
	x = 0
	lastresdato = 0
	while not oRec3.EOF 
		
		if x > 0 AND (lastresdato <> oRec3("rdato"))then%>
		</table></td><td valign="top">
			<%if lastresdato <> oRec3("rdato") then%>
			<table cellspacing="2" cellpadding="2" border="0"><tr><td bgcolor="#5582d2" colspan=2 style="padding-left:5px;" class=alt><b><%=weekdayname(weekday(oRec3("rdato"), 1), 1) &" "&formatdatetime(oRec3("rdato"),2)%></b></td></tr>
			<%end if
		else
			if lastresdato <> oRec3("rdato") then%>
			<table cellspacing="2" cellpadding="2" border="0"><tr><td bgcolor="#5582d2" colspan=2 style="padding-left:5px;" class=alt><b><%=weekdayname(weekday(oRec3("rdato"), 1), 1) &" "&formatdatetime(oRec3("rdato"),2)%></b></td></tr>
			<%end if
		end if
		%>
		<tr><td bgcolor="#ffffff" width=200 style="padding-left:5px;"><%=oRec3("jobnavn")%> (<%=oRec3("jnr")%>)<br>
		<font class=lillegray><%=oRec3("kkundenavn")%> (<%=oRec3("kkundenr")%>)</font></td>
		<td width=40 class=lille bgcolor="#ffffff" style="padding-left:5px;" valign=top><%=formatnumber(oRec3("sumtimer"),2)%></td></tr>
	<%
	timertottildelt = timertottildelt + oRec3("sumtimer")
	lastjobid = oRec3("jid")
	lastresdato = oRec3("rdato")
	x = x + 1
	oRec3.movenext
	wend
	oRec3.close
	%>
	</table></td></tr>
	<%if x = 0 then%>
	<tr><td colspan=2>
	<font color=red><b>!</b></font> <b>Der er ikke tildelt ressourcetimer på din profil i denne uge.</b>
	<ul>
	<li>  Der kan tildeles timer ved at redigere et job og klikke på fanebladet "ressourcer".
	 <li>  Eller ved at klikke på menupunktet "Ressourcer" i hovedmenuen og<br> derefter tildele timer under "Ressource timer".
	 </ul>  <br><br>&nbsp;</td></tr>
	<%end if%>
	</table>
	</div>

 



<%
end if 'user session
%>
<!--#include file="../inc/regular/footer_inc.asp"-->





