<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	if len(request("year")) <> 0 then
	useyear = request("year")
	else
	useyear = year(now)
	end if
	
	if len(request("usejob")) <> 0 then
	usejob = request("usejob")
	strJobKri = " ORDER BY projektgrupper.id"
	else
	usejob = 0
	strJobKri = " GROUP BY job.id ORDER BY jobnavn"
	end if
	
	
	
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	
	<%
	if request("print") = "j" then
	leftPos = 10
	topPos = 20
	else
	leftPos = 190
	topPos = 80
	%>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%
	end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    	<td valign="top"><img src="../ill/header_pgallok.gif" alt="" border="0" width="758" height="45"></TD>
	</TR>
	</TABLE>
	<%if request("print") <> "j" then%>
	<br>
	Vis for <a href="pg_allokering.asp?menu=job&year=<%=useyear%>">job</a> eller <a href="pg_allokering.asp?menu=job&year=<%=useyear%>&usejob=1">projektgrupper</a>
	<br><br>
	<a href="pg_allokering.asp?menu=job&year=2002">2002</a>&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&year=2003">2003</a>&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&year=2004">2004</a>&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&year=2005">2005</a>&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&year=2006">2006</a>
	<img src="../ill/blank.gif" width="350" height="1" alt="" border="0"><a href="javascript:NewWin('pg_allokering.asp?menu=job&print=j&year=<%=useyear%>&usejob=<%=usejob%>')" target="_self" class='rmenu'>&nbsp;Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<br>
	<%end if%>
	Oversigt er for år: <font class=roed><%=useyear%></font><br>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582d2">
	<tr bgcolor="#5582D2">
	<td width="150" class=alt>&nbsp;<b>Projektgruppe / job</b>&nbsp;&nbsp;</td>
	<td width=31 class=alt style="border:1px #000000 solid;">&nbsp;<b>Jan</b></td>
	<td width=30 class=alt style="border:1px #000000 solid;">&nbsp;<b>Feb</b></td>
	<td width=31 class=alt style="border:1px #000000 solid;">&nbsp;<b>Mar</b></td>
	<td width=30 class=alt style="border:1px #000000 solid;">&nbsp;<b>Apr</b></td>
	<td width=31 class=alt style="border:1px #000000 solid;">&nbsp;<b>Maj</b></td>
	<td width=30 class=alt style="border:1px #000000 solid;">&nbsp;<b>Jun</b></td>
	<td width=31 class=alt style="border:1px #000000 solid;">&nbsp;<b>Jul</b></td>
	<td width=31 class=alt style="border:1px #000000 solid;">&nbsp;<b>Aug</b></td>
	<td width=30 class=alt style="border:1px #000000 solid;">&nbsp;<b>Sep</b></td>
	<td width=31 class=alt style="border:1px #000000 solid;">&nbsp;<b>Okt</b></td>
	<td width=30 class=alt style="border:1px #000000 solid;">&nbsp;<b>Nov</b></td>
	<td width=31 class=alt style="border:1px #000000 solid;">&nbsp;<b>Dec</b></td>
	<td width=30>&nbsp;</td>
	</tr>
	</table>
	<table cellspacing="0" cellpadding="0" border=0">
	
	<%
	strSQL = "SELECT job.id AS jobid, jobnavn, jobstartdato, jobslutdato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, projektgrupper.id AS Pid, projektgrupper.navn AS pnavn FROM job LEFT JOIN projektgrupper ON (projektgrupper.id = projektgruppe1 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe2 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe3 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe4 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe5 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe6 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe7 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe8 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe9 AND projektgrupper.id > 1 OR projektgrupper.id = projektgruppe10 AND projektgrupper.id > 1) WHERE fakturerbart = 1 AND jobstatus = 1 "& strJobKri &""
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	'*** Hvis startdato = valg år eller slut dato >= end valgt år 
	if cint(datepart("yyyy",oRec("jobstartdato"))) = cint(useyear) OR cint(datepart("yyyy",oRec("jobslutdato"))) >= cint(useyear) then
		%>
		<tr>
		<%'** Hvis for job eller proj. grupper. **
		if cint(usejob) <> 1 then %>
		<td class="blaa" valign=top bgcolor="#ffffff" height="10" width="150">&nbsp;&nbsp;
			<%if request("print") = "j" then%>
			<%=left(oRec("jobnavn"), 16)%></td>
			<%else%>
			<a href="jobs.asp?menu=job&func=red&int=1&id=<%=oRec("jobid")%>" class="vmenu"><%=left(oRec("jobnavn"), 16)%></a></td>
			<%end if%>
		<%else%>
		<td bgcolor="#ffffff" class="blaa">&nbsp;&nbsp;<%=oRec("pnavn")%></td>
		<%end if%>
		
		
			<%'*******************************************************************
			'Finder projektlængen og strækker giffen efter denne. 
			'*********************************************************************
			
			'** startår = valgtår
			if cint(datepart("yyyy",oRec("jobstartdato"))) = cint(useyear) then
			projektlaengde = Datediff("d", oRec("jobstartdato"), oRec("jobslutdato"))
			stdayofyear = datepart("y", oRec("jobstartdato")) 
			
			
			d = 1
			for d = 1 to 364
				if cint(stdayofyear) >= cint(d) then%>
				<td width="1"><div style="position:relative; background-color:#FFCC00; height:5; width:1;"></div></td>
				<%else%>
				<td width="1" bgcolor="#CCCCCC"></td>
				<%end if
			next
			
			end if%>
		<td width=30>&nbsp;</td>
		</tr>
		<%
	end if
	
	oRec.movenext
	wend
	
	oRec.close
	%>	
	</table>
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
