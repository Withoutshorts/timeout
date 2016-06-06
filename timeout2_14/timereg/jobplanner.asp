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
	%>
	
	<%
	'*** Loader bar ****
	if request("print") = "j" then%>
	<div id="loading" name="loading" style="position:absolute; top:65; left:360; width:300; height:20; z-index:1000; visibility:visible; padding-left:4px; padding-top:5px;">
	<img src="../ill/loaderbar.gif" width="174" height="13" alt="" border="0">
	</div>
	<%end if%>
	
	<%	
	func = request("func")
	if func = "upd" then
	
	stdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&request("FM_start_dag")
	sldato = request("FM_slut_aar")&"/"&request("FM_slut_mrd")&"/"&request("FM_slut_dag")
	thisaid = request("jaid")
	thisjid = request("id")
	strAffaktSTART = ""
	strAffaktSLUT = ""
	chaktdatesST = "n"
	chaktdatesSL = "n"
	godk = request("godk")
	
	'**** Er der valgt gyldige datoer?? ***
	if not isdate(stdato) OR not isdate(sldato) then
	%>
	<div id="er3" name="er3" style="position:absolute; width:400; left:200; top:150; visibility:visible; display:; background-color:#FFFFE1; filter:alpha(opacity=90); border-left:1 #990200 solid; border-top:1 #990200 solid; border-bottom:1 #990200 solid; border-right:1 #990200 solid; z-index:5000;">
	<table border=0 cellpadding=0 cellspacing=0 width="400" height="100">
	<tr><td style="padding-left:5; padding-top:4;"><b>Fejl!</b><br>
	Den valgte <b>startdato</b> eller den valgte <b>slutdato</b> er en ikke gyldig dato.<br>
	F.eks 31/11 
	<br><br>
	<a href="jobplanner.asp?menu=job&id=<%=thisjid%>" class="vmenu">Prøv igen...</a>
	</td></tr>
	</table>
	</div>
	<%	
	
	else
		'** hvis akt start og slut dato er større end job start og slut dato ***
		if (cdate(stdato) < cdate(request("jstdato")) OR cdate(request("jsldato")) < cdate(sldato)) AND thisaid <> 0 then
		%>
		<div id="er1" name="er1" style="position:absolute; width:400; left:200; top:150; visibility:visible; display:; background-color:#FFFFE1; filter:alpha(opacity=90); border-left:1 #990200 solid; border-top:1 #990200 solid; border-bottom:1 #990200 solid; border-right:1 #990200 solid; z-index:5000;">
		<table border=0 cellpadding=0 cellspacing=0 width="400" height="100">
		<tr><td style="padding-left:5; padding-top:4;"><b>Fejl!</b><br>
		A) Den valgte <b>startdato</b> er en tidligere dato end <b>startdatoen</b> på jobbet!<br>
		B) Den valgte <b>slutdato</b> er en senere dato end <b>slutdatoen</b> på jobbet!.<br><br>
		Aktiviteten er <b>ikke</b> blevet opdateret!
		<br><br>
		<a href="#" onclick=closeer1(); class="vmenu">Luk dette vindue og prøv igen. [x]</a>
		</td></tr>
		</table>
		</div>
		<%
		else
				
			'*** er startdato større end slut datoen?? ***	
			if (cdate(stdato) > cdate(sldato)) then
			%>
			<div id="er3" name="er3" style="position:absolute; width:400; left:200; top:150; visibility:visible; display:; background-color:#FFFFE1; filter:alpha(opacity=90); border-left:1 #990200 solid; border-top:1 #990200 solid; border-bottom:1 #990200 solid; border-right:1 #990200 solid; z-index:5000;">
			<table border=0 cellpadding=0 cellspacing=0 width="400" height="100">
			<tr><td style="padding-left:5; padding-top:4;"><b>Fejl!</b><br>
			Den valgte <b>startdato</b> er en senere dato end <b>slutdatoen</b>. 
			<br><br>
			<a href="jobplanner.asp?menu=job&id=<%=thisjid%>" class="vmenu">Prøv igen...</a>
			</td></tr>
			</table>
			</div>
			<%	
			else
		
			'**** er det en aktivitet eller jobbet der opdateres?? ***
			if thisaid <> 0 then
			strSQL2 = "UPDATE aktiviteter SET aktstartdato = '"&stdato&"', aktslutdato = '"&sldato&"' WHERE id=" & thisaid
			oConn.execute(strSQL2)
			else
			
					strSQL3 = "SELECT id, aktstartdato, aktslutdato, navn FROM aktiviteter WHERE aktiviteter.job ="& thisjid 
					oRec.open strSQL3, oConn, 3
					while not oRec.EOF 
							
							if cdate(oRec("aktstartdato")) < cdate(stdato) then
								if godk = "t" then
								strSQL4 = "UPDATE aktiviteter SET aktstartdato = '"&stdato&"' WHERE id=" & oRec("id")
								oConn.execute(strSQL4)
								end if
							chaktdatesST = "y"
							strAffaktSTART = strAffaktSTART & oRec("navn") & "<br>"
							end if
							
							if cdate(oRec("aktslutdato")) > cdate(sldato) then
								if godk = "t" then
								strSQL5 = "UPDATE aktiviteter SET aktslutdato = '"&sldato&"'  WHERE id=" & oRec("id")
								oConn.execute(strSQL5)
								end if
							chaktdatesSL = "y"
							strAffaktSLUT = strAffaktSLUT & oRec("navn") & "<br>"
							end if
							
					oRec.movenext
					wend
					oRec.close
					
					if (chaktdatesST = "y" OR chaktdatesSL = "y") AND godk <> "t" then
					%>
					<div id="er2" name="er2" style="position:absolute; width:400; left:200; top:150; visibility:visible; display:; background-color:#FFFFE1; filter:alpha(opacity=90); border-left:1 #990200 solid; border-top:1 #990200 solid; border-bottom:1 #990200 solid; border-right:1 #990200 solid; z-index:5000;">
					<table border=0 cellpadding=0 cellspacing=0 width="400" height="100">
					<%
					'*** Form til at gemme datoer hvis opdatering fortrydes. ***
					'*** Jobdato er rykket så det har indflydelse ppå aktiviteter ****
					%>
					<form action="jobplanner.asp?menu=job&func=upd&godk=t" method="post">
					<input type="hidden" name="FM_start_dag" value="<%=request("FM_start_dag")%>">
					<input type="hidden" name="FM_start_mrd" value="<%=request("FM_start_mrd")%>">
					<input type="hidden" name="FM_start_aar" value="<%=request("FM_start_aar")%>">
					<input type="hidden" name="FM_slut_dag" value="<%=request("FM_slut_dag")%>">
					<input type="hidden" name="FM_slut_mrd" value="<%=request("FM_slut_mrd")%>">
					<input type="hidden" name="FM_slut_aar" value="<%=request("FM_slut_aar")%>">
					<input type="hidden" name="jaid" value="<%=request("jaid")%>">
					<input type="hidden" name="id" value="<%=request("id")%>">
					<tr><td style="padding-left:5; padding-top:4;"><b>Info!</b><br>
					En eller flere af de tilhørende aktiviteter har en start eller
					slut dato der ligger uden for det netop angivet tidsinterval.
					<br><br>
					<%if chaktdatesST = "y" then%>
					<b>Følgende aktiviteter vil få rykket deres startdato til d. <%=request("FM_start_dag")&"/"&request("FM_start_mrd")&"/"&request("FM_start_aar")%></b>:<br>
					<%=strAffaktSTART%><br><br>
					<%end if%>
					
					<%if chaktdatesSL = "y" then%>
					<b>Følgende aktiviteter vil få rykket deres slutdato til d. <%=request("FM_slut_dag")&"/"&request("FM_slut_mrd")&"/"&request("FM_slut_aar")%></b>:<br>
					<%=strAffaktSLUT%>
					<br><br>
					<%end if%>
					Skal job start- og slutdato rykkes og skal datoerne på ovenstående aktiviteter ændres til at følge denne start- og slutdato?
					<br><br>
						</td>
					</tr>
					<tr>
						<td style="padding-left:5; padding-top:2;" align=center>
						<a href="jobplanner.asp?menu=job&id=<%=thisjid%>" class="vmenu"><img src="../ill/annulerpil.gif" width="116" height="30" alt="" border="0"></a>
						&nbsp;&nbsp;&nbsp;<input type="image" src="../ill/opdaterpil.gif"><br><br>
						</td>
					</tr>
					</form>	
					</table>
					</div>
					<%
					end if
				
					'**** Opdaterer job ***
					if godk = "t" OR (chaktdatesST = "n" AND chaktdatesSL = "n") then
					strSQL2 = "UPDATE job SET jobstartdato = '"&stdato&"', jobslutdato = '"&sldato&"' WHERE id=" & thisjid
					oConn.execute(strSQL2)
					end if	
				 end if '*** validering
			end if
		end if
		end if
	end if
	
	%>
	<html>
		<head>
		<title>TimeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
		<script LANGUAGE="javascript">
		function NewWin_popupcal(url) {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
		</script>
		
		<script language="javascript">
		function showeditdate(jaid, stdag, stmd, staar, sldag, slmd, slaar){
		var usestdag = stdag;
		var usestmd = stmd;
		var usestaar = staar;
		var usesldag = sldag;
		var useslmd = slmd;
		var useslaar = slaar;
		
		document.all["editdate"].style.display = "";
		document.all["editdate"].style.visibility = "visible";
		document.all["jaid"].value = jaid;
		document.all["editdate"].style.left = window.event.x - 20
		document.all["editdate"].style.top = window.event.y - 50
		document.all["FM_start_dag"].value = usestdag;
		document.all["FM_start_mrd"].value = usestmd;
		document.all["FM_start_aar"].value = usestaar;
		document.all["FM_slut_dag"].value = usesldag;
		document.all["FM_slut_mrd"].value = useslmd;
		document.all["FM_slut_aar"].value = useslaar
		}
		
		function closeeditdate(){
		document.all["editdate"].style.display = "none";
		document.all["editdate"].style.visibility = "hidden";
		}
		
		function closeer1(){
		document.all["er1"].style.display = "none";
		document.all["er1"].style.visibility = "hidden";
		}
		
		//function closeer3(){
		//document.all["er3"].style.display = "none";
		//document.all["er3"].style.visibility = "hidden";
		//}
		</script>
		</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<div id="loading" name="loading" style="position:absolute; top:60; left:340; width:300; height:20; z-index:1000; visibility:visible; padding-left:4px; padding-top:5px;">
	<!--<img src="../ill/info.gif" width="42" height="38" alt="" border="0">--><b>Henter information (kan tage optil 15-20 sek.)...</b><br><br>
	<img src="../ill/loaderbar.gif" width="174" height="13" alt="" border="0">
	<!--Denne handling kan tage op til 20 sek. alt efter din forbindelse.-->
	</div>
	
	
	<!--#include file="inc/convertDate.asp"-->
	
	<%
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
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
	
	
	leftPos = 0
	topPos = 0
	
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	
	<table cellspacing="0" cellpadding="0" border="0" width="880">
		<tr>
			<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
			<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
		</tr>
		</table>
	<br><br>
	
    <div style="position:relative; left:20px;">
			
	<table cellspacing="1" cellpadding="0" border="0" bgcolor="silver">
	<%
	strSQL = "SELECT job.id AS jobid, jobnavn, jobnr, job.fakturerbart, jobstartdato, jobslutdato, jobstatus, aktiviteter.job, aktiviteter.navn AS anavn, aktstartdato, aktslutdato, aktstatus, fakturerbar, aktiviteter.id AS aid FROM job RIGHT JOIN aktiviteter ON (aktiviteter.job = "& id &" AND fakturerbar <> 2) WHERE job.id = "& id & " ORDER BY job.id, aktstartdato"
	'Response.write strSQL
	
	oRec.open strSQL, oConn, 3
	x = 0
	while not oRec.EOF
	
	'******** udskriver det valgte job ************
	if x = 0 then
	
	thisjobnavn = oRec("jobnavn")
	thisjid = oRec("jobid")
	useLowerdatoKri = cdate(oRec("jobstartdato"))
	useUpperdatoKri = cdate(oRec("jobslutdato"))
	
	strDofyear = datepart("y", oRec("jobstartdato"))
	strDofyear_slut = datepart("y", oRec("jobslutdato"))
	
	jobdiff = datediff("d", useLowerdatoKri, useUpperdatoKri) 
	
	'strAktStDato = datepart("yyyy", oRec("aktstartdato"),2,2)
	'strAktStDato_slut = datepart("yyyy", oRec("aktslutdato"),2,2) 
	
	%>
		
			<tr bgcolor="#d6dff5">
			<td style="padding-left:2;">&nbsp;</td>
			<%
			'** td'er over kalender popup ***
			if request("print") <> "j" then%>
			<td>&nbsp;</td>
			<%end if
			
			'for firstyear = firstyear to lastyear
				cnt = 0 
				for cnt = 0 to jobdiff
				getnyDatoMrd = dateadd("d", cnt, useLowerdatoKri)
				if lastm <> month(getnyDatoMrd) then%>
				<td bgcolor="#8caae6" style="padding:2;"><b><%=left(monthname(month(getnyDatoMrd)), 3)%>&nbsp;<%=right(year(getnyDatoMrd),2)%></b></td>
				<%end if%>
				
				<%	
					thisDay = Weekday(getnyDatoMrd)
					thisDayn = day(getnyDatoMrd)
						
						
						Response.write "<td align=center class=lille style=""width:14px;"">"& thisDayn & "</td>"%>
						<%
						
				
				lastm = month(getnyDatoMrd) 
				'cnt = cnt + 1
				next
			'next
			%>
			<td valign="top" width=20 colspan="2" class=alt align="right">&nbsp;</td>
		</tr>
		
		<tr>
			<td bgcolor="#FFFFFF" style="padding:3px; width:300px;"><b><%=left(oRec("jobnavn"), 20)%></b></td>
			<%if request("print") <> "j" then%>
			<td bgcolor="#FFFFFF" style="padding-left:1; padding-right:1;"><a href="#" onClick="showeditdate('0','<%=day(oRec("jobstartdato"))%>','<%=month(oRec("jobstartdato"))%>','<%=datepart("yyyy", oRec("jobstartdato"))%>','<%=day(oRec("jobslutdato"))%>','<%=month(oRec("jobslutdato"))%>','<%=datepart("yyyy", oRec("jobslutdato"))%>');"><img src="../ill/popupcal.gif" width="16" height="15" alt="" border="0"></a></td>
			<%end if
			
			'** reset værdier **
			strDofyear = datepart("y", oRec("jobstartdato"))
			strDofyear_slut = datepart("y", oRec("jobslutdato")) 
			jobdiff = datediff("d", useLowerdatoKri, useUpperdatoKri) 
			cnt = 0
			'd = strDofyear
			for cnt = 0 to jobdiff
			getnyDato = dateadd("d", cnt, useLowerdatoKri)
			thisDayn = day(getnyDato)
			thisWeekDay = Weekday(getnyDato)		
					
					
				if cnt <> 0 then%>
				<td bgcolor="#ffffff" align=center style="width:14px; padding:4px;"><img src="../ill/g-prik.gif" width="6" height="6" alt="" border="0"></td>
				<%end if
				
				if lastm <> month(getnyDato) OR cnt = 0 then%>
					<td bgcolor="#8caae6" style="width:14px;" class=alt>&nbsp;</td>
				<%end if
					
				lastm = month(getnyDato)
			next
			
			'end if%>
			<td valign="top" bgcolor="snow" width=20 colspan="2" class=alt align="right">&nbsp;</td>
		</tr>
		
		
	<%end if
	
	
	'************** udskriver aktiviteter *********************
	%>
	<tr>
		<td bgcolor="#ffffff" style="padding:3px; width:300px;" class=lille>
		<%if cint(thisaid) = cint(oRec("aid")) then%>
			<u><%=left(oRec("anavn"), 20)%></u>
		<%else%>
			<%=left(oRec("anavn"), 20)%>
		<%end if%></td>
			
			<%if request("print") <> "j" then
			'***** Popup kalender Akt. ***%>
			<td bgcolor="#ffffff" style="padding-left:1; padding-right:1;"><a href="#" onClick="showeditdate('<%=oRec("aid")%>','<%=day(oRec("aktstartdato"))%>','<%=month(oRec("aktstartdato"))%>','<%=datepart("yyyy", oRec("aktstartdato"))%>','<%=day(oRec("aktslutdato"))%>','<%=month(oRec("aktslutdato"))%>','<%=datepart("yyyy", oRec("aktslutdato"))%>');"><img src="../ill/popupcal.gif" width="16" height="15" alt="" border="0"></a></td>
			<%end if
			strADofyear = datepart("y", oRec("aktstartdato"))
			strADofyear_slut = datepart("y", oRec("aktslutdato"))  
			
			cnt = 0
			for cnt = 0 to jobdiff
			
			nextAktDato = dateAdd("d", cnt, oRec("jobstartdato"))
				
				
				
				if cnt <> 0 then
				
					if (cdate(nextAktDato) >= cdate(oRec("aktstartdato")) AND cdate(nextAktDato) <= cdate(oRec("aktslutdato"))) then%>
					<td bgcolor="#ffffff" align=center style="width:14px; padding:4px;"><img src="../ill/g-prik.gif" width="6" height="6" alt="" border="0"></td>
					<%else%>
					<td bgcolor="#FFFFFF" align=center style="width:14px; padding:4px;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					<%end if
				
				end if
				
				if lastm <> month(nextAktDato) OR cnt = 0 then%>
				<td bgcolor="#8caae6" class=alt style="padding-left:2; padding-right:2;">&nbsp;</td>
				<%end if
				
				lastm = month(nextAktDato)
				'cnt = cnt + 1
			next
		%>
		<td valign="top" bgcolor="#FFFFFF" width=20 colspan="2" class=alt align="right">&nbsp;</td>
	</tr>
	<%
	x = x + 1
	oRec.movenext
	wend
	
	oRec.close
	%>	
	</table>
	</div>
	
	<br><br>
	<%if request("print") <> "j" then%>
	&nbsp;&nbsp;<a href="Javascript:window.close()" class=vmenu>Luk vindue</a>
	<%end if%>
	<script>
	document.all["loading"].style.visibility = "hidden";
	</script>
	<br><br>
	<br>
	</div>
	
	<div id="editdate" name="editdate" style="position:absolute; width:200; left:300; top:250; visibility:hidden; display:none; background-color:#FFFFE1; filter:alpha(opacity=90); border-left:1 #5582d2 solid; border-top:1 #5582d2 solid; border-bottom:1 #5582d2 solid; border-right:1 #5582d2 solid;">
	<table border=0 cellpadding=0 cellspacing=0 width="400" height="100">
	<form action="jobplanner.asp?menu=job&func=upd" target="_self" method="post" name="updated" id="updated">
	<input type="hidden" name="id" id="id" value="<%=thisjid%>">
	<input type="hidden" name="jaid" id="jaid" value="0">
	<input type="hidden" name="jstdato" value="<%=useLowerdatoKri%>">
	<input type="hidden" name="jsldato" value="<%=useUpperdatoKri%>">
	<tr>
		<td colspan="2" style="Padding-left:5; Padding-right:5; Padding-top:4;"><b>Vælg start og slut dato:</b><br>
		Jobperiode:&nbsp;<%=formatdatetime(useLowerdatoKri,1)%> - <%=formatdatetime(useUpperdatoKri, 1)%></td>
	</tr>
	
	<tr>
	<td style="Padding-left:5; Padding-right:5; Padding-top:3;">
	<!-- start dato kal -->
	<select name="FM_start_dag">
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>
		
		<select name="FM_start_mrd">
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
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_start_aar">
		<%for x = -5 to 5 
		useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		<option value="<%=useY%>"><%=right(useY, 2)%></option>
		<%next %></select>
		
		
		&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
				
	</td>
	<td style="Padding-left:5; Padding-right:5; Padding-top:3;">
	
	<select name="FM_slut_dag">
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>
		
		<select name="FM_slut_mrd">
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
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_slut_aar">
		<%for x = -5 to 5
		useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		<option value="<%=useY%>"><%=right(useY, 2)%></option>
		<%next %></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=4')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		
	
	
	</td>
	</tr>
	<tr>
		<td colspan="2" align=center style="Padding-left:5; Padding-right:5;"><br>
		<input type="image" src="../ill/opretpil.gif">
		</td>
	</tr>
	<tr>
		<td colspan="2" align=right style="Padding-left:5; Padding-right:5;">
		<br><a href="#" onclick=closeeditdate(); class="vmenu">Luk dette vindue [x]</a></td>
	</tr>
	
	</form>
	</table>	
	</div>
	<script>
document.all["loading"].style.visibility = "hidden";
</script>
	<%
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
