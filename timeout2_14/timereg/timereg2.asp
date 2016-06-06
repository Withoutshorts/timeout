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
	<SCRIPT language=javascript src="inc/timereg_func.js"></script>
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
	
	thisfile = "timereg"
	
	func = request("func")
	
	'**************** opdaterer guiden og viser timeregsiden ************************* 
	if func = "updguide" then
		oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& usemrn &"")
		
		useJob = request("FM_use_job")
		j = 0
		intuseJob = Split(useJob, ", ")
	   	For j = 0 to Ubound(intuseJob)
		oConn.execute("INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& usemrn &", "& intuseJob(j) &")")
		next
		
		'if request("FM_visguide") = "j" then
			oConn.execute("UPDATE medarbejdere SET visguide = 1 WHERE Mid = "& usemrn &"")
			'else
			'oConn.execute("UPDATE medarbejdere SET visguide = 0 WHERE Mid = "& usemrn &"")
		'end if
	end if
	
	'Response.write "searchstring:" & request("searchstring")
	
	if len(request("searchstring")) = "0" then
		if len(request.cookies("kundetim")) <> 0 then
		searchstring = request.cookies("kundetim")
		else
		searchstring = 0
		end if
	else
	searchstring = cint(request("searchstring"))
	response.cookies("kundetim") = searchstring 
	response.cookies("kundetim").Expires = date + 20
	end if
	
	if len(request("filterA_B")) <> 0 then
	filterABsel = request("filterA_B")
	response.cookies("filterA_B") = filterABsel
	response.cookies("filterA_B").Expires = date + 20
	else
		if len(request.cookies("filterA_B")) <> 0 then
		filterABsel = request.cookies("filterA_B")
		else
		filterABsel = "a"
		end if
	end if
	
	
	'*** Kun hvis filter a er valgt benyttes kundeselect
	if filterABsel = "a" then
	searchstring = searchstring
	else
	searchstring = 0
	end if
	
	 if searchstring = 0 then
	 varSelectedJob = ""
	 else
	 varSelectedJob = " jobknr = "& searchstring &" AND " 
 	 end if	
	
	'*********** Globale Functions ********************************************
	'Bruges både af guiden og af timereg.
	
	public strMnavn
	public strMnr
	public strMType
	public brugergruppeId()
	public f
	public strSQLkri
	
	public function hentbrugergrupper(medarbid)
	f = 0
		Redim brugergruppeId(f)
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid="& medarbid &" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid" 
			oRec.Open strSQL, oConn, 0, 1 
			
			While Not oRec.EOF
			Redim preserve brugergruppeId(f) 
				strMnavn = oRec("Mnavn")
				strMnr = oRec("Mid")
				strMType = oRec("type")
				brugergruppeId(f) = oRec("ProjektgruppeId")
				f = f + 1
				oRec.MoveNext
			Wend
		
		oRec.Close
		
		
		'Rettigheds tjeck på job
		'***********************************************************************
	  	strSQLkri =  varSelectedJob &" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND ("
		
		for intcounter = 0 to f - 1  
	  
	  	strSQLkri = strSQLkri &" job.projektgruppe1 = "&brugergruppeId(intcounter)&""_
		&" OR "_
		&" job.projektgruppe2 = "&brugergruppeId(intcounter)&""_
		&" OR "_
		&" job.projektgruppe3 = "&brugergruppeId(intcounter)&""_
		&" OR "_
		&" job.projektgruppe4 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe5 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe6 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe7 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe8 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe9 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe10 = "&brugergruppeId(intcounter)&""_
		&" OR "
		
		next
	  	
		'** Trimmer de 2 sql states ***
		strSQLkri_len = len(strSQLkri)
		strSQLkri_left = strSQLkri_len - 3
		strSQLkri_use = left(strSQLkri, strSQLkri_left) 
	  	strSQLkri = strSQLkri_use & ") "
		
		
	end function
	'**************************************************************
	
	
	  '*** Vis Guiden ell. timereg
	  if func <> "updguide" then
		strSQL = "SELECT visguide FROM medarbejdere WHERE Mid = "& usemrn &""
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
			if oRec("visguide") = 0 then
			visGuide = 0
			else
			visGuide = 1
			end if
		end if
		oRec.close
	  else
	  visGuide = 1
	  end if
	
	if func = "showguide" then
	visGuide = 0
	end if
	
	if cint(visGuide) = 0 then 
	
	'********************************************
	'Henter de brugergrupper medarb. er medlem af 
	call hentbrugergrupper(usemrn)
	
	'************************************************************************************************
	'**** Guiden... Vis Kunder ********
	'************************************************************************************************
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
	<br><font class="stor-blaa">Guiden vælg dine aktive job:</font><br>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF">
	<form action="timereg.asp?func=updguide&FM_use_me=<%=usemrn%>" method="post" name="usejobguide">
 	<tr bgcolor="#5582D2">
		<td width="8" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="553" height="1" alt="" border="0"></td>
		<td align=right  valign=top><img src="../ill/tabel_top_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
		<td colspan="2" class='alt' valign=top>Dette er din <b>Guide</b> til at vælge hvilke <b>job</b> du arbejder på iøjeblikket.<br>
		Det betyder at du næste gang du logger ind i TimeOut, <b>kun</b> vil få vist job der er tilvalgt her i <b>Guiden dine aktive job</b>.<br></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td height="30" class='alt' width="450"><b>Kunder og job</b></td>
		<td class='alt' width="100"><b>Medtag Job?</b></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="2" align="right" valign=bottom>&nbsp;<a href="#" name="CheckAll" onClick="checkAll(document.usejobguide.FM_use_job)"><img src="../ill/alle.gif" border="0"></a>
		&nbsp;&nbsp;<a href="#" name="UnCheckAll" onClick="uncheckAll(document.usejobguide.FM_use_job)"><img src="../ill/ingen.gif" border="0"></a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	lastKid = 0
	strSQL = "SELECT id, jobstatus, jobnavn, jobnr, jobknr, Kid, Kkundenavn, Kkundenr, useasfak FROM job, kunder WHERE "& strSQLkri &" ORDER BY Kkundenavn, jobnavn"
	'Response.write (strSQL)	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
			
		if lastKid <> oRec("Kid") then
		%>
		<tr>
			<td valign="top" rowspan=2><img src="../ill/tabel_top.gif" width="1" height="15" alt="" border="0"></td>
			<td bgcolor="#5582D2" colspan="2"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align="right" rowspan=2><img src="../ill/tabel_top.gif" width="1" height="15" alt="" border="0"></td>
		</tr>
		<tr>
			<td colspan="2" bgcolor="#5582D2" class=alt>&nbsp;&nbsp;<%=oRec("Kkundenr")%>&nbsp;&nbsp;<b><%=oRec("Kkundenavn")%></b></td>
		</tr>
		<%else%>
		<tr>
			<td bgcolor="#5582D2" colspan="4"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		</tr>
		<%end if
		
		'**** Er jobbet allerede valgt?
		strSQL3 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE medarb = "& usemrn &" AND jobid = " & oRec("id")
		oRec3.open strSQL3, oConn, 3
		
		if not oRec3.EOF then
		selJob = "CHECKED"
		else
		selJob = ""
		end if
		
		oRec3.close 
		%>
		<tr>
			<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			<td height=20>&nbsp;&nbsp;&nbsp;<b><%=oRec("jobnr")%>&nbsp;&nbsp;<%=oRec("jobnavn")%></b></td>
			<td><input type="checkbox" name="FM_use_job" value="<%=oRec("id")%>" <%=selJob%>></td>
			<td valign="top" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<%
		
		
		lastKid = oRec("Kid")
		oRec.movenext
		wend
		oRec.close
		%>
		<tr bgcolor="#5582D2">
			<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
			<td colspan="2" valign="bottom"><img src="../ill/tabel_top.gif" width="553" height="1" alt="" border="0"></td>
			<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
		</tr>
		<!--
		<tr bgcolor="#D6DFF5">
			<td colspan="4" align="right"><br>Vis denne guide næste gang du logger på?&nbsp;Ja:<input type="checkbox" name="FM_visguide" value="j"></td>
		</tr>-->
		<tr bgcolor="#D6DFF5">
			<td colspan="4" align="right"><br>
			<input type="hidden" name="FM_use_job" id="FM_use_job" value="0">
			<input type="image" src="../ill/opdaterpil.gif"></td>
		</tr>
		</form>
		</table>
	
	<%
	'************************************************************************************************
	else
	'************************************************************************************************
	'*** Viser Timereg *******
	'************************************************************************************************
	%>
	<!--<div id="loading" name="loading" style="position:absolute; top:60; left:740; width:300; height:20; z-index:1000; visibility:visible; padding-left:4px; padding-top:5px;">
	<img src="../ill/info.gif" width="42" height="38" alt="" border="0"><b>Henter information...</b><br>
	<img src="../ill/loaderbar.gif" width="174" height="13" alt="" border="0">
	Denne handling kan tage op til 20 sek. alt efter din forbindelse.
	</div>-->


	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/calender.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	<!--#include file="inc/timereg_func_inc.asp"-->
	
<%

function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
end function

function SQLBless2(s)
		dim tmp2
		tmp2 = s
		tmp2 = replace(tmp2, ".", ",")
		SQLBless2 = tmp2
end function


%>
	<div id="sindhold" style="position:absolute; left:190; top:60; visibility:visible; width:530;">
	<table cellspacing="0" cellpadding="0" border="0" width="550">
	<tr>
    <td valign="top">
	<!-------------------------------Sideindhold------------------------------------->
	<img src="../ill/header_reg.gif" alt="" border="0">
	<hr align="left" width="330" size="1" color="#003399" noshade>
	<br><br><br><br><br><br><br><br><br><br><br>
	<font class=megetlillesort>Rediger din jobliste her:</font> <a href='timereg.asp?func=showguide&FM_use_me=<%=usemrn%>' class='vmenu'>Guiden dine aktive job&nbsp;<font color='darkred'>!</font></a>
		
	<%
	
	if strA3 = "on" then
				'***********************************************************************************
				'Henter mulighed for indtaste timer for andre medarbejdere.
				'***********************************************************************************
				if level <= 3 OR level = 6 then%> 
				<div id="usm" style="position:absolute; left:300; top:198; visibility:visible;">
				<form action="timereg.asp" name="selmed" method="POST">
				Indtast timer for:&nbsp;<select name="FM_use_me" style="font-family: arial,helvetica,sans-serif; font-size: 10px; width:120;">
				<%
					strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere ORDER BY Mnavn"
					oRec.open strSQL, oConn, 3
					
					while not oRec.EOF 
					'if oRec("Brugergruppe") = 4 OR cint(oRec("Mid")) = cint(session("Mid")) then
					if cint(oRec("Mid")) = cint(usemrn) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if%>
					<option value="<%=oRec("Mid")%>" <%=rchk%>><%=left(oRec("mnavn"), 16)%></option>
					<%
					'end if
					
					oRec.movenext
					wend
					oRec.close
				%></select>&nbsp;&nbsp;<input type="image" src="../ill/pillillexp_tp.gif"></form>
				</div><br>
				<%
				end if
				'*********************** Slut ********************************************************
	end if
		
		
		
	
	'Henter de brugergrupper medarb. er medlem af 
	call hentbrugergrupper(usemrn)

%>
	</td>
</tr>
</table>

<img src="../ill/blank.gif" alt="" height="4" width="1" border="0"><br>
<table cellspacing="0" cellpadding="0" border="0" width="530">
<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="74" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Job</font></td><td style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid; padding-left:6; padding-top:1; padding-right:6;" class=alt>&nbsp;</td>
	<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
</tr>
</table>

<%
'***** Henter uge dage i den valgte uge ****
%>
<table cellspacing="0" cellpadding="0" border="0" width="530">
<form action="timereg_db.asp?FM_use_me=<%=usemrn%>" method="POST" name="timereg">
<input type="Hidden" name="FM_start_dag" value="<%=strdag%>">
<input type="Hidden" name="FM_start_mrd" value="<%=strmrd%>">
<input type="Hidden" name="FM_start_aar" value="<%=straar%>">
<!--<input type="Hidden" name="searchstring" value="<=searchstring%>">-->
<input type="Hidden" name="FM_mnr" value="<%=strMnr%>">
<!--#include file="inc/timereg_dage_inc.asp"-->
<%
call dageDatoer()
%>
<input type="hidden" name="year" value="<%=strRegAar%>">
<input type="hidden" name="datoMan" value="<%=tjekdag(1)%>">
<input type="hidden" name="datoTir" value="<%=tjekdag(2)%>">
<input type="hidden" name="datoOns" value="<%=tjekdag(3)%>">
<input type="hidden" name="datoTor" value="<%=tjekdag(4)%>">
<input type="hidden" name="datoFre" value="<%=tjekdag(5)%>">
<input type="hidden" name="datoLor" value="<%=tjekdag(6)%>">
<input type="hidden" name="datoSon" value="<%=tjekdag(7)%>">
</table>

<table cellspacing="0" cellpadding="0" border="0" width="530">
<tr>
	<td colspan=2 bgcolor="#5582d2" style="border-left:1px #003399 solid; padding-left:2; padding-top:2;">
		<table border=0 cellspacing=1 cellpadding=0><tr>
		<td align=center width=14 bgcolor="#FFFFe1"><font class='megetlillesort'>M</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>T</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>O</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>T</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>F</td>
		<td align=center width=13 bgcolor="#FFDFDF"><font class='megetlillesort'>L</td>
		<td align=center width=13 bgcolor="#FFDFDF"><font class='megetlillesort'>S</td>
		</tr></table>
	</td>
	<td width=200 bgcolor="#5582d2" style="border-right:1px #003399 solid; padding-left:2; padding-top:2;">
		<font class='megetlillehvid'>Timer Forkalkuleret / <font class='megetlilleblaa'>Realiseret
	</td>
</tr>
<!--
<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
	<td colspan="3" bgcolor="#d6dff5" height=1><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
</tr>-->

<%
		
 
 '** Sætter jobnr på det sidst udskrevne job. ***
  'areJobPrinted = 0
  'strLastjobnr = 0
  'jobPrinted = "0#"
  'aktPrinted = "0#"
  intMnr = strMnr
  
  '** Array nr. ***
  de = 1

 '************************************************************************************************** 
  'Opbygger timereg SQL states
  '**************************************************************************************************
 	
	'*** Henter job fra guide ****
	Select case filterABsel
	case "a" 
			'*** Filter A ***
			if searchstring = 0 then 'Hvis Vis kunde ikke er brugt ell. der er valgt Alle!
				varUseJob = "("
				
				strSQL3 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE medarb = "& usemrn &""
				oRec3.open strSQL3, oConn, 3
				
				while not oRec3.EOF 
				varUseJob = varUseJob & " job.id = "& oRec3("jobid") & " OR "
				oRec3.movenext
				wend 
				
				oRec3.close 
				
				if varUseJob = "(" then
				varUseJob = " job.id = 0 AND "
				else
				varUseJob_len = len(varUseJob)
				varUseJob_left = varUseJob_len - 3
				varUseJob_use = left(varUseJob, varUseJob_left) 
				varUseJob = varUseJob_use & ") AND "
				end if
				
			else
				varUseJob = " "
			end if
	
	case "b"
		
		'*** Filter B ***
		varUseJob = "("
		
		'strSQL3 = "SELECT tmnr, tjobnr, tdato, timer, j.id AS jid FROM timer"_
		'&" LEFT JOIN  job j ON (jobnr = tjobnr)"_
		'&" RIGHT JOIN timereg_usejob ON (medarb = "& usemrn &" AND jobid = j.id)"_
		'&" WHERE tmnr = medarb GROUP BY jobnr ORDER BY tdato DESC LIMIT 0, 5" 
		
		strSQL3 = "SELECT tmnr, tjobnr, tdato, timer, j.id AS jid, jobnr, "_
		&" jobnavn FROM timereg_usejob tu"_
		&" LEFT JOIN job j ON (j.id = tu.jobid)"_ 
		&" LEFT JOIN timer ON (tmnr = 1 AND tjobnr = j.jobnr)"_
		&" WHERE  medarb = 1 ORDER BY tdato DESC LIMIT 0, 50" 
		
		'Response.write strSQL3 & "<br>"
		'Response.flush
		oRec3.open strSQL3, oConn, 3 
		x = 0
		isused = "#,"
		while not oRec3.EOF AND x < 5 
		
		tjekisused = 0
		tjekisused = instr(isused, "#"& oRec3("jid") &"#,")
		
		
		if len(oRec3("jid")) <> 0 AND cint(tjekisused) = 0 then
		varUseJob = varUseJob & " job.id = "& oRec3("jid") & " OR "
		isused = isused & "#"& oRec3("jid") &"#,"
		x = x + 1
		end if
		
		
		oRec3.movenext
		wend
		oRec3.close 
		
		varUseJob = varUseJob & " job.id = 0) AND "
	
	case "c"
	
		'*** Filter C ***
		varUseJob = "("
		
		strSQL3 = "SELECT jobnavn, j.id AS jid FROM timereg_usejob "_
		&" RIGHT JOIN job j ON (j.id = jobid) "_
		&" WHERE medarb = "& usemrn &" GROUP BY jobnr ORDER BY j.id DESC LIMIT 0, 8"
		oRec3.open strSQL3, oConn, 3 
		x = 0
		while not oRec3.EOF 'AND x < 5 
		
		if len(oRec3("jid")) <> 0 then
		varUseJob = varUseJob & " job.id = "& oRec3("jid") & " OR "
		end if
		
		x = x + 1
		oRec3.movenext
		wend
		oRec3.close 
		
		varUseJob = varUseJob & " job.id = 0) AND "	
		
	end select
	
	
	
		'**** Finder kunder til dd ****
		strKundeKri = " AND ("
		strSQL3 = "SELECT j.id, medarb, jobid, kid, jobknr FROM timereg_usejob"_
		&" RIGHT JOIN job j ON (j.id = jobid)"_
		&" LEFT JOIN kunder ON (kid = j.jobknr) WHERE medarb = "& usemrn &" GROUP BY kid"
		'response.write strSQL3
		'Response.flush
		oRec3.open strSQL3, oConn, 3
		
		while not oRec3.EOF 
		strKundeKri = strKundeKri & " kid = "& oRec3("kid") & " OR "
		oRec3.movenext
		wend 
		
		oRec3.close 
	
	
 	'**************************** Rettigheds tjeck aktiviteter **************************
  	strSQLkri3 = " aktiviteter.job = job.id AND aktstatus = 1 AND ( "
	for intcounter = 0 to f - 1  
  
	strSQLkri3 = strSQLkri3 &" aktiviteter.projektgruppe1 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.projektgruppe2 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe3 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe4 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe5 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe6 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe7 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe8 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe9 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe10 = "& brugergruppeId(intcounter) &" OR "
	
	next
  	
	'** Trimmer sql states ***
	strSQLkri3_len = len(strSQLkri3)
	strSQLkri3_left = strSQLkri3_len - 3
	strSQLkri3_use = left(strSQLkri3, strSQLkri3_left)  
  	strSQLkri3 = strSQLkri3_use &") "
	'*************************************************************************************
	
  	'*******************************************************************
	'Finder startdato på den viste uge for at tjekke
	'om der er en faktura af nyere dato så 
	'indtastningen af timer kan lukkes 
	
	if datepart("ww", tjekdag(7)) = 1 then
		'if datepart("y", tjekdag(1), 0) > 7 then
		'varTjDatoUS_start = dateadd("yyyy", -1, tjekdag(1))
		'else
		'varTjDatoUS_start = varTjDatoUS_start
		'end if	
		
		'*** rettet 26/12-2004 *** 
		varTjDatoUS_start = cdate(convertDate(tjekdag(1)))
	else
	varTjDatoUS_start = cdate(convertDate(tjekdag(1)))
	end if	
	
	use_varTjDatoUS_start = convertDateYMD(varTjDatoUS_start)
	
	'Response.write varTjDatoUS_start & "<br>"
	
	'********************************************************************
	
	%>
	<!--#include file="inc/timereg_setvalues_a_inc.asp"-->
	<%
	
	lastfakdato = ""
	dtimeTtidspkt = ""
	antalguidenjobs = 0
	
	'******************************************************************** 
	'Gennemløber de job som brugeren har adgang til. 
	'Udskrivning af side starter her
	'0 jobid
	'1 jobnr
	'2 jobnavn
	'4 kundenavn
	'5 aktid
	'6 aktnavn
	'10 timer
	'12 dato
	'15 kommentar
	'16 fakturerbar
	'17 budgettimer
	'18 ikkebudgettimer
	'19 aktbudgettimer
	'20	offentlig
	'21 Kundeid (kid)
	'22 Jobans1
	'23 Jobans2
	
	'************** Main SQL Call ******************************************
	strSQL = "SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn, aktiviteter.id AS akt,"_
	&" aktiviteter.navn AS navn, aktiviteter.job, fakdato, f.tidspunkt AS ftidspkt, timer.timer, "_
	&" timer.TAktivitetId, timer.tdato, timer.tjobnr, timer.tidspunkt AS ttidspkt, timer.Timerkom, "_
	&" aktiviteter.fakturerbar, job.budgettimer, job.ikkebudgettimer,"_
	&" aktiviteter.budgettimer AS aktbtimer, timer.offentlig, kid, jobans1, jobans2 FROM job, kunder"_
	&" LEFT JOIN aktiviteter ON ("& strSQLkri3 &")"_
	&" LEFT JOIN fakturaer f ON (f.jobid = job.id AND fakdato >= '"& use_varTjDatoUS_start &"')"_
	&" LEFT JOIN timer ON"_
	&" (TAktivitetId = aktiviteter.id AND Tmnr = "& intMnr &" AND (Tdato = '"& varTjDatoUS_son &"'"_
	&" OR Tdato = '"& varTjDatoUS_man &"'"_
	&" OR Tdato = '"& varTjDatoUS_tir &"'"_
	&" OR Tdato = '"& varTjDatoUS_ons &"'"_
	&" OR Tdato = '"& varTjDatoUS_tor &"'"_
	&" OR Tdato = '"& varTjDatoUS_fre &"'"_
	&" OR Tdato = '"& varTjDatoUS_lor &"')"_
	&") WHERE "& varUseJob & strSQLkri & " ORDER BY kkundenavn, jobnavn, id, navn" 
	
	'Response.write strSQL
	'Response.flush
	'oRec.cachesize = 150
	oRec.Open strSQL, oConn, 3
	
	Dim aTable1Values
	lastknavn = ""
	
	
	if not oRec.EOF Then
	aTable1Values = oRec.GetRows()
	oRec.close
	
	ctime = 0
	'*** Udskriver records ****
	Dim iRowLoop, iColLoop
	For iRowLoop = 0 to UBound(aTable1Values, 2)
	editok = 0
	
		
		'*** Hvis det er første gennemløb ****
		'*** Finder om der er flere akt. på det aktuelle job ****
		if iRowLoop > 0 then
		prevrecord = iRowLoop-1
		else
		prevrecord = iRowLoop
		end if
		
		'***** Hvis det er sidste gennemløb *******
		'*** Finder om der er flere job ****
		if iRowLoop < UBound(aTable1Values, 2) then
		nextrecord = iRowLoop + 1
		else
		nextrecord = "0"	
		end if
		
		if aTable1Values(1,iRowLoop) <> aTable1Values(1,(prevrecord)) OR iRowLoop = 0 then
		
		if len(aTable1Values(8,iRowLoop)) <> 0 then
		lastfakdato = aTable1Values(8,iRowLoop)
		lastFaktime = formatdatetime(aTable1Values(9,iRowLoop), 3)
		else
		lastfakdato = "1/13/2000"
		lastFaktime = "10:00:01 AM"
		end if
		
					if (aTable1Values(1,iRowLoop) <> aTable1Values(1,(prevrecord)) OR iRowLoop > 0) then 'AND len(aTable1Values(5,prevrecord)) <> 0 then				
					'Response.write "<tr><td colspan=9><b>Slut akt div -- "& aTable1Values(2,prevrecord) &" --</b></td></tr>"
					Response.write "</table></div></td></tr>"
					antalguidenjobs = antalguidenjobs + 1
					end if	
					
					if lastknavn <> aTable1Values(4,iRowLoop) then
					%>
					<tr>
						<td colspan=9 bgcolor="#8caae6" style="border-left:1px #003399 solid; solid; border-right:1px #003399 solid; padding-left:10; padding-top:3; border-top:1px #d6dff5 solid;" class=lysblaastor><b><%=aTable1Values(4,iRowLoop)%></b></td>
					</tr>
					<tr>
						<td bgcolor="#8caae6" height=4 colspan="9" style="border-left: 1px solid #003399; border-right: 1px solid #003399; border-bottom:1px #d6dff5 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					</tr>
					<%end if%>
					<tr>
						<td valign=top width=30 height=20 style="background-color: #5582D2; padding-top:2px; padding-left : 2px; border-left: 1px solid #003399">
						<%
						'*** henter timer på job denne uge *** 
						call timerdage(aTable1Values(1,iRowLoop), intMnr)
						%></td>
						
						<td width=380 valign=top style="background-color: #5582D2; padding-left : 4px;padding-right : 4px;">
						<%
						'** Er der nogen aktiviteter **
						strSQL2 = "SELECT count(id) AS antalAkt FROM aktiviteter WHERE job = "& aTable1Values(0,iRowLoop) &""
						oRec2.Open strSQL2, oConn, 3
						if not oRec2.EOF then
						antalAkt = oRec2("antalAkt")
						else
						antalAkt = 0
						end if
						oRec2.close
							
						if antalAkt <> 0 then%>
						<a href="javascript:expand('<%=iRowLoop%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=iRowLoop%>" id="Menub<%=iRowLoop%>">&nbsp;</a>&nbsp;<font class='lilleblaa'>(<%=antalAkt%>)</font>&nbsp;&nbsp;
						<%else%>
						&nbsp;
						<%end if%>
						
						
						<%
						if len(aTable1Values(2,iRowLoop)) > 30 then
						jobnavnthis = left(aTable1Values(2,iRowLoop), 30) &".."
						else
						jobnavnthis = aTable1Values(2,iRowLoop)
						end if
						
						
						if level = 1 then
						editok = 1
						else
								if cint(session("mid")) = aTable1Values(22,iRowLoop) OR cint(session("mid")) = aTable1Values(23,iRowLoop) OR (cint(aTable1Values(22,iRowLoop)) = 0 AND cint(aTable1Values(23,iRowLoop)) = 0) then
								editok = 1
								end if
						end if
						
						'if level <= 3 OR level = 6 then 
						if editok = 1 then
						%>
						<a href="jobs.asp?menu=job&func=red&id=<%=aTable1Values(0,iRowLoop)%>&int=1&rdir=treg" class="alt"><%=jobnavnthis%></a>
						<%else%>
						<font class='storsilver'><b><%=jobnavnthis%></b></font>
						<%end if%>
						
						&nbsp;<font class='lillehvid'><b>(<%=aTable1Values(1,iRowLoop)%>)</b>
						
						
						
						<!--&nbsp;
						<font class='lilleblaa'>(
						<if len(aTable1Values(4,iRowLoop)) > 20 then
						Response.write left(aTable1Values(4,iRowLoop), 20) &".."
						else
						Response.write aTable1Values(4,iRowLoop)
						end if%>
						)</font>-->
							
						<%
						'* finder antal brugte timer på job ***
						strSQL2 = "SELECT sum(timer) AS sumtimer FROM timer WHERE tjobnr = "& aTable1Values(1,iRowLoop) &" AND tfaktim <> 5 ORDER BY timer"
						oRec2.open strSQL2, oConn, 3
						if not oRec2.EOF then 
						timerbrugtthis = oRec2("sumtimer")
						end if
						oRec2.close
						
						if len(timerbrugtthis ) > 0 then
						timerbrugtthis = timerbrugtthis 
						else
						timerbrugtthis = 0
						end if
						
						%>
						</td>
						<td width=150 align=right style='padding-right:3px; background-color: #5582D2; border-right: 1px solid #003399;'>
						
						<font class='megetlillehvid'><%=formatnumber(aTable1Values(17,iRowLoop) + aTable1Values(18,iRowLoop), 2)%> / <font class='megetlilleblaa'><%= formatnumber(timerbrugtthis, 2) %>
						<input type="checkbox" name="FM_flyttilguide" value="<%=aTable1Values(0,iRowLoop)%>">
						<%
						if lto = "dencker" OR lto = "demo" OR lto = "intranet" then%>
						<a href="javascript:popUp('materialer_indtast.asp?id=<%=aTable1Values(0,iRowLoop)%>','600','620','250','20');" target="_self" class=alt><img src="../ill/ac0038-16_blaa.gif" width="18" height="18" border="0" alt="Materiale forbrug / udlæg"></a>
						<%end if%>
						
						</td>
					</tr>
						
							
		<%
		'**** Starter div uanset om der findes aktiviteter ***
		%>
		<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
			<td colspan="3" bgcolor="#8caae6" height=1><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
		<td colspan="3" width="530" bgcolor="#5582d2"><div ID="Menu<%=iRowLoop%>" Style="position: relative; display: none; width:530;">
		<table cellspacing="1" cellpadding="0" border="0" width=100% bgcolor="#ffffff">
		
		<%				
		end if
		
		'*** Hvis aktid <> 0 ****
		if len(aTable1Values(5,iRowLoop)) <> 0 then
		
			'**** Finder ugens timer på aktid ***%>
			<!--#include file="inc/timereg_setvalues_b_inc.asp"-->
			<%	
				
				'************** Aktivitetsnavn og timeforbrug i alt *********************''	
				if aTable1Values(5, iRowLoop) <> aTable1Values(5, nextrecord) OR aTable1Values(2, iRowLoop) <> aTable1Values(2, nextrecord) OR nextrecord = "0" then
					%>
					<input type="hidden" name="FM_jobnr_<%=iRowLoop%>" value="<%=aTable1Values(1,iRowLoop)%>">
					<input type="hidden" name="FM_Aid_<%=iRowLoop%>" value="<%=aTable1Values(5,iRowLoop)%>">
					
					<tr bgcolor=#D6DFF5>
						<td width=120 style='padding-left:2px; padding-top:2px;'>
						<%
						if len(aTable1Values(6, iRowLoop)) > 40 then
							Response.write "<font class=lille-kalender>"& left(aTable1Values(6, iRowLoop), 40) &"..</font>"
							else
							Response.write "<font class=lille-kalender>"& aTable1Values(6, iRowLoop) &"</font>"
						end if
						
						if aTable1Values(16,iRowLoop) = 2 then
						%>
						<img src="../ill/blank.gif" width="60" height="1" alt="" border="0">(Km)
						<%
						end if
						
						'* finder antal brugte timer på akt ***
							strSQL3 = "SELECT sum(timer) AS sumakttimer FROM timer WHERE taktivitetid = "& aTable1Values(5,iRowLoop) &" ORDER BY timer"
							oRec3.open strSQL3, oConn, 3
							if not oRec3.EOF then 
							timerbrugtAktthis = oRec3("sumakttimer")
							end if
							oRec3.close
							
							if len(timerbrugtAktthis) > 0 then
							timerbrugtAktthis = timerbrugtAktthis
							else
							timerbrugtAktthis = 0
							end if
						
						Response.write "<br><font class='megetlillesilver'>" & formatnumber(aTable1Values(19,iRowLoop), 2) &" / "& formatnumber(timerbrugtAktthis, 2)&"</font>"
						%></td>	
					<%
				end if
			
			
			
			
			if aTable1Values(5, iRowLoop) <> aTable1Values(5, nextrecord) OR aTable1Values(2, iRowLoop) <> aTable1Values(2, nextrecord) OR nextrecord = "0" then
				'**** Tjekker om der er fundet timeindtastninger på de enkelte dage ***
				%>
				<!--#include file="inc/timereg_setvalues_c_inc.asp"-->
				<%				
			 	'**** Henter feltfarver alt efter om der findes fakturaer på jobbet.
				call fakfarver()
				
				
					
				'*** Udskriver timefelter til siden ***********************	
				for w = 1 to 7
				
				select case w
				case 1
				dtype = "man"
				bgcolUse = fmbg_man
				val = manTimerVal(manRLoop)
				maxl = maxl_man
				case 2
				dtype = "tir"
				bgcolUse = fmbg_tir
				val = tirTimerVal(tirRLoop)
				maxl = maxl_tir
				case 3
				dtype = "ons"
				bgcolUse = fmbg_ons
				val = onsTimerVal(onsRLoop)
				maxl = maxl_ons
				case 4
				dtype = "tor"
				bgcolUse = fmbg_tor
				val = torTimerVal(torRLoop)
				maxl = maxl_tor
				case 5
				dtype = "fre"
				bgcolUse = fmbg_fre
				val = freTimerVal(freRLoop)
				maxl = maxl_fre
				case 6
				dtype = "lor"
				bgcolUse = fmbg_lor
				val = lorTimerVal(lorRLoop)
				maxl = maxl_lor
				case 7
				dtype = "son"
				bgcolUse = fmbg_son
				val = sonTimerVal(sonRLoop)
				maxl = maxl_son
				end select
				%>
				<td align="center" id="specialtest_<%=dtype%>_<%=iRowLoop%>" name="specialtest_<%=dtype%>_<%=iRowLoop%>" style="border:2px #d6dff5 solid; width:50px;" bgcolor="<%=bgcolUse%>">
				<a href="#" onClick="specialrp('<%=iRowLoop%>','<%=SQLBless2(val)%>','<%=maxl%>','<%=dtype%>')">
				<%if len(val) <> 0 then%>
				<%=SQLBless2(val)%>
				<%else%>
				<font color=#cccccc>0</font>
				<%end if%>
				</a>
				<%call showcommentfunc(w)%>
				</td>
				<%next
				
				Response.flush%>
				
				
				
				
				
				
			</tr>
			
			
				<%'kommentarer%>
				<input type="hidden" name="FM_kom_man_<%=iRowLoop%>" id="FM_kom_man_<%=iRowLoop%>" value="<%=manKomm%>">
				<input type="hidden" name="FM_kom_tir_<%=iRowLoop%>" id="FM_kom_tir_<%=iRowLoop%>" value="<%=tirKomm%>">
				<input type="hidden" name="FM_kom_ons_<%=iRowLoop%>" id="FM_kom_ons_<%=iRowLoop%>" value="<%=onsKomm%>">
				<input type="hidden" name="FM_kom_tor_<%=iRowLoop%>" id="FM_kom_tor_<%=iRowLoop%>" value="<%=torKomm%>">
				<input type="hidden" name="FM_kom_fre_<%=iRowLoop%>" id="FM_kom_fre_<%=iRowLoop%>" value="<%=freKomm%>">
				<input type="hidden" name="FM_kom_lor_<%=iRowLoop%>" id="FM_kom_lor_<%=iRowLoop%>" value="<%=lorKomm%>">
				<input type="hidden" name="FM_kom_son_<%=iRowLoop%>" id="FM_kom_son_<%=iRowLoop%>" value="<%=sonKomm%>">
				
				<%'Komm. offentlig tilgængelig for kunde%>
				<%
				if len(manKomm_off) <> 0 then
				manKomm_off = manKomm_off
				else
				manKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_man_<%=iRowLoop%>" id="FM_off_man_<%=iRowLoop%>" value="<%=manKomm_off%>">
				<%
				if len(tirKomm_off) <> 0 then
				tirKomm_off = tirKomm_off
				else
				tirKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_tir_<%=iRowLoop%>" id="FM_off_tir_<%=iRowLoop%>" value="<%=tirKomm_off%>">
				<%
				if len(onsKomm_off) <> 0 then
				onsKomm_off = onsKomm_off
				else
				onsKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_ons_<%=iRowLoop%>" id="FM_off_ons_<%=iRowLoop%>" value="<%=onsKomm_off%>">
				<%
				if len(torKomm_off) <> 0 then
				torKomm_off = torKomm_off
				else
				torKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_tor_<%=iRowLoop%>" id="FM_off_tor_<%=iRowLoop%>" value="<%=torKomm_off%>">
				
				<%
				if len(freKomm_off) <> 0 then
				freKomm_off = freKomm_off
				else
				freKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_fre_<%=iRowLoop%>" id="FM_off_fre_<%=iRowLoop%>" value="<%=freKomm_off%>">
				<%
				if len(lorKomm_off) <> 0 then
				lorKomm_off = lorKomm_off
				else
				lorKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_lor_<%=iRowLoop%>" id="FM_off_lor_<%=iRowLoop%>" value="<%=lorKomm_off%>">
				
				<%
				if len(sonKomm_off) <> 0 then
				sonKomm_off = sonKomm_off
				else
				sonKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_son_<%=iRowLoop%>" id="FM_off_son_<%=iRowLoop%>" value="<%=sonKomm_off%>">
				
					
			
			<%
			
			manKomm = ""
			tirKomm = ""
			onsKomm = ""
			torKomm = ""
			freKomm = ""
			lorKomm = ""
			sonKomm = ""
			
			
			manKomm_off = 0
			tirKomm_off = 0
			onsKomm_off = 0
			torKomm_off = 0
			freKomm_off = 0
			lorKomm_off = 0
			sonKomm_off = 0
			
			
			manRLoop = ""
			tirRLoop = ""
			onsRLoop = ""
			torRLoop = ""
			freRLoop = ""
			lorRLoop = ""
			sonRLoop = ""
			
			Response.flush
			
			end if
		de = iRowLoop
		end if 
		lastknavn = aTable1Values(4,iRowLoop)
	Next 
	
	'*** afslutter den sidtste div 
	if iRowLoop > 0 then 
	'Response.write "<tr><td><b>final akt div</b></td></tr>"%>
	</table></div></td></tr>
	<tr>
		<td bgcolor="#5582d2" height=5 colspan="9" style="border-left: 1px solid #003399; border-right: 1px solid #003399; border-bottom: 1px solid #003399;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td colspan=9 align=right style="padding-right:2; padding-top:2;"><font class="megetlillesort">Marker job for at fjerne fra time-registreringen og overføre til "Guiden dine aktive job!".</td>
	</tr></table>
	<%end if%>
	
	<%
	'********************** slut gennemløb af eksterne job *****************************
	antal_de = de 
	%>
	<input type="hidden" name="FM_de" value="<%=antal_de%>">
	<%
	
	'*** Hvis der oRec.EOF true => Ikke adgang til nogen job ***************************
	else
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="530">
	<tr>
		<td colspan=3 style='border-left: 1px solid #003399; border-right: 1px solid #003399; border-bottom: 1px solid #003399; background-color: #5582D2; padding-left : 38px;padding-right : 4px;' class='alt'>		
		<br>Du er <b>ikke</b> tildelt adgang til nogen eksterne-job.<br>
		Eller der er ikke nogen aktive job på den valgte kunde.<br><br>
		Kontakt din projektleder/ system administrator.<br>
		Eller brug <a href='timereg.asp?func=showguide' class='alt'><u>Guiden dine aktive job&nbsp;<font color='darkred'>!</font></u></a> for at få vist de job der er aktive på din profil.<br>&nbsp;</td>
	</tr>
	</table><%
	oRec.close
	end if
	
	
	
	
'*** Antal eksterne / interne ***	
di = antal_de + 1

'*** Antal interne job til timreg_db.asp
antal_di = di - 1
%>
<input type="hidden" name="FM_di" value="<%=antal_di%>">
<input type="hidden" name="FM_strweek" value="<%=strWeek%>">


<table width='530' border='0' cellspacing='1' cellpadding='0'>
<tr>
   
	<td width=300>&nbsp;
	<%
	'**** Tjekker om medarbejder har slået smileyordningen til ****
	strSQLord = "SELECT mid, smilord FROM medarbejdere WHERE mid = "& usemrn
	oRec.open strSQLord, oConn, 3 
	
	if not oRec.EOF then
		if cint(oRec("smilord")) = 1 then
		smilaktiv = smilaktiv
		else
		smilaktiv = 0
		end if
	end if
	oRec.close 
	
	
	if cint(smilaktiv) = 1 then
	
		showAfsuge = 1
		sidstedagiuge = year(tjekdag(7))&"/"&month(tjekdag(7))&"/"&day(tjekdag(7))
		thisMid = usemrn
		strSQLafslut = "SELECT status, afsluttet FROM ugestatus WHERE uge = '"& sidstedagiuge &"' AND mid = "& thisMid
		oRec.open strSQLafslut, oConn, 3 
		while not oRec.EOF 
			showAfsuge = 0
			cdAfs = oRec("afsluttet")
		oRec.movenext
		wend
		oRec.close 
		
		%>
		<br><b>Afslut uge?</b><br>
		<%
		if showAfsuge = 1 then%>
				<%if datepart("ww", useDate, 2, 3) <= datepart("ww", now, 2, 3) then%>
				<input type="checkbox" name="FM_afslutuge" id="FM_afslutuge" value="1">
				<input type="hidden" name="FM_afslutuge_sidstedag" id="FM_afslutuge_sidstedag" value="<%=tjekdag(7)%>">
				Marker denne når du vil afslutte uge: <b><%=datepart("ww", useDate, 2, 3)%></b>.<br>
				Ved markering inden søndag kl. 24:00 vil denne uge betragtes som rettidigt afsluttet.
				<%else%>
				Ugen kan først afsluttes fra uge: <b><%=datepart("ww", useDate, 2, 3)%></b> 
				<%end if%>
		<%else%>
		Uge: <b><%=datepart("ww", useDate, 2, 3)%></b> afsluttet: <b><%=weekdayname(weekday(cdAfs))%> d. <%=formatdatetime(cdAfs, 0)%> </b>
		<%end if%>
	
	<%end if '** smilaktiv %>
	</td>
	
	<td align="right">
	<br><input type="image" src="../ill/frem.gif"></td>
	</tr>
</table> 
	
	


<!--
</td>
</tr>
</table>-->




	<%'***** kommentar layer *** 
	
	if instr(Request.ServerVariables("HTTP_USER_AGENT"), "Firefox") <> 0 then
	hgt = 170
 	else
	hgt = 140
	end if
	%>
	<div id="kom" name="kom" Style="position: absolute; display: none; left:20; top:170; z-index:2994; width:500; height:<%=hgt%>; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">
	Kommentar:
	<input type="hidden" name="daytype" id="daytype" value="son">
	<input type="hidden" name="rowcounter" id="rowcounter" value="0">
	<img src="../ill/blank.gif" width="190" height="1" alt="" border="0">
	(maks. 255 karakterer)&nbsp;
	<input type="text" name="antch" id="antch" size="3" maxlength="4"><br>
	<textarea cols="50" rows="6" id="FM_kom" name="FM_kom" onKeyup="antalchar();"></textarea><br>
	Skjul kommentar på kundeside:
	<select name="FM_off" id="FM_off" style="font-size:9px;">
	<option value="0" SELECTED>Nej</option>
	<option value="1">Ja</option>
	</select><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">
	<a href="#" onclick="closekomm();" class=vmenu>Gem og Luk&nbsp;</a><br>&nbsp;
	</div>

<br>
<br>
<br><br>
</div>



<% '***************************** Smiley ******************************
	
	if cint(smilaktiv) = 1 then '** Smiley aktiv
	
	'**** Sur Smiley overruler ******
	antalAfsDato = 0
	
	'*** startdato: 1 DEC 2005 ***
	if year(useDate) <= 2005 then
	surDatoSQLSTART = "2005/11/1" '12
		if lto = "execon" then
		antalAfsDato = 47 '47 '47 Sidste hele uge inden man går ind i december.
		else
		antalAfsDato = 43 '43 Sidste hele uge inden man går ind i december.
		end if
	else
	surDatoSQLSTART = year(useDate)&"/1/1"
	antalAfsDato = 0
	end if
	
	surDatoSQLEND = year(useDate)&"/"&month(useDate)&"/"&day(useDate)
	strSQL3 = "SELECT uge, count(afsluttet) AS antalafs FROM ugestatus "_
	&" WHERE mid = "& thisMid &" AND uge BETWEEN '"& surDatoSQLSTART &"' AND '"& surDatoSQLEND &"'"_
	&" GROUP BY mid ORDER BY afsluttet DESC"
	
	'Response.write  strSQL3
	'Response.flush
	oRec3.open strSQL3, oConn, 3 
	if not oRec3.EOF then
		antalAfsDato = antalAfsDato + oRec3("antalafs")
	end if
	oRec3.close  
	
	useDateWeek = datepart("ww", useDate, 2, 3)
	'Response.write "lastAfsDatoWeek: " & antalAfsDato & " = "&  useDateWeek
	
	
	'** Glade ***
	if (antalAfsDato + 1) >= useDateWeek then
	surSmil = 0
	
				'*** Viser Smiley ***
				if weekday(useDate, 2) = 1 then
				periodeDage = 6
				minusenuge = 1
				else
				periodeDage = datediff("d", mandag, useDate)
				minusenuge = 0
				end if
				
				
				smileyok = 1
				antaldage = 1
				
				if weekday(useDate, 2) = 1 then '** Mandag altid glad Smiley
					
					antaldage = antaldage
					smileyok = smileyok
				
				else
				
							for i = 1 to periodeDage + 1
								
							thisDato = dateadd("ww", -(minusenuge), tjekdag(i))
							'Response.write thisDato & "# "
							thisDatoSQL = year(thisDato)&"/"&month(thisDato)&"/"&day(thisDato)
								
							select case weekday(thisDato, 2)
							case 7
							normdag = "normtimer_son"
							case 1
							normdag = "normtimer_man" 
							case 2
							normdag = "normtimer_tir"
							case 3
							normdag = "normtimer_ons"
							case 4
							normdag = "normtimer_tor"
							case 5
							normdag = "normtimer_fre"
							case 6
							normdag = "normtimer_lor"
							case else
							normdag = "normtimer_man"
							end select
				
							
							timerThis = 0
							strSQL = "SELECT medarbejdertype, "& normdag &" AS timer FROM medarbejdere m"_
							&" LEFT JOIN medarbejdertyper t ON (t.id = m. medarbejdertype)"_
							&" WHERE m.mid = " & thisMid
							
							'Response.write strSQL & "<br><br>"
							
							oRec.open strSQL, oConn, 3 
							if not oRec.EOF then
							timerThis = oRec("timer")
									
									if timerThis <> 0 then
										
										strSQL2 = "SELECT sum(timer) AS stimer FROM timer WHERE tdato = '"& thisDatoSQL &"' AND tmnr = "& thisMid &" GROUP BY tdato"
										'Response.write strSQL2 & "<br>"
										oRec2.open strSQL2, oConn, 3 
										while not oRec2.EOF 
										
										if oRec2("stimer") > 0 then
										smileyok = smileyok + 1
										end if
										
										oRec2.movenext
										wend
										oRec2.close 
									else
										smileyok = smileyok + 1
									end if
							oRec.movenext
							end if
							oRec.close 
							
							antaldage = antaldage + 1
							next
				end if
	else
	surSmil = 1
	end if
	
	if instr(Request.ServerVariables("HTTP_USER_AGENT"), "Firefox") <> 0 then
	wdt = 127
 	else
	wdt = 160
	end if
	%>
	<div id="smil" name="smil" Style="position:absolute; visibility:visible; display:; left:10px; top:390px; z-index:2000; width:<%=wdt%>px; background-color:#ffffff; border:1px #8caae6 solid; padding:15px;">
	<%
	'Response.write Request.ServerVariables("HTTP_USER_AGENT") 
	
	'Response.write "<br>(antalAfsDato + 1) >= useDateWeek "& (antalAfsDato + 1) &">="& useDateWeek 
	'Response.write "-->(Glad/mellem)<br>antaldage "& (antaldage - 1)
	'Response.write " <= Smileyok "& smileyok & " == Glad "
	'Response.write "<br><br>SurSmil "& surSmil & "<br><br>"
	
	smileyRnd = right(cint(Second(now)), 1)
	
	select case smileyRnd
	case 0, 1 
	smilVal = 1
	case 2, 3
	smilVal = 2
	case 4, 5
	smilVal = 3
	case 6, 7
	smilVal = 4
	case else
	smilVal = 5
	end select
			
		%>
		 
		<%=left(weekdayname(weekday(useDate)), 3) & " d. "& formatdatetime(useDate, 1) %><br>
					
		<%
		if surSmil = 1 then
				%>
				<b>Sur Smiley.</b> <br>
				<img src="../ill/sur_<%=smilVal%>.gif" alt="" border="0"><br>
				<font class=megetlillesort>En eller flere uger, frem til indeværende uge, er ikke afsluttet.<br></font>
				<%
		else
		
				if smileyok >= (antaldage-1) then
				%>
				<b>Glad Smiley.</b> <br>
				<img src="../ill/gladsmil_<%=smilVal%>.gif" alt="" border="0"><br>
				<!--<font class=megetlillesort>Du mangler ikke nogen timeregistreringer i denne uge.<br>-->
				<%
				else
				%>
				<b>Mellemfornøjet Smiley.</b> <br>
				<img src="../ill/mellemsmil_<%=smilVal%>.gif" alt="" border="0"><br>
				<font class=megetlillesort>En eller flere dage i indeværende uge er ikke udfyldt.<br></font>
				<%end if
		end if%>
		
		
		<font class=megetlillesort><br>Nb: Smileyordning kan slås til og fra i medarbejder-profilen.</font>
		</div>
	
	<%end if '*** Smiley aktiv%>
	

<%'***** ressourcetimer + ugetotaler *** %>
<div name=ressourcetimer id=ressourcetimer style="position:absolute; left:45; top:180; visible:hidden; display:none; border:1px #003399 solid; overflow:auto; background-color:#FFFFF1; filter:alpha(opacity=95); z-index:10000; padding-left:3;">
<!--#include file="inc/timereg_restimer_inc.asp"-->
</div>

<%'***** dagsseddel *** %>
<!--#include file="inc/timereg_dagsseddel_inc.asp"-->



<%'***** Kunde kriterie *** %>
<div name=kundesel id=kundesel style="position:absolute; left:190; top:117; visibility:visible; display:; z-index:200; border:1px #003399 solid; padding:2px; background-color:#FFFFFF;">
<%
ketypeKri = " ketype <> 'e' "
'actionLine = "timereg.asp?menu=timereg"
'actionLine = "timereg_db.asp?menu=timereg"
'name1 = "searchstring"
imgsearch = "../ill/v-menu_soeg.gif"
writethis = "(Dine aktive job)"
'hiddenfield1 = "<input type='hidden' name='FM_use_me' value='"& request("FM_use_me") &"'>"
	
	'If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	'hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_dag' value='"&request("FM_start_dag")&"'>"
	'hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_mrd' value='"&request("FM_start_mrd")&"'>"
	'hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_aar' value='"&request("FM_start_aar")&"'>"
	'else
	'hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_dag' value='"&request("strdag")&"'>"
	'hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_mrd' value='"&request("strmrd")&"'>"
	'hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_aar' value='"&request("straar")&"'>"
	'end if

	strKundeKri = strKundeKri &" kid = 0)"
	
	select case filterABsel
	case "a"
	aSel = "CHECKED"
	bSel = ""
	cSel = ""
	case "b"
	aSel = ""
	bSel = "CHECKED"
	cSel = ""
	case "c"
	aSel = ""
	bSel = ""
	cSel = "CHECKED"
	end select
	
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="530">
	<tr><td colspan=3><b>Vælg hvilke job der skal vises på listen udfra filter-kriterie A, B eller C.</b></td></tr>
	<tr>
		<td><input type="radio" name="filterA_B" value="a" <%=aSel%>> <b>A)</b> <u>Kundefilter.</u>
		&nbsp;&nbsp;<select name="searchstring" size="1" style="font-size : 9px; width:185px;">
		<option value="0">Alle <%=writethis%></option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(searchstring) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=left(oRec("Kkundenavn"), 18)%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select><br>
		<input type="radio" name="filterA_B" value="b" <%=bSel%>> <b>B)</b>  <u>Seneste 5 job du har registreret timer på.</u><br>
		
		<input type="radio" name="filterA_B" value="c" <%=cSel%>> <b>C)</b> <u>De 8 senest oprettede job.</u><br></td>
	</tr>
	<tr>
		<td style="padding-top:4px;">
		<input type="image" src="../ill/brug-filter-kunde.gif" border="0"></td>
		<!--</form>-->
	</tr>
	<tr><td style="padding-top:5px;"><font class=megetlillesort>Når du vælger filter indlæses timer, inden det nye filter eksekveres. 
		Følg med i dine timeregistreringer i kalenderen til højre. Kun job der er slået "til" i "Guiden dine aktive job!" vises uanset om filter A, B eller C benyttes.
	</td></tr>
	</table>
	
</div>

</form>






<!--<script>
document.all["loading"].style.visibility = "hidden";
</script>-->
<%end if 'Viskunder Wiz eller Timereg
end if 'user session
%>
<!--#include file="../inc/regular/footer_inc.asp"-->





