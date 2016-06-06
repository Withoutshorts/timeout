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
	
	if len(request("searchstring")) = "0" then
	searchstring = 0
	else
	searchstring = cint(request("searchstring"))
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
	<form action="timereg.asp?func=updguide" method="post" name="usejobguide">
 	<tr bgcolor="#5582D2">
		<td width="8" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="553" height="1" alt="" border="0"></td>
		<td align=right  valign=top><img src="../ill/tabel_top_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="130" alt="" border="0"></td>
		<td colspan="2" class='alt'>Dette er din guide til at vælge hvilke <b>job</b> du arbejder på iøjeblikket.<br><br>
		Det betyder at du næste gang du logger ind kun vil få vist job der tilhører disse kunder.<br>
		Dette vil give dig et forbedret overblik over dine job når du skal indtaste timer,<br>
		samt en øget performance på timeregistrerings siden.<br><br>
		Du kan selvfølgelig altid ændre disse indstillinger ved at klikke på "Guiden dine aktive job !"<br><br>
		Herunder følger alle de kunder der har aktive job, hvor de projektgrupper du er medlem af har adgang til:<br></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="130" alt="" border="0"></td>
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
			<td valign="top" rowspan=2><img src="../ill/tabel_top.gif" width="1" height="20" alt="" border="0"></td>
			<td bgcolor="#5582D2" colspan="2"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align="right" rowspan=2><img src="../ill/tabel_top.gif" width="1" height="20" alt="" border="0"></td>
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
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
			<td height="20">&nbsp;&nbsp;&nbsp;<b><%=oRec("jobnr")%>&nbsp;&nbsp;<%=oRec("jobnavn")%></b></td>
			<td><input type="checkbox" name="FM_use_job" value="<%=oRec("id")%>" <%=selJob%>></td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
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
			<td colspan="4" align="right"><br><input type="image" src="../ill/opdaterpil.gif"></td>
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
	<div id="loading" name="loading" style="position:absolute; top:60; left:740; width:300; height:20; z-index:1000; visibility:visible; padding-left:4px; padding-top:5px;">
	<!--<img src="../ill/info.gif" width="42" height="38" alt="" border="0">--><b>Henter information...</b><br>
	<img src="../ill/loaderbar.gif" width="174" height="13" alt="" border="0">
	<!--Denne handling kan tage op til 20 sek. alt efter din forbindelse.-->
	</div>
	
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


%>
	<div id="sindhold" style="position:absolute; left:190; top:60; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="550">
	<tr>
    <td valign="top">
	<!-------------------------------Sideindhold------------------------------------->
	<br>
	<img src="../ill/header_reg.gif" alt="" border="0"><hr align="left" width="750" size="1" color="#000000" noshade>
	<%
	if strA3 = "A3 Skift medarbejder on" then
				'***********************************************************************************
				'Henter mulighed for indtaste timer for andre medarbejdere.
				'***********************************************************************************
				if level <= 3 OR level = 6 then%> 
				<div id="usm" style="position:absolute; left:240; top:65; visibility:visible;">
				<form action="timereg.asp" name="selmed" method="POST">
				Indtast timer for (skift medarbejder):&nbsp;<select name="FM_use_me" style="font-family: arial,helvetica,sans-serif; font-size: 10px;">
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
<br>



<table cellspacing="0" cellpadding="0" border="0" width="530">
<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="74" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Eksterne job</font></td><td style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid; padding-left:6; padding-top:1; padding-right:6;" class=alt><img src="../ill/eksternt_job_trans.gif" width="15" height="15" alt="" border="0"></td>
	<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
</tr>
</table>

<%
'***** Henter uge dage i den valgte uge ****
%>
<table cellspacing="0" cellpadding="0" border="0" width="530">
<form action="timereg_db.asp" method="POST" name="timereg">
<input type="Hidden" name="FM_start_dag" value="<%=strdag%>">
<input type="Hidden" name="FM_start_mrd" value="<%=strmrd%>">
<input type="Hidden" name="FM_start_aar" value="<%=straar%>">
<input type="Hidden" name="searchstring" value="<%=searchstring%>">
<input type="Hidden" name="FM_mnr" value="<%=strMnr%>">
<!--#include file="inc/timereg_dage_inc.asp"-->
<%
call dageDatoer()
%>
<input type="hidden" name="year" value="<%=strRegAar%>">
<input type="hidden" name="datoSon" value="<%=tjekdag(1)%>">
<input type="hidden" name="datoMan" value="<%=tjekdag(2)%>">
<input type="hidden" name="datoTir" value="<%=tjekdag(3)%>">
<input type="hidden" name="datoOns" value="<%=tjekdag(4)%>">
<input type="hidden" name="datoTor" value="<%=tjekdag(5)%>">
<input type="hidden" name="datoFre" value="<%=tjekdag(6)%>">
<input type="hidden" name="datoLor" value="<%=tjekdag(7)%>">
</table>

<table cellspacing="0" cellpadding="0" border="0" width="530">
<tr>
	<td colspan=2 bgcolor="#5582d2" style="border-left:1px #003399 solid; padding-left:2; padding-top:2;">
		<table border=0 cellspacing=1 cellpadding=0><tr>
		<td align=center width=14 bgcolor="#FFFFe1"><font class='megetlillesort'>S</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>M</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>T</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>O</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>T</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>F</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>L</td>
		</tr></table>
	</td>
	<td bgcolor="#5582d2" width=120 style="border-right:1px #003399 solid; padding-left:2; padding-top:2;">
		<font class='megetlillehvid'>Timer tildelt / <font class='megetlilleblaa'>Forbrugt
	</td>
</tr>

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
	
	
 	'**************************** Rettigheds tjeck aktiviteter **************************
  	strSQLkri3 = " aktiviteter.job = job.id AND aktstatus = 1 AND ( "
	for intcounter = 0 to f - 1  
  
	strSQLkri3 = strSQLkri3 &" aktiviteter.projektgruppe1 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.projektgruppe2 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe3 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe4 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe5 = "& brugergruppeId(intcounter) &" OR "
	
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
		if datepart("y", tjekdag(1), 0) > 7 then
		varTjDatoUS_start = dateadd("yyyy", -1, tjekdag(1))
		else
		varTjDatoUS_start = varTjDatoUS_start
		end if	
	else
	varTjDatoUS_start = cdate(convertDate(tjekdag(1)))
	end if	
	use_varTjDatoUS_start = convertDateYMD(varTjDatoUS_start)
	
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
	'17 budgettimer
	'18 ikkebudgettimer
	'19 aktbudgettimer
	
	'************** Main SQL Call ******************************************
	strSQL = "SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn, aktiviteter.id AS akt,"_
	&" aktiviteter.navn AS navn, aktiviteter.job, fakdato, f.tidspunkt AS ftidspkt, timer.timer, "_
	&" timer.TAktivitetId, timer.tdato, timer.tjobnr, timer.tidspunkt AS ttidspkt, timer.Timerkom,"_
	&" aktiviteter.fakturerbar, job.budgettimer, job.ikkebudgettimer,"_
	&" aktiviteter.budgettimer AS aktbtimer FROM job, kunder"_
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
	&") WHERE "& varUseJob & strSQLkri & " ORDER BY jobnavn, id, navn" 
	
	'Response.write (strSQL)
	'oRec.cachesize = 150
	oRec.Open strSQL, oConn, 3
	
	Dim aTable1Values
	if not oRec.EOF Then
	aTable1Values = oRec.GetRows()
	oRec.close
	
	ctime = 0
	'*** Udskriver records ****
	Dim iRowLoop, iColLoop
	For iRowLoop = 0 to UBound(aTable1Values, 2)
	
		
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
					
					%>
					<tr>
						<td valign=top style='background-color: #5582D2; padding-top:2px; padding-left : 2px; border-left: 1px solid #003399'>
						<%
						'*** henter timer på job denne uge *** 
						call timerdage(aTable1Values(1,iRowLoop), intMnr)
						%></td>
						<td valign=top height=20 style='background-color: #5582D2; padding-left : 4px;padding-right : 4px; width:380px;'>
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
						<a href="javascript:expand('<%=iRowLoop%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=iRowLoop%>">&nbsp;</a>&nbsp;<font class='lilleblaa'>(<%=antalAkt%>)</font>&nbsp;&nbsp;
						<%else%>
						&nbsp;
						<%end if%>
						
						<font class='lillehvid'><b><%=aTable1Values(1,iRowLoop)%></b>&nbsp;
						<%
						if len(aTable1Values(2,iRowLoop)) > 25 then
						jobnavnthis = left(aTable1Values(2,iRowLoop), 25) &".."
						else
						jobnavnthis = aTable1Values(2,iRowLoop)
						end if
						%>
						
						<%if level <= 3 OR level = 6 then%>
						<a href="jobs.asp?menu=job&func=red&id=<%=aTable1Values(0,iRowLoop)%>&int=1&rdir=treg" class="alt"><%=jobnavnthis%></a>
						<%else%>
						<font class='storhvid'><%=jobnavnthis%></font>
						<%end if%>
						&nbsp;
						<font class='lilleblaa'>(
						<%if len(aTable1Values(4,iRowLoop)) > 20 then
						Response.write left(aTable1Values(4,iRowLoop), 20) &".."
						else
						Response.write aTable1Values(4,iRowLoop)
						end if%>
						)</font>
							
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
						<td valign=top align=right width=120 style='padding-right:3px; background-color: #5582D2; border-right: 1px solid #003399;'><font class='megetlillehvid'><%=formatnumber(aTable1Values(17,iRowLoop) + aTable1Values(18,iRowLoop), 2)%> / <font class='megetlilleblaa'><%= formatnumber(timerbrugtthis, 2) %>
						<input type="checkbox" name="FM_flyttilguide" value="<%=aTable1Values(0,iRowLoop)%>"></td>
					</tr>
						
							
		<%
		'**** Starter div uanset om der findes aktiviteter ***
		%>
		<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
		<td colspan="3"><!--<b>Start akt div -- <=aTable1Values(2,iRowLoop)%> --</b><br>-->
		<div ID="Menu<%=iRowLoop%>" Style="position: relative; display: none;">
		<table cellspacing="0" cellpadding="0" border="0" width=100% bgcolor="#ffffff">
		
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
					
					<tr style='background-color: #D6DFF5;'>
						<td width=130 style='border-left: 1px solid #003399; border-bottom: 1px solid #FFFFFF; padding-left:2px; padding-top:2px;'>
						<%
						if len(aTable1Values(6, iRowLoop)) > 22 then
							Response.write "<font class=lille-kalender>"& left(aTable1Values(6, iRowLoop), 22) &"..</font>"
							else
							Response.write "<font class=lille-kalender>"& aTable1Values(6, iRowLoop) &"</font>"
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
				%>
				<!--#include file="inc/fakfarver_timereg.asp"-->
				<%
					
				'*** Udskriver timefelter til siden ***********************	
				%>
				<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
				<input type="hidden" name="FM_son_opr_<%=iRowLoop%>" value="<%=sonTimerVal(sonRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then ' 2 = Kørsel (2 i akt.fakbar 5 i timer.tfaktim)/ alm aktivitet %>
				<input type="Text" name="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=sonTimerVal(sonRLoop)%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'son'), tjektimer('son',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_son%>; border-color: <%=fakbgcol_son%>; border-style: solid; width : 35px;">
				<%else%>
				<input type="Text" name="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=sonTimerVal(sonRLoop)%>" onkeyup="tjekkm('son',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_son%>; border-color: Sienna; border-style: dashed; width : 35px;">
				<%end if%>
				<%call showcommentfunc(1)%>
				<a href="#" onclick="expandkomm('son',<%=iRowLoop%>,'<%=sonKomm%>');"><%=kommtrue%></a>
				</td>
				<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
				<input type="hidden" name="FM_man_opr_<%=iRowLoop%>" value="<%=manTimerVal(manRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=manTimerVal(manRLoop)%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'man'), tjektimer('man',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_man%>; border-color: <%=fakbgcol_man%>; border-style: solid; width : 35px;">
				<%else%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=manTimerVal(manRLoop)%>" onkeyup="tjekkm('man',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_man%>; border-color: #5582D2; border-style: dashed; width : 35px;">
				<%end if%>
				<%call showcommentfunc(2)%>
				<a href="#" onclick="expandkomm('man',<%=iRowLoop%>,'<%=manKomm%>');"><%=kommtrue%></a>
				</td>
				<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
				<input type="hidden" name="FM_tir_opr_<%=iRowLoop%>" value="<%=tirTimerVal(tirRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=tirTimerVal(tirRLoop)%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'tir'), tjektimer('tir',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_tir%>; border-color: <%=fakbgcol_tir%>; border-style: solid; width : 35px;">
				<%else%>
				<input type="Text" name="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=tirTimerVal(tirRLoop)%>" onkeyup="tjekkm('tir',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_tir%>; border-color: #5582D2; border-style: dashed; width : 35px;">
				<%end if%>
				<%call showcommentfunc(3)%>
				<a href="#" onclick="expandkomm('tir',<%=iRowLoop%>,'<%=tirKomm%>');"><%=kommtrue%></a>
				</td>
				<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
				<input type="hidden" name="FM_ons_opr_<%=iRowLoop%>" value="<%=onsTimerVal(onsRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=onsTimerVal(onsRLoop)%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'ons'), tjektimer('ons',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_ons%>; border-color: <%=fakbgcol_ons%>; border-style: solid; width : 35px;">
				<%else%>
				<input type="Text" name="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=onsTimerVal(onsRLoop)%>" onkeyup="tjekkm('ons',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_ons%>; border-color: #5582D2; border-style: dashed; width : 35px;">
				<%end if%>
				<%call showcommentfunc(4)%>
				<a href="#" onclick="expandkomm('ons',<%=iRowLoop%>,'<%=onsKomm%>');"><%=kommtrue%></a>
				</td>
				<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
				<input type="hidden" name="FM_tor_opr_<%=iRowLoop%>" value="<%=torTimerVal(torRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=torTimerVal(torRLoop)%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'tor'), tjektimer('tor',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_tor%>; border-color: <%=fakbgcol_tor%>; border-style: solid; width : 35px;">
				<%else%>
				<input type="Text" name="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=torTimerVal(torRLoop)%>" onkeyup="tjekkm('tor',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_tor%>; border-color:  #5582D2; border-style: dashed; width : 35px;">
				<%end if%>
				<%call showcommentfunc(5)%>
				<a href="#" onclick="expandkomm('tor',<%=iRowLoop%>,'<%=torKomm%>');"><%=kommtrue%></a>
				</td>
				<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
				<input type="hidden" name="FM_fre_opr_<%=iRowLoop%>" value="<%=freTimerVal(freRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=freTimerVal(freRLoop)%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'fre'), tjektimer('fre',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_fre%>; border-color: <%=fakbgcol_fre%>; border-style: solid; width : 35px;">
				<%else%>
				<input type="Text" name="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=freTimerVal(freRLoop)%>" onkeyup="tjekkm('fre',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_fre%>; border-color: #5582D2; border-style: dashed; width : 35px;">
				<%end if%>
				<%call showcommentfunc(6)%>
				<a href="#" onclick="expandkomm('fre',<%=iRowLoop%>,'<%=freKomm%>');"><%=kommtrue%></a>
				</td>
				<td align="center" style="border-right: 1px solid #003399; border-bottom: 1px solid #FFFFFF;">
				<input type="hidden" name="FM_lor_opr_<%=iRowLoop%>" value="<%=lorTimerVal(lorRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=lorTimerVal(lorRLoop)%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'lor'), tjektimer('lor',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_lor%>; border-color: <%=fakbgcol_lor%>; border-style: solid; width : 35px;">
				<%else%>
				<input type="Text" name="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=lorTimerVal(lorRLoop)%>" onkeyup="tjekkm('lor',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_lor%>; border-color: Sienna; border-style: dashed; width : 35px;">
				<%end if%>
				<%call showcommentfunc(7)%>
				<a href="#" onclick="expandkomm('lor',<%=iRowLoop%>,'<%=lorKomm%>');"><%=kommtrue%></a>
				</td>
			</tr>
			
				<%'kommentarer%>
				<input type="hidden" name="FM_kom_son_<%=iRowLoop%>" id="FM_kom_son_<%=iRowLoop%>" value="<%=sonKomm%>">
				<input type="hidden" name="FM_kom_man_<%=iRowLoop%>" id="FM_kom_man_<%=iRowLoop%>" value="<%=manKomm%>">
				<input type="hidden" name="FM_kom_tir_<%=iRowLoop%>" id="FM_kom_tir_<%=iRowLoop%>" value="<%=tirKomm%>">
				<input type="hidden" name="FM_kom_ons_<%=iRowLoop%>" id="FM_kom_ons_<%=iRowLoop%>" value="<%=onsKomm%>">
				<input type="hidden" name="FM_kom_tor_<%=iRowLoop%>" id="FM_kom_tor_<%=iRowLoop%>" value="<%=torKomm%>">
				<input type="hidden" name="FM_kom_fre_<%=iRowLoop%>" id="FM_kom_fre_<%=iRowLoop%>" value="<%=freKomm%>">
				<input type="hidden" name="FM_kom_lor_<%=iRowLoop%>" id="FM_kom_lor_<%=iRowLoop%>" value="<%=lorKomm%>">
			<%
			sonKomm = ""
			manKomm = ""
			tirKomm = ""
			onsKomm = ""
			torKomm = ""
			freKomm = ""
			lorKomm = ""
			
			sonRLoop = ""
			manRLoop = ""
			tirRLoop = ""
			onsRLoop = ""
			torRLoop = ""
			freRLoop = ""
			lorRLoop = ""
			
			end if
		de = iRowLoop
		end if 
	Next 
	
	'*** afslutter den sidtste div 
	if iRowLoop > 0 then 
	'Response.write "<tr><td><b>final akt div</b></td></tr>"%>
	</table></div></td></tr>
	<tr>
		<td bgcolor="#5582d2" height=5 colspan="9" style="border-left: 1px solid #003399; border-right: 1px solid #003399; border-bottom: 1px solid #003399;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td colspan=9 align=right style="padding-right:2; padding-top:2;"><font class="megetlillesort">Marker job for at fjerne det fra denne oversigt.</td>
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
	
	
	'**** Interne Job ****
	%>
	<br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="530">
	<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="74" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Interne job</font></td><td style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid; padding-left:2; padding-top:1;"><img src="../ill/internt_job_trans.gif" width="15" height="15" alt="" border="0"></td>
	<td width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
	</tr>
	</table>
	<table width='530' border='0' cellspacing='0' cellpadding='0'>
	<%
	call dageDatoer()
	'** Udskriver Interne Jobnr **
	strLastjobnr = ""
	strSQL = "SELECT job.id, jobnr, jobnavn, kkundenavn FROM job, kunder WHERE jobstatus = 1 AND fakturerbart = 0 AND kunder.Kid = jobknr ORDER BY jobnavn"
	oRec.cachesize = 10
	oRec.Open strSQL, oConn, 0, 1
		
		di = antal_de + 1
		
		While Not oRec.EOF
		
		
					'*** Finder timer på interne job. ****
					sonTimerValint = ""
					manTimerValint = ""
					tirTimerValint = ""
					onsTimerValint = ""
					torTimerValint = ""
					freTimerValint = ""
					lorTimerValint = ""
		
					
					strSQLtimer = "SELECT timer, tdato FROM timer WHERE Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "& strMnr &" AND Tdato = '"& varTjDatoUS_son &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_man &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_tir &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_ons &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_tor &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_fre &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_lor &"'"
					
					
					oRec2.Open strSQLtimer, oConn, 0, 1  
					while not oRec2.EOF 
					
						select case DatePart("w", oRec2("tdato")) 
						case 1	
						sonTimerValint = SQLBless(oRec2("timer"))
						case 2
						manTimerValint = SQLBless(oRec2("timer"))
						case 3
						tirTimerValint = SQLBless(oRec2("timer"))
						case 4
						onsTimerValint = SQLBless(oRec2("timer"))
						case 5
						torTimerValint = SQLBless(oRec2("timer"))
						case 6
						freTimerValint = SQLBless(oRec2("timer"))
						case 7
						lorTimerValint = SQLBless(oRec2("timer"))
						end select
						
					oRec2.movenext
					wend
					oRec2.close
					'***************************************************
					
					
		Response.write "<tr><td style='width : 160px; !background-color: #ffffff; padding-left : 4px; padding-right : 4px; border-left: 1px solid #003399;'><font class='lillesort'>"_
		& oRec("Jobnr") &"&nbsp;"& left(oRec("Jobnavn"), 15)&"</td>"
		%><input type="hidden" name="FM_jobnr_<%=di%>" value="<%=oRec("jobnr")%>">
		<td align="center"><input type="hidden" name="FM_son_opr_<%=di%>" value="<%=sonTimerValint%>"><input type="Text" name="Timer_son_<%=di%>" maxlength="5" value="<%=sonTimerValint%>" onkeyup="setTimerTot(<%=di%>, 'son'), tjektimer('son',<%=di%>)"; style="!border: 1px; background-color: #FFDFDF; border-color: #cd853f; border-style: solid; width : 52px;"></td>
		<td align="center"><input type="hidden" name="FM_man_opr_<%=di%>" value="<%=manTimerValint%>"><input type="Text" name="Timer_man_<%=di%>" maxlength="5" value="<%=manTimerValint%>" onkeyup="setTimerTot(<%=di%>, 'man'), tjektimer('man',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 52px;"></td>
		<td align="center"><input type="hidden" name="FM_tir_opr_<%=di%>" value="<%=tirTimerValint%>"><input type="Text" name="Timer_tir_<%=di%>" maxlength="5" value="<%=tirTimerValint%>" onkeyup="setTimerTot(<%=di%>, 'tir'), tjektimer('tir',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 52px;"></td>
		<td align="center"><input type="hidden" name="FM_ons_opr_<%=di%>" value="<%=onsTimerValint%>"><input type="Text" name="Timer_ons_<%=di%>" maxlength="5" value="<%=onsTimerValint%>" onkeyup="setTimerTot(<%=di%>, 'ons'), tjektimer('ons',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 52px;"></td>
		<td align="center"><input type="hidden" name="FM_tor_opr_<%=di%>" value="<%=torTimerValint%>"><input type="Text" name="Timer_tor_<%=di%>" maxlength="5" value="<%=torTimerValint%>" onkeyup="setTimerTot(<%=di%>, 'tor'), tjektimer('tor',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 52px;"></td>
		<td align="center"><input type="hidden" name="FM_fre_opr_<%=di%>" value="<%=freTimerValint%>"><input type="Text" name="Timer_fre_<%=di%>" maxlength="5" value="<%=freTimerValint%>" onkeyup="setTimerTot(<%=di%>, 'fre'), tjektimer('fre',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 52px;"></td>
		<td align="center" style="border-right: 1px solid #003399;"><input type="hidden" name="FM_lor_opr_<%=di%>" value="<%=lorTimerValint%>"><input type="Text" name="Timer_lor_<%=di%>" maxlength="5" value="<%=lorTimerValint%>" size="6" onkeyup="setTimerTot(<%=di%>, 'lor'), tjektimer('lor',<%=di%>)"; style="!border: 1px; background-color: #FFDFDF; border-color: #cd853f; border-style: solid; width : 52px;"></td>
		</tr>
		<%
		showWeekInt = 1
		di = di + 1
		strLastjobnr = oRec("Jobnr")
	oRec.MoveNext
	Wend 
	oRec.Close
	
	antal_di = di - 1%>
<input type="hidden" name="FM_di" value="<%=antal_di%>">
<input type="hidden" name="FM_strweek" value="<%=strWeek%>">
<tr>
    <td colspan="9" align="right"><br><input type="image" src="../ill/frem.gif"></td>
</tr>
</table>


<!--
</td>
</tr>
</table>-->




	<%'***** kommentar layer *** %>
	<div id="kom" name="kom" Style="position: absolute; display: none; left:20; top:170; z-index:2994; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">
	Kommentar:
	<input type="hidden" name="daytype" value="son">
	<input type="hidden" name="rowcounter" value="0">
	<img src="../ill/blank.gif" width="190" height="1" alt="" border="0">
	(maks. 255 karakterer)&nbsp;
	<input type="text" name="antch" id="antch" size="3" maxlength="4"><br>
	<textarea cols="50" rows="6" id="FM_kom" name="FM_kom" onKeyup="antalchar();"></textarea><br>
	<a href="#" onclick="closekomm();" class=vmenu>Gem og Luk&nbsp;</a><br>&nbsp;
	</div>

	
<br>
<br>
<br><br>
</div>

<%'***** ressourcetimer + ugetotaler *** %>
<div name=ressourcetimer id=ressourcetimer style="position:absolute; left:85; top:185; visible:hidden; display:none; border:1px #003399 solid; overflow:auto; background-color:#FFFFF1; filter:alpha(opacity=95); z-index:2000; padding-left:3;">
<!--#include file="inc/timereg_restimer_inc.asp"-->
</div>

<%'***** dagsseddel *** %>
<!--#include file="inc/timereg_dagsseddel_inc.asp"-->


</form>



<%
'Response.write antalguidenjobs
if cint(antalguidenjobs) > 15 AND (right(day(date),1) = 3) then
%>
<div id="tomanyjob" name="tomanyjob" style="position:absolute; left:240; top:203; visibility:visible; background-color:#d6dff5; filter:alpha(opacity=90); border-left:1 #003399 solid; border-top:1 #003399 solid; border-bottom:1 #003399 solid; border-right:1 #003399 solid;">
<table border=0 cellpadding=0 cellspacing=0>
<tr bgcolor="#FFFFFF"><td width="400" style="Padding-left:5; Padding-right:5;">
<img src="../ill/info.gif" width="42" height="38" alt="" border="0">
Hej <b><%=session("user")%></b> !<br>
Har du overvejet at sætte antallet dine aktive job ned? <br><br>
Du får i øjeblikket vist <b><%=antalguidenjobs%></b> job her på siden. Hvis du vil have et bedre overblik 
kan du ændre dette i <a href="timereg.asp?func=showguide">Guiden dine aktive job!</a>, eller klikke dem fra i checkboksen ude til højre for hvert job.
<br><br>&nbsp; 
</td><td align=right valign=top><a href="#" onclick=closetomanyjob();><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0"></a></td></tr></table> 
</div>
<%
end if
%>


<script>
document.all["loading"].style.visibility = "hidden";
</script>
<%end if 'Viskunder Wiz eller Timereg
end if 'user session
%>
<!--#include file="../inc/regular/footer_inc.asp"-->





