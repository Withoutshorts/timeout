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
	<html>
		<head>
		<title>TimeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
		<script type="text/javascript">
			function visf5(){
			window.opener.document.all["f5"].innerHTML = "<b>Status:</b> <font color=red>Opdater siden [F5].</font>"; 
			} 
		</script>
		<SCRIPT language=javascript src="inc/timereg_func.js"></script>
		</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<!--#include file="inc/helligdage_func.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	<%
	
	id = request("jid")
	intMnr = request("mnr")
	strdag = request("FM_start_dag")
	strmrd = request("FM_start_mrd")
	straar = request("FM_start_aar")
	searchstring = request("searchstring")
	showeksterne = request("showeks")
	
	'**** sætter datoer, når kalender ikke bruges ***
	strMrd = request("strmrd")
	select case strMrd
	case 2
		if request("strdag") > 28 then
		strDag = 28
		else
		strDag = request("strdag")
		end if
	case 1, 3, 5, 7, 8, 10, 12
		strDag = request("strdag")
	case else
		if request("strdag") > 30 then
		strDag = 30
		else
		strDag = request("strdag")
		end if
	end select
	strAar = request("straar")
	
	daynow = formatdatetime(day(now) & "/" & month(now) & "/" & year(now), 0)
	useDate = formatdatetime(strDag & "/" & strMrd & "/" & strAar, 0)
	
	function SQLBless(s)
			dim tmp
			tmp = s
			tmp = replace(tmp, ",", ".")
			SQLBless = tmp
	end function
	
	
	
	'********************************
	'Udksriver side til skærm.
	'********************************
	
	
	'** Henter ugedage i den valgte uge 
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="633">
		<tr>
			<td bgcolor="#003399" width="633"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
			<td bgcolor="#FFFFFF" align=right>&nbsp;</td>
		</tr>
		</table>
	<div style="position:absolute; left:0; top:43;">
	<table cellspacing="0" cellpadding="0" border="0" width="633">
	<!--<tr bgcolor="5582D2">
		<td width="8" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="21" alt="" border="0"></td>
		<td valign="top" colspan=9><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
		<td align=right valign=top><img src="../ill/tabel_top_right.gif" width="8" height="21" alt="" border="0"></td>
	</tr>
	</table>
	<table cellspacing="0" cellpadding="0" border="0" width="590">-->
	<form action="timereg_db.asp" method="POST" name="timereg" onSubmit="visf5();">
	<input type="Hidden" name="FM_start_dag" value="<%=strdag%>">
	<input type="Hidden" name="FM_start_mrd" value="<%=strmrd%>">
	<input type="Hidden" name="FM_start_aar" value="<%=straar%>">
	<input type="Hidden" name="searchstring" value="<%=searchstring%>">
	<input type="Hidden" name="FM_mnr" value="<%=intMnr%>">
	<!--#include file="inc/timereg_dage_inc.asp"-->
	<%
	call dageDatoer()
	%>
	
	<!--?--><input type="hidden" name="year" value="<%=strRegAar%>">
	<input type="hidden" name="datoSon" value="<%=tjekdag(1)%>">
	<input type="hidden" name="datoMan" value="<%=tjekdag(2)%>">
	<input type="hidden" name="datoTir" value="<%=tjekdag(3)%>">
	<input type="hidden" name="datoOns" value="<%=tjekdag(4)%>">
	<input type="hidden" name="datoTor" value="<%=tjekdag(5)%>">
	<input type="hidden" name="datoFre" value="<%=tjekdag(6)%>">
	<input type="hidden" name="datoLor" value="<%=tjekdag(7)%>">
	<%
	
	
	
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
	
	
	'*** sql datoer ***
	varTjDatoUS_son = convertDateYMD(tjekdag(1))
	varTjDatoUS_man = convertDateYMD(tjekdag(2))
	varTjDatoUS_tir = convertDateYMD(tjekdag(3))
	varTjDatoUS_ons = convertDateYMD(tjekdag(4)) 
	varTjDatoUS_tor = convertDateYMD(tjekdag(5)) 
	varTjDatoUS_fre = convertDateYMD(tjekdag(6))
	varTjDatoUS_lor = convertDateYMD(tjekdag(7))
	
	
	'******************************************************************
	'Sætter farver på første row 
	fakbgcol_son = "#cd853f"
	fakbgcol_man = "#7F9DB9"	
	fakbgcol_tir = "#7F9DB9"	
	fakbgcol_ons = "#7F9DB9"	
	fakbgcol_tor = "#7F9DB9"	
	fakbgcol_fre = "#7F9DB9"	
	fakbgcol_lor = "#cd853f"
	
	maxl_son = 5
	maxl_man = 5
	maxl_tir = 5
	maxl_ons = 5
	maxl_tor = 5
	maxl_fre = 5
	maxl_lor = 5
	
	fmbg_son = "#FFDFDF" 
	fmbg_man = "#FFFFFF" 
	fmbg_tir = "#FFFFFF" 
	fmbg_ons = "#FFFFFF" 
	fmbg_tor = "#FFFFFF" 
	fmbg_fre = "#FFFFFF" 
	fmbg_lor = "#FFDFDF"  
	'******************************************************************
	
	
	
	
	'*******************************************
	'**** Viser Eksterne eller Interne Job. 
	'*******************************************
	if showeksterne = "1" then
	
	
	'**** funktion til at vise kommentar ikoner ****
		public kommTrue
		public function showcommentfunc(dag)
		dag = dag
		select case dag
		case 1
		useKomm = sonKomm
		case 2
		useKomm = manKomm
		case 3
		useKomm = tirKomm
		case 4
		useKomm = onsKomm
		case 5
		useKomm = torKomm
		case 6
		useKomm = freKomm
		case 7
		useKomm = lorKomm
		end select
		
		'Response.write "usek: "& useKomm & "<br>"
		
			if len(useKomm) <> 0 then
			kommtrue = "+"
			else
			kommtrue = "<font color='#999999'>+</font>"
			end if 
		end function
		'***************************************************
	
	
	'**************************** Rettigheds tjeck aktiviteter **************************
	f = 0
		dim brugergruppeId
		Redim brugergruppeId(f)
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid="& intMnr &" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid" 
			oRec.Open strSQL, oConn, 0, 1 
			
			While Not oRec.EOF
			Redim preserve brugergruppeId(f) 
				strMnavn = oRec("Mnavn")
				strMnr = intMnr
				strMType = oRec("type")
				brugergruppeId(f) = oRec("ProjektgruppeId")
				f = f + 1
				oRec.MoveNext
			Wend
		
	oRec.Close
	
	
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
	
	
	Redim sonTimerVal(iRowLoop)
	Redim manTimerVal(iRowLoop)
	Redim tirTimerVal(iRowLoop)
	Redim onsTimerVal(iRowLoop)
	Redim torTimerVal(iRowLoop)
	Redim freTimerVal(iRowLoop)
	Redim lorTimerVal(iRowLoop)
	
	Redim dtimeTtidspkt_son(iRowLoop)
	Redim dtimeTtidspkt_man(iRowLoop)
	Redim dtimeTtidspkt_tir(iRowLoop)
	Redim dtimeTtidspkt_ons(iRowLoop)
	Redim dtimeTtidspkt_tor(iRowLoop)
	Redim dtimeTtidspkt_fre(iRowLoop)
	Redim dtimeTtidspkt_lor(iRowLoop)
	
	sonRLoop = ""
	manRLoop = ""
	tirRLoop = ""
	onsRLoop = ""
	torRLoop = ""
	freRLoop = ""
	lorRLoop = ""
	
	lastfakdato = ""
	dtimeTtidspkt = ""
	
	
	'******************************************************************** 
	'Gennemløber de job som brugeren har adgang til. 
	'Udskrivning af side starter her
	'1 jobid
	'2 jobnr
	'4 jobnavn
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
	&") WHERE job.id = "& id &" " & strSQLkri & " AND Kid = jobknr ORDER BY jobnavn, id, navn" 
	
	'Response.write (strSQL)
	'oRec.cachesize = 150
	
	oRec.Open strSQL, oConn, 3
	
	Dim aTable1Values
	if not oRec.EOF Then
	aTable1Values = oRec.GetRows()
	oRec.close
			
		
		
		'*** Udskriver records (aktiviteter) ****
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
		
						if aTable1Values(1,iRowLoop) <> aTable1Values(1,(prevrecord)) OR iRowLoop > 0 then				
						Response.write "</table></div></td></tr>"
						
						end if	
			
			
			end if
		
			'*** Hvis aktid <> 0 ****
			if len(aTable1Values(5,iRowLoop)) <> 0 then
				
				
				
				'**********************************************************************
				'**** Finder ugens timer på aktid ***
				Redim preserve sonTimerVal(iRowLoop)
				Redim preserve manTimerVal(iRowLoop)
				Redim preserve tirTimerVal(iRowLoop)
				Redim preserve onsTimerVal(iRowLoop)
				Redim preserve torTimerVal(iRowLoop)
				Redim preserve freTimerVal(iRowLoop)
				Redim preserve lorTimerVal(iRowLoop)
				
				Redim preserve dtimeTtidspkt_son(iRowLoop)
				Redim preserve dtimeTtidspkt_man(iRowLoop)
				Redim preserve dtimeTtidspkt_tir(iRowLoop)
				Redim preserve dtimeTtidspkt_ons(iRowLoop)
				Redim preserve dtimeTtidspkt_tor(iRowLoop)
				Redim preserve dtimeTtidspkt_fre(iRowLoop)
				Redim preserve dtimeTtidspkt_lor(iRowLoop)
				
				Select case datepart("w", aTable1Values(12, iRowLoop))
				case 1
				sonTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				sonRLoop = iRowLoop
				dtimeTtidspkt_son(sonRLoop) = aTable1Values(14,sonRLoop)
					
				sonKomm = aTable1Values(15,iRowLoop)
				case 2
				manTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				manRLoop = iRowLoop
				dtimeTtidspkt_man(manRLoop) = aTable1Values(14,manRLoop)
					
				manKomm = aTable1Values(15,iRowLoop)
				case 3
				tirTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				tirRLoop = iRowLoop
				dtimeTtidspkt_tir(tirRLoop) = aTable1Values(14,tirRLoop)
					
				tirKomm = aTable1Values(15,iRowLoop)
				case 4
				onsTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				onsRLoop = iRowLoop
				dtimeTtidspkt_ons(onsRLoop) = aTable1Values(14,onsRLoop)
					
				onsKomm = aTable1Values(15,iRowLoop)
				case 5
				torTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				torRLoop = iRowLoop
				dtimeTtidspkt_tor(torRLoop) = aTable1Values(14,torRLoop)
					
				torKomm = aTable1Values(15,iRowLoop)
				case 6
				freTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				freRLoop = iRowLoop
				dtimeTtidspkt_fre(freRLoop) = aTable1Values(14,freRLoop)
					
				freKomm = aTable1Values(15,iRowLoop)
				case 7
				lorTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				lorRLoop = iRowLoop
				dtimeTtidspkt_lor(lorRLoop) = aTable1Values(14,lorRLoop)
					
				lorKomm = aTable1Values(15,iRowLoop)
				end select
				'**************************************************************************
				
				
			if aTable1Values(5, iRowLoop) <> aTable1Values(5, nextrecord) OR aTable1Values(2, iRowLoop) <> aTable1Values(2, nextrecord) OR nextrecord = "0" then%>
				<input type="hidden" name="FM_jobnr_<%=iRowLoop%>" value="<%=aTable1Values(1,iRowLoop)%>">
				<input type="hidden" name="FM_Aid_<%=iRowLoop%>" value="<%=aTable1Values(5,iRowLoop)%>">
				
				<tr style='background-color: #D6DFF5;'>
					<td width=106 style='border-left: 1px solid #003399; border-bottom: 1px solid #FFFFFF; padding-left:2px; padding-top:2px;'>
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
								if len(sonRLoop) <> 0 then
								sonRLoop = sonRLoop
								else
								sonRLoop = iRowLoop
								end if
								
								if len(manRLoop) <> 0 then
								manRLoop = manRLoop
								else
								manRLoop = iRowLoop
								end if
								
								if len(tirRLoop) <> 0 then
								tirRLoop = tirRLoop
								else
								tirRLoop = iRowLoop
								end if
								
								if len(onsRLoop) <> 0 then
								onsRLoop = onsRLoop
								else
								onsRLoop = iRowLoop
								end if
								
								if len(torRLoop) <> 0 then
								torRLoop = torRLoop
								else
								torRLoop = iRowLoop
								end if
								
								if len(freRLoop) <> 0 then
								freRLoop = freRLoop
								else
								freRLoop = iRowLoop
								end if
								
								if len(lorRLoop) <> 0 then
								lorRLoop = lorRLoop
								else
								lorRLoop = iRowLoop
								end if
				 			
					
						'************************************************************
						' Henter feltfarver alt efter
						' om der findes fakturaer på jobbet.
						'************************************************************
						%>
						<!--#include file="inc/fakfarver_timereg.asp"-->
						<%
						'************************************************************
						
						
						
					'*** Udskriver timefelter til siden ***********************	
					%>
					<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
					<input type="hidden" name="FM_son_opr_<%=iRowLoop%>" value="<%=sonTimerVal(sonRLoop)%>">
					<%if aTable1Values(16,iRowLoop) <> 2 then ' 2 = Kørsel (2 i akt.fakbar 5 i timer.tfaktim)/ alm aktivitet %>
					<input type="Text" name="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=sonTimerVal(sonRLoop)%>" onkeyup="tjektimer('son',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_son%>; border-color: <%=fakbgcol_son%>; border-style: solid; width : 35px;">
					<%else%>
					<input type="Text" name="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=sonTimerVal(sonRLoop)%>" onkeyup="tjekkm('son',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_son%>; border-color: Sienna; border-style: dashed; width : 35px;">
					<%end if%>
					<%call showcommentfunc(1)%>
					<a href="#" onclick="expandkomm('son',<%=iRowLoop%>);"><%=kommtrue%></a>
					</td>
					<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
					<input type="hidden" name="FM_man_opr_<%=iRowLoop%>" value="<%=manTimerVal(manRLoop)%>">
					<%if aTable1Values(16,iRowLoop) <> 2 then%>
					<input type="Text" name="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=manTimerVal(manRLoop)%>" onkeyup="tjektimer('man',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_man%>; border-color: <%=fakbgcol_man%>; border-style: solid; width : 35px;">
					<%else%>
					<input type="Text" name="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=manTimerVal(manRLoop)%>" onkeyup="tjekkm('man',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_man%>; border-color: #5582D2; border-style: dashed; width : 35px;">
					<%end if%>
					<%call showcommentfunc(2)%>
					<a href="#" onclick="expandkomm('man',<%=iRowLoop%>);"><%=kommtrue%></a>
					</td>
					<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
					<input type="hidden" name="FM_tir_opr_<%=iRowLoop%>" value="<%=tirTimerVal(tirRLoop)%>">
					<%if aTable1Values(16,iRowLoop) <> 2 then%>
					<input type="Text" name="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=tirTimerVal(tirRLoop)%>" onkeyup="tjektimer('tir',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_tir%>; border-color: <%=fakbgcol_tir%>; border-style: solid; width : 35px;">
					<%else%>
					<input type="Text" name="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=tirTimerVal(tirRLoop)%>" onkeyup="tjekkm('tir',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_tir%>; border-color: #5582D2; border-style: dashed; width : 35px;">
					<%end if%>
					<%call showcommentfunc(3)%>
					<a href="#" onclick="expandkomm('tir',<%=iRowLoop%>);"><%=kommtrue%></a>
					</td>
					<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
					<input type="hidden" name="FM_ons_opr_<%=iRowLoop%>" value="<%=onsTimerVal(onsRLoop)%>">
					<%if aTable1Values(16,iRowLoop) <> 2 then%>
					<input type="Text" name="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=onsTimerVal(onsRLoop)%>" onkeyup="tjektimer('ons',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_ons%>; border-color: <%=fakbgcol_ons%>; border-style: solid; width : 35px;">
					<%else%>
					<input type="Text" name="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=onsTimerVal(onsRLoop)%>" onkeyup="tjekkm('ons',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_ons%>; border-color: #5582D2; border-style: dashed; width : 35px;">
					<%end if%>
					<%call showcommentfunc(4)%>
					<a href="#" onclick="expandkomm('ons',<%=iRowLoop%>);"><%=kommtrue%></a>
					</td>
					<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
					<input type="hidden" name="FM_tor_opr_<%=iRowLoop%>" value="<%=torTimerVal(torRLoop)%>">
					<%if aTable1Values(16,iRowLoop) <> 2 then%>
					<input type="Text" name="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=torTimerVal(torRLoop)%>" onkeyup="tjektimer('tor',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_tor%>; border-color: <%=fakbgcol_tor%>; border-style: solid; width : 35px;">
					<%else%>
					<input type="Text" name="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=torTimerVal(torRLoop)%>" onkeyup="tjekkm('tor',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_tor%>; border-color:  #5582D2; border-style: dashed; width : 35px;">
					<%end if%>
					<%call showcommentfunc(5)%>
					<a href="#" onclick="expandkomm('tor',<%=iRowLoop%>);"><%=kommtrue%></a>
					</td>
					<td align="center" width=56 style="border-bottom: 1px solid #FFFFFF;">
					<input type="hidden" name="FM_fre_opr_<%=iRowLoop%>" value="<%=freTimerVal(freRLoop)%>">
					<%if aTable1Values(16,iRowLoop) <> 2 then%>
					<input type="Text" name="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=freTimerVal(freRLoop)%>" onkeyup="tjektimer('fre',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_fre%>; border-color: <%=fakbgcol_fre%>; border-style: solid; width : 35px;">
					<%else%>
					<input type="Text" name="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=freTimerVal(freRLoop)%>" onkeyup="tjekkm('fre',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_fre%>; border-color: #5582D2; border-style: dashed; width : 35px;">
					<%end if%>
					<%call showcommentfunc(6)%>
					<a href="#" onclick="expandkomm('fre',<%=iRowLoop%>);"><%=kommtrue%></a>
					</td>
					<td align="center" style="border-right: 0px solid #003399; border-bottom: 1px solid #FFFFFF;">
					<input type="hidden" name="FM_lor_opr_<%=iRowLoop%>" value="<%=lorTimerVal(lorRLoop)%>">
					<%if aTable1Values(16,iRowLoop) <> 2 then%>
					<input type="Text" name="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=lorTimerVal(lorRLoop)%>" onkeyup="tjektimer('lor',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_lor%>; border-color: <%=fakbgcol_lor%>; border-style: solid; width : 35px;">
					<%else%>
					<input type="Text" name="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=lorTimerVal(lorRLoop)%>" onkeyup="tjekkm('lor',<%=iRowLoop%>);" style="!border: 1px; background-color: <%=fmbg_lor%>; border-color: Sienna; border-style: dashed; width : 35px;">
					<%end if%>
					<%call showcommentfunc(7)%>
					<a href="#" onclick="expandkomm('lor',<%=iRowLoop%>);"><%=kommtrue%></a>
					</td>
				</tr>
				
				<%'kommentarer%>
				<div id="kom_son_<%=iRowLoop%>" name="kom_son_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2994; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_son)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_son_<%=iRowLoop%>" id="antch_son_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_son_<%=iRowLoop%>" name="FM_kom_son_<%=iRowLoop%>" onKeyup="antalchar('son',<%=iRowLoop%>);"><%=sonKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('son',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_man_<%=iRowLoop%>" name="kom_man_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2995; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_man)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_man_<%=iRowLoop%>" id="antch_man_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_man_<%=iRowLoop%>" name="FM_kom_man_<%=iRowLoop%>" onKeyup="antalchar('man',<%=iRowLoop%>);"><%=manKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('man',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_tir_<%=iRowLoop%>" name="kom_tir_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2996; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_tir)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_tir_<%=iRowLoop%>" id="antch_tir_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_tir_<%=iRowLoop%>" name="FM_kom_tir_<%=iRowLoop%>" onKeyup="antalchar('tir',<%=iRowLoop%>);"><%=tirKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('tir',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_ons_<%=iRowLoop%>" name="kom_ons_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2997; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_ons)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_ons_<%=iRowLoop%>" id="antch_ons_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_ons_<%=iRowLoop%>" name="FM_kom_ons_<%=iRowLoop%>" onKeyup="antalchar('ons',<%=iRowLoop%>);"><%=onsKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('ons',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_tor_<%=iRowLoop%>" name="kom_tor_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2998; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_tor)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_tor_<%=iRowLoop%>" id="antch_tor_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_tor_<%=iRowLoop%>" name="FM_kom_tor_<%=iRowLoop%>" onKeyup="antalchar('tor',<%=iRowLoop%>);"><%=torKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('tor',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_fre_<%=iRowLoop%>" name="kom_fre_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2999; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_fre)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_fre_<%=iRowLoop%>" id="antch_fre_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_fre_<%=iRowLoop%>" name="FM_kom_fre_<%=iRowLoop%>" onKeyup="antalchar('fre',<%=iRowLoop%>);"><%=freKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('fre',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_lor_<%=iRowLoop%>" name="kom_lor_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:3000; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_lor)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_lor_<%=iRowLoop%>" id="antch_lor_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_lor_<%=iRowLoop%>" name="FM_kom_lor_<%=iRowLoop%>" onKeyup="antalchar('lor',<%=iRowLoop%>);"><%=lorKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('lor',<%=iRowLoop%>);">Luk&nbsp;</a></div>
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
	
	end if
	
	'**** Antal aktiviteter ****
	antal_de = de 
	%>
	<input type="hidden" name="FM_de" value="<%=antal_de%>">
	<input type="hidden" name="FM_di" value="<%=antal_de%>">
	<%
	else
	antal_de = 0
	%>
	<input type="hidden" name="FM_de" value="<%=antal_de%>">
	<%			
	'********************************************************
	'**** Interne ****
	'********************************************************
	
				
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
					
								
								strSQLtimer = "SELECT timer, tdato FROM timer WHERE Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "& intMnr &" AND Tdato = '"& varTjDatoUS_son &"'"_
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
								
								'Response.write strSQLtimer
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
					<td align="center"><input type="hidden" name="FM_lor_opr_<%=di%>" value="<%=lorTimerValint%>"><input type="Text" name="Timer_lor_<%=di%>" maxlength="5" value="<%=lorTimerValint%>" size="6" onkeyup="setTimerTot(<%=di%>, 'lor'), tjektimer('lor',<%=di%>)"; style="!border: 1px; background-color: #FFDFDF; border-color: #cd853f; border-style: solid; width : 52px;"></td>
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
		<%end if%>
		<input type="hidden" name="FM_strweek" value="<%=strWeek%>">
	<tr>
		<td colspan="9" align="right" style="padding-right:40;"><br><input type="image" src="../ill/frem.gif"></td>
	</tr>
	</table>
	
	<%
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
