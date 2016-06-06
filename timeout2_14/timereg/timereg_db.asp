 	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
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
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="inc/isint_func.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	

<%
'Henter variable fra timereg siden
'Response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;usm:" & request("FM_use_me")

filterA_B = request("filterA_B")
searchstring = request("searchstring")
strdag = request("FM_start_dag")
strmrd = request("FM_start_mrd")
straar = request("FM_start_aar")

strMnr = request("FM_mnr")
usemrn = strMnr '* Bruges i kalenderen
antal_to = request("FM_di") 
antal_de = request("FM_de")
antal_di = (antal_to - antal_de)
errorPrinted = 0
strWeek = request("FM_strweek")
thisfile = "timereg_db"

strYear = Request("year")

'*** Sætter lokal dato/kr format. *****
Session.LCID = 1030
				
				
	'*** flytter job til guide ****
	moveJobtoguide = request("FM_flyttilguide")
	if len(moveJobtoguide) <> 0 then
		j = 0
		intuseJob = Split(moveJobtoguide, ", ")
		For j = 0 to Ubound(intuseJob)
		oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& strMnr &" AND jobid = "& intuseJob(j) &"")
		next
	end if
		
strDatoSon = formatdatetime(Request("datoSon"), 2)
strDatoMan = formatdatetime(Request("datoMan"), 2)
strDatoTir = formatdatetime(Request("datoTir"), 2)
strDatoOns = formatdatetime(Request("datoOns"), 2)
strDatoTor = formatdatetime(Request("datoTor"), 2)
strDatoFre = formatdatetime(Request("datoFre"), 2)
strDatoLor = formatdatetime(Request("datoLor"), 2)

errors = 0


function SQLBless(s)
dim tmp
tmp = s
tmp = replace(tmp, ",", ".")
SQLBless = tmp
end function

function SQLBless2(s2)
dim tmp2
tmp2 = s2
tmp2 = replace(tmp2, "'", "''")
SQLBless2 = tmp2
end function

public intTimepris, foundone
function alttimepris(useaktid)

							'*** Finder evt. alternativ timepris på medarbejder ****
							'*** Hvis der ikke findes alternative timepriser bruges default ****
							foundone = "n"
							strSQL = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris FROM timepriser WHERE jobid = "& intjobid &" AND aktid = "& useaktid &" AND medarbid =  "& strMnr
							oRec2.open strSQL, oConn, 3 
							if not oRec2.EOF then
							timeprisAlt = oRec2("timeprisalt")
							Select case timeprisAlt
							case 1
							timeprisalernativ = "timepris_a1"
							case 2
							timeprisalernativ = "timepris_a2"
							case 3
							timeprisalernativ = "timepris_a3"
							case 4
							timeprisalernativ = "timepris_a4"
							case 5
							timeprisalernativ = "timepris_a5"
							case 6
							timeprisalernativ = "timepris_a6"
							case else
							timeprisalernativ = "timepris"
							end select
								
								if timeprisalernativ = "timepris_a6" then '* Bruger de vejl. alternative timepriser 
									
									intTimepris = oRec2("6timepris")
									foundone = "y"
								
								else
									strSQL3 = "SELECT mid, "&timeprisalernativ&" AS useTimepris, medarbejdertype FROM medarbejdere, medarbejdertyper WHERE mid =" & strMnr & " AND medarbejdertyper.id = medarbejdertype"
									oRec3.open strSQL3, oConn, 3 
									if not oRec3.EOF then
									intTimepris = oRec3("useTimepris")
									foundone = "y"
									end if
									oRec3.close
								end if
								
							end if 
							oRec2.close 
		
end function

'** Henter timepris ***
oRec.Open "SELECT medarbejdertype, timepris, kostpris, mnavn FROM medarbejdere, medarbejdertyper WHERE Mid = "& strMnr &" AND medarbejdertyper.id = medarbejdertype", oConn, 3
		
		if Not oRec.EOF then
		 	if oRec("kostpris") <> "" then
			dblkostpris = oRec("kostpris")
			else
			dblkostpris = 0
			end if
		intTimepris = oRec("timepris")
		strMnavn = oRec("mnavn")
		else
		strMnavn = ".."
		intTimepris = 0
		dblkostpris = 0
		end if

oRec.close
'** Slut timepris **
	
Redim strJobnr(cint(antal_to)), strAktId(cint(antal_to))
Redim strTimer_son(cint(antal_to)), strTimer_man(cint(antal_to)), strTimer_tir(cint(antal_to)) 
Redim strTimer_ons(cint(antal_to)), strTimer_tor(cint(antal_to)), strTimer_fre(cint(antal_to)), strTimer_lor(cint(antal_to))

Redim strTimer_kom_son(cint(antal_to))
Redim strTimer_kom_man(cint(antal_to))
Redim strTimer_kom_tir(cint(antal_to))
Redim strTimer_kom_ons(cint(antal_to))
Redim strTimer_kom_tor(cint(antal_to))
Redim strTimer_kom_fre(cint(antal_to))
Redim strTimer_kom_lor(cint(antal_to))

Redim off_son(cint(antal_to))
Redim off_man(cint(antal_to))
Redim off_tir(cint(antal_to))
Redim off_ons(cint(antal_to))
Redim off_tor(cint(antal_to))
Redim off_fre(cint(antal_to))
Redim off_lor(cint(antal_to))



'********************************************************************************
'**** Loop **********
'********************************************************************************
for antal_to = 0 to cint(antal_to) 

Redim preserve strJobnr(cint(antal_to)), strAktId(cint(antal_to))
Redim preserve strTimer_son(cint(antal_to)), strTimer_man(cint(antal_to)), strTimer_tir(cint(antal_to)) 
Redim preserve strTimer_ons(cint(antal_to)), strTimer_tor(cint(antal_to)), strTimer_fre(cint(antal_to)), strTimer_lor(cint(antal_to))
Redim preserve strTimer_kom_son(cint(antal_to))
Redim preserve strTimer_kom_man(cint(antal_to))
Redim preserve strTimer_kom_tir(cint(antal_to))
Redim preserve strTimer_kom_ons(cint(antal_to))
Redim preserve strTimer_kom_tor(cint(antal_to))
Redim preserve strTimer_kom_fre(cint(antal_to))
Redim preserve strTimer_kom_lor(cint(antal_to))
Redim preserve off_son(cint(antal_to))
Redim preserve off_man(cint(antal_to))
Redim preserve off_tir(cint(antal_to))
Redim preserve off_ons(cint(antal_to))
Redim preserve off_tor(cint(antal_to))
Redim preserve off_fre(cint(antal_to))
Redim preserve off_lor(cint(antal_to))
	
		'Sætter antal registrerede timer for datoen.
		call erDetInt(Request("Timer_son_"& cint(antal_to)))
		if len(trim(Request("Timer_son_"& cint(antal_to)))) > 0 AND isInt = "0" then 
		strTimer_son(cint(antal_to)) = SQLBless(Request("Timer_son_"& cint(antal_to) &""))
		strTimer_kom_son(cint(antal_to)) = Replace(SQLBless2(request("FM_kom_son_"& cint(antal_to) &"")),Chr(34),"&#34;")
		off_son(cint(antal_to)) = request("FM_off_son_"& cint(antal_to) &"")
		else
		strTimer_kom_son(cint(antal_to)) = ""
		strTimer_son(cint(antal_to)) = -1
		off_son(cint(antal_to)) = 0
			if isInt <> "0" then
			varWchar = 1
			end if 
		end if
		
		call erDetInt(Request("Timer_man_"& cint(antal_to)))
		if len(trim(Request("Timer_man_"& cint(antal_to)))) > 0 AND isInt = "0" then  
		strTimer_man(cint(antal_to)) = SQLBless(Request("Timer_man_"& cint(antal_to) &""))
		strTimer_kom_man(cint(antal_to)) = Replace(SQLBless2(request("FM_kom_man_"& cint(antal_to) &"")),Chr(34),"&#34;")
		off_man(cint(antal_to)) = request("FM_off_man_"& cint(antal_to) &"")
		else
		strTimer_kom_man(cint(antal_to)) = ""
		strTimer_man(cint(antal_to)) = -1
		off_man(cint(antal_to)) = 0
			if isInt <> "0" then
			varWchar = 1
			end if 
		end if
		
		call erDetInt(Request("Timer_tir_"& cint(antal_to)))
		if len(trim(Request("Timer_tir_"& cint(antal_to)))) > 0 AND isInt = "0" then 
		strTimer_tir(cint(antal_to)) = SQLBless(Request("Timer_tir_"& cint(antal_to) &""))
		strTimer_kom_tir(cint(antal_to)) = Replace(SQLBless2(request("FM_kom_tir_"& cint(antal_to) &"")),Chr(34),"&#34;")
		off_tir(cint(antal_to)) = request("FM_off_tir_"& cint(antal_to) &"")
		else
		strTimer_kom_tir(cint(antal_to)) = ""
		strTimer_tir(cint(antal_to)) = -1
		off_tir(cint(antal_to)) = 0
			if isInt <> "0" then
			varWchar = 1
			end if 
		end if
		
		call erDetInt(Request("Timer_ons_"& cint(antal_to)))
		if len(trim(Request("Timer_ons_"& cint(antal_to)))) > 0 AND isInt = "0" then 
		strTimer_ons(cint(antal_to)) = SQLBless(Request("Timer_ons_"& cint(antal_to) &""))
		strTimer_kom_ons(cint(antal_to)) = Replace(SQLBless2(request("FM_kom_ons_"& cint(antal_to) &"")),Chr(34),"&#34;")
		off_ons(cint(antal_to)) = request("FM_off_ons_"& cint(antal_to) &"")
		else
		strTimer_kom_ons(cint(antal_to)) = ""
		strTimer_ons(cint(antal_to)) = -1
		off_ons(cint(antal_to)) = 0
			if isInt <> "0" then
			varWchar = 1
			end if 
		end if
		
		call erDetInt(Request("Timer_tor_"& cint(antal_to)))
		if len(trim(Request("Timer_tor_"& cint(antal_to)))) > 0 AND isInt = "0" then 
		strTimer_tor(cint(antal_to)) = SQLBless(Request("Timer_tor_"& cint(antal_to) &""))
		strTimer_kom_tor(cint(antal_to)) = Replace(SQLBless2(request("FM_kom_tor_"& cint(antal_to) &"")),Chr(34),"&#34;")
		off_tor(cint(antal_to)) = request("FM_off_tor_"& cint(antal_to) &"")
		else
		strTimer_kom_tor(cint(antal_to)) = ""
		strTimer_tor(cint(antal_to)) = -1
		off_tor(cint(antal_to)) = 0
			if isInt <> "0" then
			varWchar = 1
			end if 
		end if
		
		call erDetInt(Request("Timer_fre_"& cint(antal_to)))
		if len(trim(Request("Timer_fre_"& cint(antal_to)))) > 0 AND isInt = "0" then 
		strTimer_fre(cint(antal_to)) = SQLBless(Request("Timer_fre_"& cint(antal_to) &""))
		strTimer_kom_fre(cint(antal_to)) = Replace(SQLBless2(request("FM_kom_fre_"& cint(antal_to) &"")),Chr(34),"&#34;")
		off_fre(cint(antal_to)) = request("FM_off_fre_"& cint(antal_to) &"")
		else
		strTimer_kom_fre(cint(antal_to)) = ""
		strTimer_fre(cint(antal_to)) = -1
		off_fre(cint(antal_to)) = 0
			if isInt <> "0" then
			varWchar = 1
			end if 
		end if
		
		call erDetInt(Request("Timer_lor_"& cint(antal_to)))
		if len(trim(Request("Timer_lor_"& cint(antal_to)))) > 0 AND isInt = "0" then 
		strTimer_lor(cint(antal_to)) = SQLBless(Request("Timer_lor_"& cint(antal_to) &""))
		strTimer_kom_lor(cint(antal_to)) = Replace(SQLBless2(request("FM_kom_lor_"& cint(antal_to) &"")),Chr(34),"&#34;")
		off_lor(cint(antal_to)) = request("FM_off_lor_"& cint(antal_to) &"")
		else
		strTimer_kom_lor(cint(antal_to)) = ""
		strTimer_lor(cint(antal_to)) = -1
		off_lor(cint(antal_to)) = 0
			if isInt <> "0" then
			varWchar = 1
			end if 
		end if
		'Slut sæt antal timer og kommentarer
		
	strJobnr(cint(antal_to)) = Request("FM_jobnr_"& cint(antal_to) &"")
	
	
	
	'**** Henter oplysninger om det job der er indtastet timer på. ****
	if cint(antal_de) >= cint(antal_to) then
			strAktId(cint(antal_to)) = request("FM_Aid_"& cint(antal_to) &"")
			if len(strAktId(cint(antal_to))) <> 0 then
				strSQL = "SELECT job.id AS jobid, navn, fakturerbar, jobnavn, kkundenavn, jobknr, fastpris, job.serviceaft FROM aktiviteter, job, kunder WHERE aktiviteter.id = "& strAktId(cint(antal_to)) &" AND Jobnr = "& strJobnr(cint(antal_to)) &" AND kunder.Kid = jobknr"
				oRec.Open strSQL, oConn, 3
				if Not oRec.EOF then
					 intjobid = oRec("jobid")
					 strJobnavn = oRec("Jobnavn")
			 		 strJobknavn = oRec("kkundenavn")
					 strJobknr = oRec("Jobknr") 	
					 strAktNavn = oRec("navn")
					 strFakturerbart = oRec("fakturerbar")
					 intServiceAft = oRec("serviceaft")
					 
					 select case strFakturerbart
					 '** fakbar aktivitet
					 case 1
					 tfaktimvalue = 1
					 '** Kørsels aktivitet
					 case 2
					 tfaktimvalue = 5
					 case else
					 '** ikke fakbar aktivitet
					 '** Interne job sættes til 0 (se nedenfor)
					 tfaktimvalue = 2
					 end select
					 strFastpris = oRec("fastpris")
					 		
							'*** tjekker for alternativ timepris på aktivitet
							call alttimepris(strAktId(cint(antal_to)))
							
							'** Er der alternativ timepris på jobbet
							if foundone = "n" then
								call alttimepris(0)
							end if
							
							
							
				end if
				oRec.Close
			end if
	else 
		'** internt job ******** 
		'** interne job bruges ikke mrer pr. 5/10-2005 men koden beholdes
		'** for at holde konsistensen og undgå fejl.
		
		strAktId(cint(antal_to)) = "0" '"909" & strJobnr(cint(antal_to))
			oRec.Open "SELECT jobnavn, kkundenavn, jobknr FROM job, kunder WHERE Jobnr ="& strJobnr(cint(antal_to)) &" AND kunder.Kid = jobknr", oConn, 3
			if Not oRec.EOF then
				 strJobnavn = oRec("Jobnavn")
		 		 strJobknavn = oRec("kkundenavn")
				 strJobknr = oRec("Jobknr") 	
				 strAktNavn = strJobnavn
				 tfaktimvalue = 0
				 intServiceAft = 0
			end if
			oRec.Close
	end if
			
	
		dagTypeNr = 7
		Redim dag(7)
		dag(1) = "son"
		dag(2) = "man"
		dag(3) = "tir"
		dag(4) = "ons"
		dag(5) = "tor"
		dag(6) = "fre"
		dag(7) = "lor"
		
		Redim Dato(7)
		dato(1) = strDatoSon
		dato(2) = strDatoMan
		dato(3) = strDatoTir
		dato(4) = strDatoOns
		dato(5) = strDatoTor
		dato(6) = strDatoFre
		dato(7) = strDatoLor
		
		Redim antTimer(7)
		Redim strTKomm(7)
		Redim intOff(7)
		antTimer(1) = strTimer_son(cint(antal_to))
		antTimer(2) = strTimer_man(cint(antal_to))
		antTimer(3) = strTimer_tir(cint(antal_to))
		antTimer(4) = strTimer_ons(cint(antal_to))
		antTimer(5) = strTimer_tor(cint(antal_to))
		antTimer(6) = strTimer_fre(cint(antal_to))
		antTimer(7) = strTimer_lor(cint(antal_to))
		
		strTKomm(1) = strTimer_kom_son(cint(antal_to))
		strTKomm(2) = strTimer_kom_man(cint(antal_to))
		strTKomm(3) = strTimer_kom_tir(cint(antal_to))
		strTKomm(4) = strTimer_kom_ons(cint(antal_to))
		strTKomm(5) = strTimer_kom_tor(cint(antal_to))
		strTKomm(6) = strTimer_kom_fre(cint(antal_to))
		strTKomm(7) = strTimer_kom_lor(cint(antal_to))
		
		intOff(1) = off_son(cint(antal_to))
		intOff(2) = off_man(cint(antal_to))
		intOff(3) = off_tir(cint(antal_to))
		intOff(4) = off_ons(cint(antal_to))
		intOff(5) = off_tor(cint(antal_to))
		intOff(6) = off_fre(cint(antal_to))
		intOff(7) = off_lor(cint(antal_to))
		
	for dagTypeNr = 1 to 7
		
		if len(intOff(dagTypeNr)) <> 0 then
		intOff(dagTypeNr) = intOff(dagTypeNr)
		else
		intOff(dagTypeNr) = 0
		end if
		
		if antTimer(dagTypeNr) <> "-1" then
		
		'*** tilpasse dato format til mySQL ****
		datoUS = ConvertDateYMD(dato(dagTypeNr))
		
			'Tjek om job/akt allerede er registreret for den valgte dato
			oRec.Open "SELECT timer, Tjobnr, TAktivitetId FROM timer WHERE TAktivitetId = "& strAktId(cint(antal_to)) &" AND Tjobnr = "&strJobnr(cint(antal_to))&" AND Tmnr = "& strMnr &" AND Tdato = '"& datoUS &"'", oConn, 3  
			
			'son, man, tir, ons ,tor, fre, lor
			if oRec.EOF then
				if antTimer(dagTypeNr) <> "0" then
				oConn.Execute("INSERT INTO timer (TJobnr, TJobnavn, Tmnr, Tmnavn, Tdato, "_
				&" Timer, Timerkom, Tknavn, Tknr, TAktivitetId, TAktivitetNavn, "_
				&" Tfaktim, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
				&" editor, kostpris, offentlig, seraft) VALUES"_
				& "(" & strJobnr(cint(antal_to)) & ", '"&strJobnavn&"', " & strMnr & ", '" & Cstr(strMnavn) & "', '"& datoUS &"', "_
				&" "& antTimer(dagTypeNr) &", '"& strTKomm(dagTypeNr) &"', "_
				&" '" & SQLBless2(strJobknavn) & "', " & strJobknr & ", "_
				&" "& strAktId(cint(antal_to)) &", '"& SQLBless2(strAktNavn) &"', "_
				&" "& tfaktimvalue &", "& strYear &", "& SQLBless(intTimepris) &", "_
				&" '"&year(now)&"/"&month(now)&"/"&day(now)&"', '"& strFastpris  &"', "_
				&" '" & time & "', '"& session("user") &"', "& dblkostpris &", "& intOff(dagTypeNr) &", "& intServiceAft &")")
				end if
			else
				if antTimer(dagTypeNr) = "0" then
				oConn.execute("DELETE FROM timer"_
				&" WHERE Tjobnr = "& strJobnr(cint(antal_to)) & ""_
				&" AND Tmnr = "& strMnr & ""_
				&" AND Tdato = '" & datoUS & "' AND TAktivitetId = "& strAktId(cint(antal_to))& "")
				else
					if cdbl(oRec("timer")) <> cdbl(antTimer(dagTypeNr)) then
					oConn.Execute("UPDATE timer SET"_
					&" Timer = "& antTimer(dagTypeNr) &", Timerkom = '"& strTKomm(dagTypeNr) &"', "_
					&" Taar = "& strYear &", TimePris = "& SQLBless(intTimepris) &", "_
					&" seraft = "& intServiceAft &", kostpris = "& dblkostpris &", "_
					&" tastedato = '"&year(now)&"/"&month(now)&"/"&day(now)&"', "_
					&" editor = '"& session("user") &"', offentlig = "& intOff(dagTypeNr) &""_
					&" WHERE Tjobnr = "& strJobnr(cint(antal_to)) & ""_
					&" AND Tmnr = "& strMnr & ""_
					&" AND Tdato = '" & datoUS & "' AND TAktivitetId = "& strAktId(cint(antal_to))& "")
					else
					strSQLupd = "UPDATE timer SET"_
					&" Timer = "& antTimer(dagTypeNr) &", Timerkom = '"& strTKomm(dagTypeNr) &"', "_
					&" editor = '"& session("user") &"', offentlig = "& intOff(dagTypeNr) &""_
					&" WHERE Tjobnr = "& strJobnr(cint(antal_to)) & ""_
					&" AND Tmnr = "& strMnr & ""_
					&" AND Tdato = '" & datoUS & "' AND TAktivitetId = "& strAktId(cint(antal_to))
					oConn.Execute(strSQLupd)
					end if
				end if
			end if
			oRec.close
		
		else '** Bruges ikke efter redirect ***
			
			'if antTimer(dagTypeNr) > 24 then
			'errortype = 18
			'call showError(errortype)>
			'<
			'errorPrinted = 1
			'end if
			
			if varWchar = "1" AND errorPrinted <> "1" then
			errortype = 17
			call showError(errortype)%>
			<%
			errorPrinted = 1
			end if
			
		end if
	next
next				

'*** Afslutter uge ****

if len(request("FM_afslutuge")) <> 0 then
cDateUge = year(request("FM_afslutuge_sidstedag"))&"/"&month(request("FM_afslutuge_sidstedag"))&"/"&day(request("FM_afslutuge_sidstedag"))
cDateAfs = year(now)&"/"&month(now)&"/"&day(now)

'Response.write cDate(cDateUge) &" >= " & cdate(cDateAfs) 

if cDate(cDateUge) >= cdate(cDateAfs) then
intStatusAfs = 2 '** Afsluttet korrekt
else
intStatusAfs = 1 '** Afsluttet forsent
end if

cDateAfs = cDateAfs & " " & time

strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status) VALUES ('"& cDateUge &"', '"& cDateAfs &"', "& strMnr &", "& intStatusAfs &")" 
oConn.execute(strSQLafslut)

end if

'******************************** output på side **********************************

Response.redirect "timereg.asp?filterA_B="&filterA_B&"&searchstring="&searchstring&"&FM_use_me="&request("FM_use_me")&"&strdag="&strdag&"&strmrd="&strmrd&"&straar="&straar

%>
<!--include file="../inc/regular/vmenu.asp"-->
<!--include file="../inc/regular/calender.asp"-->
<!--<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
<br>
<img src="../ill/header_reg.gif" alt="" border="0"><hr align="left" width="750" size="1" color="#000000" noshade>
<br>	
<if errorPrinted = "0" then%>
<table cellspacing="0" cellpadding="0" border="0" width="520" bgcolor="#FFFFFF">
	<tr bgcolor="5582D2">
		<td width="8" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="21" alt="" border="0"></td>
		<td valign="top"><img src="../ill/tabel_top.gif" width="507" height="1" alt="" border="0"></td>
		<td align=right valign=top><img src="../ill/tabel_top_right.gif" width="8" height="21" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="270" alt="" border="0"></td>
		<td bgcolor="#FFFFFF" style="padding-left:20;"><img src="../ill/info.gif" alt="" border="0"><br><br>Der er netop indtastet timer for <b><=strMnavn%></b> 
		<br>Time-registreringerne er <font color="LimeGreen"><b>godkendt</b></font> og gemt i TimeOut. 
		<br>Du kan altid vende tilbage til timeregistreringssiden og ændre de indtastede timer for den valgte medarbejder.<br><br>
		Hvis du ønsker at slette en indtastning skal du bare skrive et <b>"0" (nul)</b> istedet <br>for det allerede indtastede timeantal.<br>
		<br><br><br>
		<a href="timereg.asp?strdag=<=strdag%>&strmrd=<=strmrd%>&straar=<=straar%>&searchstring=<=searchstring%>&FM_use_me=<=strMnr%>"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0">&nbsp;&nbsp;Timeregistrering</a>
		<br><br><br>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="270" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td valign="bottom"><img src="../ill/tabel_top.gif" width="507" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr></table>
	<br><br><br>
	<br>
	<br>
<end if%>
</div>
<br>
<br>-->

<!--#include file="../inc/regular/footer_inc.asp"-->
<%end if%>