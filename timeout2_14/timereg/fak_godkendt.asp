<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/functions_inc.asp"-->  
<%

a = 0
dim aval
redim aval(a)

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<!--#include file="inc/fak_layout_inc.asp"-->
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	id = request("id")
	showmoms = 0
	
	if len(request("jobid")) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
	if len(request("aftid")) <> 0 then
	aktid = request("aftid")
	else
	aktid = 0
	end if
	
	'*** aftale side variable ****
	
	if len(request("aftaleid")) <> 0 then
	aftid = request("aftaleid")
	else
	aftid = 0
	end if
	
	if len(request("FM_aftnr")) <> 0 then
	sogaftnr = request("FM_aftnr")
	else
	sogaftnr = 0
	end if
	
	if len(request("thiskid")) <> 0 then
	thiskid = request("thiskid")
	else
	thiskid = 0
	end if
	
	'***
	
	'*** Faktura info ****
	if jobid <> 0 then
	strSQL_sel = " jobid, jobnavn, jobnr,"
	strSQL_lftJ = " job ON (id = jobid) "
	else
	strSQL_sel = " s.navn AS aftnavn, s.aftalenr, "
	strSQL_lftJ = " serviceaft s ON (s.id = aftaleid) "
	end if
	
	strSQL = "SELECT f.editor, tidspunkt, fid, faknr, fakdato, timer, beloeb, kommentar, "_
	&" tidspunkt, fakadr, att, faktype, konto, modkonto, "& strSQL_sel &" "_
	&" vismodland, vismodatt, vismodtlf, vismodcvr, visafstlf, visafsemail,"_
	&" visafsswift, visafsiban, visafscvr, moms, enhedsang, f.varenr, jobbesk, valuta, kurs "_
	&" FROM fakturaer f LEFT JOIN "& strSQL_lftJ &" WHERE fid = "& id
	
	'Response.write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	strEditor = oRec("editor")
	fakTidspunkt = oRec("tidspunkt")
	intFaktype = oRec("faktype")
	intKundeid = oRec("fakadr")
	intAtt = oRec("att")
	intKonto = oRec("konto")
	intModKonto = oRec("modkonto")
	intFaktureretTimer = oRec("timer")
	intFaktureretBelob = oRec("beloeb")
	strKomm = oRec("kommentar")
	
	if intFaktype = 1 then
	intFaktureretBelob = -(intFaktureretBelob)
	end if
	
	if jobid <> 0 then
	strJobnavn = oRec("jobnavn")
	intJobnr = oRec("jobnr")
	else
	strAftNavn = oRec("aftnavn")
	strAftVarenr = oRec("varenr")
	intAftnr = oRec("aftalenr")
	end if
	
	fakdato = oRec("fakdato") 
	varFaknr = oRec("faknr")
	strKom = oRec("kommentar")
	
	intVismodland = oRec("vismodland")
	intVismodatt = oRec("vismodatt")
	intVismodtlf = oRec("vismodtlf")
	intVismodcvr = oRec("vismodcvr")
	
	
	intVisafstlf = oRec("visafstlf")
	intVisafsemail = oRec("visafsemail")
	intVisafsswift = oRec("visafsswift")
	intVisafsiban = oRec("visafsiban")
	intVisafscvr = oRec("visafscvr")
	
	intMoms = oRec("moms")
	if intFaktype = 1 then
	intMoms = -(intMoms)
	end if
	
	select case oRec("enhedsang")
	case 1	
	enhed = "Stk. pris"
	case 2
	enhed = "Enhedspris"
	case else
	enhed = "Timepris"
	end select
	
	strJobBesk = oRec("jobbesk")
	
	fakValuta = oRec("valuta")
	fakValutaKurs = oRec("kurs")
	
	end if
	oRec.close 
	
	
	
	
	
	 
	
	' *** Modtager ****
	strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr FROM kunder WHERE Kid =" & intKundeid		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			intKnr = oRec("kkundenr")
			strKnavn = oRec("kkundenavn")
			strKadr = oRec("adresse")
			strKpostnr = oRec("postnr")
			strBy = oRec("city")
			strLand = oRec("land")
			intKid = oRec("Kid")
			intCVR = oRec("cvr")
			intTlf = oRec("telefon")
		end if
		
		oRec.close
		
	'*** Att ******
	select case intAtt
	case 991
	showAtt = "Den økonomi ansvarlige"
	case 992
	showAtt = "Regnskabs afd."
	case 993
	showAtt = "Administrationen"
	case else
	
	strSQL3 = "SELECT id, navn FROM kontaktpers WHERE id="& intAtt
	oRec3.open strSQL3, oConn, 3
	if not oRec3.EOF then
	showAtt = oRec3("navn")
	end if
	oRec3.close
	
	if len(showAtt) <> 0 then
	showAtt = showAtt
	else
	showAtt = ""
	end if
	
	end select	
	
	
	'*** Afsender ***
	
		strSQL = "SELECT adresse, postnr, city, land, telefon, email, regnr, kkundenavn, kontonr, cvr, bank, swift, iban, kid FROM kunder WHERE useasfak = 1"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			yourbank = oRec("bank")
			yourRegnr = oRec("regnr")
			yourKontonr = oRec("kontonr")
			yourCVR = oRec("cvr")
			yourNavn = oRec("kkundenavn")
			yourAdr = oRec("adresse")
			yourPostnr = oRec("postnr")
			yourCity = oRec("city")
			yourLand = oRec("land")
			yourEmail = oRec("email")
			yourTlf = oRec("telefon")
			yourSwift = oRec("swift")
			yourIban = oRec("iban")
			afskid = oRec("kid")
		end if
		oRec.close
		
		
		
		'**** Konto / Modkonto ***
		
		'** modtagers konto **
		if len(intKonto) <> 0 then
		intKonto = intKonto
		else
		intKonto = 0
		end if 
		
		
		strKontoNavn = "-"
		intMomskode = "-"
		intKvotient = "-"
		intKontonr = "0"
		
		strSQL2 = "SELECT momskode, kvotient, kontoplan.navn, kontonr FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontoplan.kontonr = "& intKonto
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		strKontoNavn = oRec2("navn")
		intMomskode = oRec2("momskode")
		
		select case oRec2("kvotient")
		case 5 
		intKvotient = "0,25 %"
		case 21
		intKvotient = "0,5 %"
		case 1
		intKvotient = "0"
		end select
		
		intKontonr = oRec2("kontonr")
		end if
		oRec2.close
		
		
		
		'** afsenders konto **
		if len(intModKonto) <> 0 then
		intModKonto = intModKonto
		else
		intModKonto = 0
		end if 
		
		strModKontoNavn = "-"
		intModKontonr = "0"
		
		strSQL2 = "SELECT kontoplan.navn, kontonr FROM kontoplan WHERE kontoplan.kontonr = "& intModKonto
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		strModKontoNavn = oRec2("navn")
		intModKontonr = oRec2("kontonr")
		end if
		oRec2.close
		
		
		'*********************************************************************************
		'**** Fakturaen ***
		'*********************************************************************************
		%>
		<div id="sindhold" style="position:absolute; left:10; top:5; visibility:visible;">
		<br><br>
		<%
		'*** fak header layout *******
		call fakheader() 
	
		'*** Maintable layout ********
		call maintable
		
		'*** Modtager boks layout ****
		call modtager_layout()
		
		'*** Maintable 2 layout ********
		call maintable_2
		
		'*** Afsender boks layout ****										
		call afsender_layout()
		
		'*** Maintable 3 layout ********
		call maintable_3
		
												
		'*** Vedr. (upspecificering top) ***
		call vedr
		
		
		
		
		
		
		'**** Udspecificering ******
		
		if jobid <> 0 then
					'*** Fakturering af job / Aktiviteter ***
					h = 100
					
					strSQL = "SELECT antal, beskrivelse, aktpris, fakid, "_
					&" enhedspris, aktid, showonfak, valuta, kurs FROM faktura_det WHERE fakid = "& id &" AND showonfak = 1 ORDER BY beskrivelse"
					
					oRec.open strSQL, oConn, 3 
					while not oRec.EOF 
						
						aktEhpris = oRec("enhedspris")
						aktPris = oRec("aktpris")
						if intFaktype = 1 then
						aktPris = -(aktPris)
						aktEhpris = -(aktEhpris)
						end if
					
					strFakdet = strFakdet &"<tr><td valign='top'>&nbsp;</td>"_
					&"<td><img src='../ill/blank.gif' width='20' height='1' alt='' border='0'></td>"_
					&"<td valign=top align=right style='padding-right:2;'>"& formatnumber(oRec("antal"), 2) &"</td>"_
					&"<td colspan=2 valign=top style='padding-left:5;'>" & oRec("beskrivelse") &"</td>"_
					&"<td valign=top align=right width=100 style='padding-right:30;'>" & formatcurrency(aktEhpris, 2) &"</td>"_
					&"<td valign='top' align=right colspan=2 width=150 style='padding-right:30;'>"& formatcurrency(aktPris,2) &"</td>"_
					&"<td valign='top' align='right'>&nbsp;</td>"_
					&"</tr>"
					
					h = h + 20			
								
								fakDetValuta = oRec("valuta")
								fakDetValutaKurs = oRec("kurs")
								
								ehpris = replace(oRec("enhedspris"), ",", ".")
								
								'**** Medarbejder timer der vises på print ***
								strSQL2 = "SELECT fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, showonfak FROM fak_med_spec WHERE fakid = "& id &" AND aktid = "& oRec("aktid") &" AND enhedspris = "& ehpris &" AND showonfak = 1"
								'Response.write strSQL2
								'Response.flush
								oRec2.open strSQL2, oConn, 3 
								while not oRec2.EOF 
								
								medEhpris = oRec2("enhedspris")
								medBelob = oRec2("beloeb")
								if intFaktype = 1 then
								medBelob = -(medBelob)
								medEhpris = -(medEhpris)
								end if
								
								strFakdet = strFakdet &"<tr><td valign='top'>&nbsp;</td>"_
								&"<td><img src='../ill/blank.gif' width='20' height='1' alt='' border='0'></td>"_
								&"<td valign=top align=right style='padding-right:2;'><font class=megetlillesort>"& formatnumber(oRec2("fak"), 2) &"</td>"_
								&"<td colspan=2 valign=top style='padding-left:5;'><font class=megetlillesort>" & oRec2("tekst") &"</td>"_
								&"<td valign=top align=right style='padding-right:30;'><font class=megetlillesort>" & formatcurrency(medEhpris, 2) &"</td>"_
								&"<td valign='top' align=right colspan=2 style='padding-right:30;'><font class=megetlillesort>"& formatcurrency(medBelob,2) &"</td>"_
								&"<td valign='top' align='right'>&nbsp;</td>"_
								&"</tr>"
								
								h = h + 20
								
								oRec2.movenext
								wend
								oRec2.close 
								
								'**** til ikke printbar medarbjeder timeoversigt (højre side) *** 
								if lastaktid <> oRec("aktid") then 
								strMedarbFakdet = strMedarbFakdet &"<tr><td colspan=4><b>"& oRec("beskrivelse") &"</b></td></tr>"
								end if
								
								mf = 0
								
								strSQL2 = "SELECT fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, showonfak, "_
								&" valuta, kurs FROM fak_med_spec WHERE fakid = "& id &" AND aktid = "& oRec("aktid") &" AND enhedspris = "& ehpris &" AND fak <> 0"
								oRec2.open strSQL2, oConn, 3 
								while not oRec2.EOF 
								
								medEhpris = oRec2("enhedspris")
								medBelob = oRec2("beloeb")
								if intFaktype = 1 then 'kreditnota
								medBelob = -(medBelob)
								medEhpris = -(medEhpris)
								end if
								
								fmsValuta = oRec("valuta")
								fmsValutaKurs = oRec("kurs")
								
								strMedarbFakdet = strMedarbFakdet &"<tr>"_
								&"<td valign=top align=right style='padding-right:5;' width=30><font class=megetlillesort>"& formatnumber(oRec2("fak"), 2) &"</td>"_
								&"<td valign=top style='padding-left:5;' width=70><font class=megetlillesort>" & left(oRec2("tekst"), 10) &"</td>"_
								&"<td valign=top align=right style='padding-right:10;' width=70><font class=megetlillesort>" & formatcurrency(medEhpris, 2) &"</td>"_
								&"<td valign=top align=right style='padding-right:10;' width=70><font class=megetlillesort>"& formatcurrency(medBelob,2) &"</td>"_
								&"</tr>"
								
								medarbejdertimer = medarbejdertimer + oRec2("fak")
								medarbejdersum = medarbejdersum + oRec2("beloeb")
								
								mf = 1
								oRec2.movenext
								wend
								oRec2.close 
								
								if mf = 0 then
								strMedarbFakdet = strMedarbFakdet &"<tr><td colspan=4 style='padding-left:20;'>-</td></tr>"
								end if
									
					lastaktid = oRec("aktid")			
					oRec.movenext
					wend
					oRec.close 
					
					
					if len(medarbejdertimer) <> 0 then
					medarbejdertimer = medarbejdertimer
					else
					medarbejdertimer = 0
					end if
					
					if len(medarbejdersum) <> 0 then
					medarbejdersum = medarbejdersum
					else
					medarbejdersum = 0
					end if
					
					if medarbejdertimer <> intFaktureretTimer then
					bgthis = "#ffffe1"
					else
					bgthis = "#ffffff"
					end if
					
					'** Er der overensstemmelse mellem antal fakturatimer og antal medarbjeder timer og beløb?
					if intFaktype = 1 then
					medarbejdersum = -(medarbejdersum)
					end if
					
					strMedarbFakdet = strMedarbFakdet &"<tr><td bgcolor="& bgthis &" colspan=4 class=lille><br>"_
					&" Medarbejder timer: "& formatnumber(medarbejdertimer, 2) &"<br>"_
					&" Medarbejder beløb: "& formatcurrency(medarbejdersum, 2) &"</td></tr>"
					
					
		
		else '**** Fakturering af Aftaler ***
		
				
				'** Betalingsbetingelder **
				strKom = "Netto kontant 8 dage."
				
				strFakdet = strFakdet &"<tr><td valign='top'>&nbsp;</td>"_
				&"<td><img src='../ill/blank.gif' width='20' height='1' alt='' border='0'></td>"_
				&"<td valign=top align=right style='padding-right:2;'>"& formatnumber(intFaktureretTimer, 2) &"</td>"_
				&"<td colspan=2 valign=top style='padding-left:5;'>" & strKomm &"</td>"
			
				if intFaktureretTimer <> 0 AND intFaktureretBelob <> 0 then
				intstkpris = (intFaktureretBelob/intFaktureretTimer)
				else
				intstkpris = 0
				end if
				
				strFakdet = strFakdet & "<td valign=top align=right width=100 style='padding-right:30;'>" & formatcurrency(intstkpris, 2) &"</td>"_
				&"<td valign='top' align=right colspan=2 width=150 style='padding-right:30;'>"& formatcurrency(intFaktureretBelob,2) &"</td>"_
				&"<td valign='top' align='right'>&nbsp;</td>"_
				&"</tr>" 
		
				
				strMedarbFakdet = "Der kan ikke angives timer på en medarbejder når det er en aftale der faktureres."
				
		end if
		
		call udspecificering(strFakdet)
		
		'**** Totaler og moms ***
		call totogmoms()
		
		
		'*** Komm. og bet. betingelser ***
		call komogbetbetingelser									
		%>
		</div>
											
		<%
		'**** Højreside (ikke printbar) ***
		call ikkeprintbar
		
		
											
											
											
											
											
											
											
	
	
	
	
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->


