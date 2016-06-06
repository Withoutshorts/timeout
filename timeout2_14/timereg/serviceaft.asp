<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<script language=javascript>
function showhidefordeldiv(){
		if (document.getElementById("FM_fordel").checked == true) {
		//document.getElementById("FM_enhederrest").value = document.getElementById("FM_antal").value
		document.getElementById("fordeldiv").style.visibility = "visible"
		document.getElementById("fordeldivpil").style.visibility = "visible"
		} else {
		document.getElementById("fordeldiv").style.visibility = "hidden"
		document.getElementById("fordeldivpil").style.visibility = "hidden"
		}
} 

function isNum(passedVal){
	invalidChars = " /:;<>abcdefghijklmnopqrstuvwxyzæøå"
	if (passedVal == ""){
	return false
	}
	
	for (i = 0; i < invalidChars.length; i++) {
	badChar = invalidChars.charAt(i)
		if (passedVal.indexOf(badChar, 0) != -1){
		return false
		}
	}
	
		for (i = 0; i < passedVal.length; i++) {
			if (passedVal.charAt(i) == ".") {
			return true
			}
			else {
				if (passedVal.charAt(i) < "0") {
				return false
				}
					if (passedVal.charAt(i) > "9") {
					return false
					}
				}
			return true
			} 
			
}

function validZip(inZip){
	if (inZip == "") {
	return true
	}
	if (isNum(inZip)){
	return true
	}
	return false
}


function beregnenhederrest(mth, yearthis){
var valtotopr = 0;
var newvaltot = 0;
var newvaltotberegnet = 0;
var newval = 0;
var newvalopr = 0;

	if (!validZip(document.getElementById("FM_"+mth+"_"+yearthis).value)){
	alert("Der er angivet et ugyldigt tegn.")
	document.getElementById("FM_"+mth+"_"+yearthis).value = document.getElementById("FM_"+mth+"_"+yearthis+"_opr").value;
	return false
	} else {
		
		newvalopr = document.getElementById("FM_"+mth+"_"+yearthis+"_opr").value;
		if (document.getElementById("FM_"+mth+"_"+yearthis).value != ""){
		newval = document.getElementById("FM_"+mth+"_"+yearthis).value.replace(",",".");
		document.getElementById("FM_"+mth+"_"+yearthis).value = newval;
		}
		
		valtotopr = parseFloat(document.getElementById("FM_enhederrest").value)
		newvaltot = parseFloat(valtotopr) - parseFloat(newvalopr) 
		newvaltotberegnet = parseFloat(newvaltot) + parseFloat(newval) 
		document.getElementById("FM_enhederrest").value = parseFloat(newvaltotberegnet)
		document.getElementById("FM_"+mth+"_"+yearthis+"_opr").value = parseFloat(newval)
		
	}
}
</script>

<%

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	session.lcid = 1030
	
	func = request("func")
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	kid = request("kundeid")
	
	forny = request("forny")
	
	
	select case func
	
	case "dbopr", "dbred"
	errorthis = 0
	%>
	<!--#include file="inc/isint_func.asp"-->
	<%
	'*** Her indsættes en ny type i db ****
	call erDetInt(request("FM_antal"))
	if isInt > 0 then
	errorthis = 1
	end if
	isInt = 0
	
	call erDetInt(request("FM_pris"))
	if isInt > 0 then
	errorthis = 2
	end if
	isInt = 0
	
	call erDetInt(request("FM_aftnr"))
	if isInt > 0 then
	errorthis = 3
	end if
	isInt = 0
	
	if len(request("FM_navn")) = 0 then
	errorthis = 4
	end if
		
		
		if len(request("FM_aftnr")) <> 0 then
		intAftnr = request("FM_aftnr")
		else
		intAftnr = 0
		end if
		
		if errorthis = 0 then
			'*** Findes aftale nr allerede?? ***
			strSQL = "SELECT aftalenr FROM serviceaft WHERE id <> "& id & " AND aftalenr = "& intAftnr 
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
				errorthis = 5
			end if
			oRec.close 
		end if
		
		'**** Tjekker dato format ******
		if len(request("FM_start_aar")) = 2 then
		usestaar = 2000 + request("FM_start_aar")
		else
		usestaar = request("FM_start_aar")
		end if
		
		if len(request("FM_slut_aar")) = 2 then
		useslaar = 2000 + request("FM_slut_aar")
		else
		useslaar = request("FM_slut_aar")
		end if
		
		dStdato = usestaar &"/"& request("FM_start_mrd") &"/"& request("FM_start_dag")
		dSldato = useslaar &"/"& request("FM_slut_mrd") &"/"& request("FM_slut_dag")
		
		
		if isdate(dStdato) <> true then
		errorthis = 6
		end if
		
		if isdate(dSldato) <> true then
		errorthis = 6
		end if
		'**** End if *****
		
		if request("FM_fordel") = "1" AND (request("FM_enhederrest") <> replace(request("FM_antal"), ",", ".")) then
		errorthis = 7
		end if
		
		if request("kundeid") = 0 then
		errorthis = 8
		end if
		
		if errorthis <> 0 then
		%>
		<!--#include file="../inc/regular/header_hvd_inc.asp"-->
		<div id="sindhold" style="position:absolute; left:40px; top:40px; visibility:visible;">
		<table cellspacing="0" cellpadding="0" border="0" width="220" bgcolor="#ffffff"><tr>
		<td><font color=red><b>Fejl!</b>
		<ul>
		<%
		select case errorthis
		case 4  
		%>
		<li>Der mangler at blive angivet et <b>aftale navn</b>.
		<%
		case 2
		%>
		<li><b>Pris</b> er ikke angivet som et tal.
		<%
		case 1
		%>
		<li><b>Enheder</b> er ikke angivet som et tal.
		<%
		case 3
		%>
		<li><b>Aftale nr.</b> er ikke angivet som et tal.
		<%
		case 5
		%>
		<li>En aftale med det valgte <b>aftale nr.</b> findes allerede.
		<%
		case 6
		%>
		<li>En af de valgte <b>datoer</b> (start eller slutdato) er <b>ikke en gyldig dato</b>.<br> F.eks 31 feb. 2006
		<%
		case 7
		%>
		<li><b>Enheder tildelt på aftale</b> og enheder fordelt på måneder stemmer ikke overens.
		<%
		'Response.write request("FM_fordel") & " " & (request("FM_enhederrest") &"<>"& replace(request("FM_antal"), ",", ".")) 
		case 8
		%>
		<li>Der er ikke valgt en kontakt.
		<%
		end select%>
		</ul> </font><br><br>
		<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
		</td>
		</tr>
		</table>
		</div>
		<%
		else
		
		
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		intAftnr = intAftnr 'request("FM_aftnr")
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		intStatus = request("FM_status")
		
		if len(request("FM_antal")) <> 0 then
		dblAntal = replace(request("FM_antal"), ",", ".")
		else
		dblAntal = 0
		end if
		
		if len(request("FM_pris")) <> 0 then
		dblpris = replace(request("FM_pris"), ",", ".")
		else
		dblpris = 0
		end if
		
		
		strVarenr = request("FM_varenr")
		
		intAdvitype = request("FM_advitype")
		intAdvihvor = request("FM_advihvor")
		
		if len(request("FM_udlob_aldrig")) <> 0 then
		intAfg = request("FM_udlob_aldrig")
		else
		intAfg = 0
		end if
		
		strBesk = request("FM_besk")
		
		if len(request("FM_fordel")) <> 0 then
		intFordel = request("FM_fordel")
		else
		intFordel = 0
		end if
		
		if len(request("FM_overforsaldo")) <> 0 then
		intOverforSaldo = replace(request("FM_overforsaldo"), ",", ".") 
		else
		intOverforSaldo = 0
		end if
		
		
		valuta = request("FM_valuta")
		
		if func = "dbopr" then
		strSQL = "INSERT INTO serviceaft "_
		&" (navn, editor, dato, "_
		&" stdato, sldato, perafg, status, "_
		&" kundeid, enheder, pris, "_
		&" varenr, advitype, advihvor, besk, aftalenr, fordel, overfortsaldo, valuta) "_
		&" VALUES ('"& strNavn &"', "_
		&" '"& strEditor &"', "_
		&" '"& strDato &"', "_
		&" '"& dStdato &"', "_
		&" '"& dSldato &"', "_
		&" "& intAfg &", "_
		&" "& intStatus &", "_
		&" "& kid &", "_
		&" "& dblAntal &", "_
		&" "& dblpris &", "_
		&" '"& strVarenr & "', "_
		&" "& intAdvitype &", "_
		&" "& intAdvihvor &", '"& strBesk &"', "& intAftnr &", "_
		&" "& intFordel &", "& intOverforSaldo &", "& valuta &" "_
		&")"
		
		'Response.write strSQL
		oConn.execute(strSQL)
				

                '** finder ID ***'
                strSQL3 = "SELECT id FROM serviceaft ORDER BY id DESC "
						oRec.open strSQL3, oConn, 3 
						if not oRec.EOF then
							newseraftid = oRec("id") 
						end if
						oRec.close 


				'*** Slår "er fornyet til" ****
				if forny = 1 then
					useoldid = request("id")
					oConn.execute("UPDATE serviceaft SET erfornyet = 1 WHERE id =" & useoldid)
					
					'** Luk gammelt aftale ***
					if request("FM_luk") = 1 then
					oConn.execute("UPDATE serviceaft SET status = 0 WHERE id =" & useoldid)
					end if
					
					
					'** Luk gamle job ***
					if request("FM_luk_job") = 1 then
					oConn.execute("UPDATE job SET jobstatus = 0 WHERE serviceaft =" & useoldid)
					end if
					
					
					'*** Overfør alle gamle job der kører på gammel aftale til denne nye aft.
					if request("FM_overfor") = 1 then
					oConn.execute("UPDATE job SET serviceaft = "& newseraftid &" WHERE serviceaft =" & useoldid)
					end if
					
				end if
			
			'*** Opretter folder ****
			strSQL3 = "INSERT INTO foldere (editor, dato, jobid, kundeid, kundese, navn) "_
			&" VALUES ('"& strEditor &"', '"& strDato &"', 0, "& kid &", 1, '"& strNavn &"')"
			oConn.Execute(strSQL3)
			
			
			
		else
		
		strSQL = "UPDATE serviceaft SET navn ='"& strNavn &"', "_
		&" editor = '" &strEditor &"', "_
		&" dato = '" & strDato &"', "_
		&" stdato = '"& dStdato &"', "_
		&" sldato = '"& dSldato &"', "_
		&" status = "& intStatus &", "_
		&" kundeid = "& kid &", "_
		&" enheder = "& dblAntal &", "_
		&" pris = "& dblpris &", "_
		&" varenr = '"& strVarenr &"', "_
		&" advitype = "& intAdvitype &", "_
		&" advihvor = "& intAdvihvor &", "_
		&" perafg = "& intAfg &", besk = '"& strBesk &"', "_
		&" aftalenr = "& intAftnr &", fordel = "& intFordel &", "_
		&" overfortsaldo = "& intOverforSaldo &", valuta = "& valuta &" "_
		&" WHERE id = "& id 
		
		'Response.write strSQL
		
		oConn.execute(strSQL)
								
								
								
								'*** Hvis der er ændret kunde tilknytning ****
								oldkundeid = request("oldkundeid")
								if oldkundeid <> kid then
								'** Finder kunde ***
								strSQL = "SELECT kkundenavn, kkundenr FROM kunder WHERE kid = " & kid
								oRec.open strSQL, oConn, 3
								while not oRec.EOF
									
										
										'** Opdaterer job der er tilknyttet denne aftale til at følge samme kunde ***
										strSQLjob = "UPDATE job SET jobknr = " & kid & " WHERE serviceaft = "& id
										oConn.execute(strSQLjob)  
										
										
										'*** Opdaterer relaterede timeregisteringer ***
										strSQLtimer = "UPDATE timer SET "_
										& " Tknavn = '"& oRec("kkundenavn") &"', Tknr = "& kid &""_
										& " WHERE seraft = "& id
										oConn.execute(strSQLtimer)
									
								oRec.movenext
								wend
								
								oRec.close
								
								
								end if 
								
		
        newseraftid = id

		end if
		
		
		
		'****************************************
		'****  Opretter fordeling på måneder ****
		'****************************************
			
            aftidthis = newseraftid '0
			
            
            'if func = "dbred" then
		    '		aftidthis = id
			'else
			'		strSQL3 = "SELECT id FROM serviceaft ORDER BY id DESC "
		    '			oRec.open strSQL3, oConn, 3 
			'		if not oRec.EOF then
			'			aftidthis = oRec("id") 
			'		end if
			'		oRec.close
			'end if
			
			'Response.write aftidthis
			'response.flush 
			
			strSQL3 = "DELETE FROM aft_enh_fordeling WHERE aft_id =" & aftidthis
			oConn.Execute(strSQL3)
			
			'Response.write request("FM_useyear") &"<br><br>"
			
			dim valgteaar
			valgteaar = split(cstr(request("FM_useyear")), ",")
			for t = 0 to UBOUND(valgteaar)
				
				'vaar = valgteaar(t)
				'Response.write 
				
				for m = 1 to 12
				timerThis = 0
				
				if len(trim(request("FM_"&m&"_"&trim(valgteaar(t))))) <> 0 then
				timerThis = request("FM_"&m&"_"&trim(valgteaar(t)))
				else
				timerThis = 0
				end if
				
				'Response.write request("FM_2_2007")
				'Response.write valgteaar(t) &" - "& m &" t: "& timerThis &"<br>"
					
					if timerThis <> 0 then
					
					fordelDatoThis = valgteaar(t)&"/"&m&"/1"
					
					strSQL3 = "INSERT INTO aft_enh_fordeling (aft_id, fordeldato, aar, maned, enheder, editor, dato)"_
					&" VALUES "_
					&" ("& aftidthis &", '"& fordelDatoThis &"' ,"& valgteaar(t) &", "& m &", "& request("FM_"&m&"_"&trim(valgteaar(t))) &", '" &strEditor &"', '" &strDato &"') "
					oConn.Execute(strSQL3)
					
					'Response.write strSQL3 & "<br>"
					'Response.flush
					end if
				next
			
			next


		
        '**************************'
        '**** Job tilknytning *****'
        '**************************'

        '*** Fjernet ****'

        'if len(trim(request("FM_job"))) <> 0 then 'et eller flere job valgt eller "Ingen (fjern)" valgt
        'jobids = split(request("FM_job"), ",")

        ''Response.Write request("FM_job")
        ''Response.end

                 '*** Nulstiller ***'    
        '         strSQLjobnulstil = "UPDATE job SET serviceaft = 0 WHERE jobknr = " & kid & " AND serviceaft = " & aftidthis
        '         oConn.Execute(strSQLjobnulstil)'

        '         if len(trim(request("FM_overforGamleTimereg"))) <> 0 then
        '         overfor = 1
        '         else 
        '         overfor = 0
        '         end if
                
        '        if cint(overfor) = 1 then
                
        '        strSQLtimer = "UPDATE timer SET seraft = 0 WHERE seraft = " & aftidthis
        '        oConn.Execute(strSQLtimer)

        '        strSQLmat_forbrug = "UPDATE materiale_forbrug SET serviceaft = 0 "_
	'			& " WHERE serviceaft = " & aftidthis
       '         oConn.execute(strSQLmat_forbrug)

        '        end if
         '       '****'



        'for j = 0 TO UBOUND(jobids)

        'strSQLjob = "UPDATE job SET serviceaft = "& aftidthis & " WHERE id = " & jobids(j)
        ''Response.Write strSQLjob
        'Response.end
        'oConn.Execute(strSQLjob)

                
                'if cint(overfor) = 1 then

                'strSQLjid = "SELECT jobnr FROM job WHERE id = " & aftidthis
                'oRec3.open strSQLjid, oConn, 3
                'if not oRec3.EOF then  

                'strSQLtimer = "UPDATE timer SET seraft = "& aftidthis &" WHERE tjobnr = " & oRec3("jobnr")
                'oConn.Execute(strSQLtimer)

                'strSQLmat_forbrug = "UPDATE materiale_forbrug SET serviceaft = " & aftidthis &""_
				'& " WHERE jobid = " & jobids(j)
                'oConn.execute(strSQLmat_forbrug)

                'end if
                'oRec3.close

                'end if

        'next
		
        'end if
        	
		'**************
		'**************
		
		Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
		Response.Write("<script language=""JavaScript"">window.close();</script>")
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny "
	dbfunc = "dbopr"
	dblPris = 0
	dblEnheder = 0
	intOverforSaldo = 0
	
		strSQL = "SELECT aftalenr FROM serviceaft WHERE id <> 0 AND id <> 130 ORDER BY aftalenr DESC"
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
		intAftalenr = oRec("aftalenr") + 1
		end if
		oRec.close 
	
	else
	
	strSQL = "SELECT navn, editor, dato, status, enheder, pris, stdato, sldato, "_
	&" varenr, advitype, advihvor, perafg, besk, aftalenr, fordel, overfortsaldo, valuta"_
	&" FROM serviceaft WHERE id = " & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strStatus = oRec("status")
	StrTdato = oRec("stdato")
	StrUdato = oRec("sldato")
	dblPris = oRec("pris")
	dblEnheder = oRec("enheder")
	strVarenr = oRec("varenr")
	intAdvitype = oRec("advitype")
	intAdvihvor = oRec("advihvor")
	perafg = oRec("perafg")
	strBesk = oRec("besk")
	intAftalenr = oRec("aftalenr")
	intFordel = oRec("fordel")
	intOverforSaldo = oRec("overfortsaldo")
	valuta = oRec("valuta")
		
	end if
	oRec.close
		
		if forny = 0 then
		dbfunc = "dbred"
		varbroedkrumme = "Rediger "
		varSubVal = "opdaterpil"
		else
		varSubVal = "opretpil" 
		varbroedkrumme = "Forny "
		dbfunc = "dbopr"
		
		antaldage = datediff("d", StrTdato, StrUdato)
		
		'Response.write antaldage
		StrTdato = dateadd("d", 1, StrUdato)
		StrUdato = dateadd("d", antaldage, StrTdato)
		
		strSQL = "SELECT aftalenr FROM serviceaft WHERE id <> 0 ORDER BY aftalenr DESC"
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
		intAftalenr = oRec("aftalenr") + 1
		end if
		oRec.close 
		
		end if 
	
	end if
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
	<!--#include file="inc/dato2.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	
	
	<div id="sindhold" style="position:absolute; left:10px; top:10px; visibility:visible;">
	<table cellspacing="0" cellpadding="2" border="0" width="320" bgcolor="#ffffff">
	<form action="serviceaft.asp?menu=kun&func=<%=dbfunc%>&forny=<%=forny%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="oldkundeid" value="<%=kid%>">
	<!--<input type="hidden" name="kundeid" value="<kid%>">-->
	
	
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right width=8 valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
    	<td colspan="2" class=alt><b><%=varbroedkrumme%> aftale</b></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=formatdatetime(strDato, 1)%></b> af <b><%=strEditor%></b></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%end if%>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan="2">&nbsp;</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top><font color=red>*</font> <b>Kontakt:</b></td>
		<td><select name="kundeid" size="1" style="width:195px;">
		<%if forny = 1 OR dbfunc = "dbred" then
		kundeWhere = "Kid = " & kid
		else
		kundeWhere = "ketype <> 'e' ORDER BY Kkundenavn"
		%><option value="0">Vælg kontakt</option><%
		end if
		
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE " & kundeWhere
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
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
		</select></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><font color=red>*</font> <b>Aftale:</b></td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:195;"></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF">
					    <td style="border-left:1px #003399 solid;">&nbsp;</td>
					    <td>Valuta:&nbsp;</td>
						<td>
							<select name="FM_valuta" style="width:195px;">
							<%strSQL = "SELECT id, valuta, valutakode, kurs FROM valutaer WHERE id <> 0 ORDER BY id " 
							oRec.open strSQL, oConn, 3
							while not oRec.EOF 
							 if oRec("id") = cint(valuta) then
							 vSEL = "SELECTED"
							 else
							 vSEL = ""
							 end if%>
							<option value="<%=oRec("id") %>" <%=vSEL %>><%=oRec("valuta") %> | <%=oRec("valutakode") %> | kurs: <%=oRec("kurs")%></option>
							<%
							oRec.movenext
							wend
							oRec.close%>
							
							</select>
							
					    </td>
					    <td style="border-right:1px #003399 solid;">&nbsp;</td>
					</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td>Aftalenr:</td>
		<td><input type="text" name="FM_aftnr" value="<%=intAftalenr%>" style="width:195px;"></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td>Varenr:</td>
		<td><input type="text" name="FM_varenr" value="<%=strVarenr%>" style="width:195px;"></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
    <tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td style="padding-top:4px; padding-bottom:4px;"><b>Status:</b></td>	
		<td style="padding-top:4px; padding-bottom:4px;"><select name="FM_status">
					<%if dbfunc = "dbred" then 
					select case strStatus
					case 1
					strStatusNavn = "Aktiv"
					case 0
					strStatusNavn = "Lukket"
					end select
					%>
					<option value="<%=strStatus%>" SELECTED><%=strStatusNavn%></option>
					<%end if%>
					<option value="1">Aktiv</option>
					<option value="0">Lukket</option>
					</select>
					
			</td>
			<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=2><br /><b>Beskrivlese:</b><br>
		<textarea cols="35" rows="3" name="FM_besk" id="FM_besk" id="FM_besk"><%=strBesk%></textarea></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=2><br><b>Aftale håndtering</b></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top style="padding-top:5px;">Enheder / klip:</td>
		<td><input type="text" name="FM_antal" value="<%=dblenheder%>" size="4" style="border:2px green solid;"> stk.&nbsp;&nbsp;
		<%if intFordel <> 0 then
		fdCHK = "CHECKED"
		else
		fdCHK = ""
		end if%>
		<input type="checkbox" name="FM_fordel" id="FM_fordel" value="1" onClick="showhidefordeldiv();" <%=fdCHK%>> Fordel enheder
		
		</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top style="padding-top:5px;">Overfør saldo:</td>
		<td><input type="text" name="FM_overforsaldo" value="<%=intOverforSaldo%>" size="4" style="border:2px Crimson solid;"> &nbsp;(fra eksisterende aftaler)</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=2 style="padding-top:10px;"><b>Periode</b></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td style="padding-top:4px;" valign=top>Start dato:</td>
		<td style="padding-top:4px;"><select name="FM_start_dag" style="font-size:9px;">
		<option value="<%=strDag%>"><%=strDag%></option> 
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
		<option value="31">31</option></select>&nbsp;
		
		<select name="FM_start_mrd" style="font-size:9px;">
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
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_start_aar" style="font-size:9px;">
		<option value="<%=strAar%>">
		<%if id <> 0 then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		
		<%for t = -2 to 5
		thisStrAar = dateadd("yyyy", t, "1-1-"& strAar)%>
		<option value="<%=year(thisStrAar)%>"><%=year(thisStrAar)%></option>
		<%next%>
		</select></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
		</tr>
		<tr>
			<td style="border-left:1px #003399 solid;">&nbsp;</td>
			<td style="padding-bottom:4px;" valign=top>Slut dato:</td>
			<td style="padding-bottom:4px;"><select name="FM_slut_dag" style="font-size:9px;">
		<option value="<%=strDag_slut%>"><%=strDag_slut%></option> 
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
		<option value="31">31</option></select>&nbsp;
		
		<select name="FM_slut_mrd" style="font-size:9px;">
		<option value="<%=strMrd_slut%>"><%=strMrdNavn_slut%></option>
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
		
		
		<select name="FM_slut_aar" style="font-size:9px;">
		<option value="<%=strAar_slut%>">
		<%if id <> 0 then%>
		20<%=strAar_slut%>
		<%else%>
		<%=strAar_slut%>
		<%end if%></option>
		<%for t = -2 to 10
		thisStrAar = dateadd("yyyy", t, "1-1-"& strAar)%>
		<option value="<%=year(thisStrAar)%>"><%=year(thisStrAar)%></option>
		<%next%>
		
		</select></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=2>
		<%
		select case perafg 
		case 1
		chkud = "CHECKED"
		case else
		chkud = ""
		end select
		%>
		<input type="checkbox" name="FM_udlob_aldrig" id="FM_udlob_aldrig" value="1" <%=chkud%>>
		 <b>Nej</b>, aftalen skal ikke have nogen udløbsdato.<br>&nbsp;
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><b>Pris:</b></td>
		<td><input type="text" name="FM_pris" value="<%=dblPris%>" size="12" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;kr.</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<!--<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top colspan=2><br><b>Automatisk fornyelse og email advisering vedr. aftaleudløb?</b>
		Ved automatisk fornyelse bliver periode/klip altid forlænget med det oprindeligt angivet.</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>-->
	<!--<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top colspan=2><br>
		<select name="FM_auto" style="font-size:9px; width:160px;">
		<option value="1">Ja</option>
		<option value="2">Nej</option>
		<option value="3">Nej - Men send email om udløb.</option>
		</select></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>-->
    <%if func = "red" then %>
    <tr>
    <td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top colspan=2><br />
        <b>Følgende job er tilknyttet til aftalen:</b><br />
         <%
        
        'if kundeid <> 0 then
	    strjobKidSQL = "jobknr = " & kid
	    'else
	    'strjobKidSQL = "jobknr <> 0 "
	    'end if
	    
	    strSQLjob = "SELECT j.id AS jid, j.jobnavn, j.jobnr, serviceaft, jobstatus "_
	    &" FROM job j "_
	    &" WHERE "& strjobKidSQL &" AND jobstatus <> 3 AND serviceaft = "& id &" GROUP BY j.id ORDER BY j.jobnavn"
	  
	  'Response.Write strSQLjob
	  'Response.flush
	
          
            
            oRec.open strSQLjob, oConn, 3
				
                j = 0
                while not oRec.EOF
				
			

                stTxt = ""
                select case oRec("jobstatus")
                case 0
                stTxt = "- Lukket"
                case 2
                stTxt = "- Passivt"
                end select
				
				
				%>
				<%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) <%=stTxt%><br />
				<%
                j = j + 1
				oRec.movenext
				wend
				oRec.close

        if j = 0 then
		%>
        (Ingen - tilknyt job til aftalen under job)
        <%end if %>
         
       
        </td>

        <td style="border-right:1px #003399 solid;">&nbsp;</td>
    </tr>
    <%end if %>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top colspan=2><br><b>Skal aftalen være enheds eller periode baseret?</b>
		
		<%select case intAdvitype
		case 0, 1
		chk1 = "checked"
		chk2 = ""
		'chk3 = ""
		case 2
		chk1 = ""
		chk2 = "checked"
		'chk3 = ""
		case 0
		chk1 = ""
		chk2 = ""
		'chk3 = "checked"
		case else
		chk1 = "checked"
		chk2 = ""
		'chk1 = ""
		'chk2 = ""
		'chk3 = "checked"
		end select%>
		<br>
		<!--<input type="radio" name="FM_advitype" value="0" <=chk3%>>Nej - Ingen automatisk advisering/fornyelse<br>-->
		<input type="radio" name="FM_advitype" value="1" <%=chk1%>>Enheder/klip&nbsp;&nbsp;
		<input type="radio" name="FM_advitype" value="2" <%=chk2%>>Periode
		</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>	
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top colspan=2><br>
		<b>Advisering vedr. aftaleudløb:</b><br>
		
		<%select case intAdvihvor
		case 1
		chk1 = "SELECTED"
		chk2 = ""
		chk3 = ""
		chk4 = ""
		chk5 = ""
		case 2
		chk1 = ""
		chk2 = "SELECTED"
		chk3 = ""
		chk4 = ""
		chk5 = ""
		case 3
		chk1 = ""
		chk2 = ""
		chk3 = "SELECTED"
		chk4 = ""
		chk5 = ""
		case 4
		chk1 = ""
		chk2 = ""
		chk3 = ""
		chk4 = "SELECTED"
		chk5 = ""
		case 5
		chk1 = ""
		chk2 = ""
		chk3 = ""
		chk4 = ""
		chk5 = "SELECTED"
		case else
		chk1 = ""
		chk2 = ""
		chk3 = "SELECTED"
		chk4 = ""
		chk5 = ""
		end select%>
		
		<select name="FM_advihvor" style="font-size:9px; width:160px;">
		<option value="1" <%=chk1%>>3 månder før / 50% rest.</option>
		<option value="2" <%=chk2%>>30 dage før / 25% rest.</option>
		<option value="3" <%=chk3%>>10 dage før / 5% rest.</option>
		<option value="4" <%=chk4%>>1 dag før / 1% rest.</option>
		<option value="5" <%=chk5%>>(Aldrig)</option>
		</select></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	
    <!--
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=2>
            <a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=0&aftid=&rdir=aftaler&typ=1','650','500','250','120');" target="_self">Opret termin</a>
        </td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
    -->
	<%if forny = 1 then%>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=2><br>
		<b>Status på den oprindelige aftale.</b><br>
		<input type="checkbox" name="FM_luk" id="FM_luk" value="1" CHECKED> Luk aftale.<br>
		<input type="checkbox" name="FM_overfor" id="FM_overfor" value="1"> Overfør job fra den oprindelige aftale, til den nye aftale.<br>
		<input type="checkbox" name="FM_luk_job" id="FM_luk_job" value="1"> Luk job der er tilknyttet den oprindelige aftale.</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%end if%>
	
	<%'* Opret automatisk folder ***
	if func = "opret" then
	%>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=2><br>
		<b>Opret folder til denne aftale i filarkivet?</b><br>
		<input type="checkbox" name="FM_oprfolder" id="FM_oprfolder" value="1"> 
		Ja, opret folder med samme navn som denne aftale.<br>
		<font class=megetlillesort>(Kunde får adgang til denne folder)</font></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	end if
	%>
	<tr>
		<td style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan="2" align=center style="border-bottom:1px #003399 solid; padding:5px;"><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
		<td style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	</table>
    <br /><br />&nbsp;&nbsp;
	</div>
    <br /><br />&nbsp;&nbsp;
	
	
	<%
	if intFordel <> 0 then
	vzb = "visible"
	else
	vzb = "hidden"
	end if
	
	if dbfunc = "dbopr" then
	topfordel = 297
	else
	topfordel = 317
	end if
	%>
	<div id="fordeldivpil" name="fordeldivpil" style="position:absolute; visibility:<%=vzb%>; display:; left:290; top:<%=topfordel%>; z-index:2000;">
	<img src="../ill/aftaler_pil_top.gif" width="350" height="20" alt="" border="0">
	</div>
	
	<div id="fordeldiv" name="fordeldiv" style="position:absolute; visibility:<%=vzb%>; display:; left:350; top:<%=topfordel+28%>; width:420px; height:380px; overflow:auto; border:2px #999999 solid; padding:10px; background-color:#ffffff;">
	<h3>Fordel enheder (og pris) på måneder:</h3><br>
	<table>
	
	<tr>
		<td>&nbsp;</td>
		<td align=center>Jan</td>
		<td align=center>Feb</td>
		<td align=center>Mar</td>
		<td align=center>Apr</td>
		<td align=center>Maj</td>
		<td align=center>Jun</td>
		<td align=center>Jul</td>
		<td align=center>Aug</td>
		<td align=center>Sep</td>
		<td align=center>Okt</td>
		<td align=center>Nov</td>
		<td align=center>Dec</td>
	</tr>
	
	<%for y = 0 to 8%>
	<tr>
		<%
		
		'timerThis = 0
		timerJan = 0
		timerFeb = 0
		timerMar = 0
		timerApr = 0
		timerMaj = 0
		timerJun = 0
		timerJul = 0
		timerAug = 0
		timerSep = 0
		timerOkt = 0
		timerNov = 0
		timerDec = 0
		
 
		useyear = cint(2007+y)
		
		strSQL = "SELECT aft_id, aar, dag, maned, enheder FROM aft_enh_fordeling WHERE aft_id = "& id &" AND aar = "& useyear
		'Response.write strSQL
		'Response.flush
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
			
			select case oRec("maned")
			case 1
			timerJan = oRec("enheder")
			case 2
			timerFeb = oRec("enheder")
			case 3
			timerMar = oRec("enheder")
			case 4
			timerApr = oRec("enheder")
			case 5
			timerMaj = oRec("enheder")
			case 6
			timerJun = oRec("enheder")
			case 7
			timerJul = oRec("enheder")
			case 8
			timerAug = oRec("enheder")
			case 9
			timerSep = oRec("enheder")
			case 10
			timerOkt = oRec("enheder")
			case 11
			timerNov = oRec("enheder")
			case 12
			timerDec = oRec("enheder")
			end select
			
		oRec.movenext
		wend
		
		oRec.close
		%>
		
		<td><b><%=useyear%></b><input type="hidden" name="FM_useyear" id="FM_useyear" value="<%=useyear%>"></td>
			
			<%for m = 1 to 12
			
			select case m
			case 1
			timerThis = timerJan
			case 2
			timerThis = timerFeb
			case 3
			timerThis = timerMar
			case 4
			timerThis = timerApr
			case 5
			timerThis = timerMaj
			case 6
			timerThis = timerJun
			case 7
			timerThis = timerJul
			case 8
			timerThis = timerAug
			case 9
			timerThis = timerSep
			case 10
			timerThis = timerOkt
			case 11
			timerThis = timerNov
			case 12
			timerThis = timerDec
			end select
			
			timerForDeltTotalt = timerForDeltTotalt + timerThis
			
			if len(timerThis) = 0 OR timerThis = 0 then
			timerThis = ""
			timerThis_opr = 0
			else
			timerThis = timerThis
			timerThis_opr = timerThis
			end if
			%>
			<td><input type="text" style="width:25px; font-size:9px;" name="FM_<%=m%>_<%=useyear%>" id="FM_<%=m%>_<%=useyear%>" value="<%=timerThis%>" onkeyup="beregnenhederrest('<%=m%>','<%=useyear%>');">
			<input type="hidden" name="FM_<%=m%>_<%=useyear%>_opr" id="FM_<%=m%>_<%=useyear%>_opr" value="<%=timerThis_opr%>"></td>
			<%next%>
	</tr>
	<%next%>
	<!--<tr>
		<td colspan="13" align=center><br><input type="image" src="../ill/<=varSubVal%>.gif"><br>&nbsp;</td>
	</tr>-->
	</table>
	Der er fordelt: <input type="text" name="FM_enhederrest" id="FM_enhederrest" value="<%=timerForDeltTotalt%>" size=3 style="border:2px red solid; font-size:9px;"> enheder.
	</form>
	</div>
	<%end select%>
	
	

<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
