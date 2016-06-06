<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/konto_inc.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->
<!--#include file="inc/functions_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
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
	
	kundeid = request("kid")
	
	if len(request("kontonr")) <> 0 then
	kontonr = request("kontonr")
	        
	        
	        %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(kontonr)
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
			<%
			errortype = 95
			call showError(errortype)
			
			isInt = 0
			Response.End
			end if
	        
	        
	else
	kontonr = 0
	end if
	
	if len(request("FM_soeg")) <> 0 then 
	thiskri = request("FM_soeg")
	useKri = 1
	else
	thiskri = ""
	useKri = 0
	end if
	
	'usePeriodefilter = request("useper")
	
	call finddatoer()
	
	
	
	select case func
	case "bog"
	oConn.execute("UPDATE posteringer SET status = 1 WHERE status = 0")
	Response.redirect "posteringer.asp?menu=kon&id=0"
	case "slet"
	'*** Her spørges om det er ok at der slette ***'
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>slette</b> en <b>postering</b>. Er dette korrekt?<br>"
	slturl = "posteringer.asp?menu=kon&func=sletok&id="&id&"&kontonr="&kontonr&"&oprid="&request("oprid")
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,200)
	
	
	
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM posteringer WHERE oprid = "& request("oprid") &"")
	if kontonr <> 0 then
	Response.redirect "posteringer.asp?menu=kon&id="&id&"&kontonr="&kontonr
	else
	Response.redirect "posteringer.asp?menu=kon&id=0"
	end if
	
	case "dbopr", "dbred"
	    
	    
	    '*** Her indsættes en ny type i db ****
		if len(request("FM_total")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		useleftdiv = "t"
		errortype = 50
		call showError(errortype)
		
		else
			
			call erDetInt(request("FM_total"))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<% 
			useleftdiv = "t"
			errortype = 49
			call showError(errortype)
			
			isInt = 0
			
			else
		
		
		if func = "dbred" then
		opridOld = request("oprid")
		strSQLDel = "DELETE FROM posteringer WHERE oprid = "& opridOld
		oConn.execute(strSQLDel)
		end if
		
		
		
		intkontonr = request("FM_kontonr_sel")
		modkontonr = request("FM_modkonto_sel")
		
		
		response.cookies("erp")("kontonr_1") = intkontonr
		response.cookies("erp")("kontonr_2") = modkontonr
		
		
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		function SQLBless2(s)
		dim tmp2
		tmp2 = s
		tmp2 = replace(tmp2, ",", ".")
		SQLBless2 = tmp2
		end function
		
		function SQLBless3(s)
		dim tmp3
		tmp3 = s
		tmp3 = replace(tmp3, ".", ",")
		SQLBless3 = tmp3
		end function
		
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		vorref = request("FM_ref")
		varBilag = request("FM_bilagsnr")
		strTekst = SQLBless(request("FM_tekst"))
		posteringsdato = request("FM_aar")&"/"&request("FM_md")&"/"&request("FM_dag")
		
		kundeid = request("FM_kundeid")
		
		intTotal = request("FM_total")
		
		debkre = "pos"
		
		rdirKontonr = request("rdirkontonr")
		
		if len(request("status")) <> 0 AND request("status") <> 0 then
		intStatus = 1
		else
		intStatus = 0
		end if
		
		
		'*** Beregner moms Debit Konto ***'
	    call beregnmoms(debkre, intTotal, intkontonr)
	    
	    intNetto = SQLBless2(intNetto)
		intMoms = SQLBless2(intMoms)
		intTotal = SQLBless2(intTotal)
		
		
	    '**** Postering debit konto ****'
		'call opretpos("1", func, modkontonr, intkontonr,intTotal_konto,intTotal_modkonto,intMoms,strEditor,varBilag,strTekst,posteringsdato,intStatus,vorref)
		
		oprid = 0
		call opretPosteringSingle(oprid,"1", func, intkontonr, modkontonr, intNetto, intNetto, intMoms, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorref)
		
		
		
		'*** Beregner moms Kredit konto (Modkonto) ***'
	    call beregnmoms(debkre, intTotal, modkontonr)
	    
	   
		
		intNetto = SQLBless2(-intNetto)
		intMoms = SQLBless2(-intMoms)
		intTotal = SQLBless2(-intTotal) 
		
		
		
	    '**** Postering kredit konto ****'
		'call opretpos("1", func, modkontonr, intkontonr,intTotal_konto,intTotal_modkonto,intMoms,strEditor,varBilag,strTekst,posteringsdato,intStatus,vorref)

		
		call opretPosteringSingle(oprid, "1", func, modkontonr, intkontonr, intNetto, intNetto, intMoms, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorref)
		
		
		'.end
		
		%>
        <script>
		window.opener.location.href = "posteringer.asp?menu=kon&kid=<%=kundeid%>&kontonr=<%=rdirKontonr%>&id=<%=id%>"
		window.close();
		</script>
		<%
		
		                                   
		
		
            'if kontonr <> 0 then
			'Response.redirect "posteringer.asp?menu=kon&kid="&kundeid&"&kontonr="&intkontonr&"&id="&id
			'else
			'Response.redirect "posteringer.asp?menu=kon&id=0"
			'end if
		
		end if 'validering
	end if'validering
	
	
	
	
	
	
	
	
	case "opret", "red"
            %>
            <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
            <%
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	intMomskode = 1
	varSubVal = "Opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	intDag = day(now)
	intMd = month(now)
	intAar = year(now) 
	intRef = session("Mid")
	intBelob_opr = 0
	oprid = 0
	
	varBilagsnr = 0 
	foundone = 0
	strSQL = "SELECT bilagsnr FROM posteringer WHERE bilagsnr IS NOT NULL ORDER BY id DESC"
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF AND foundone = 0
		
        
        if IsNull(oRec("bilagsnr")) <> true AND len(trim(oRec("bilagsnr"))) <> "" then

        call erDetInt(oRec("bilagsnr"))
		
		if isInt > 0 then
		varBilagsnr = varBilagsnr
		else
		varBilagsnr = oRec("bilagsnr") + 1  'Sidst oprettede 
		foundone = 1
		end  if

        else

        varBilagsnr = 0

        end if
	
	isInt = 0	
	oRec.movenext
	wend 
	oRec.close
	
	'if len(request("kontonr")) <> 0 AND request("kontonr") <> 0 then
	'varKontonr = request("kontonr")
	'else
	varKontonr = request.cookies("erp")("kontonr_1")
	'end if
	
	varModkontonr = request.cookies("erp")("kontonr_2")
	
	    if len(request("status")) <> 0 AND request("status") <> 0 then
		intStatus = 1
		else
		intStatus = 0
		end if
	
	else
	strSQL = "SELECT posteringer.dato AS dato, posteringer.editor AS editor, id, bilagstype, kontonr, modkontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, mid, mnavn, mnr, oprid FROM posteringer LEFT JOIN medarbejdere ON (mid = att) WHERE posteringer.id = "& id &" ORDER BY posteringsdato"
	oRec.open strSQL, oConn, 3
	
	if not oRec.EOF then
	varBilagsnr = oRec("bilagsnr")
	intBeloeb = oRec("beloeb")
	if intBeloeb < 0 then
	intBelob_opr = intBeloeb
	lenintBeloeb = len(intBeloeb)
	rightintBeloeb = right(intBeloeb, (lenintBeloeb-1))
	intBeloeb = rightintBeloeb
	else
	intBelob_opr = intBeloeb
	intBeloeb = intBeloeb
	end if
	
	intNetto = oRec("nettobeloeb")
	intMoms = oRec("moms")
	strTekst = oRec("tekst")
	posdato = oRec("posteringsdato")
	intRef = oRec("mid")
	strMnavn = oRec("mnavn") 
	intMedmnr = oRec("mnr") 
	
	intDag = day(posdato)
	intMd = month(posdato)
	intAar = year(posdato) 
	
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intStatus = oRec("status")
	
	varKontonr = oRec("kontonr")
	varModkontonr = oRec("modkontonr")
	
	oprid = oRec("oprid")
	
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdaterpil" 
	end if
	%>
	
	<script>
	function fjerndot(){
	document.all["FM_total"].value = document.all["FM_total"].value.replace(".",",")
	}
	
	
	function rensdag(){
	document.getElementById("FM_dag").value = ""
	}
	
	function rensmd(){
	document.getElementById("FM_md").value = ""
	}
	
	function rensaar(){
	document.getElementById("FM_aar").value = ""
	}
	
	function bodyonload(){
	document.getElementById("FM_total").focus();
	}
	
	function valgmodkont(){
	document.getElementById("FM_modkonto").value = document.getElementById("FM_modkonto_sel").value
	}
	
	function valgkont(){
	document.getElementById("FM_kontonr").value = document.getElementById("FM_kontonr_sel").value
	//alert(document.getElementById("FM_kontonr_sel").selectedIndex)
	}
	
	function valgkont_sel(){
	
    var newval = document.getElementById("FM_kontonr").value
    var x=document.getElementById("FM_kontonr_sel")
    for(i=0; i<=250; i++)
        {
            if(x.options[i].value == newval)
            {
                x.selectedIndex = i;
                break;
            }
        }
	}
	
	function valgmodkont_sel(){
    var newval = document.getElementById("FM_modkonto").value
	var x=document.getElementById("FM_modkonto_sel")
    for(i=0; i<=250; i++)
        {
            if(x.options[i].value == newval)
            {
                x.selectedIndex = i;
                break;
             }
        }
    }
	
	</script>
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:22px; visibility:visible;">
	
	
	
	<table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff">
	<form action="posteringer.asp?menu=kon&func=<%=dbfunc%>&status=<%=intStatus%>&rdirkontonr=<%=kontonr%>&oprid=<%=oprid%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2 style="border-top:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=2 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2 style="border-top:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 class=alt><b> <%=varbroedkrumme%> Postering</b></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr bgcolor="#5582D2">
		<td style="border-left:1px #003399 solid; border-right:1px #003399 solid; padding-left:7;" colspan=4 class=alt>Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td width="5" style="border-left:1px #003399 solid;" rowspan=6>&nbsp;</td>
		<td><br>Dato:</td>
		<td><br><input type="text" name="FM_dag" id="FM_dag" value="<%=intDag%>" size="2" maxlength=2 onFocus="rensdag()">
		<input type="text" name="FM_md" id="FM_md" value="<%=intMd%>" size="2" maxlength=2  onFocus="rensmd()">
		<input type="text" name="FM_aar" id="FM_aar" value="<%=intAar%>" size="4" maxlength=4 onFocus="rensaar()">&nbsp;(dd-mm-åååå)</td>
	    <td width="5" style="border-right:1px #003399 solid;" rowspan=6>&nbsp;</td>
	</tr>
	<tr>
	<td>Brutto beløb.</td>
		<td><input type="text" name="FM_total" id="FM_total" onkeyup="fjerndot();" value="<%=intBeloeb%>" size="10"> kr.</td>
	  
	</tr>
	<tr>
		<td>Bilagsnr.</td>
		<td><input type="text" name="FM_bilagsnr" value="<%=varBilagsnr%>" style="width:200px;" ></td>
	</tr>
	<tr>
		<td>Ref.</td><td>
		<select name="FM_ref" style="width:200px;">
		<%
		strSQL = "SELECT Mnavn, Mnr, Mid FROM medarbejdere ORDER BY mnavn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		if oRec("mid") = cint(intRef) then
		selth = "SELECTED"
		else
		selth = ""
		end if
		%>
		<option value="<%=oRec("mid")%>" <%=selth%>><%=oRec("mnavn")%> (<%=oRec("mnr")%>)</option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		</select></td>
	</tr>
	<tr>
		<td valign=top>Tekst</td><td>
            <textarea id="FM_tekst" name="FM_tekst" cols="61" rows="3"><%=strTekst%></textarea>
         </td>
	</tr>
	<tr>
		<td colspan=2>
		<table width=100% border=0 cellspacing=1 cellpadding=2>
		<tr><td><b>Debiter</b></td><td>
            &nbsp;
		</td><td><b>Krediter (modkonto)</b></td></tr>
		<tr>
		    <td>Vælg: 
		    <% 
		    strSQL = "SELECT kontonr, kontoplan.navn AS kontonavn, kontoplan.id, kid, debitkredit, "_
		    &" keycode, type, nk.navn, m.navn AS momskode FROM kontoplan "_
		    &" LEFT JOIN nogletalskoder nk ON (nk.id = keycode) "_
		    &" LEFT JOIN momskoder m ON (m.id = kontoplan.momskode) "_
		    &" WHERE status = 1 ORDER BY kontonr, kontoplan.navn"
		
		'Response.Write strSQL
		'Response.Flush
		
		%>
		    <select name="FM_kontonr_sel" id="FM_kontonr_sel" style="width:250px; background-color:#ffffe1;" onchange="valgkont()";>
		<%
		
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		if cstr(varKontonr) = cstr(oRec("kontonr")) then
		selkon = "SELECTED"
		else
		selkon = ""
		end if
		
		'if oRec("debitkredit") = 1 then
		'debitkredit = "Kre"
		'else
		'debitkredit = "Deb"
		'end if
		
		if oRec("type") = 1 then
		ktype = "Drift"
		else
		ktype = "Status"
		end if
		
		%>
		<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;<%=oRec("kontonavn")%> - <%=ktype %> - <%=oRec("momskode") %></option>
		<%
		oRec.movenext
		Wend 
		oRec.close%>
		</select><br />
		Ell. angiv kontonr: 
            <input id="FM_kontonr" name="FM_modkontonr" value="<%=varKontonr%>" style="width:195px;" type="text" onkeyup="valgkont_sel()" /></td>
		<td>
            &nbsp;
		</td>
		<td>
		<%
		    strSQL = "SELECT kontonr, kontoplan.navn AS kontonavn, kontoplan.id, kid, debitkredit, "_
		    &" keycode, type, nk.navn, m.navn AS momskode FROM kontoplan "_
		    &" LEFT JOIN nogletalskoder nk ON (nk.id = keycode) "_
		    &" LEFT JOIN momskoder m ON (m.id = kontoplan.momskode) "_
		    &" WHERE status = 1 ORDER BY kontonr, kontoplan.navn"
		
		%>
		<select name="FM_modkonto_sel" id="FM_modkonto_sel" style="width:275px; background-color:pink;" onchange="valgmodkont()">
		<%
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		
		if cstr(varModKontonr) = cstr(oRec("kontonr")) then
		selkon = "SELECTED"
		else
		selkon = ""
		end if
		
		if oRec("debitkredit") = 1 then
		debitkredit = "Kre"
		else
		debitkredit = "Deb"
		end if
		
		if oRec("type") = 1 then
		ktype = "Drift"
		else
		ktype = "Status"
		end if
		
		%>
		<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;<%=oRec("kontonavn")%> - <%=ktype %> - <%=oRec("momskode") %></option>
		<%
		oRec.movenext
		Wend 
		oRec.close%>
		</select><br />
            <input id="FM_modkonto" name="FM_modkonto" value="<%=varModKontonr%>" style="width:275px;" type="text" onkeyup="valgmodkont_sel()" /></td>
		</td>
		</tr>
		
		</table>
		</td>
	
		
	</tr>
	<tr>
		<td colspan="4" align="center" style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;"><br><br>
            <input id="Submit1" type="submit" value="Opret postering" /><br><br>&nbsp;</td>
	</tr>
	
	</form>
	</table>
	</div>
	
	<%
	Response.Write("<script language=""JavaScript"">bodyonload();</script>")
	%>
	
	<div style="position:absolute; top:20px; left:680px; width:200px; background-color:#FFFFFF; padding:20px;">
	<b>Info:</b><br />
	<table cellspacing=2 cellpadding=1 border=0>
	<tr>
	<td colspan=4><b>A) Ved faktura til kunde</b></td>
	</tr>
	<tr>
	<td>Konto</td>
	<td>Debit</td>
	<td>Kredit</td>
	<td>Moms</td>
	</tr>
	<tr>
	<td>Debitor konto</td>
	<td>1250</td>
	<td>0</td>
	<td>0</td>
	</tr>
	<tr>
	<td>Omsætningskonto</td>
	<td>0</td>
	<td>1000</td>
	<td>250</td>
	</tr>
	
	<tr>
	<td colspan=4><b>B) Betaling af udgiftsbilag (telefon)</b></td>
	</tr>
	<tr>
	<td>Konto</td>
	<td>Debit</td>
	<td>Kredit</td>
	<td>Moms</td>
	</tr>
	<tr>
	<td>Telefon konto</td>
	<td>800</td>
	<td>0</td>
	<td>200</td>
	</tr>
	<tr>
	<td>Bankkonto</td>
	<td>0</td>
	<td>1000</td>
	<td>0</td>
	</tr>
	
	<tr>
	<td colspan=4><b>C) Registrering af indbetalt faktura</b></td>
	</tr>
	<tr>
	<td>Konto</td>
	<td>Debit</td>
	<td>Kredit</td>
	<td>Moms</td>
	</tr>
	<tr>
	<td>Debitor konto</td>
	<td>0</td>
	<td>1250</td>
	<td>0</td>
	</tr>
	<tr>
	<td>Bankkonto</td>
	<td>1250</td>
	<td>0</td>
	<td>0</td>
	</tr>
	</table>
	
	
	</div>
	
	<%case else%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%call menu_2014() %>
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	
	    
	    <%call filterheader_2013(0,0,1010,pTxt)%>
	    <table cellspacing="0" cellpadding="0" border="0" width=100%>
		
			<tr>
			<form action="posteringer.asp?menu=kon&kid=<%=kundeid%>" method="POST">
			<td><b>Konto nr:</b></td><td><input type="text" name="kontonr" id="kontonr" value="<%=kontonr%>" size="25"></td>
			<td><b>Bilagsnr:</b></td><td><input type="text" name="FM_soeg" id="FM_soeg" value="<%=thiskri%>" size="10"></td>
			<td align=right><input type="hidden" name="useper" value="j"></td>
			<!--#include file="inc/weekselector_b.asp"-->
			<td align=center><input type="image" src="../ill/pilstorxp.gif" border="0"></td>
			</tr></form>
	    </table>
		<!-- filter header sLut -->
	</td></tr></table>
	</div>
	
	
	

	
	
	
	
	
	
	<%if kontonr = 0 then
	bgthis = "#cccccc"
	txt = "Kassekladde balance:"
	txt2 = "Kassekladde"
	ext = "_gray"
	cspan = 12
	else
		select case kontonr
		case "-1" 
		bgthis = "pink"
		txt = "Saldo på konto:"
		txt2 = "Posteringer på konto: -1 - Indgående moms" 
		ext = ""
		cspan = 10
		case "-2"
		bgthis = "pink"
		txt = "Saldo på konto:"
		txt2 = "Posteringer på konto: -2 - Udgående moms" 
		ext = ""
		cspan = 10
		case else
		bgthis = "#5582d2"
		txt = "Saldo på konto:"
		
		
		knavn = ""
		'*** Henter navn på konto ***'
		strSQLknavn = "SELECT navn FROM kontoplan WHERE kontonr = " & kontonr
		oRec.open strSQLknavn, oConn, 3
		if not oRec.EOF then
		
		knavn = oRec("navn")
		
		end if
		oRec.close
		
		
		txt2 = "Posteringer på konto: " & kontonr & " - "& knavn 
		ext = ""
		cspan = 12
		end select
	end if
	
	if kontonr <> 0 then
	intStatus = 1
	else
	intStatus = 0
	end if
	%>
	
	
	
	<h4><%=txt2%></h4>
	
	<br>Periode: <b><%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1)%></b> til
	<b><%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%></b> <br> 
	
		
		<%if kontonr = 0 then %>
	    
	    <div style="position:absolute; left:832px; top:105px;">
	   

             <% 
               nWdt = 187
                nTxt = "Opret ny postering"
                nLnk = "posteringer.asp?menu=kon&func=opret&kid="&kundeid&"&kontonr="&kontonr&"&status="&intStatus
                nTgt = "_blank"
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) 
                %>

	    </div>
	    <%end if %>
	
	
	
	
	<% 
    tTop = 20
	tLeft = 0
	tWdth = 1004
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%> 
	
	
	<table cellspacing="0" cellpadding="1" border="0" width=100%>
	<tr bgcolor="#5582d2">
	<td width="5" style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid;">&nbsp;</td>
	<td height="30" class=alt style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid; width:70px;"><b>Dato</b></td>
	<td class=alt align=right style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid;"><b>Bilagsnr</b></td>
	<td class=alt style="padding-left:20px; border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid;"><b>Tekst</b></td>
	<td class=alt style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid; width:150px;"><b>Konto</b></td>
	<td class=alt style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid; width:150px;"><b>Modkonto</b></td>
	<td class=alt style="border-bottom:1px #8cAAe6 solid; padding-left:10px; padding-right:3px; border-right:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid; width:100px;"><b>Vor ref.</b></td>
	<%if kontonr >= "0" then %>
	<td class=alt align="right" style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid; border-right:1px #8cAAe6 solid; padding-right:3px;"><b>Debit</b></td>
	<td class=alt align="right" style="border-bottom:1px #8cAAe6 solid; border-right:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid; padding-right:3px;"><b>Kredit</b></td>
	<%end if %>
	
	<%if kontonr <= "0" then %>
    <td class=alt align="right" style="border-bottom:1px #8cAAe6 solid; border-right:1px #8cAAe6 solid; padding-right:3px; border-top:1px #8cAAe6 solid;"><b>Moms</b></td>
	<%else %>
	<td class=alt align="right" style="border-bottom:1px #8cAAe6 solid; padding-right:3px; border-top:1px #8cAAe6 solid;">
      &nbsp;</td>
    <%end if %>
	<td class=alt align="right" style="border-bottom:1px #8cAAe6 solid; border-right:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid; padding-right:3px;"><b>Saldo</b></td>
	
	<td class=alt align=center style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid;"><b>Status</b></td>
	<td class=alt style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid;">&nbsp;</td>
	<td width="5" class=alt style="border-bottom:1px #8cAAe6 solid; border-top:1px #8cAAe6 solid;">&nbsp;</td>
	</tr>
	<%
	if useKri <> 0 then
	useSQLKri = " posteringer.bilagsnr = "& thiskri  
	else
	useSQLKri = " posteringer.id > 0 "
	end if
	
	'if usePeriodefilter = "j" then
	periodefilter = " AND posteringsdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	'else
	'periodefilter = ""
	'end if
	
	intNetto_tot = 0
	intMoms_tot = 0
	intTotal_belob = 0
	
	select case kontonr
		case "-1", "-2"
		
			if kontonr = "-2" then
			kassekladeorkonto = "(posteringer.kontonr > 0 AND posteringer.status = 1 AND posteringer.moms < 0)"
			momskri = " k.kontonr = posteringer.kontonr" 'AND k.debitkredit = 2 AND k.type = 2
			momskri2 = " k2.kontonr = posteringer.modkontonr "
			end if
			
			if kontonr = "-1" then 'AND kontonr < 0 then
			kassekladeorkonto = "(posteringer.kontonr > 0 AND posteringer.status = 1 AND posteringer.moms > 0)"
			momskri = " k.kontonr = posteringer.kontonr"  ' AND k.debitkredit = 1 AND k.type = 1
			momskri2 = " k2.kontonr = posteringer.modkontonr "
			end if
			
			useSQLKri = useSQLKri &" AND k.type <> -1"
			
		case 0
			kassekladeorkonto = "(posteringer.kontonr > 0 AND posteringer.status = 0)"
			momskri = " k.kontonr = posteringer.kontonr "	
			momskri2 = " k2.kontonr = posteringer.modkontonr "	
		case else
			
			kassekladeorkonto = "(posteringer.kontonr = "& kontonr &" AND posteringer.status = 1)"
			momskri = " k.kontonr = posteringer.kontonr "
			momskri2 = " k2.kontonr = posteringer.modkontonr "
			
			
	end select 
	
	
	
	strSQL = "SELECT posteringer.id, posteringer.bilagstype, posteringer.kontonr, "_
	&" posteringer.modkontonr, posteringer.bilagsnr, posteringer.beloeb, posteringer.nettobeloeb, "_
	&" posteringer.moms, posteringer.tekst, posteringsdato, posteringer.status, mid, mnavn, mnr, "_
	&" posteringer.oprid, k.debitkredit AS debitkredit, k.navn AS kontonavn, k2.debitkredit AS debitkredit2, k2.navn AS modkontonavn FROM posteringer "_
	&" LEFT JOIN kontoplan k ON ("& momskri &")  "_
	&" LEFT JOIN kontoplan k2 ON ("& momskri2 &")  "_
	&" LEFT JOIN medarbejdere ON (mid = att) WHERE "& kassekladeorkonto &" "& periodefilter &" AND "& useSQLKri &" GROUP BY posteringer.id ORDER BY posteringsdato, id"
	
	'Response.write strSQL
	'Response.Flush
	
	oRec.open strSQL, oConn, 3
	c = 0
	debitTot = 0
	kreditTot = 0
	
	
	
	while not oRec.EOF 
	
	if cdbl(id) = oRec("id") then
	bgcol = "#ffffe1"
	else
	    
	    select case right(c, 1)
	    case 0,2,4,6,8
	    bgcol = "#EFF3FF"
	    case else
    	bgcol = "#ffffff"
	    end select
    	
	
	
	end if%>
	<!--<tr>
		<td bgcolor="#ffffff" colspan="12" style="border-right:1px #003399 solid; border-right:1px #003399 solid; border-bottom:1px #999999 dashed;"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>-->
	<tr bgcolor="<%=bgcol%>">
		<td style="border-bottom:1px #cccccc solid;">&nbsp;<!--<=oRec("id") %>--></td>
		<td style="border-bottom:1px #cccccc solid;"><%=oRec("posteringsdato")%></td>
		<%if kontonr = 0 AND oRec("nettobeloeb") > 0 then%>
		<td height="20" align=right style="border-bottom:1px #cccccc solid;">
		<!--<a href="posteringer.asp?menu=kon&func=red&id=<%=oRec("id")%>&kontonr=<%=kontonr%>">-->
		<a href="posteringer.asp?menu=kon&func=red&id=<%=oRec("id")%>&kontonr=<%=kontonr%>" class=vmenu target="_blank"><%=oRec("bilagsnr")%></a></td>
	    
		<%else%>
		<td height="20" align=right style="border-bottom:1px #cccccc solid;"><b><%=oRec("bilagsnr")%></b></td>
		<%end if%>
		<td style="padding-left:20px; border-bottom:1px #cccccc solid;"><%=oRec("tekst")%>
            &nbsp;</td>
		<td style="border-bottom:1px #cccccc solid;"><%=oRec("kontonr")%> - <%=oRec("kontonavn") %>
            &nbsp;</td>
		<td class=lillegray style="border-bottom:1px #cccccc solid;"><%=oRec("modkontonr")%> - <%=oRec("modkontonavn") %>
            &nbsp;</td>
		<td style="border-bottom:1px #cccccc solid; padding-left:10px; padding-right:3px; border-right:1px #8cAAe6 solid;"><%=oRec("mnavn")%> (<%=oRec("mnr")%>)&nbsp;</td>
		<%select case kontonr
		case "-1", "-2"%>
		<td align=right style="border-bottom:1px #cccccc solid; border-right:1px #8cAAe6 solid; padding-right:3px;"><%=formatnumber(oRec("moms"), 2)%></td>
		<%
		SaldoTot = saldoTot + (oRec("moms"))
		%>
		<td align=right style="border-bottom:1px #cccccc solid; border-right:1px #8cAAe6 solid; padding-right:3px;"><%=formatnumber(SaldoTot, 2) %></td>
        <%
        SaldoGrandTot = SaldoTot
		case else%>
		    
		    <%if oRec("nettobeloeb") > 0 then %>
		    <td align=right style="border-bottom:1px #cccccc solid; border-right:1px #8cAAe6 solid; padding-right:3px;">
		    <%=formatnumber(oRec("nettobeloeb"), 2)%></td>
		    <td style="border-bottom:1px #cccccc solid; border-right:1px #8cAAe6 solid; padding-right:3px;">
            &nbsp;</td>
            <%
            debitTot = debitTot + oRec("nettobeloeb")
            else %>
            <td style="border-bottom:1px #cccccc solid; border-right:1px #8cAAe6 solid; padding-right:3px;">
                &nbsp;
		    </td>
		    <td style="border-bottom:1px #cccccc solid; border-right:1px #8cAAe6 solid; padding-right:3px;" align=right>
            <%=formatnumber(oRec("nettobeloeb"), 2)%></td>
            <%
            kreditTot = kreditTot + oRec("nettobeloeb")
            end if %>
            
           
		                <%if kontonr = "0" then %>
						 <td align=right style="border-bottom:1px #cccccc solid; border-right:1px #8cAAe6 solid; padding-right:3px;"><%=formatnumber(oRec("moms"), 2)%></td>
		                <%else %>
						<td align=right style="border-bottom:1px #cccccc solid; padding-right:3px;">
                            &nbsp;</td>
						<%end if%>
		    
		    
		    <%
		    momsTot = momsTot + (oRec("moms"))
		    if kontonr = "0" then
		    SaldoTot = debitTot + (kreditTot) + (momsTot)
		    else
		    SaldoTot = debitTot + (kreditTot)
		    end if 
		    
		    
		    %>
		    <td align=right style="border-bottom:1px #cccccc solid; border-right:1px #8cAAe6 solid; padding-right:3px;"><%=formatnumber(SaldoTot, 2) %></td>
        <%
		SaldoGrandTot = SaldoTot
		
		end select%>
		
		<!--<td align=right><=formatnumber(oRec("moms"), 2)%></td>-->
		<td align=center style="border-bottom:1px #cccccc solid;"><%select case oRec("status") 
		case 0
		Response.write "<font color=#cccccc>Kladde</font>"
		case 1
		Response.write "Bogført"
		case else
		Response.write "<font color=#cccccc>Kladde</font>"
		end select%></td>
		
		<%if kontonr = 0 then%>
		<td style="border-bottom:1px #cccccc solid;"><a href="posteringer.asp?menu=kon&func=slet&id=<%=oRec("id")%>&kontonr=<%=kontonr%>&oprid=<%=oRec("oprid")%>"><img src="../ill/slet_16.gif" alt="" border="0"></a></td>
		<%else%>
		<td style="border-bottom:1px #cccccc solid;">&nbsp;</td>
		<%end if%>
		<td style="border-bottom:1px #cccccc solid;">&nbsp;</td>
	</tr>
	<%
	'intNetto_tot = intNetto_tot + oRec("nettobeloeb")
	'intMoms_tot = intMoms_tot + oRec("moms")
	'intTotal_belob = intTotal_belob + oRec("beloeb")
	'Response.write intTotal_belob & "<br>"
	x = 0
	c = c + 1
	oRec.movenext
	wend
	
	
	
	%>
	
	<tr bgcolor="#ffffff">
	    <td height=40 colspan=7 align=right><b><%=txt%></b>
            &nbsp;</td>
			
			<%if kontonr = "-1" OR kontonr = "-2" then%>
			<td align=right>
                &nbsp;</td>
			    <td align=right style="width:80px;"><b><%=formatnumber(SaldoGrandTot, 2)%></b></td>
			<%else%>
				
				
						
						<td align=right style="width:80px;"><b><%=formatnumber(debitTot, 2)%></b></td>
						<td align=right style="width:80px;"><b><%=formatnumber(kreditTot, 2)%></b></td>
						<%if kontonr = "0" then %>
						<td align=right style="width:80px;"><b><%=formatnumber(momsTot, 2)%></b></td>
						<%else %>
						<td align=right>
                            &nbsp;</td>
						<%end if%>
					    <td align=right style="width:80px;"><b><%=formatnumber(SaldoGrandTot, 2)%></b></td>
					
				
			<%end if%>
			
	<td colspan=3 align=right>&nbsp;
	<%if cdbl(formatnumber(SaldoGrandTot, 2)) = 0 AND kontonr = 0 then%>
	<a href="posteringer.asp?menu=kon&func=bog" class=vmenu>Bogfør >></a>
	<%end if%></td>
	</tr>
	<tr><td bgcolor="<%=bgthis%>" colspan=<%=cspan+2 %>>&nbsp;</td></tr>	
	</table>
	
	
	</div><!-- call table div-->
	<br /><br />
	Alle beløb og posteringer er angivet i <%=basisValISO %>
	
	
	<br><br><br>
	<a href="kontoplan.asp?menu=kon" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage til kontoplan</a><br><br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
