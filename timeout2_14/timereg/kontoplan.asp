<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/konto_inc.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->

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
	
	if len(request("FM_soeg")) <> 0 then 
	thiskri = request("FM_soeg")
	useKri = 1
	else
	thiskri = ""
	useKri = 0
	end if
	
	if len(request("FM_kd")) <> 0 then
	filterKD = request("FM_kd")
	else
	filterKD = 0
	end if
	
	if len(request("FM_tp")) <> 0 then
	filterTP = request("FM_tp")
	else
	filterTP = 0
	end if
	
	
	if len(request("FM_status")) <> 0 then
	filterStatus = request("FM_status")
	else
	filterStatus = 1
	end if
	
	
	if len(request("FM_nogletal")) <> 0 THEN 
	filterNogletal = request("FM_nogletal")
	else
	filterNogletal = 0
	end if
	
	
	
	
	'usePeriodefilter = request("useper") Altid til.
	
	call finddatoer()
	

    





	select case func
	case "slet"
	
	'*** Her spørges om det er ok at der slette ***'
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>slette</b> en <b>konto</b>. Er dette korrekt?<br>"_
	&"Kun konti uden posteringer kan slettes."
	slturl = "kontoplan.asp?menu=kon&func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,200)
	
	
	case "sletok"
	'*** Her slettes en konto ***
	
	kontonr = -3
	strSQL = "SELECT kontonr FROM kontoplan WHERE id = "& id
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	kontonr = oRec("kontonr")
	end if
	oRec.close
	
	'*** Sletter posteringer ***
	oConn.execute("DELETE FROM posteringer WHERE kontonr = "& kontonr &" OR modkontonr = "& kontonr &"")
	
	'**** Sletter konto ****
	oConn.execute("DELETE FROM kontoplan WHERE id = "& id &"")
	Response.redirect "kontoplan.asp?menu=erp"
	
	case "dbopr", "dbred"
	
	showmenu = request("showmenu")
	
	'*** Her indsættes en ny konto i db ****
		if len(request("FM_navn")) = 0 OR len(request("FM_kontonr")) = 0 then
		
		
		if showmenu = "no" then
		%>
		<!--#include file="../inc/regular/header_hvd_inc.asp"-->
		<%
		useleftdiv = "j"
		else
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		
		<%
		end if
		
		errortype = 46
		call showError(errortype)
		
		else		
				
				%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(request("FM_kontonr")) 
			if isInt > 0 OR instr(request("FM_kontonr"), ".") <> 0 then
				
				if showmenu = "no" then
				%>
				<!--#include file="../inc/regular/header_hvd_inc.asp"-->
				<%
				useleftdiv = "j"
				else
				%>
				<!--#include file="../inc/regular/header_inc.asp"-->
				
				<%
				end if
				
				errortype = 62
				call showError(errortype)
				
				isInt = 0
		
				else
				
							intkontonr = request("FM_kontonr")
							
							'**** Findes kontonr allerede ?? **
							errKontonrFindes = 0
							strSQL = "SELECT kontonr FROM kontoplan WHERE kontonr = "& intkontonr & " AND id <> "& id
							oRec.open strSQL, oConn, 3 
							while not oRec.EOF 
							
							errKontonrFindes = 1
							
							oRec.movenext
							wend
							oRec.close 
				
							if errKontonrFindes = 1 then
							
							if showmenu = "no" then
							%>
							<!--#include file="../inc/regular/header_hvd_inc.asp"-->
							<%
							useleftdiv = "j"
							else
							%>
							<!--#include file="../inc/regular/header_inc.asp"-->
							
							<%
							end if
							
							errortype = 46
							call showError(errortype)
		
					
					
					else
		
		
					function SQLBless(s)
					dim tmp
					tmp = s
					tmp = replace(tmp, "'", "''")
					SQLBless = tmp
					end function
					
					strNavn = SQLBless(request("FM_navn"))
					strEditor = session("user")
					strDato = year(now)&"/"&month(now)&"/"&day(now)
					
					kundeid = request("FM_kundeid")
					intMomskode = request("FM_momskode")
					intKeycode = request("FM_keycode")
					intDebkre = request("FM_debkre")
					intStatus = request("FM_status")
					intType = request("FM_type")
					
					if func = "dbopr" then
					oConn.execute("INSERT INTO kontoplan (navn, editor, dato, kontonr, kid, "_
					&" momskode, debitkredit, status, keycode, type) "_
					&" VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "_
					&" '"& intkontonr &"', '"& kundeid &"', "& intMomskode &", "&intDebkre&", "_
					&" "&intStatus&", "&intKeycode&", "& intType &")")
					else
					oConn.execute("UPDATE kontoplan SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', kontonr = '"& intkontonr &"', "_
					&" momskode = "& intMomskode &", kid = '"& kundeid &"', debitkredit = "& intDebkre &", "_
					&" status = "&intStatus&", keycode = "&intKeycode&", type = "& intType &" WHERE id = "&id&"")
					end if
					
					
					
                    if showmenu = "no" then
					Response.Write("<script language=""JavaScript"">window.close();</script>")
					else
					Response.redirect "kontoplan.asp?menu=erp&id="&id
					end if
					
					
					end if '** Validering 
				end if '** Validering 
			end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	intMomskode = 1
	varSubVal = "Opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	intdebkre = 2
	intType = 1
	intStatus = 1
	intNogletalskode = 0
	
	else
	strSQL = "SELECT navn, editor, dato, id, kid, kontonr, momskode, keycode, debitkredit, status, type FROM kontoplan WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intKid = oRec("kid")
	intdebkre = oRec("debitkredit")
	intStatus = oRec("status")
	intMomskode = oRec("momskode")
	varKontonr = oRec("kontonr")
	intNogletalskode = oRec("keycode")
	intType = oRec("type")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdaterpil" 
	end if
	
	if len(request("showmenu")) <> 0 then
	showmenu = "no"
	leftdiv = 20
	topdiv = 80
	else
	leftdiv = 90
	topdiv = 102
	showmenu = ""
	end if 
	%>
	
	
	
	
	<%if showmenu <> "no" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
    <%call menu_2014() %>
	<%else%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if%>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftdiv%>px; top:<%=topdiv%>px; visibility:visible;">
	
	<table cellspacing="0" cellpadding="2" border="0" width="600" bgcolor="#ffffff">
	<form action="kontoplan.asp?menu=kon&func=<%=dbfunc%>&showmenu=<%=showmenu%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	 <tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2 style="border-left:1px #003399 solid; border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=2 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td width="8" valign=top rowspan=2 style="border-right:1px #003399 solid; border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 class=alt><b><%=varbroedkrumme%> Konto.</b></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr bgcolor="#5582D2">
		<td style="border-left:1px #003399 solid; border-right:1px #003399 solid; padding-left:15;" colspan=4 class=alt>Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
    <tr>
		<td width="5" style="border-left:1px #003399 solid;" rowspan=8>&nbsp;</td>
		<td height=20><br /><font class=roed><b>*</b></font>&nbsp;<b>Kontonr:</b></td>
		<td><br /><%if dbfunc = "dbred" then%>
		<b><%=varKontonr%></b>
		<input type="hidden" name="FM_kontonr" value="<%=varKontonr%>">
		<%else%>
		<input type="text" name="FM_kontonr" value="<%=varKontonr%>" size="30" style="width:100px; border:1px #86b5e4 solid;">
		<%end if%></td>
		<td width="5" style="border-right:1px #003399 solid;" rowspan=8>&nbsp;</td>
	</tr>
	<tr>
		
		<td width="120"><font class=roed><b>*</b></font>&nbsp;<b>Kontonavn:</b></td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:300px; border:1px #86b5e4 solid;"></td>
		
	</tr>
	<tr>
		<td><b>Kontoindehaver:</b></td>
		<td><select name="FM_kundeid" style="width:300px;">
		
		<%
		if showmenu = "no" then '*** Fra Kontakt opr.
			if len(request("kundeid")) <> 0 then
			thisKid = request("kundeid")
			else
			thisKid = 0
			end if
		sqlWhere = " WHERE kid = "& thisKid
		else
		sqlWhere = "" 
		end if
		
		strSQL = "SELECT kkundenavn, kkundenr, kid, useasfak FROM kunder "& sqlWhere &" ORDER BY kkundenavn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		
		
			if func <> "red" then
				if oRec("useasfak") = 1 then
				selth = "SELECTED"
				else
				selth = ""
				end if
			else
				if oRec("kid") = cint(intKid) then
				selth = "SELECTED"
				else
				selth = ""
				end if
			end if
		
		
		
		%>
		<option value="<%=oRec("kid")%>" <%=selth%>><%=oRec("kkundenavn")%> (<%=oRec("kkundenr")%>)</option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		</select>&nbsp;
		
	</tr>
	<tr>
	<td><b>Egenskab:</b></td>
	<td>
		<select name="FM_debkre" style="width:300px;">
		<%
		select case intdebkre 
		case 1 
		selth1 = "SELECTED"
		selth2 = ""
		selth3 = ""
		selth4 = ""
		selth5 = ""
		selth6 = ""
		case 2
		selth1 = ""
		selth2 = "SELECTED"
		selth3 = ""
		selth4 = ""
		selth5 = ""
		selth6 = ""
		case 3
		selth1 = ""
		selth2 = ""
		selth3 = "SELECTED"
		selth4 = ""
		selth5 = ""
		selth6 = ""
		case 4
		selth1 = ""
		selth2 = ""
		selth3 = ""
		selth4 = "SELECTED"
		selth5 = ""
		selth6 = ""
		case 5
		selth1 = ""
		selth2 = ""
		selth3 = ""
		selth4 = ""
		selth5 = "SELECTED"
		selth6 = ""
		case 6
		selth1 = ""
		selth2 = ""
		selth3 = ""
		selth4 = ""
		selth5 = ""
		selth6 = "SELECTED"
		end select
		%>
		<option value="1" <%=selth1%>>Kreditor</option>
		<option value="2" <%=selth2%>>Debitor</option>
		<option value="3" <%=selth3%>>Samlekonto for omsætning</option>
		<option value="4" <%=selth4%>>Samlekonto for udgifter</option>
		<option value="5" <%=selth5%>>Overskrift for gruppe</option>
		<option value="6" <%=selth6%>>Ingen</option>
		
		</select></td>
	</tr>
	<tr>
		<td><b>Type:</b></td>
		<td><select name="FM_type" id="FM_type">
		<%select case cint(intType)
		case 1
		selth1 = "SELECTED"
		selth2 = ""
		selth3 = ""
		case 2
		selth1 = ""
		selth2 = "SELECTED"
		selth3 = ""
		case else
		selth1 = ""
		selth2 = ""
		selth3 = "SELECTED"
		end select
		%>
		<option value="1" <%=selth1%>>Drift</option>
		<option value="2" <%=selth2%>>Status</option>
		<option value="3" <%=selth3%>>Ingen</option>
		</select>
		</td>
	</tr>
	
	<tr>
		<td valign=top><b>Momskode:</b></td><td>
		<select name="FM_momskode" id="FM_momskode">
		<%
		'if func = "red" then
		'strSQL = "SELECT navn, id, kvotient FROM momskoder WHERE id="&intMomskode
		'else
		strSQL = "SELECT navn, id, kvotient FROM momskoder ORDER BY navn"
		'end if
		
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		if intMomskode = oRec("id") then
		momSel = "SELECTED"
		else
		momSel = ""
		end if
		%>
		<option value="<%=oRec("id")%>" <%=momSel%>><%=oRec("navn")%> (<%=oRec("kvotient")%>)</option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		</select></td>
	</tr>
	
	<tr>
		<td><b>Nøgletalskode:</b></td>
		<td><select name="FM_keycode" style="width:300px;">
		<option value="0">Ingen</option>
		<%
		strSQL = "SELECT navn, id, ap_type FROM nogletalskoder ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		if intNogletalskode = oRec("id") then
		keySel = "SELECTED"
		else
		keySel = ""
		end if
		
		select case oRec("ap_type")
		case 1
		ap_type = "Aktiver" 
		case 2
		ap_type = "Passiver"
		case else
		ap_type = "Ingen"
		end select
		
		
		%>
		<option value="<%=oRec("id")%>" <%=keySel %>><%=oRec("navn")%> (<%=ap_type%>)</option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		</select></td>
	</tr>
	<tr>
	<%
	select case intStatus
	case 1
	stSEL1 = "SELECTED"
	stSEL2 = ""
	case 2
	stSEL1 = ""
	stSEL2 = "SELECTED"
	case else
	stSEL1 = "SELECTED"
	stSEL2 = ""
	end select
	%>
		<td><b>Status:</b></td>
		<td><select name="FM_status">
		<option value="1" <%=stSEL1 %>>Åben</option>
		<option value="2" <%=stSEL2 %>>Lukket</option>
		</select></td>
	</tr>
	<tr>
		<td colspan="4" align="center" style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;"><input type="image" src="../ill/<%=varSubVal%>.gif"><br><br></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<%if showmenu <> "no" then%>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<%end if%>
	<br>
	<br>
	</div>
	<%case else
	
	print = request("print")
	
	%>
	
	
	<%if print <> "j" then 
	dtop = 102
	dleft = 90
	tw = "770"
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

    
 <SCRIPT Language=JavaScript>
     function BreakItUp() {
         //Set the limit for field size.
         var FormLimit = 302399

         //Get the value of the large input object.
         var TempVar = new String
         TempVar = document.theForm.BigTextArea.value

         //If the length of the object is greater than the limit, break it
         //into multiple objects.
         if (TempVar.length > FormLimit) {
             document.theForm.BigTextArea.value = TempVar.substr(0, FormLimit)
             TempVar = TempVar.substr(FormLimit)

             while (TempVar.length > 0) {
                 var objTEXTAREA = document.createElement("TEXTAREA")
                 objTEXTAREA.name = "BigTextArea"
                 objTEXTAREA.value = TempVar.substr(0, FormLimit)
                 document.theForm.appendChild(objTEXTAREA)

                 TempVar = TempVar.substr(FormLimit)
             }
         }
     }

	</SCRIPT>

	
	
	<%
        
           
        call menu_2014() 
        
         %>


   

	
	<%else
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
	<%
	dtop = 60
	dleft = 20
	tw = "580"
	end if
	
	
	 
	    if len(trim(request("sort"))) <> 0 then
	    sort = request("sort")
	            
	            if sort = 2 then
	            sortCHK2 = "CHECKED"
	            sortCHK1 = ""
	            sortOrderKri = "kontoplan.kontonr, kkundenavn"
	            else
	            sortCHK2 = ""
	            sortCHK1 = "CHECKED"
	            sortOrderKri = "kkundenavn"
	            end if
	   
	    else
	    sort = 2
	    sortCHK2 = "CHECKED"
	    sortCHK1 = ""
	    sortOrderKri = "kontoplan.kontonr, kkundenavn"
	    end if
        
        
        if len(request("sort")) <> 0 then
            
            if len(request("visikkenul")) <> 0 then
            visikkenul = 1
            visikkenulCHK = "CHECKED"
            else
            visikkenul = 0
            visikkenulCHK = ""
            end if
        else
            if len(request.cookies("erp")("viskunkontimpos")) <> 0 then
                if request.cookies("erp")("viskunkontimpos") = 1 then
                visikkenul = 1
                visikkenulCHK = "CHECKED"
                else
                visikkenul = 0
                visikkenulCHK = ""
                end if
             else
                visikkenul = 0
                visikkenulCHK = ""
             end if
       end if 
       
       response.cookies("erp")("viskunkontimpos") = visikkenul
        
       expTxt = ""
	
        if len(trim(request("nolist"))) <> 0 then
        nolist = 1
        else
	    nolist = 0
        end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=dleft%>px; top:<%=dtop%>px; visibility:visible;">
	
	
	
	<%if print <> "j" then %>
	        <%call filterheader_2013(0,0,850,pTxt)%>
	        <h4>Kontoplan</h4>
	        <table cellspacing="0" cellpadding="2" border="0" width=100%>
			<tr><form action="kontoplan.asp?menu=kon" method="POST">
			<td colspan=2><b>Søg</b> (Kontonavn / nr., % vis alle)<br /><input type="text" name="FM_soeg" id="FM_soeg" value="<%=thiskri%>" style="width:300px;"></td>
			<td align=right valign=bottom style="padding-right:10px"><b>Konto type:</b><br /><select name="FM_tp" id="FM_tp" >
			<%
			select case filterTP 
			case 0
			sel0 = "SELECTED"
			sel1 = ""
			sel2 = ""
			case 1
			sel0 = ""
			sel1 = "SELECTED"
			sel2 = ""
			case 2
			sel0 = ""
			sel1 = ""
			sel2 = "SELECTED"
			case else
			sel0 = "SELECTED"
			sel1 = ""
			sel2 = ""
			end select
			%>
			<option value="0" <%=sel0%>>Alle</option>
			<option value="2" <%=sel2%>>Drift</option>
			<option value="1" <%=sel1%>>Status</option>
			</select></td>
			<td align=right valign=bottom style="padding-right:10px"><b>Egenskab:</b><br /><select name="FM_kd" id="FM_kd" >
			<%
			select case filterKD 
			case 0
			sel0 = "SELECTED"
			sel1 = ""
			sel2 = ""
			case 1
			sel0 = ""
			sel1 = "SELECTED"
			sel2 = ""
			case 2
			sel0 = ""
			sel1 = ""
			sel2 = "SELECTED"
			case else
			sel0 = "SELECTED"
			sel1 = ""
			sel2 = ""
			end select
			%>
			<option value="0" <%=sel0%>>Alle</option>
			<option value="2" <%=sel2%>>Debitorer</option>
			<option value="1" <%=sel1%>>Kreditorer</option>
			</select></td>
			
			
			
			<td align=right valign=bottom style="padding-right:10px"><b>Status:</b><br /><select name="FM_status" id="FM_status" >
			<%
			select case filterStatus 
			case 0
			sel0 = "SELECTED"
			sel1 = ""
			sel2 = ""
			case 1
			sel0 = ""
			sel1 = "SELECTED"
			sel2 = ""
			case 2
			sel0 = ""
			sel1 = ""
			sel2 = "SELECTED"
			case else
			sel0 = "SELECTED"
			sel1 = ""
			sel2 = ""
			end select
			%>
			<option value="0" <%=sel0%>>Alle</option>
			<option value="1" <%=sel1%>>Åbne</option>
			<option value="2" <%=sel2%>>Lukkede</option>
			</select></td>
			
		
			
			<td align=right valign=bottom><b>Nøgletalskode:</b><br /><select name="FM_nogletal" id="FM_nogletal" >
			<option value="0" <%=sel0%>>Alle</option>
			<%
			
			strSQLnt = "SELECT id, navn, ap_type FROM nogletalskoder WHERE id <> 0 ORDER BY navn"
			oRec.open strSQLnt, oConn, 3
			while not oRec.EOF
			
			if cint(filterNogletal) = oRec("id") then
			selNt = "SELECTED"
			else
			selNt = ""
			end if 
			
			%>
			<option value="<%=oRec("id") %>" <%=selNt%>><%=oRec("navn") %></option>
			
			<%
			oRec.movenext
			wend
			oRec.close
			
			
			%>
		
			</select></td>
			
			</tr>
			
			
			<tr>
			<td colspan=6><br /><b>Periode:</b></td>
			</tr>
			<tr>
			<!--#include file="inc/weekselector_b.asp"-->
			
			<td align=right colspan=4>
                <input id="Submit1" type="submit" value="Vis konti" />
                </td>
			</tr>
			<tr><td colspan=4><b>Sortér</b><br />
             <input id="sort" name="sort" type="radio" value="1" <%=sortCHK1 %> />Kontoindehaver,  
            <input id="sort" name="sort" type="radio" value="2" <%=sortCHK2 %> />Kontonr</td>
            <td colspan=2 align=right>
                <input id="visikkenul" name="visikkenul" type="checkbox" <%= visikkenulCHK %> />Vis kun konti med posteringer i valgte periode.</td></tr>
			<input type="hidden" name="useper" value="j">
			</form>
			</table>
			<!-- filter header sLut -->
	</td></tr></table>
	</div>
			
	
		
	
	
	
	
	    
	
<%else %>	
<h4>Kontoplan</h4>
	<%end if%>

   
      <% 
    if cint(nolist) = 1 then
        response.end
	end if%>


	<br><br><br>Periode: <b><%=formatdatetime(strDag&"/"&strMrd&"/"&strAar,1)%></b> til
	<b><%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%></b> <br> 
	
	
   

	<table cellspacing="0" cellpadding="0" border="0" width=100% bgcolor="#FFFFFF">
	 <tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2 style="border-top:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=11 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2 style="border-top:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
	<%
	strKnavn = ""
	strSQL3 = "SELECT kkundenavn FROM kunder WHERE useasfak = 1"
	oRec3.open strSQL3, oConn, 3
	if not oRec3.EOF then
	
	strKnavn = oRec3("kkundenavn")
	
	end if
	oRec3.close
	%>
	
		<td colspan=11 class=alt style="padding:15px 10px 0px 10px;"><h3>Kontoplan <%=strKnavn %></h3></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td width="5" style="border-left:1px #003399 solid;">&nbsp;</td>
	<td height="30" align=right style="padding-right:10px;"><b>Kontonr</b></td>
	<td><b>Kontonavn</b></td>
	<td><b>Kontoindehaver</b></td>
	<td><b>Egenskab</b></td>
	
	<!--<td width="90"><b>Nøgletalskode</b></td>-->
	<td width=90><b>Momskode</b></td>
	<td width=60><b>Type</b></td>
	<td width=90><b>Nøgletalskode</b></td>
	
	
	<%if print <> "j" then%>
	<td align="center"><b>Status</b></td>
	<%else%>
	<td>&nbsp;</td>
	<%end if%>
	
	<td bgcolor="#FFFFFF" style="border-left:1px #cccccc dashed; border-right:1px #cccccc dashed; padding-right:10px;" align="right"><b>Saldo</b></td>
	<%if print <> "j" then%>
	<td width="20">&nbsp;</td>
	<td width="20">&nbsp;</td>
	<%else%>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<%end if%>
	
	<td width="5" style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	
	expTxt = expTxt &"Kontonr;Kontonavn;Kontoindehaver;Debitor/kreditor;Momskode;Type;Nøgletalskode;Status;Saldo"
	expTxt = expTxt &";Dato;Bilagsnr;Tekst;Debit;Kredit"
	
	
	if useKri <> 0 then
	useSQLKri = " WHERE kontoplan.navn LIKE '"&thiskri&"%' OR kontoplan.kontonr = '"&thiskri&"' OR kkundenavn LIKE '"&thiskri&"%'" 
	else
	useSQLKri = " WHERE kontoplan.id > 0 "
	end if
	
	select case filterKD
	case 0
	dkkri = " AND debitkredit <> 99 "
	case 1
	dkkri = " AND debitkredit = 1 "
	case 2
	dkkri = " AND debitkredit = 2 "
	case else
	dkkri = " AND debitkredit <> 99 "
	end select
	
	select case filterTP
	case 0
	typkri = " AND type <> 99 "
	case 1
	typkri = " AND type = 2 "
	case 2
	typkri = " AND type = 1 "
	case else
	typkri = " AND type <> 99 "
	end select
	
	select case filterStatus
	case 0
	statusKri = " AND kontoplan.status <> -1 "
	case 1
	statusKri = " AND kontoplan.status = " & filterStatus
	case 2
	statusKri = " AND kontoplan.status = " & filterStatus
	end select
	
	select case filterNogletal
	case 0
	ntKri = " AND keycode <> -1"
	case else
	ntKri = " AND keycode = " & filterNogletal
	end select
	
	if visikkenul <> 1 then
	grpby = "kontoplan.kontonr"
	else
	grpby = "p.kontonr"
	end if
	
	imoms = 0
	umoms = 0
	saldo = 0
	
	if print <> "j" then
	btop = 0
	else
	btop = 1
	end if
	
	periodefilter = " posteringsdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	

	strSQL = "SELECT kontoplan.id AS id, kontoplan.navn AS navn, kontoplan.kontonr, kontoplan.kid AS kid, "_
	&" kontoplan.type, kkundenavn, kkundenr, kunder.kid AS kunderid, momskode, "_
	&" kontoplan.status, debitkredit, momskoder.navn AS momskodenavn, "_
	&" nogletalskoder.navn AS nogletalkodenavn, sum(p.nettobeloeb) AS saldo "_
	&" FROM kontoplan LEFT JOIN kunder ON (kunder.kid = kontoplan.kid) "_
	&" LEFT JOIN momskoder ON (momskoder.id = momskode) "_
	&" LEFT JOIN nogletalskoder ON (nogletalskoder.id = keycode) "_
	&" LEFT JOIN posteringer p ON (p.kontonr = kontoplan.kontonr AND "& periodefilter &" AND p.status = 1)"_
    &  useSQLKri &" "& dkkri &" "& typkri &" "& statusKri &" "& ntKri &" GROUP BY "& grpby &" ORDER BY "& sortOrderKri 
	
	'Response.Write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	if print = "j" then
	bgcol = "#8caae6" '"#FFFF99"
	else
	
		if cint(id) = oRec("id") then
		bgcol = "#ffff99"
		else
			
			'if oRec("debitkredit") = 5 then 'Overskrift
		    'bgcol = "#cccccc"
		    'else
		    bgcol = "#FFFFFF"
		    'end if
			
		
		end if
	end if
	%>
	
	<tr bgcolor="<%=bgcol%>">
		<td style="border-left:1px #003399 solid; border-top:<%=btop%>px #8cAAe6 solid; border-bottom:1px #cccccc dashed;">&nbsp;</td>
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid; padding-right:10px;" align=right><b><%=oRec("kontonr")%></b></td>
		<td height="25" style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;">
		<%if print <> "j" then%>
		<a href="posteringer.asp?menu=erp&id=<%=oRec("id")%>&kontonr=<%=oRec("kontonr")%>&kid=<%=oRec("kunderid")%>" class=vmenu><%=oRec("navn")%> </a>
		<%else%>
		<b><%=oRec("navn")%></b>
		<%end if%></td>
		
		<%expTxt = expTxt &"xx99123sy#z"&oRec("kontonr")&";"&oRec("navn")
		
		if oRec("debitkredit") <> 5 then
		expTxt = expTxt &";"&oRec("kkundenavn") 
		else
		expTxt = expTxt &";"
		end if
		%>
		
		
		
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;"><i>
		<%if oRec("debitkredit") <> 5 then %>
		<%=oRec("kkundenavn")%>
		<%end if%>
        &nbsp;</i></td>
        
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;"><%
		
		dbst = ""
		
		'if oRec("type") = 2 then 'status
			
			if oRec("debitkredit") = 1 then
			dbst = "Kreditor"
			end if
			
			if oRec("debitkredit") = 2 then
			dbst = "Debitor"
			end if
			
	    'else
		    
		    if oRec("debitkredit") = 3 then
		    dbst = "Samlekonto for oms."
		    end if
		    
		    if oRec("debitkredit") = 4 then
		    dbst = "Samlekonto for udg."
		    end if
		    
		    if oRec("debitkredit") = 5 then
		    dbst = "Overskrift"
		    end if
		    
		    if oRec("debitkredit") = 6 then
		    dbst = ""
		    end if
		    
		 'end if%>
		
		<%=dbst&"&nbsp;"%>
		
		</td>
		
		
		
		
		
		<!--<td><=oRec("nogletalkodenavn")%></td>-->
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;">
		<%if oRec("debitkredit") <> 5 then %>
		<%
		Response.write oRec("momskodenavn")
		%>
		<%end if %>
		&nbsp;
		</td>
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;"><%
		tpe = ""
		
		select case oRec("type")
		case 1
		tpe = "Drift"
		case 2
		tpe = "Status"
		case else
		tpe = ""
		end select%>
		
		<%if oRec("debitkredit") <> 5 then %>
		<b><%=tpe %></b>
		<%end if %>
            &nbsp;</td>
		
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;" class=lillegray><%=oRec("nogletalkodenavn") %>
            &nbsp;</td>
		
		
		<%
		
		sta = ""
		if print <> "j" AND oRec("debitkredit") <> 5 then%>
			
			<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;" align="center">
			<%if oRec("status") = 1 then
			sta = "<font color=LimeGreen>Åben</font>"
			staexp = "åben"
			else
			sta = "<font color=darkred>Lukket</font>"
			staexp = "lukket"
			end if%>
			
			<%=sta %></td>
		
		<%else%>
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;">&nbsp;</td>
		<%end if%>
		
		
		
		<td align=right style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid; padding-left:10px; padding-right:10px; border-left:1px silver dashed; border-right:1px silver dashed;">
		<%
		
		saldo = oRec("saldo")
		
		if len(saldo) <> 0 OR saldo > 0 then
		saldo = saldo
		else
		saldo = 0
		end if
		
		
		totalSaldo = totalSaldo + (saldo)
		
		if oRec("debitkredit") <> 5 then
		expTxt = expTxt &";"&dbst&";"&oRec("momskodenavn")&";"&tpe&";"&oRec("nogletalkodenavn")&";"&staexp&";"&saldo 
		else
		expTxt = expTxt &";"&dbst&";;;;;"
		end if
		%>
		
	    <%if oRec("debitkredit") <> 5 then %>
		<b><u><%=formatnumber(saldo,2)%></u></b>
		<%else %>
		&nbsp;
		<%end if%>
		</td>
		<%if print <> "j" then%>
		<td style="padding-left:3px; border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;"><a href="kontoplan.asp?menu=erp&func=red&id=<%=oRec("id")%>"><img src="../ill/blyant.gif" width="12" height="11" alt="Rediger konto" border="0"></a></td>
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;">
		<%if level = 1 then 
		        
		        
		       posFindes = 0
		       '*** Tjekker om der findes posteringer på konto ***'
		       strSQLpos = "SELECT p.kontonr FROM posteringer p WHERE p.kontonr = "& oRec("kontonr") & " GROUP BY p.kontonr" 
		       oRec2.open strSQLpos, oConn, 3
		       
		       if not oRec2.EOF then
		       
		       posFindes = 1
		       
		       end if
		       oRec2.close
    
		        
		if posFindes <> 1 then%>
		<a href="kontoplan.asp?menu=erp&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" alt="Slet" border="0"></a>
		<%else %>
		&nbsp;
		<%end if %>
		
		<%else %>
		&nbsp;
		<%end if %>
		</td>
		
		<%else%>
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;">&nbsp;</td>
		<td style="border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;">&nbsp;</td>
		<%end if%>
		<td style="border-right:1px #003399 solid; border-bottom:1px #cccccc dashed; border-top:<%=btop%>px #8cAAe6 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
						
						
						if print = "j" then
						
						%>
						
						<tr>
							<td style="border-left:1px #003399 solid;">&nbsp;</td>
							<td colspan=11>
							
							<table>
							<tr>
								<td bgcolor="#ffffff" colspan="5"><br>Posteringer:</td>
							</tr>
							<tr>
								<td width=70><b>Dato</b></td>
								<td width="90" align=right><b>Bilagsnr</b></td>
								<td width="140" style="padding-left:10;"><b>Tekst</b></td>
								<td width="70" align="right"><b>Konto</b></td>
								<td width="70" align="right"><b>Modkonto</b></td>
								<td width="80" align="right"><b>Debit</b></td>
								<td width="80" align="right"><b>Kredit</b></td>
							</tr>
							
							
						<%
						
						end if
						
						
						
						'**** Ved eksporter/print hentes posteringer ****
						kassekladeorkonto = "(p.kontonr = "& oRec("kontonr") &" AND p.status = 1)"
						momskri = " AND (kontoplan.kontonr = p.kontonr) "
						useSQLKri = " p.id > 0 "
								
						strSQL3 = "SELECT p.id, p.bilagstype, p.kontonr, "_
						&" p.modkontonr, p.bilagsnr, p.beloeb, p.nettobeloeb, "_
						&" p.moms, p.tekst, posteringsdato, p.status, "_
						&" p.oprid, debitkredit, kontoplan.navn AS kontonavn FROM posteringer p, kontoplan "_
						&" WHERE "& kassekladeorkonto &" AND "& periodefilter &" "& momskri &" AND "& useSQLKri &" ORDER BY posteringsdato DESC, id DESC"
						
						'Response.Write strSQL3
						'Response.flush
						p = 0
						oRec3.open strSQL3, oConn, 3
						while not oRec3.EOF 
						
								
								if print = "j" then%>
								
								<tr bgcolor="#ffffff">
									
									<td><%=oRec3("posteringsdato")%></td>
									<td align=right><b><%=oRec3("bilagsnr")%></b></td>
									<td style="padding-left:10;"><%=oRec3("tekst")%></td>
									
									<td align="right"><b><%=oRec3("kontonr")%></b></td>
									
									<td align="right" class=lillegray><%=oRec3("modkontonr")%></td>
									
									<%if oRec3("beloeb") >= 0 then %>
									<td align=right><%=formatnumber(oRec3("beloeb"), 2)%></td>
									<td>&nbsp;</td>
                                   <%else %>
                                        <td>&nbsp;</td>
                                        <td align=right><%=formatnumber(oRec3("beloeb"), 2)%></td>
									<%end if %>
									
								</tr>
								<%
								
								end if
								
								
								expTxt = expTxt &"xx99123sy#z"&oRec("kontonr")&";;;;;;;;;"&oRec3("posteringsdato")&";"&oRec3("bilagsnr")&";"&oRec3("tekst")
								
								
								
								if oRec3("beloeb") >= 0 then 
								expTxt = expTxt & ";"&formatnumber(oRec3("beloeb"), 2)&";;"
								else
								expTxt = expTxt & ";;"&formatnumber(oRec3("beloeb"), 2)&";"
								end if
								
								p = p + 1
								oRec3.movenext
								wend
								oRec3.close
								
								
								if print = "j" then%>
								<tr>
									<td bgcolor="#ffffff" colspan="5" height=20>&nbsp;</td>
								</tr>
								</table>
								</td>
								<td style="border-right:1px #003399 solid;">&nbsp;</td>
								</tr>
								<% end if
						
						
						
						
	oRec.movenext
	wend
	oRec.close
	%>
	
	<tr><td bgcolor="#8caae6" colspan=13 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;">&nbsp;</td></tr>	
	</table>
	<br />
	Totalsaldo på viste konti i periode: <b>
	<%=formatnumber(totalSaldo,2) %></b>
	<br><br><br>
	
	
	        <%
            if print <> "j" then
            
            ptop = 0
            pleft = 880
            pwdt = 140
            
            call eksportogprint(ptop, pleft, pwdt)

            call htmlparseCSV(expTxt)
            expTxt = htmlparseCSVtxt
            %>
            
             <form action="kontoplan_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>">
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=expTxt%>">
			<input type="hidden" name="txt20" id="txt20" value="">
			
            <tr>
                <td align=center><input type="image" src="../ill/export1.png">
                </td>
                <td>.csv fil eksport</td>
               </tr>
                </form>
                <tr>
                <td align=center>
                <a href="kontoplan.asp?menu=kon&print=j&FM_soeg=<%=thiskri%>&FM_kd=<%=filterKD%>&FM_tp=<%=filterTP%>&visikkenul=<%=visikkenul%>&FM_nogletal=<%=filterNogletal%>&FM_status=<%=filterStatus %>" class=vmenu target="_blank"><img src="../ill/printer3.png" border=0></a></td>
		       <td>Print version</td>
               </tr>
               </form>
	            
               </table>


<br /><br />
                    

<span>
	 
            <% 
               nWdt = 160
                nTxt = "Opret ny konto"
                nLnk = "kontoplan.asp?menu=kon&func=opret"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) 
                %>
	    </span>


            </div>
            
            <%
            else
            Response.Write("<script language=""JavaScript"">window.print();</script>")
            end if%>
	
	       
	<br /><br /><br />
	
	<a href="Javascript:history.back()" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage</a><br><br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
