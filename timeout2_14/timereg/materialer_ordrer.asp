<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/mat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->
<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
		if len(request("id")) <> 0 then
		id = request("id")
		else
		id = 0
		end if
	end if
	
	if len(trim(request("showmenu"))) then
	showmenu = request("showmenu")
	else
	showmenu = 1
	end if
	
	thisfile = "materialer_ordrer.asp"
	menu = request("menu")
	
	if len(trim(request("FM_medarb"))) then
	selmedarb = request("FM_medarb")
	Response.Cookies("mat")("medarb") = selmedarb
	else
	    if Request.Cookies("mat")("medarb") <> "" then
	    selmedarb = Request.Cookies("mat")("medarb")
	    else
	    selmedarb = 0
	    end if
	end if
	
	if len(trim(request("FM_sog"))) <> 0 then
	sogval = request("FM_sog")
	usesogval = 1
	else
	sogval = ""
	usesogval = 0
	end if
	
	session.lcid = 1030 'DK
	
	select case func
	
	case "slet"
	'*** Her spørges om det er ok at der slettes en ordre ***'
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:210; top:100; background-color:#ffff99; visibility:visible; border:2px red dashed; padding:15px;">
	<table cellspacing="2" cellpadding="2" border="0" bgcolor="#ffff99" width=400>
	<tr>
	    <td><img src="../ill/alert.gif" alt="" border="0">&nbsp;<h4>Slet ordre?</h4>
		Du er ved at <b>slette</b> en ordre. Du vil samtidig slette <b>alle ordrelinier</b> på denne ordre.
		Er du sikker på at du vil slette denne ordre?<br><br>
		<a href="materialer_ordrer.asp?func=sletok&id=<%=id%>" class=vmenu>Ja!, Slet ordre</a><br><br>
		<a href="javascript:history.back()" class=vmenualt>Nej, jeg vil ikke slette denne ordre!</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en ordre og oedrelinier ****'
	
    '*** Sletter ordre ****'
	oConn.execute("DELETE FROM materiale_ordrer WHERE id = "& id &"")
	
	'*** Sletter ordre linier ****'
	oConn.execute("DELETE FROM materiale_ordrer_linier WHERE ordreid = "& id &"")
	
	
	'Response.flush
	
	
	Response.redirect "materialer_ordrer.asp"
	
	
	
	
	
	case "dbopr", "dbred"
	
	editor = session("user")
	editdato = year(now) &"/"& month(now) &"/"& day(now)
	odrdato = trim(request("FM_odrdato_dag")) &"-"& trim(request("FM_odrdato_md")) &"-"& trim(request("FM_odrdato_aar"))
	odrdatoSQL = trim(request("FM_odrdato_aar")) &"/"& trim(request("FM_odrdato_md")) &"/"& trim(request("FM_odrdato_dag"))
	
	
	  '*** errors ***'
	  if isDate(odrdato) = false then
	   
	        %>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<%
			errortype = 100
			call showError(errortype)
			
			Response.end
	    
	    end if
	    
	 
	 %>
			<!--#include file="inc/isint_func.asp"-->
			<%    
	
	        odrid = request("FM_odrid")  
	        odrstatus = request("FM_odrstatus") 
	
	        
	        call erDetInt(odrid) 
			if isInt > 0 then
				
		    %>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<%
			errortype = 103
			call showError(errortype)
			isInt = 0
			
			Response.end
			end if
	
	
	if func = "dbopr" then
	    
	    
	    '*** errors ***'
	    '*** Findes ordre nr ****'
	    findesordnr = 0
	    strSQLfindes = "SELECT id, ordrenr FROM materiale_ordrer WHERE ordrenr = "& odrid
	    oRec.open strSQLfindes, oConn, 3
	    if not oRec.EOF then
	    findesordnr = 1
	    end if
	    oRec.close
	
	 
	    if findesordnr = 1 then
	        
	        lastordnr = 0
	        strSQLlastnr = "SELECT ordrenr FROM materiale_ordrer WHERE id <> 0 ORDER BY id DESC"
	        oRec.open strSQLlastnr, oConn, 3
	        if not oRec.EOF then
	        lastordnr = oRec("ordrenr")
	        end if
	        oRec.close
	        
	        %>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<%
			errortype = 102
			call showError(errortype)
			
			Response.end
	    
	    end if
	
	
	'*** Opretter ordre ***'
	strSQLodr = "INSERT INTO materiale_ordrer (editor, dato, ordrenr, ordredato, status, mid) VALUES "_
	&"('"& editor &"', '"& editdato &"', "& odrid &", '"& odrdatoSQL &"', '"& odrstatus &"', "& session("mid") &")"
	'Response.Write strSQLodr
	'Response.flush
	
	oConn.execute(strSQLodr)
	
	lastordId = 0
	strSQLordreid = "SELECT id FROM materiale_ordrer WHERE id <> 0 ORDER BY id DESC"
	oRec.open strSQLordreid, oConn, 3
	if not oRec.EOF then
	lastordId = oRec("id")
	end if
	oRec.close
	
	else
	
	strSQLupd = "UPDATE materiale_ordrer SET editor = '"& editor &"', "_
	&" dato = '"& editdato &"', ordrenr = "& odrid &", "_
	&" ordredato = '"& odrdatoSQL &"', status = '"& odrstatus &"', mid = "& session("mid") &""_
	&" WHERE id = "& id
	
	
	'Response.Write strSQLupd
	
	oConn.execute(strSQLupd)
	
	lastordId = id
	
	'*** Renser ud i ordrelinier inde indlæsning af nye ***'
	oConn.execute("DELETE FROM materiale_ordrer_linier WHERE ordreid = "& id)
	
	
	end if
	
	'******************************'
	'**** Opretter ordre linier ***'
	'******************************'
	
	thisID = 0
    varenr = 0
    varenavn = ""
    varegrp = 0
    varebetegn = ""
    vareantal = 0
    varestatus = 0
	
	
	antalordrelinier = split(request("FM_id"),",")
	
	'Response.Write "antalordrelinier: "& request("FM_id")
	
	for x = 0 to UBOUND(antalordrelinier)
	    
	    thisID = trim(antalordrelinier(x))
	    
	    varenr = request("FM_varenr_"&thisID&"")
	    varenavn = request("FM_navn_"&thisID&"")
	    varegrp = request("FM_gruppe_"&thisID&"")
	    varebetegn = request("FM_betegn_"&thisID&"")
	    vareantal = request("FM_antal_"&thisID&"")
	    
	    if len(trim(request("FM_status_"&thisID&""))) <> 0 then
	    varestatus = request("FM_status_"&thisID&"")
	    else
	    varestatus = 0
	    end if
	    
	     
			call erDetInt(vareantal) 
			if isInt > 0 then
				
		    %>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<%
			errortype = 101
			call showError(errortype)
			isInt = 0
			
			    if func = "dbopr" then
			    '*** Sletter den netop oprerttede ordre ved. fej li antal på en ordre linie, således ' 
			    '*** at der ikke optræsder tomme / ½ ordre. ***'
			    
			    oConn.execute("DELETE FROM materiale_ordrer WHERE id = "& lastordId)
			    
			    end if
			
			Response.end
			
			end if
			
			
	    
	    if cint(vareantal) > 0 then
	    
	    '*** Opretter ordre linier *'
	    strSQLodrL = "INSERT INTO materiale_ordrer_linier "_
	    &" (ordreid, matid, navn, betegnelse, varenr, gruppe, status, antal) "_
	    &" VALUES "_
	    &" ("& lastordId &", 0, '"& varenavn &"', "_
	    &" '"& varebetegn &"', '"& varenr &"', '"& varegrp &"', "& varestatus &", "& vareantal &")"
	    
	    'Response.Write strSQLodrL &"<br>"
	    'Response.flush
	    
	    oConn.execute(strSQLodrL)
	    '***'
	    
	    
	        '**** Opdaterer antal materialer i materialer lager hvis materiale er ankommet ***'
	        if varestatus = 1 then
    	    
	        matgrpID = 0
	        strSQL = "SELECT id FROM materiale_grp WHERE navn = '"& varegrp &"'"
	        oRec.open strSQL, oConn, 3
	        if not oRec.EOF then
	        matgrpID = oRec("id")
	        end if 
	        oRec.close
    	    
    	    
	        strSQL = "UPDATE materialer SET antal = (antal + "& vareantal &") WHERE varenr = '"& varenr &"' AND matgrp = "&matgrpID
	        oConn.execute(strSQL)
    	    
	        end if
	        '****'
	    
	    end if
	
	next
	
	'Response.end
	
	
    Response.Write("<script language=""JavaScript"">window.close();</script>")
    Response.Write("<script language=""JavaScript"">window.opener.location.href('materialer_ordrer.asp');</script>")
	
	
	case "opr", "red"%>
	
	<script language=javascript>
	
	function showordrelinie(){
	
	var nextn = 0;
	nextn = (document.getElementById("lastn").value / 1) - 1
	
	//alert(nextn)
	for (m=0;m<9;m++){
	document.getElementById("td_"+nextn+"_"+m+"").style.display = ""
	document.getElementById("td_"+nextn+"_"+m+"").style.visibility = "visible"
	}
	
	document.getElementById("lastn").value = nextn
	}
	
	function popUp(URL,width,height,left,top) {
		window.open(URL, 'navn', 'left='+left+',top='+top+',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
	}
    </script>
	
	
	
	<%
	if func = "opr" then
	funcval = "dbopr"
	else
	funcval = "dbred"
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:20; visibility:visible;">
	
	<h3><img src="../ill/ac0030-24.gif" width="24" height="24" alt="" border="0">&nbsp;Materialeordre</h3>
	
   
	     <%dontshowDD = "1" %>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			
		
		<% 
		
		if func <> "red" then
		
		matidSQLkri = "m.id = 0"
	    matids = split(request("matids"), ",")
	    for x = 0 to UBOUND(matids)
	    
	    if x <> 0 then
	    matidSQLkri = matidSQLkri &" OR  m.id = "& matids(x)
	    end if
	    
	    next
	
	    
	    ordrenr = 0
	    strSQLordid = "SELECT ordrenr FROM materiale_ordrer WHERE id <> 0 ORDER BY id DESC"
	    oRec.open strSQLordid, oConn, 3
	    if not oRec.EOF then
	    ordrenr = oRec("ordrenr") + 1 
	    end if
	    oRec.close
	    
	    else
	    
	    strSQL = "SELECT ordrenr, mid, editor, ordredato, status FROM materiale_ordrer WHERE id =" & id
	    oRec.open strSQL, oConn, 3
	    if not oRec.EOF then
	    ordrenr = oRec("ordrenr")
	    ordreDato = oRec("ordredato")
	    ordrestatus = oRec("status") 
	    end if
	    oRec.close
	    
	    
	    end if
	    
	    
		%>
		
		
		
		
	
	<form action="materialer_ordrer.asp?func=<%=funcval %>&id=<%=id %>" method=post>
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr><td>
	Ordrenr:&nbsp;<input name="FM_odrid" type="text" value="<%=ordrenr%>" size=10 />
	</td><td>
	 &nbsp;&nbsp;&nbsp;Ordredato:
	
	   <% if  func <> "red" then
	  ordreDato = now
	  else 
	   ordreDato = ordreDato
        end if %>
	
	<input id="FM_odrdato_dag" name="FM_odrdato_dag" type="text" size=2 value="<%=datepart("d",ordreDato)%>" /> -
	    <input id="FM_odrdato_md" name="FM_odrdato_md" type="text" size=2 value="<%=datepart("m",ordreDato)%>" /> -
	    <input id="FM_odrdato_aar" name="FM_odrdato_aar" type="text" size=4 value="<%=datepart("yyyy",ordreDato)%>" /> dd-mm-åååå</td>
	
	</td><td>
	&nbsp;&nbsp;&nbsp;Ordrestatus: 
	
	<select name="FM_odrstatus">
	    <%if func = "red" then %>
	    <option value="<%=ordrestatus %>" selected><%=ordrestatus %></option>
	    <%end if %>
           <option value="6">Kladde</option>
             <option value="5">Venter</option>
              <option value="4">Udført</option>
              <option value="3">Godkendt</option>
              <option value="2">Afsluttet</option>
               <option value="1">Igang</option>
        </select>
        </td><td>
        &nbsp;&nbsp;&nbsp;
        <a href="javascript:popUp('materialer_find.asp','600','400','100','50');" target="_self" class=vmenu>
            <img src="../ill/brugmatfinder.gif" border=0 /></a>
            </td></tr>
	</table>
            
			<br /><br />
	<table cellspacing="0" cellpadding="2" border="0" width="95%">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan=2 height=10 style="border-top:1px #003399 solid; border-bottom:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=6 class='alt' valign=middle style="border-top:1px #003399 solid;">&nbsp;</td>
		<td width="8" rowspan=2 style="border-top:1px #003399 solid; border-bottom:1px #003399 solid; border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt style="padding-right:5px; border-bottom:1px #003399 solid;"><b>Gruppe / Lager</b></td>
		<td class=alt style="padding-right:5px; border-bottom:1px #003399 solid;"><b>Navn</b></td>
		<td height=20 class=alt style="padding-right:5px; border-bottom:1px #003399 solid;"><b>Varenr.</b></td>
		<td class=alt style="padding-right:5px; border-bottom:1px #003399 solid;"><b>Betegnelse</b></td>
	    <td class=alt style="border-bottom:1px #003399 solid;"><b>Antal</b></td>
	    <td class=alt style="border-bottom:1px #003399 solid;"><b>Modtaget</b></td>
	</tr>
	<%


	'*** SQL datoer ***
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	
	if func <> "red" then
	strSQL = "SELECT m.id, m.navn, m.varenr, m.matgrp, mg.navn AS gnavn, mg.nummer, m.betegnelse FROM materialer m "_
	&" LEFT JOIN materiale_grp mg ON (mg.id = m.matgrp) WHERE "& matidSQLkri
	else
	strSQL = "SELECT id, ordreid, matid, navn, betegnelse, varenr, gruppe AS gnavn, status, antal FROM materiale_ordrer_linier"_
	&" WHERE ordreid = "&id
	end if
	
	'Response.Write strSQL
	'Response.flush
	
	x = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	'if cdbl(id) = oRec("id") then
	'bgthis = "#ffff99"
	'else
	    select case right(x,1)
	    case 0,2,4,6,8
	    bgthis = "#ffffff"
	    case else
	    bgthis = "#eff3ff"
	    end select
	'end if
	
	if func <> "red" then
	
	thisantalUse = 0
	thisantal = split(request("FM_antal_b_"&trim(oRec("id"))&""),",")
	for t = 0 to UBOUND(thisantal)
	thisantalUse = thisantalUse + thisantal(t)
	next
	
	else
	
	thisantalUse = oRec("antal")
	
	end if
	
	
	%>
	<tr bgcolor="<%= bgthis %>">
	    <input name="FM_id" type="hidden" value="<%=oRec("id")%>" />
		<td style="border-left:1px #003399 solid; border-top:1px silver solid;">&nbsp;</td>
		<td style="padding-right:5px; border-top:1px silver solid;"><input name="FM_gruppe_<%=oRec("id")%>" type="text" style="width:150px; font-size:9px;" value="<%=oRec("gnavn")%>" /></td>
		<td style="padding-right:5px; border-top:1px silver solid;"><input name="FM_navn_<%=oRec("id")%>" type="text" style="width:200px; font-size:9px;" value="<%=oRec("navn")%> " /></td>
		<td style="padding-right:5px; border-top:1px silver solid;"><input name="FM_varenr_<%=oRec("id")%>" type="text" style="width:100px; font-size:9px;" value="<%=oRec("varenr")%>" /></td>
		<td style="padding-right:5px; border-top:1px silver solid;"><input name="FM_betegn_<%=oRec("id")%>" type="text" style="width:200px; font-size:9px;" value="<%=oRec("betegnelse")%>" /></td>
		<td style="border-top:1px silver solid;"><input name="FM_antal_<%=oRec("id")%>" type="text" style="width:50px; font-size:9px;" value="<%=thisantalUse %>" /></td>
		<td style="border-top:1px silver solid;" class=lille>
		    
		    <%
		    
		    if func = "red" then
		        if oRec("status") = 1 then
		        %>
		        Modtaget.<br />
		        lager opdateret.
		        <%
		        else
		        %>
		        <input name="FM_status_<%=oRec("id")%>" type="checkbox" value="1" />
		        <%
		        end if
		     else
		      %>
		        <input name="FM_status_<%=oRec("id")%>" type="checkbox" value="1" />
		        <%
		     end if %>
		
            </td>
		<td style="border-right:1px #003399 solid; border-top:1px silver solid;">&nbsp;</td>
	</tr>
	<%
	x = x + 1
	oRec.movenext
	wend
	
	%>
	
	
	
	
	
	
	
	<%
	for n = 0 to 10
	    
	    select case right(n,1)
	    case 0,2,4,6,8
	    bgthis = "#ffffff"
	    case else
	    bgthis = "#eff3ff"
	    end select
	    
	    call nyordrelinie(bgthis, -n)
	
	next
	%>
	
	
	<%function nyordrelinie(bgcol, val) 
	    
	    if val = 0 then
	    dsp = ""
	    wzb = "visible"
	    else
	    dsp = "none"
	    wzb = "hidden"
	    end if
	    
	
	%>
	<!--<tr><td colspan=8>
	
	<div id="d_<%=val%>" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>;">
	<table cellspacing="0" cellpadding="2" border="0" width="100%">-->
	    <tr id="td_<%=val%>_0" bgcolor="<%=bgthis%>" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>;">
	        <input name="FM_id" type="hidden" value="<%=val%>" />
		    <td id="td_<%=val%>_1" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>; border-left:1px #003399 solid; border-top:1px silver solid;">&nbsp;</td>
		    <td id="td_<%=val%>_2" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>; padding-right:5px; border-top:1px silver solid;"><input name="FM_gruppe_<%=val%>" type="text" style="width:150px; font-size:9px;" value="" /></td>
		    <td id="td_<%=val%>_3" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>; padding-right:5px; border-top:1px silver solid;"><input name="FM_navn_<%=val%>" type="text" style="width:200px; font-size:9px;" value="" /></td>
		    <td id="td_<%=val%>_4" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>; padding-right:5px; border-top:1px silver solid;"><input name="FM_varenr_<%=val%>" type="text" style="width:100px; font-size:9px;" value="" /></td>
		    <td id="td_<%=val%>_5" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>; padding-right:5px; border-top:1px silver solid;"><input name="FM_betegn_<%=val%>" type="text" style="width:200px; font-size:9px;" value="" /></td>
		    <td id="td_<%=val%>_6" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>; border-top:1px silver solid;"><input name="FM_antal_<%=val%>" type="text" style="width:50px; font-size:9px;" value="0" /></td>
		    <td id="td_<%=val%>_7" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>; border-top:1px silver solid;">
                <input name="FM_status_<%=val%>" type="checkbox" value="1" /></td>
		    <td id="td_<%=val%>_8" style="position:relative; display:<%=dsp %>; visibility:<%=wzb %>; border-right:1px #003399 solid; border-top:1px silver solid;">&nbsp;</td>
	    </tr>
	<!--</table>
	</div>
	</td></tr>-->
	<%end function %>
	
       
        <tr bgcolor="#ffffff">
		<td width="8" valign=top height=20 style="border-left:1px #5582d2 dashed;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right colspan=6 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-right:1px #5582d2 dashed;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
       
       
        
     <tr bgcolor="#ffffff">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right colspan=6 valign="top" style="border-bottom:1px #003399 solid;">
		<a href="#" onclick="showordrelinie()" class=alt>
            <img src="../ill/tilfojordrelinie.gif" border=0 /></a></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
        
        
	</table>
	
	
	<br><br>&nbsp;
	<input id="lastn" type="hidden" value="0" />
        <input id="Submit1" type="submit" value="Indlæs ordre -->" />
	</form>
	
	<div id=sidenoter style="position:relative; top:10px; left:0px; width:400px; padding:20px; background-color:#ffffe1; border:2px red dashed;">
	 <img src="../ill/ac0005-24.gif" /> <b>Side noter:</b><br />
	 Brug <b>materiale finder</b> til at søge i materiale lager, og tilføje ordre linier.<br /><br />
	 
	 Når materiale registreres som modtaget, bliver antal på lager opdateret med det bestilte antal. 
	 Dette kræver dog at både varenr og materiale gruppe kan findes.
	</div>
       
        
	
	</div>
	<%case else %>
	
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script language=javascript>

	    function BreakItUp2() {
	        //Set the limit for field size.
	        var FormLimit = 102399

	        //Get the value of the large input object.
	        var TempVar = new String
	        TempVar = document.theForm2.BigTextArea.value

	        //If the length of the object is greater than the limit, break it
	        //into multiple objects.
	        if (TempVar.length > FormLimit) {
	            document.theForm2.BigTextArea.value = TempVar.substr(0, FormLimit)
	            TempVar = TempVar.substr(FormLimit)

	            while (TempVar.length > 0) {
	                var objTEXTAREA = document.createElement("TEXTAREA")
	                objTEXTAREA.name = "BigTextArea"
	                objTEXTAREA.value = TempVar.substr(0, FormLimit)
	                document.theForm2.appendChild(objTEXTAREA)

	                TempVar = TempVar.substr(FormLimit)
	            }
	        }
	    }

	</script>
	<%if showmenu <> 0 then %>
	
	
	<%
        call menu_2014()
	
	dtop = 102
	dleft = 90
	
	else 
	
	dtop = 20
	dleft = 20
	
	end if %>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=dleft %>px; top:<%=dtop %>px; visibility:visible;">
	
	

    <%call filterheader_2013(0,0,600,pTxt)%>

        <h4>Materialeordrer</h4>

	<table cellspacing=0 cellpadding=2 border=0 bgcolor="#FFFFFF" width=100%>
	
	<form action="materialer_ordrer.asp" method="post">
	<tr>
	    <td><b>Medarbejder:</b>
	    <br /><select name="FM_medarb" style="font-size : 11px; width:205px;">
	    <option value="0">Alle</option>
	   <%
	    strSQL = "SELECT Mid, Mnavn, Mnr, mansat, init FROM medarbejdere WHERE mansat <> 2 ORDER BY Mnavn"
	    oRec.open strSQL, oConn, 3
	    while not oRec.EOF 
	
		
		if cint(selmedarb) = oRec("mid") then
		thisChecked = "SELECTED"
		else
		thisChecked = ""
		end if
		%>
		<option value="<%=oRec("mid")%>" <%=thisChecked%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>) - <%=oRec("init")%></option>
	<%
	oRec.movenext
	wend
	oRec.close
	%></select>
	    </td>
	    <td align=right><b>Status:</b><br />
	    <select name="FM_odrstatus">
	    
	     <option value="Igang">Igang</option>
	    
	    <%
	    sqlStatusKri = " AND mo.status = 1"
	    
	    if len(trim(request("FM_odrstatus"))) <> 0 then
	    
	        if request("FM_odrstatus") <> "Alle" then
	        sqlStatusKri = " AND mo.status = '"& request("FM_odrstatus") &"'"
	        else
	        sqlStatusKri = ""
	        end if
	        
	        Response.Cookies("mat")("status") = request("FM_odrstatus")
	    
	    %>
	    <option value="<%=request("FM_odrstatus") %>" SELECTED><%=request("FM_odrstatus") %></option>
	    <%
	    else
	        
	       
	        
	        if Request.Cookies("mat")("status") <> "" then
	        
	      
	        
	        if Request.Cookies("mat")("status") <> "Alle" then
	        sqlStatusKri = " AND mo.status = '"& Request.Cookies("mat")("status") &"'"
	        else
	        sqlStatusKri = ""
	        end if
	        
	        %>
	        <option value="<%=Request.Cookies("mat")("status") %>" SELECTED><%=Request.Cookies("mat")("status") %></option>
	        <%
	        
	        else
	        
	        sqlStatusKri = sqlStatusKri 
	        
	        end if
	    
	    
	    end if
	    
	    
	    Response.Cookies("mat").expires = date + 10
	    %>
	    
	            <option value="Kladde">Kladde</option>
             <option value="Venter">Venter</option>
              <option value="Udført">Udført</option>
              <option value="Godkendt">Godkendt</option>
              <option value="Afsluttet">Afsluttet</option>
              <option value="Alle">Alle</option>
        </select> 
        
       </td>
	</tr>
	
	
	       <tr>
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td colspan=2><br /><br /><b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			</td>
		</tr>
		<tr>
		   <td><br />
		   <b>Søg på ordrenr.:</b>&nbsp;<input name="FM_sog" id="FM_sog" value="<%=sogval%>" type="text" style="width:250px; border:2px yellowgreen solid;" /></td>
		<td align=right>
            <input id="Submit1" type="submit" value="Vis ordrer >>" /></td></tr>
	
		</form>
		</table><br>
		
		<!-- filter header sLut -->
	</td></tr></table>
	</div>

	<br /><br />
	
	<a href="materialer_ordrer.asp?func=opr" target="_blank"><img src="../ill/opretnymatord.gif" border=0 /></a>
	<br /><br />
	<table cellspacing=1 cellpadding=1 border=0 width=700 bgcolor="silver">
	<tr bgcolor="#5582d5">
	    <td height=20 class=alt align=center><b>Ordredato</b></td>
	    <td style="padding-left:5px;" class=alt><b>Ordrenr.</b></td>
	    <td style="padding-left:5px;" class=alt><b>Antal / Ordrelinier</b></td>
	    <td style="padding-left:5px;" class=alt><b>Oprettet af / oprettet dato</b></td>
	    <td style="padding-left:5px;" class=alt><b>Status</b></td>
	    <td class=alt align=center><b>Slet</b></td>
	</tr>
	<%
	
	strExport = "Ordredato;Ordrenr;Status;Oprettet Af;Gruppe;Varenavn;Varenr;Betegnelse;Antal;Modtaget"
	
	
    '*** SQL datoer ***
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	if selmedarb <> 0 then
	medidKri = " AND mo.mid = " & selmedarb
	else
	medidKri = ""
	end if
	
	if cint(usesogval) = 1 then
	whclaus = " mo.ordrenr = '" & sogval &"'"
	else
	whclaus = " ordredato "_
	&" BETWEEN '"& strStartDato &"' AND "_
	&" '"& strSlutDato &"' "& medidKri &""& sqlStatusKri &""  
	end if
	
	strSQL = "SELECT mo.ordrenr, mo.ordredato, mo.dato, mo.id, mo.mid, "_
	&" mo.status AS ordrestatus, m.mnavn, m.mnr, m.init, SUM(mol.antal) AS antalmat "_
	&" FROM materiale_ordrer mo "_
	&" LEFT JOIN medarbejdere m ON (m.mid = mo.mid) "_
	&" LEFT JOIN materiale_ordrer_linier mol ON (mol.ordreid = mo.id) "_
	&" WHERE "& whclaus &" GROUP BY mo.id ORDER BY ordredato DESC, ordrenr DESC"
	
	'Response.Write strSQL
	'Response.flush
	
	
	oRec.open strSQL, oConn, 3
	x = 0
	while not oRec.EOF 
	
	select case right(x, 1)
	case 2,4,6,8,0
	bgcol = "#ffffff" 
	case else
	bgcol = "#eff3ff"
	end select
	
	%>
	<tr bgcolor=<%=bgcol%>>
	    <td align=center><%=formatdatetime(oRec("ordredato"), 1)%></td>
	    <td style="padding-left:5px;"><a href="materialer_ordrer.asp?func=red&id=<%=oRec("id") %>" target="_blank" class=vmenu><%=oRec("ordrenr")%></a></td>
	    <td style="padding-left:5px;"><%=oRec("antalmat") %></td>
	    <td style="padding-left:5px;"><i><%=oRec("mnavn") %>
            &nbsp;<%=oRec("mnr") %>
            <%if len(trim(oRec("init"))) <> 0 then %>
             - <%=oRec("init") %>
             <%end if %></i>
            &nbsp;&nbsp;<font class=megetlillesilver> <%=formatdatetime(oRec("dato"),2) %></font></td>
             <td style="padding-left:5px;"><%=oRec("ordrestatus") %></td>
        <td align=center><a href="materialer_ordrer.asp?func=slet&id=<%=oRec("id") %>" class=red>
            <img src="../ill/slet.gif" border=0 /></a></td>
	</tr>
	
	<%
	
	strSQLmol = "SELECT gruppe, navn, betegnelse, varenr, antal, status FROM materiale_ordrer_linier WHERE ordreid = " & oRec("id")
	oRec3.open strSQLmol, oConn, 3
	while not oRec3.EOF
	
	strExport = strExport & "xx99123sy#z"& oRec("ordredato") &";"& oRec("ordrenr") &";"& oRec("ordrestatus") &";"& oRec("mnavn") &"("& oRec("mnr") &")"&";"& oRec3("gruppe") &";"& oRec3("navn") &";"& oRec3("varenr") &";"& oRec3("betegnelse") &";"& oRec3("antal") &";"& oRec3("status") 
	
    oRec3.movenext
	wend
	oRec3.close 
	
	
	
	x = x + 1
	oRec.movenext
	wend
	oRec.close
	
	%>
	</table>
	
	<br /><br />
		<form action="materiale_ordrer_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=strExport%>">
			<input type="hidden" name="txt20" id="txt20" value="">
			
			<input type="submit" value="Eksporter">
			
			</form>
	
	</div>
    <%end select %>	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
