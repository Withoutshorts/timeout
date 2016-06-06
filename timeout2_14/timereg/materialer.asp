<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/mat_func.asp"-->

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
		if len(request("id")) <> 0 then
		id = request("id")
		else
		id = 0
		end if
	end if
	
	menu = request("menu")
	
	session.lcid = 1030 'DK
	
	select case func
	case "flytlager"
	
	nytLager = request("FM_nymgrp")
	
	if len(request("FM_antal")) <> 0 then
	antal = replace(request("FM_antal"), ",", ".")
	else
	antal = 0
	end if
	
	
	                call erDetInt(antal)
					if isInt > 0 then
			        %>
					<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					    
					<%
					errortype = 76
					call showError(errortype)
					isInt = 0
					response.end
					end if
	
	
	call flytlager(id, antal, nytLager, 0)
	
	response.redirect "materialer.asp?menu="&menu&"&id="&id&"&FM_sog="&request("FM_sog")
	
	
	
	
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:210px; top:150px; visibility:visible; border:10px #999999 solid; padding:20px;"">
	<h3>Materialer - slet?</h3>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> et materiale. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="materialer.asp?menu=mat&func=sletok&id=<%=id%>&menu=<%=menu%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM materialer WHERE id = "& id &"")
	Response.redirect "materialer.asp?menu=mat&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
		
		else
		
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(request("FM_nr"))
			if isInt > 0 then
			sortorder = 1000
			else
			sortorder = request("FM_nr")
			end if
			
			isInt = 0
			
	        if len(trim(request("FM_nr"))) = 0 OR request("FM_nr") = "0"  then
		    %>
		    <!--#include file="../inc/regular/header_inc.asp"-->
		    <%
		    errortype = 73
		    call showError(errortype)
		    
		
		else
			
			
			call erDetInt(request("FM_antal"))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			errortype = 54
			call showError(errortype)
			
			isInt = 0
		
		else
			
		
			call erDetInt(request("FM_minlager"))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			errortype = 75
			call showError(errortype)
			
			isInt = 0
		
		else
			
			
			call erDetInt(request("FM_indkobspris"))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			errortype = 55
			call showError(errortype)
			
			isInt = 0
		
		else
			
			call erDetInt(request("FM_salgspris"))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			errortype = 56
			call showError(errortype)
			
			isInt = 0
		
		else
			
			call erDetInt(request("FM_ptid"))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			errortype = 57
			call showError(errortype)
			
			isInt = 0
		
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
		
        strVarenr = SQLBless(request("FM_nr"))
		
		if len(request("FM_antal")) <> 0 then
		intAntal = replace(request("FM_antal"), ",", ".")
		else
		intAntal = 0
		end if
		
		if len(request("FM_indkobspris")) <> 0 then
		dblIndkobspris = replace(request("FM_indkobspris"), ",", ".")
		else
		dblIndkobspris = 0
		end if
		
		if len(request("FM_salgspris")) <> 0 then
		dblSalgspris = replace(request("FM_salgspris"), ",", ".")
		else
		dblSalgspris = 0
		end if
		
		dtStart_dag = request("FM_start_dag")
		dtStart_mrd = request("FM_start_mrd")
		dtStart_aar = request("FM_start_aar")
		
		sqlDato = dtStart_aar & "/" & dtStart_mrd & "/" & dtStart_dag  
		
		intMatGrp = request("FM_mgrp")
		intPtid = request("FM_ptid")
		
        if len(trim(request("FM_pic_navn"))) <> 0 then
		intPic = request("FM_pic")
        else
        intPic = 1 'blank gif
        end if


                'response.write "len(trim(request(FM_pic_navn))): " & len(trim(request("FM_pic_navn")))
                'response.end
		
		if len(request("FM_enhed")) <> 0 then
		strEnhed = request("FM_enhed")
		else
		strEnhed = "Stk."
		end if
		
		if len(request("FM_ptid_arr")) <> 0 then
		ptid_arr = request("FM_ptid_arr")
		else
		ptid_arr = 0
		end if
		
		intLev_a = request("FM_lev_a")
		intLev_b = request("FM_lev_b")
		intSer_a = request("FM_ser_a")
		
		if len(trim(request("FM_minlager"))) <> 0 then
		intMinLager = replace(request("FM_minlager"), ",", ".")
		else
		intMinLager = 0
		end if
		
		strLokal = SQLBless(request("FM_lokal"))
		
		strBetegn = SQLBless(request("FM_betegn"))
		
		varenrFindes = 0
		strSQL = "SELECT id FROM materialer WHERE matgrp = "& intMatGrp &" AND varenr = '"& strVarenr &"' AND id <> "& id
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		varenrFindes = 1
		oRec.movenext
		wend
		
		oRec.close
		
		
		if varenrFindes = 1 then
		
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			errortype = 74
			call showError(errortype)
		
		
		else
		
		if func = "dbopr" then
		
		strSQL = "INSERT INTO materialer "_
		&" (navn, editor, dato, varenr, antal, arrestordredato, salgspris, "_
		&" indkobspris, matgrp, ptid, ptid_arr, enhed, pic, leva, levb, sera, minlager, "_
		&" sortorder, betegnelse, lokation) VALUES "_
		&" ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "_
		&" '"& strVarenr &"', "& intAntal &", '"& sqlDato &"', "& dblSalgspris &", "_
		&" "& dblIndkobspris &", "& intMatGrp &", "_
		&" "& intPtid &", "& ptid_arr &", '"& strEnhed &"', "& intPic &", "_
		&" "& intLev_a &", "& intLev_b &", "& intSer_a &", "& intMinLager &", "_
		&" "& sortorder &", '"& strBetegn &"', '"& strLokal &"')"
		
		'Response.write strSQL
		'Response.flush
		
		oConn.execute(strSQL)
		
		
		strSQL2 = "SELECT id FROM materialer WHERE id <> 0 ORDER BY id DESC"
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			lastid = oRec2("id")
		end if
		oRec2.close 
		
		
		else
		
		strSQL = "UPDATE materialer SET navn ='"& strNavn &"', "_
		&" varenr = '"& strVarenr &"', editor = '" &strEditor &"', dato = '" & strDato &"', "_
		&" antal = "& intAntal &", arrestordredato =  '"& sqlDato &"', "_
		&" salgspris = "& dblSalgspris &", indkobspris = "& dblIndkobspris &", "_
		&" matgrp = "& intMatGrp &", ptid = "& intPtid &", ptid_arr = "& ptid_arr &", "_
		&" enhed = '"& strEnhed &"', pic = "& intPic &", "_
		&" leva = "& intLev_a &", levb = "& intLev_b &", sera = "& intSer_a &", "_
		&" minlager = "& intMinLager &", sortorder = "& sortorder &", "_
		&" betegnelse = '"& strBetegn &"', lokation = '"& strLokal &"'"_
		&" WHERE id = "&id
		
		lastid = id
		'Response.write strSQL
		'Response.flush
		
		oConn.execute(strSQL)
		
		strSQL = "UPDATE materiale_forbrug SET matnavn = '"& strNavn &"', matvarenr = '"& strVarenr &"' WHERE matid = " & id
		oConn.execute(strSQL)
		
		end if
		
	
		Response.redirect "materialer.asp?menu="&menu&"&shokselector=1&soglev="&request("soglev")&"&FM_mgrp="&request("FM_mgrp")&"&FM_sog="&request("FM_sog")&"&id="&lastid
		
		
		end if 'varenr findes i forvejen i gruppe.
		end if 'Navn
		end if 'Varenr
		end if 'salgsspris
		end if 'indkbspris
		end if 'minlager
		end if 'antal
		end if 'Produktionstid
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret nyt materiale (vare)"
	dbfunc = "dbopr"
	sidst_redigeret = ""
	
	'strDag = day(now)
	'strMrd = month(now)
	'strAar = year(now)
	
	matGrpSel = request("FM_mgrp")
	

	intPtid = 0
	ptid_arr = 0
	
	strPic = 0
	intlev_a = 0
	intlev_b = 0
	intSer_a = 0
	intMinLager = 0
	strLokal = ""
	
	else
	
	strSQL = "SELECT navn, editor, dato, varenr, antal, "_
	& "ptid_arr, arrestordredato, salgspris, "_
	&" indkobspris, matgrp, ptid, enhed, pic, leva, levb, sera, minlager, betegnelse, lokation FROM materialer WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strVarenr = oRec("varenr")
	dblKobsPris = oRec("indkobspris")
	dblSalgsPris = oRec("salgspris")
	intAntal = oRec("antal")
	StrTdato = oRec("arrestordredato")
	matGrpSel = oRec("matgrp")
	intPtid = oRec("ptid")
	ptid_arr = oRec("ptid_arr")
	strEnhed = oRec("enhed")
	strPic = oRec("pic")
	intlev_a = oRec("leva")
	intlev_b = oRec("levb")
	intSer_a = oRec("sera")
	intMinLager = oRec("minlager")
	strBetegn = oRec("betegnelse")
	strLokal = oRec("lokation")
	
	end if
	oRec.close
	
	sidst_redigeret = "Sidst opdateret den <b>"&formatdatetime(strDato, 1)&"</b> af <b>"&strEditor&"</b>"
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
	
	
	strPicNavn = "../ill/blank.gif"
	
	if ptid_arr = 0 then
	chkArr_ptid1 = ""
	else
	chkArr_ptid1 = "CHECKED"
	end if
	%>
	
	<%if menu <> "0" then
	
	topD = "102"
	lDiv = "90"
	
	 %>
	 <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
    <!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">

	<%'call tsamainmenu(11)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call mattopmenu()
	'end if
	%>
	</div>
        -->
	<%
        
        call menu_2014()
        else 
	
	topD = "20"
	lDiv = "20"
	
	%>
	 <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%end if %>
	
	 <!--#include file="inc/dato2.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=lDiv %>px; top:<%=topD %>px; visibility:visible; padding:20px; background-color:#ffffff;">
	<h4>Materialer <span style="font-size:9px;">- <%=varbroedkrumme%></span></h4>
	<table cellspacing="0" cellpadding="2" border="0" width="600" bgcolor="#ffffff">
	<form action="materialer.asp?menu=<%=menu%>&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="FM_sog" value="<%=request("FM_sog")%>">
	
  	<tr>
		
		<td colspan="2" valign="middle" style="height:30;"><%=sidst_redigeret%>&nbsp;</td>
		
	</tr>
	<tr>
		
		<td width=150 align=right><font color=red><b>*</b></font>&nbsp; Navn:</td>
		<td width=332 style="padding-left:5px;"><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:300px;" ></td>
		
	</tr>
	<tr>
		
		<td align=right><font color=red><b>*</b></font>&nbsp;Mat.nr. / Varenr.:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_nr" value="<%=strVarenr%>" size="20" ></td>
		
	</tr>
	<tr>
		
		<td align=right valign=top>&nbsp;Betegnelse / Note:</td>
		<td style="padding-left:5px;">
		<!--<input type="text" name="FM_betegn" value="" size="20" maxlength="20" >-->
            <textarea name="FM_betegn" id="FM_betegn" cols="35" rows="8"><%=strBetegn%></textarea></td>
		
	</tr>
	<tr>
		
		<td align=right>&nbsp;Lokation:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_lokal" value="<%=strLokal%>" size="20" maxlength="20" ></td>
		
	</tr>
	<tr>
		<%
		strSQL = "select id, filnavn FROM filer WHERE id = " & strPic
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
			strPicNavn = "../inc/upload/"&lto&"/"&oRec("filnavn")
			strPicNavn_only = oRec("filnavn")
		end if
		oRec.close 
		%>
		
		<td width=150 align=right valign=top>Billede:<br>
		Maks str: 150*100 px<br>
		(Nyt billede vises først efter opdatering)</td>
		<td width=332 style="padding-left:5px;">
		<%if strPic <> 1 then %>
		<img src="<%=strPicNavn%>" width="150" height="100" alt="" border="0"> <br>
	    <%end if %>
	    <input type="text" size=40 name="FM_pic_navn" id="FM_pic_navn" value="<%=strPicNavn_only%>"><br>
		<input type="hidden" name="FM_pic" id="FM_pic" value="<%=strPic %>">
		<br>
		<a href="javascript:popUp('upload.asp?FM_filtype=6&func=opret&type=mat&id=<%=id%>&kundeid=0&jobid=0&FM_folderid=-10&FM_adg_alle=1','600','620','250','20');" target="_self" class=vmenu>Upload / Skift billede. &nbsp;<img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0"></a>
		<br><br>&nbsp;</td>
		
	</tr>
	<tr>
		
		<td align=right>Materialegruppe (avance%):</td>
		<td style="padding-left:5px;">
		<!--onChange="visgrp()"-->
		<select name="FM_mgrp" id="FM_mgrp" style="width:200px;">
		<option value="0">Ingen</option>
		<%
		strSQL = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		if cint(matGrpSel) = oRec("id") then
		selThis = "SELECTED"
		else
		selThis = ""
		end if%>
		<option value="<%=oRec("id")%>" <%=selThis%>><%=oRec("navn")%> (<%=oRec("av") %>%)</option>
		<%
		oRec.movenext
		wend
		oRec.close %>
		</select>
		
	</tr>
	<tr>
		
		<td align=right>Antal på lager:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_antal" value="<%=intAntal%>" size="4" > (hel-tal)</td>
		
	</tr>
	<tr>
		
		<td align=right>Minimuslager:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_minlager" value="<%=intMinLager%>" size="4" > (hel-tal, -1 = Lagerstatus udsendes ikke)</td>
		
	</tr>
	<tr>
		
		<td align=right>Enhed:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_enhed" value="<%=strEnhed%>" size=20" > stk (standard), m, kg, etc..</td>
		
	</tr>
	<tr>
		
		<td align=right>Indkøbspris:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_indkobspris" value="<%=dblKobsPris%>" size="10" > kr.</td>
		
	</tr>
	<tr>
		
		<td align=right>Salgspris:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_salgspris" value="<%=dblSalgsPris%>" size="10" > kr.</td>
		
	</tr>
	<tr>
		
		<td align=right><input type="checkbox" name="FM_ptid_arr" value="1" <%=chkArr_ptid1%>> Brug Restordre dato:</td>
		<td style="padding-left:5px;"><select name="FM_start_dag">
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
		
		<select name="FM_start_mrd">
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
		
		
		<select name="FM_start_aar">
		<option value="<%=strAar%>">
		<%if id <> 0 then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		
		<%for x = -2 to 10
		tyear = year(now)+(x)
		%>
		<option value="<%=tyear%>"><%=tyear%></option>
		<%
		next%>
		
		</select>
		&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a></td>
		
	</tr>
	<tr>
		
		<td align=right>Produktionstid:</td>
		<td style="padding-left:5px;">&nbsp;<input type="text" name="FM_ptid" value="<%=intPtid%>" size="4" > (Antal hele dage efter restordredato eller dagsdato)</td>
		
	</tr>
	<tr>
		
		<td align=right><br>Leverandør A:</td>
		<td style="padding-left:5px;"><br>
			<select name="FM_lev_a" id="FM_lev_a" style="width:200px;">
			<option value="0" SELECTED>Ingen</option>
			<%strSQL = "SELECT id, navn, levnr FROM leverand WHERE id <> 0 AND type = 1 OR type = 3 ORDER BY navn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			
			if cint(intLev_a) = oRec("id") then
			selLevA = "SELECTED"
			else
			selLevA = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=selLevA%>><%=oRec("navn")%> (<%=oRec("levnr") %>)</option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			</select>
			
		</td>
		
	</tr>
	<tr>
		
		<td align=right>Leverandør B:</td>
		<td style="padding-left:5px;">
			<select name="FM_lev_b" id="FM_lev_b" style="width:200px;">
			<option value="0" SELECTED>Ingen</option>
			<%strSQL = "SELECT id, navn, levnr FROM leverand WHERE id <> 0 AND type = 1 OR type = 3 ORDER BY navn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			
			if cint(intLev_b) = oRec("id") then
			selLevB = "SELECTED"
			else
			selLevB = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=selLevB%>><%=oRec("navn")%> (<%=oRec("levnr") %>)</option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			</select>
			
		</td>
		
	</tr>
	<tr>
		
		<td align=right>Servicepartner:</td>
		<td style="padding-left:5px;">
			<select name="FM_ser_a" id="FM_ser_a" style="width:200px;">
			<option value="0" SELECTED>Ingen</option>
			<%strSQL = "SELECT id, navn, levnr FROM leverand WHERE id <> 0 AND type = 2 OR type = 3 ORDER BY navn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			if cint(intSer_a) = oRec("id") then
			selSerA = "SELECTED"
			else
			selSerA = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=selSerA%>><%=oRec("navn")%> (<%=oRec("levnr") %>)</option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			</select>
			
		</td>
		
	</tr>
	<tr>
		
		<td colspan=2 height=30>&nbsp;</td>
		
	</tr>
	
	<tr>
		<td colspan="4" align=right><br><br><input type="submit" value="<%=varSubVal %> >>" /></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><< Tilbage</a>
	
	</div>

	
	<%case else
	
	
	
	
	
	if menu <> "0" then
	
	topD = "102"
	lDiv = "90"
	
	 %>
	 <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
		<%'call tsamainmenu(11)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call mattopmenu()
	'end if
	%>
	</div>
    -->

	<%
        
        call menu_2014()
                
    else 
	
	level = session("rettigheder")
	topD = "20"
	lDiv = "20"
	
	%>
	 <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%end if %>
	
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	//function visgrp(){
	//grp = document.getElementById("FM_mgrp").value
	//memu = document.getElementById("menuuse").value
	////limitnr = document.getElementById("FM_limitvarenr").value
	//window.location.href = "materialer.asp?menu="+memu+"&FM_mgrp="+grp
	//}	
	
	function rensfelt(felt){
	document.getElementById(felt).value = ""
	}
	
	

	
	function BreakItUp2()
	{
	  //Set the limit for field size.
	  var FormLimit = 102399
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm2.BigTextArea.value
	
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm2.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm2.appendChild(objTEXTAREA)
	
	      TempVar = TempVar.substr(FormLimit)
	    }
	  }
	}
	
	//-->
	</script>
	<%
	        
	        
	        if len(request("FM_mgrp")) <> 0 then
        	    
        	    matGrpSel = request("FM_mgrp")
		        Response.cookies("mat")("matgrp") = matGrpSel
		       
		    else
        		
		        if len(Request.cookies("mat")("matgrp")) <> 0 then
		        matGrpSel = Request.cookies("mat")("matgrp")
		        else
		        matGrpSel = 0
		        end if
        		
	        end if
	        
	        if cint(matGrpSel) <> 0 then
            matgrpSQLKri = " m.matgrp = "& matGrpSel & " AND " 
            odrbyMG = ""
            else
            matgrpSQLKri = " m.matgrp <> -1 AND " 
            odrbyMG = "m.matgrp, "
            end if
	        
	        
	
	        stLimit = 0
	        stLimitfundet = 0
	        
	        
	           if len(trim(request("FM_mgrp"))) <> 0 then
        				
				       stLimitKri = trim(request("FM_limitvarenr"))
				       response.cookies("mat")("stvarenr") = stLimitKri
        		else
				       if request.cookies("mat")("stvarenr") <> "0" then
				       stLimitKri = request.cookies("mat")("stvarenr")
        			   else
        			   stLimitKri = ""
        			   end if	
		        end if
	        
	                    if stLimitKri <> "" then 
	                    
	                    strSQL = "SELECT m.varenr FROM materialer m "_
				        &" WHERE m.varenr >= '"& stLimitKri &"'"
        				
				        oRec.open strSQL, oConn, 3
				        if not oRec.EOF then
				        stLimit = oRec("varenr") 
				        stLimitfundet = 1
				       
				        end if
				        oRec.close
				        
				       
				       end if
				        
				        if cint(stLimitfundet) = 1 then
				        useStLimitSQLkri = " AND m.varenr >= '"& stLimit & "'"
				        odrbySL = "m.varenr, "
				        else
				        useStLimitSQLkri = ""
				        odrbySL = ""
				        end if
	            
	            
	        
	        
	         
        	
        	
	        call erDetInt(request("FM_limit"))
	        if len(trim(request("FM_limit"))) <> 0 then
        	        
	                if isInt > 0 OR request("FM_limit") > 1000 then
	                useLimit = 1000
	                else
	                useLimit = request("FM_limit")
	                end if
        	        
	                Response.Cookies("mat")("mat_uselimit") = useLimit
	                
	        else
        	
	            if Request.Cookies("mat")("mat_uselimit") <> "" then
	            useLimit = Request.Cookies("mat")("mat_uselimit")
	            else
	            useLimit = 100
	            end if
        	    
	        end if 
	        isInt = 0
	
	
            if len(request("FM_mgrp")) <> 0 then
			sogVal = trim(request("FM_sog"))
			response.cookies("mat")("sogval") = sogVal
			odrbyMN = "m.navn"
			else
			   
			    if request.cookies("mat")("sogval") <> "0" then
			    sogVal = request.cookies("mat")("sogval")
			    odrbyMN = "m.navn"
			    else
			    sogVal = ""
			    odrbyMN = ""
			    end if
			    
			end if
			
			orderBYSQL = odrbySL & odrbyMG & odrbyMN 
			
			'*** fra Print **
			'if instr(sogVal, "99ogprocent99") <> 0 then
			'sogVal = replace(sogVal, "99ogprocent99", "%")
			'else
			'sogVal = sogVal
			'end if
			
			'if instr(sogVal, "%") <> 0 then
			'sogValPrint = replace(sogVal, "%", "99ogprocent99")
			'else
			'sogValPrint = sogVal
			'end if
			
			if len(request("FM_minlager")) <> 0 THEN
			minlager = 1
			minLagerCHK = "CHECKED"
			else
			minlager = 0
			minLagerCHK = ""
			end if
			
			
			 if len(request("soglev")) <> 0 then 
                sogKrilev = request("soglev")
                Response.cookies("mat")("soglev") = sogKrilev
                
                if request("soglev") <> "0" then
                sogLevSQLkri = " (m.leva = " & sogKrilev & " OR m.levb = " & sogKrilev &") AND "
                else
                sogLevSQLkri = " m.leva <> -1 AND "
                end if
                
                else
                    
                    if request.cookies("mat")("soglev") <> "0" AND request.cookies("mat")("soglev") <> "" then
                    sogKrilev = request.cookies("mat")("soglev")
                    sogLevSQLkri = " (m.leva = " & sogKrilev & " OR m.levb = " & sogKrilev &") AND "
                    else
                    sogKrilev = 0
                    sogLevSQLkri = " m.leva <> -1 AND "
                    end if
                    
                    
            end if
            
            
            
            Response.Cookies("mat").expires = date + 45
            
            
            
           
			
	
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=lDiv %>px; top:<%=topD %>px; visibility:visible;">
	
	
	<%'call filterheader(0,0,700,pTxt)

        pTxt = "Materialer/lager"
        call filterheader_2013(0,0,670,pTxt)%>
	<table cellspacing="2" cellpadding="2" border="0" width="100%">
	<form action="materialer.asp?menu=mat" method="post">
	<tr>
		<td colspan=2><b>Materialegruppe:</b>&nbsp;&nbsp;<br>
		
		<select name="FM_mgrp" id="FM_mgrp" style="width:350px;">
		<option value="0">Alle</option>
		<%
		strSQL = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		if cint(matGrpSel) = oRec("id") then
		selThis = "SELECTED"
		else
		selThis = ""
		end if%>
		<option value="<%=oRec("id")%>" <%=selThis%>><%=oRec("navn")%>
		<%if level <= 2 OR level = 6 then %>
		&nbsp;(<%=oRec("av") %>%)
		<%end if %>
		</option>
		<%
		oRec.movenext
		wend
		oRec.close %>
		</select></td>
		</tr>
		
		 <tr>
	    <td><b>Leverandør:</b> <br />
            
            
            
            <select id="soglev" name="soglev" style="width:350px;">
                <option value="0" SELECTED>Alle</option>
            <% 
	     
                   strSQL3 = "SELECT id, navn FROM leverand WHERE id <> 0 ORDER BY navn"
           
                   oRec3.open strSQL3, oConn, 3 
                   while not oRec3.EOF 
                   
                        if cint(sogKrilev) = oRec3("id") then
                        slevSEL = "SELECTED"
                        else
                        slevSEL = ""
                        end if
                   
                   %>
                   <option value="<%=oRec3("id") %>" <%=slevSEL %>><%=oRec3("navn") %></option>
                   <%
                   oRec3.movenext
                   wend
                   oRec3.close
            %> 
            </select>       
            </td>
	 </tr>
		
		<tr><td style="padding-top:6px;">
		Vis fra Mat./vare nr.: <input type="text" name="FM_limitvarenr" id="FM_limitvarenr" size="6" value="<%=stLimitKri%>" style="font-size:9px;"> og de næste 
            <input id="FM_limit" name="FM_limit" type="text" size="3" value="<%=useLimit%>" style="font-size:9px;"/> (maks. 1000) Mat./vare nr.
            <br />
           <input id="FM_minlager" name="FM_minlager" type="checkbox" <%=minLagerCHK %> style="font-size:9px;"/>Vis kun materialer hvor antal på lager er mindre end minimuslager.
            <td valign=bottom>
                &nbsp;</td>
            
		</tr>
		<input id="menuuse" value="<%=menu %>" type="hidden" />
		
		<tr>
		<td style="padding-top:10px;">
		<b>Materiale:</b><br>
		Søg på navn eller varenr. <br />
		Søg på flere materialer ved at opdele søgningen med ";". Brug % som wildcard.<br>
		<input type="text" name="FM_sog" id="FM_sog" value="<%=sogVal%>" style="width:470px;"> </td>
		<td valign=bottom><input
               id="Submit1" type="submit" value="Søg materialer >>" /></td>
	</tr></form>
	</table>

	
	<!-- filter header sLut -->
	</td></tr></table>
	</div>
	
	
	
	
	
		
		
		
		<% 
		if level <= 2 OR level = 6 then 
		cspan = 13
		else
		cspan = 11
		end if
		
		
	tTop = 50
	tLeft = 0
	tWdth = 1184
	
	
	call tableDiv(tTop,tLeft,tWdth)
		
		%>
		
		
		
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan=2 height=10>&nbsp;</td>
		<td colspan=<%=cspan %> class='alt' valign=middle>&nbsp;</td>
		<td width="8" rowspan=2>&nbsp;</td>
	</tr>
       
	<tr bgcolor="#5582D2">
		<td class=alt style="padding-right:5px; width:100px;"><b>Gruppe</b>
		<%if level <= 2 OR level = 6 then %>
		<br />(Ava. %)
		<%end if %></td>
		<td class=alt style="padding-right:5px; width:50px;"><b>Lok.</b></td>
        <td class="alt"><b>Billede</b></td>
	    <td class=alt style="padding-right:5px; width:150px;"><b>Navn</b><br />
	    Betegnelse / Note</td>
		<td height=20 class=alt style="padding-right:5px; width:50px;"><b>Varenr.</b></td>
		
		<td class=alt align=right style="padding-right:5px; width:60px;"><b>Opd. dato</b></td>
		<td class=alt align=right style="padding-right:5px;"><b>På lager</b> (ordre)</td>
		<td class=alt align=right style="padding-right:5px;"><b>Min. lager</b></td>
		<%if level <= 2 OR level = 6 then %>
		<td class=alt align=right style="padding-left:5px; padding-right:5px;">Indkøbs-pris</td>
		<td class=alt align=right style="padding-left:5px; padding-right:5px;">Salgs-pris</td>
		<%end if%>
		<td class=alt style="padding-left:10px; padding-right:5px;"><b>Pri. lev.</b></td>
		
		<td class=alt style="padding-right:5px;"><b>Slet?</b></td>
		<td class=alt style="padding-right:5px;"><b>Flyt antal til.. (ny gruppe)</b></td>
	</tr>
	<%
	
	strExport = "Gruppe;Lokation;Varenr;Navn;Betegnelse;Antal på lager;Enhed;Min. Lager"
	
	if level = 1 then
	strExport = strExport &";Inkøbspris;Salgspris"
	end if
	
	
	sogKri = " m.navn <> 'xffgd123!fsd999'"
	if len(sogVal) <> 0 then
		
		sogKri = ""
		sogeKriUse = split(sogVal, ";")
		for b = 0 to UBOUND(sogeKriUse)
		sogKri = sogKri &" m.navn LIKE '"& sogeKriUse(b) &"%' OR m.varenr = '"& sogeKriUse(b) &"' OR "
		next
		
		lensogKri = len(sogKri)
		leftsogKri = left(sogKri, lensogKri - 3)
		sogKri = leftsogKri
		
	end if
	
	
	if cint(minlager) = 1 then
	lagerSQLkri = " (m.antal < minlager) AND "
	else
	lagerSQLkri = " (m.antal <> 999999999) AND "
	end if
	
	'**** Main SQL ***'
	strSQL = "SELECT m.id, m.navn, m.varenr, m.antal, m.dato, m.betegnelse, "_
	&" l.navn AS levnavn, l.id AS lid, mg.navn AS gnavn, m.enhed, leva, minlager, "_
	&" mg.id AS grpid, m.indkobspris, m.salgspris, SUM(mol.antal) AS restordre, m.lokation, mg.av, pic FROM materialer m "_
	&" LEFT JOIN materiale_grp mg ON (mg.id = matgrp) "_
	&" LEFT JOIN leverand l ON (l.id = leva) "_
	&" LEFT JOIN materiale_ordrer_linier mol ON (mol.varenr = m.varenr AND mol.gruppe = mg.navn AND mol.status = 0) "_
	&" WHERE "& lagerSQLkri &""& matgrpSQLKri &""& sogLevSQLkri &" ("& sogKri &") "& useStLimitSQLkri &""_
	&" GROUP BY m.id ORDER BY "& orderBYSQL &" LIMIT "& useLimit
	
	'Response.Write strSQL
	'Response.flush
	
	x = 0
	
	strMatids = "0"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	strMatids = strMatids &","& oRec("id")
	
	if cdbl(id) = oRec("id") then
	bgthis = "#ffff99"
	else
	    select case right(x, 1)
	    case 0,2,4,6,8
	    bgthis = "#eff3ff"
	    case else
	    bgthis = "#ffffff"
	    end select
	end if
	%>
	<tr>
		<td bgcolor="#CCCCCC" colspan="<%=cspan+2 %>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgthis%>">
		<td>&nbsp;</td>
		<td style="padding-right:5px;"><b><%=oRec("gnavn")%></b>
            <%if (level <= 2 OR level = 6) AND len(trim(oRec("av")) <> 0) then %> 
            &nbsp;(<%=oRec("av") %>%)
            <%end if %></td>
		<td style="padding-right:5px;"><%=oRec("lokation") %></td>

        <td style="padding:10px;">
            <%

            strPicNavn = ""
                strPicNavn_only = ""
		strSQL = "select id, filnavn FROM filer WHERE id = " & oRec("pic")
		oRec5.open strSQL, oConn, 3 
		if not oRec5.EOF then
			strPicNavn = "../inc/upload/"&lto&"/"&oRec5("filnavn")
			strPicNavn_only = oRec5("filnavn")
		end if
		oRec5.close 


		%>

            <%if cint(oRec("pic")) <> 1 then %>
            <!--<%=strPicNavn_only %>-->
		<img src="<%=strPicNavn%>" width="75" height="50" alt="" border="0"> 
            <%else %>
            &nbsp;
	    <%end if %>

        </td>
		
		<td style="padding-right:10px;">
		<%if level <= 2 OR level = 6 then %>
		<a href="materialer.asp?menu=<%=menu%>&func=red&id=<%=oRec("id")%>&soglev=<%=sogKrilev%>&FM_mgrp=<%=matGrpSel%>&FM_sog=<%=sogVal%>" class=vmenu><%=oRec("navn")%> </a>
		<%else %>
		<b><%=oRec("navn")%> </b>
		<%end if %>
		
		<%if len(trim(oRec("betegnelse"))) <> 0 then %>
		<br /><i>
		    <%if len(oRec("betegnelse")) > 60 then %>
		    <%=left(oRec("betegnelse"), 60) & ".."%>
		    <%else %>
		    <%=oRec("betegnelse") %>
		    <%end if %></i>
		<%end if %>
		</td>
		
		<td style="padding-right:5px;"><b><%=oRec("varenr")%></b></td>
		
		
		<td align=right style="padding-right:5px; white-space:nowrap;">
		<%if len(trim(oRec("dato"))) <> 0 then %>
 		<%=formatdatetime(oRec("dato"), 2)%>
 		<% end if%></td>
		
		<td align=right style="padding-right:5px;"><b><%=oRec("antal")%>&nbsp;<%=oRec("enhed")%></b>
		<%if oRec("antal") < oRec("minlager") then
		%>&nbsp;<font color="darkred"><b>!</b></font><%
		end if%>
		
		<%if oRec("restordre") > 0 then %>
		(<%=oRec("restordre")%>)
		
		<%end if %>
		</td>
		
		<td align=right style="padding-right:5px;"><%=oRec("minlager")%>&nbsp;<%=oRec("enhed")%></td>
		
		
		<%if level <= 2 OR level = 6 then %>
                
                <%if lto = "dencker" OR lto = "xintranet - local" then
                dicim = 3
                else
                dicim = 2
                end if %>

		<td align=right style="padding-right:5px;"><%=formatnumber(oRec("indkobspris"),dicim) %></td>
		<td align=right style="padding-right:5px;"><%=formatnumber(oRec("salgspris"),dicim) %></td>
		<%end if%>
		
		
		<%if len(oRec("levnavn")) > 30 then 
		levnavn = left(oRec("levnavn"), 30) & ".."
		else
		levnavn = oRec("levnavn") 
		end if%>
		
		<td style="padding:0px 5px 0px 10px;">
		<a href="leverand.asp?menu=<%=menu%>&func=red&id=<%=oRec("lid")%>&rdir=mat" class=vmenu>
		<%=left(levnavn, 10)%></a>
		</td>
		
		
		<td style="padding-right:5px;">
		<%if level <= 2 OR level = 6 then %>
		<a href="materialer.asp?menu=<%=menu%>&func=slet&id=<%=oRec("id")%>&soglev=<%=sogKrilev%>&FM_mgrp=<%=matGrpSel%>&FM_sog=<%=sogVal%>"><img src="../ill/slet_16.gif" alt="Slet materiale" border="0"></a>
		<%end if%>
		&nbsp;
		</td>
		
		<form method=post action="materialer.asp?menu=<%=menu%>&func=flytlager&id=<%=oRec("id")%>">
		<input type="hidden" name="FM_sog" id="<%=oRec("id")%>" value="<%=sogValPrint%>"> 
		<td width=200 style="padding-right:5px;"><input type="text" style="width:30px;" name="FM_antal" id="FM_antal" value=""> til: 
		<select name="FM_nymgrp" id="FM_nymgrp" style="width:125px;">
		<%
		strSQL2 = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		oRec2.open strSQL2, oConn, 3 
		while not oRec2.EOF 
		
		if cint(oRec2("id")) = oRec("grpid") then
		selThis2 = "SELECTED"
		else
		selThis2 = ""
		end if%>
		<option value="<%=oRec2("id")%>" <%=selThis2%>><%=oRec2("navn")%>
		<%if level <= 2 OR level = 6 then %> 
		&nbsp;(<%=oRec2("av") %>%)
		<%end if %>
		</option>
		<%
		oRec2.movenext
		wend
		oRec2.close %>
		</select>&nbsp;<input type="image" src="../ill/soeg-knap.gif"></td></form>
		<td>&nbsp;</td>
	</tr>
	<%
	
	if len(oRec("navn")) <> 0 then
	navn = replace(oRec("navn"), Chr(34), "''")
	else
	navn = ""
	end if
	
	if len(oRec("betegnelse")) <> 0 then
	betegnelse = replace(oRec("betegnelse"), Chr(34), "''")
	else
	betegnelse = ""
	end if
	
	strExport = strExport &"xx99123sy#z"&oRec("gnavn")&";"&oRec("lokation")&";"&oRec("varenr")&";"& navn &";"& betegnelse &";"&oRec("antal")&";"&oRec("enhed")&";"&oRec("minlager")
	
	if level = 1 then
	strExport = strExport &";"&formatnumber(oRec("indkobspris"),decim)&";"&formatnumber(oRec("salgspris"),decim)
	end if
	
	
	bestilantal = (oRec("antal") - oRec("minlager")) 
	if bestilantal < 0 then
	bestilantal = -(bestilantal)
	else
	bestilantal = 0
	end if
	
	strAntalFM = strAntalFM &"<input id='FM_antal_b_"&oRec("id")&"' name='FM_antal_b_"&oRec("id")&"' type=hidden value='"&bestilantal&"' />" 
	
	
	x = x + 1
	oRec.movenext
	wend
	
	if x = 0 then%>
	<tr>
		<td bgcolor="#003399" colspan="<%=cspan+2 %>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td height=50 style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=<%=cspan%>><font color="red"><b>!</b></font>&nbsp;Der er ikke nogen materialer der matcher de valgte søgekriterier.</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%end if%>	
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=<%=cspan%> valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	
	</div><!-- table div-->

	<br><br><br><br><br>&nbsp;
	
	<form action="materialer_ordrer.asp?func=opr" method="post" target="_blank"> 
			<input type="hidden" name="matids" id="matids" value="<%=strMatids%>">
			<%=strAntalFM%>
		    <input type="submit" value="Opret mat. ordre >>">
			
			</form>
			
			<br /><br />
			Der blev fundet ialt: <b><%=x%></b> materialer.
	<br><br>&nbsp;
			
    
    <%'** Eksporter & Pirnt **'
    
    ptop = 0
pleft = 710
pwdt = 120

call eksportogprint(ptop,pleft,pwdt)
%>

<form action="materialer_eksport.asp" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=strExport%>">
			<input type="hidden" name="txt20" id="txt20" value="">
			
<tr>
    
   
    <td align=center><input type="image" src="../ill/export1.png"></td>
    </td><td>.csv fil eksport</td>
   
    </tr>
   <%if level <= 2 OR level = 6 then 
	        
	'oleftpx = 0
	'otoppx = 20
	'owdtpx = 150
	
	'call opretNy("materialer.asp?menu="&menu&"&func=opret&soglev="&sogKrilev&"&FM_mgrp="&matGrpSel&"&FM_sog="&sogVal&"", "Opret nyt materiale", otoppx, oleftpx, owdtpx) 
	
	%>
     <tr><td colspan="2"><br /><br />
    <%
       
                nWdt = 120
                nTxt = "Opret nyt"
                nLnk = "materialer.asp?menu="&menu&"&func=opret&soglev="&sogKrilev&"&FM_mgrp="&matGrpSel&"&FM_sog="&sogVal
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>

         </td>
        </tr>
	
	<%end if %>
   
	</form>
   </table>
</div>
    	
	
	
			
			
	
	
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
