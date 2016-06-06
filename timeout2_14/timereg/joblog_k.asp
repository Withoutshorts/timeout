<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<SCRIPT language=javascript src="inc/serviceaft_osigt_func.js"></script>

<script>
function NewWin_large(url)    {
window.open(url, 'Help', 'width=980,height=600,scrollbars=yes,toolbar=no');    
}
</script>

<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	
	thisfile = "joblog_k"
	'**** Job ****
	
	
	kundeid = request("usekid") 'session("Mid")
	
	print = request("print")
	
	
	'*** Finder det valgte job ***
	if len(request("FM_seljob")) <> 0 then
	jobnr = request("FM_seljob")
	else
	jobnr = -1
	end if
	
	if len(request("func")) <> 0 then
	func = request("func")
	else
	func = "tim"
	end if
	
	if len(request("FM_orderby_medarb")) <> 0 AND request("FM_orderby_medarb") <> 0 then
	fordelpamedarb = request("FM_orderby_medarb")
	    
	    if fordelpamedarb = 1 then
	    orderByKri = "Tjobnavn, Tjobnr, Tmnavn"
	    else
	    orderByKri = "Tjobnavn, Tjobnr, TAktivitetId"
	    end if
	
	else
	fordelpamedarb = 0
	orderByKri = "Tjobnavn, Tjobnr"
	end if
	
	
	orderByKri = orderByKri & ", Tdato DESC"
	
	
	if len(request("FM_visgodkendte")) <> 0 then
	visgodkendte = request("FM_visgodkendte")
		select case visgodkendte
		case 1
		visGodkendtfilterKri = " AND godkendtstatus = 1"
		case 2
		visGodkendtfilterKri = " AND godkendtstatus = 0"
		case else 
		visGodkendtfilterKri = " AND godkendtstatus <> 99"
		end select
	else
	visgodkendte = 0
	visGodkendtfilterKri = " AND godkendtstatus <> 99"
	end if
	
	
	if len(request("frompost")) <> 0 OR print = "j" then
	
	        if len(request("FM_ignorerperiode")) <> 0 AND request("FM_ignorerperiode") <> 0 then
	        visheleperiode = 1
	        vishelperchk = "CHECKED"
	        else
	        visheleperiode = 0
	        vishelperchk = ""
	        end if
	else
	        
	        visheleperiode = 1
	        vishelperchk = "CHECKED"
	
	
	end if
	
	
	session.lcid = 1030
	
	'*** Indsætter kundekommentar ***
	if func = "kom" then
	
	timerid = request("FM_timerid")
	dagsdato = year(now) &"/"& month(now) &"/"& day(now)
	str_dagsdato = formatdatetime(dagsdato, 1)
	
	'if len(request("FM_godkendTimer")) <> 0 then
	'godkendtimer = 1
	'else
	godkendtimer = 0
	'end if
	
	komm = "<br><font color=#5582d2>"& str_dagsdato &", af "& session("user") &":<br></font>" & replace(request("FM_kom"), "'", "''")
	strSQLkomm = "INSERT INTO incidentnoter (editor, dato, note, timerid, todoid, status) "_
	&" VALUES ('"& session("user") &"', '"& dagsdato &"', '"& komm &"', "& timerid &", 0, "& godkendtimer &")"
	
	'Response.write strSQLkomm
	oConn.execute(strSQLkomm)
	Response.redirect "joblog_k.asp?func=tim&FM_seljob="&jobnr&"&usekid="&kundeid&"&FM_ignorerperiode="&visheleperiode&"&FM_orderby_medarb="&fordelpamedarb&"&FM_visgodkendte="&visgodkendte
	'Response.Write "kid:" & useKid
	end if
	
	
	'**** kunde/overordnet godkender *** 
	if func = "godkendt" then
	
	intGodkendtetimer = Split(request("FM_godkend"), ", ")
			For b = 0 to Ubound(intGodkendtetimer)
				strSQL = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE tid = " & intGodkendtetimer(b)
				'Response.write strSQL &"<br>"
				oConn.execute(strSQL)
			next
	
	Response.redirect "joblog_k.asp?func=tim&FM_seljob="&jobnr&"&usekid="&kundeid&"&FM_ignorerperiode="&visheleperiode&"&FM_orderby_medarb="&fordelpamedarb&"&FM_visgodkendte="&visgodkendte
	
	end if
	
	
	
	if print <> "j" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%
	else%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
		
	<%end if%>
	
	<!--#include file="inc/convertDate.asp"-->
	<script language="javascript">
	<!--
	
	function hideprint(){
	document.getElementById("printknap").style.visibility = "hidden"
	document.getElementById("printknap").style.display = "none"
	}
	
	function showkommentar(tid){
	document.getElementById("FM_timerid").value = tid
	document.getElementById("kommentar").style.visibility = "visible"
	document.getElementById("kommentar").style.display = ""
	document.getElementById("FM_kom").focus()
	document.getElementById("FM_kom").select()
	}
	
	function hidekommentar(){
	document.getElementById("FM_timerid").value = 0
	document.getElementById("kommentar").style.visibility = "hidden"
	document.getElementById("kommentar").style.display = "none"
	}
	
	
	
	function checkAll(field)
	{
	field.checked = true;
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	function uncheckAll(field){
	field.checked = false;
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
	}
	
	
	
	//-->
	</script>
	
	
	
		
	<%
		
	
	
	if print <> "j" then
		
		select case func
		case "tim"
		selMenuPkt = 1
		case "aft"
		selMenuPkt = 2
		end select 		
		
		'**** Hovedmenu ***
		%>
		<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
		<%
		call kundelogin_mainmenu(selMenuPkt, lto, kundeid)%>
		</div>
	<%
	end if 
	
	
	


	'***** Filter ****'
	
	
		if len(request("filter_per")) <> 0 then
		filterKri_kundelogin = request("filter_per")
		else
			if len(request.cookies("so_filterKri")("periode")) <> 0 then
			filterKri_kundelogin = request.cookies("so_filterKri")("periode")
			else
 			filterKri_kundelogin = 0
			end if 
		end if
		
		select case filterKri_kundelogin
		case 0
		chkfilt0 = "CHECKED"
		chkfilt1 = ""
		chkfilt2 = ""
		useFilterKri = ""
		case 1
		chkfilt0 = ""
		chkfilt1 = ""
		chkfilt2 = "CHECKED"
		useFilterKri = " AND advitype = 1 "
		case 2
		chkfilt0 = ""
		chkfilt1 = "CHECKED"
		chkfilt2 = ""
		useFilterKri = " AND advitype = 2 "
		end select
		
		
			filterStatus = request("status")
			select case filterStatus
			case "vislukkede"
			chkfilt3 = ""
			chkfilt4 = ""
			chkfilt5 = "CHECKED"
			useFilterKri = useFilterKri & " AND status = 0 "
			case "visalle"
			chkfilt3 = "CHECKED"
			chkfilt4 = ""
			chkfilt5 = ""
			useFilterKri = useFilterKri & ""
			case else
			chkfilt3 = ""
			chkfilt4 = "CHECKED"
			chkfilt5 = ""
			useFilterKri = useFilterKri & " AND status = 1 "
			end select
			
		
		   
		
		
	if len(request("FM_typer")) <> 0 then
	vartyper = request("FM_typer")
	dim typer
    typerSQL = " tfaktim = -1 "
	typer = split(request("FM_typer"), ",")
	    for t = 0 to UBOUND(typer)
	    typerSQL = typerSQL & " OR tfaktim = "& typer(t) 
    	    
	        select case typer(t) 
	        case 1
	        chktype_1_2 = "CHECKED"
	        viskorsel = 0
	        case 5
	        chktype_5 = "CHECKED"
	        viskorsel = 1
	        end select
    	    
	    next 
	else
	    vartyper = "1,2"
	    typerSQL = " tfaktim = 1 OR tfaktim = 2"
	    chktype_1_2 = "CHECKED"
        chktype_5 = ""
        viskorsel = 0
    end if 
		
		
		
	
	if print <> "j" then%>
	
	<%call filterheader(67,230,690,pTxt)%>
	<table cellspacing="2" cellpadding="2" border="0" width=100%>
	<form action="joblog_k.asp?m0enu=<%=menu%>&func=<%=func%>&FM_usedatokri=1&frompost=1" method="post">
	<input type="hidden" name="usekid" value="<%=kundeid%>">
	
	
	<%select case func%>
	
	<%case "tim"%>
	
	
	<%if len(trim(request("FM_jobstatus"))) <> 0 then
		jobstCHK = "CHECKED"
		jobstSQL = " AND jobstatus <> 99"
		else
		jobstCHK = ""
		jobstSQL = " AND jobstatus = 1"
		end if %>
	
	<tr>
		<td valign=top align=right><b>Job:</b></td>
		<td width=350 valign=top style="padding-left:15;">
		<%
		strSQL = "SELECT jobnavn, jobnr, id, budgettimer, jobslutdato, jobTpris, "_
		&" k.kkundenavn, k.kkundenr"_
		&" FROM kunder k "_
		&" LEFT JOIN job on (jobknr = " & kundeid &" AND ((fakturerbart = 1 AND kundeok = 1 "& jobstSQL &") "_
		&" OR (fakturerbart = 1 AND kundeok = 2 AND jobstatus = 0))) "_
		&" WHERE k.kid = " & kundeid &""_
		&" ORDER BY k.kkundenavn, jobnavn"		
		
		'Response.write strSQL
		'Response.flush
		
		strJobnrs = 0
		
		%>
		<select name="FM_seljob" id="FM_seljob" style="width:300; font-size:11;" onchange="submit()">
		<option value="-1">Vælg Job..</option>
		
		<%if cdbl(jobnr) = 0 then 
		allJCHK = "SELECTED"
		else
		allJCHK = ""
		end if
		%>
		<option value="0" <%=allJCHK%>>Alle</option>
		<%
		
		
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		if cdbl(jobnr) = oRec("jobnr") then
		seljid = "SELECTED"
		else
		seljid = ""
		end if 
		
		
		if len(trim(oRec("jobnr"))) <> 0 then
		jobnrss = oRec("jobnr")
		else
		jobnrss = 0
		end if
		
		if jobnr <> "-1" then
		strJobnrs = strJobnrs & "," & jobnrss
		end if
		
		%>
		<option value="<%=oRec("jobnr")%>" <%=seljid%>><%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>) | <%=oRec("jobnavn")%> &nbsp;(<%=oRec("jobnr")%>)</option>
		<%
		oRec.movenext
		wend
		oRec.close 
		
		
		
		%>
		
		
		</select>
		<br>
		
		
		
            <input id="FM_jobstatus" name="FM_jobstatus" value="1" type="checkbox" <%=jobstCHK%> /> Vis lukkede og passive job
		
		
	    <br><br><b>Vis aktivitets type(r):</b>
	     <br />
        <input name="FM_typer" id="type" value="1,2" type="checkbox" <%=chktype_1_2 %> /> Timer (fakturerbare og ikke fakturerbare)<br />
        <input name="FM_typer" id="type" value="5" type="checkbox" <%=chktype_5 %> /> Kørsel (km) <br />
        
		
		
		</td>
		
		<td valign=top style="padding-left:15; padding-top:2;">
		<b>Subtotaler:</b><br>
	
		<%if fordelpamedarb = 1 then
		chkFordelmedarb = "CHECKED"
		else
		chkFordelmedarb = ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="FM_orderby_medarb" value="1" <%=chkFordelmedarb%>>Fordel på medarbejdere<br>
	    
	    <%if fordelpamedarb = 2 then
		chkFordelpaakt = "CHECKED"
		else
		chkFordelpaakt = ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="FM_orderby_medarb" value="2" <%=chkFordelpaakt%>>Fordel på aktiviteter<br>
	    
	    <%if fordelpamedarb = 0 then
		chkFordelingen = "CHECKED"
		else
		chkFordelingen= ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="Radio1" value="0" <%=chkFordelingen%>>Ingen<br>
	
	    
	</td>
		
	</tr>
	
	<%end select%>
	
	
	<%select case func 
	case "aft", "tim"
	%>
	<tr>
		<td valign=top align=right><b>Periode:</b>
		<%if func = "aft" then%>
		<br><font class=megetlillesort>Vis kun aftaler<br> med <u>startdato</u><br> i det valgte interval.</font>
		<%end if%>
		</td>
		<td valign=top width=200 style="padding-left:15px;"><!--#include file="inc/weekselector_s.asp"--><br>
	
		<input type="checkbox" name="FM_ignorerperiode" id="FM_ignorerperiode" value="1" <%=vishelperchk%>>Ignorer periode.<br></td>
		
		
		<%if func = "aft" then%>
		<input type="hidden" name="FM_seljob" value="<%=jobnr%>">
		
		<td valign=top>
		<table><tr><td>
		<b>Type:</b><br>
			<input type="radio" name="filter_per" id="filter_per" value="0" <%=chkfilt0%>>Vis Alle&nbsp;&nbsp;<br>
			<input type="radio" name="filter_per" id="filter_per" value="1" <%=chkfilt2%>>Enheds/klip aftaler.&nbsp;&nbsp;<br>
			<input type="radio" name="filter_per" id="filter_per" value="2" <%=chkfilt1%>>Periode aftaler.
		</td><td><b>Status:</b><br>
			<input type="radio" name="status" id="status" value="visalle" <%=chkfilt3%>>Vis Alle&nbsp;&nbsp;<br>
			<input type="radio" name="status" id="status" value="visaktive" <%=chkfilt4%>>Aktive.&nbsp;&nbsp;<br>
			<input type="radio" name="status" id="status" value="vislukkede" <%=chkfilt5%>>Lukkede.&nbsp;&nbsp;
		
		</td></tr></table>
		</td>
		<%end if%>
		
		
		<td valign=bottom style="padding-right:15px;" align=right>
            <input id="Submit1" type="submit" value="Kør statistik >> " /></td>
	</tr>
	<%end select%>
	</form>
	</table>
	
	<!-- slut filter div --> 
	</td>
	</tr>
	</table>
    </div>
    
    
    
    
	
<%
        ptopher = 123
	    plefther = 950
		
		call visfirmalogo(plefther, ptopher, kundeid)
		
		
		
end if '** Print


if print <> "j" then

ptop = 250
pleft = 20
pwdt = 200

call eksportogprint(ptop,pleft,pwdt)
%>


    <tr>
    <td>
    <a href="joblog_k.asp?FM_jobstatus=<%=request("FM_jobstatus") %>&FM_orderby_medarb=<%=fordelpamedarb%>&func=<%=func%>&FM_seljob=<%=jobnr%>&print=j&usekid=<%=kundeid%>&FM_ignorerperiode=<%=visheleperiode%>&FM_visgodkendte=<%=visgodkendte%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&strJobnrs=<%=strJobnrs%>&status=<%=request("status")%>&filter_per=<%=request("filter_per")%>&FM_typer=<%=vartyper%>" target="_blank" class="vmenu">&nbsp;
	&nbsp;<img src="../ill/printer3.png" border=0 alt="" />
   Print version</a></td>
   </tr>
  </table>
</div>
<%else%>

<% 
Response.Write("<script language=""JavaScript"">window.print();</script>")
%>
<%end if




	'**** Sideindhold ****
	
	
	
	if print <> "j" then
	sideDivTop = "380px"
	sideDivLeft = "20px"
	else
	sideDivTop = "40px"
	sideDivLeft = "20px"
	end if
	
	%>
	<div id="sindhold" style="position:absolute; left:<%=sideDivLeft%>; top:<%=sideDivTop%>; visibility:visible; z-index:100; border:0px #000000 solid;">
	
	<h3><%=SideHeader%></h3>
	
	
	
	<%
	
	
	if print = "j" then
	
	strMrd = Request.Cookies("datoer")("st_md")
	strDag = Request.Cookies("datoer")("st_dag")
	strAar = Request.Cookies("datoer")("st_aar") 
	strDag_slut = Request.Cookies("datoer")("sl_dag")
	strMrd_slut = Request.Cookies("datoer")("sl_md")
	strAar_slut = Request.Cookies("datoer")("sl_aar")
	
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	else
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	end if
	
	
	if func <> "fil" then
	
		if print = "j" then
				if visheleperiode = "1" then
				Response.write "<br>Periode afgrænsning: Ingen"
				else
				Response.write "<br>Periode afgrænsning: " & formatdatetime(strDag&"/"&strMrd&"/"&strAar, 2) & " til " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2)
				end if
		end if
	
	end if
	
	
	'*** Sideindhold ***
	select case func
	case "aft"
	'****************************** Aftaler. ***********************************'
	%>
	<!--#include file="inc/serviceaft_osigt_inc.asp"-->
	<%
	case "tim"
	

	call visjoblog(showtimerkorsel) 
	
	
	
	%>
		<div id=kommentar name=kommentar style="position:absolute; left:190; top:280; z-index:100000; visibility:hidden; display:none; width:400; background-color:#EFF3FF; border:1px #8caae6 solid; padding:10px;">
		<table border=0 cellpadding=0 cellspacing=0 width=100%>
		<form action="joblog_k.asp?func=kom" method="post">
		<input type="hidden" name="FM_timerid" id="FM_timerid" value="0">
		<input type="hidden" name="usekid" value="<%=kundeid%>">
		<input type="hidden" name="FM_seljob" value="<%=jobnr%>">
		<input type="hidden" name="FM_ignorerperiode" value="<%=visheleperiode%>">
		<input type="hidden" name="FM_orderby_medarb" value="<%=fordelpamedarb%>">
		<input type="hidden" name="FM_visgodkendte" value="<%=visgodkendte%>">
		
		<tr><td>
		<b>Kommentar:</b><br>
		<textarea cols="40" rows="7" name="FM_kom" id="FM_kom" style="width:380;"></textarea><br>
		<!--<input type="checkbox" name="FM_godkendTimer" id="FM_godkendTimer" value="1">Godkend timer/enheder.-->
		</td></tr>	
		<tr><td align=center>
		<input type="submit" value="Opdater">
		</td></tr>
		<tr><td style="padding:5px;"><a href="#" onClick='hidekommentar()' class=vmenu>Luk vindue</a></td></tr>
		</form>
		</table>	
		</div>	
	<%
	end select

	
	
	
	
	'****** Subs og Funktioner ****'
	public lastmedarbnavn
	sub fordelpamedarbejdere%>
					
					
					<tr bgcolor="#EFF3FF">
					
					    <td align=right colspan="3"><br />
		
		Fakturerbare:<br />
		Ikke Fakturerbare:<br />
		<%if viskorsel = 1 then %>
		Kørsel:
		<%end if %>
		</td>
		<td align=right><br />
		<%=formatnumber(medarbFakbare, 2) %> timer<br />
		<%=formatnumber(medarbIkkeFakbare, 2) %> timer<br />
		
		<%if viskorsel = 1 then %>
		<%=formatnumber(medarbKm, 2) %> km
		<%end if %>
		</td>
		<td>
            &nbsp;</td>
		<td align=right><%=formatnumber(medarbEnh, 2)%></td>
		<td align=right style="padding-right:5px;">
            <b><%=lastmedarbnavn%></b></td>
					</tr>
	
	<%
	
	medarbEnh = 0
	medarbFakbare = 0
	medarbIkkeFakbare = 0
	medarbKm = 0
	
	end sub
	
	
	public lastaktnavn
	sub fordelpaakt%>
					
					
					<tr bgcolor="#ffffe1">
					
					    <td align=right colspan="3"><br />
		
		Fakturerbare:<br />
		Ikke Fakturerbare:<br />
		<%if viskorsel = 1 then %>
		Kørsel:
		<%end if %>
		</td>
		<td align=right><br />
		<%=formatnumber(aktFakbare, 2) %> timer<br />
		<%=formatnumber(aktIkkeFakbare, 2) %> timer<br />
		
		<%if viskorsel = 1 then %>
		<%=formatnumber(aktKm, 2) %> km
		<%end if %>
		</td>
		<td>
            &nbsp;</td>
		<td align=right><%=formatnumber(aktEnh, 2)%></td>
		<td align=right style="padding-right:5px;">
            <b><%=lastaktnavn%></b></td>
					</tr>
	
	<%
	
	aktEnh = 0
	aktFakbare = 0
	aktIkkeFakbare = 0
	aktKm = 0
	
	end sub
	
	
	sub subtotaler
	%>
	<tr>
		<td bgcolor="#d6dff5" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td bgcolor="#ffffff" align=right colspan="3"><br />
		Fakturerbare:<br />
		Ikke Fakturerbare:<br />
		<%if viskorsel = 1 then %>
		Kørsel:
		<%end if %>
		
		</td>
		<td align=right><br />
		<b><%=formatnumber(antalFakbare, 2) %></b> timer<br />
		<b><%=formatnumber(antalIkkeFakbare, 2) %></b> timer<br />
		<%if viskorsel = 1 then %>
		<b><%=formatnumber(antalKm, 2) %></b> km
		<%end if %>
		</td>
		<td>
            &nbsp;</td>
		<td align=right><%=formatnumber(antalEnh, 2)%></td>
		<td>
            &nbsp;</td>
	</tr>
	
	
	</table>
    </div>
    <!-- slut table div -->
    
    <br /><br /><br />
	
	<%
	
	antalEnh = 0
	antalFakbare = 0
	antalIkkeFakbare = 0
	antalKm = 0
	
	end sub
	
	
	
	
	
public aktFakbare, aktIkkeFakbare, aktKm, aktEnh
public medarbFakbare, medarbIkkeFakbare, medarbKm, medarbEnh
public antalFakbare, antalIkkeFakbare, antalKm, antalEnh
public jobForkalfakbare, fakbareTimerTotpajob, showtimerkorsel, enhed, lastjobnavn, emailjobans1, strJobnrs 
function visjoblog(showtimerkorsel)

    
    aktFakbare = 0
    aktIkkeFakbare = 0
    aktKm = 0 
    aktEnh = 0

    
    antalFakbare = 0
	antalIkkeFakbare = 0
	antalKm = 0
	antalEnh = 0
	
	medarbEnh = 0
	medarbFakbare = 0
	medarbIkkeFakbare = 0
	medarbKm = 0

	lastAktid = 0
	lastjobnavn = ""
	lastmedarb = 0
	lastSid = 0
	
	if visheleperiode = "1" then
	perintKri = ""
	else
	perintKri = " AND (Tdato >= '"& strStartDato &"' AND Tdato <= '"& strSlutDato &"') "
	end if
	
	if jobnr <> 0 then
	tjobnrKri = " tjobnr = '"& jobnr & "'" 
	else
	
	if print <> "j" then
	strJobnrs = strJobnrs
	else
	strJobnrs = request("strJobnrs")
	end if 
		
		if len(strJobnrs) <> 0 then
			valgtejobnr = split(strJobnrs, ",")
			
			for t = 0 to UBOUND(valgtejobnr)
				if t = 0 then
				tjobnrKri = "(tjobnr = '" & valgtejobnr(t) & "'" 
				else
				tjobnrKri = tjobnrKri & " OR tjobnr = '" & valgtejobnr(t) & "'"
				end if
			next 
			
			tjobnrKri = tjobnrKri & ")"
		else
		tjobnrKri = " tjobnr = '0'" 
		end if
	
	end if
	
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn,"_
	&" TAktivitetId, Tknavn, Tknr, Tmnr, Tmnavn, Timer, Tid, "_
	&" Tfaktim, TimePris, Timerkom, offentlig, s.navn AS aftnavn, "_
	&" sttid, sltid, s.id AS sid, a.faktor, godkendtstatus, godkendtstatusaf FROM timer"_
	&" LEFT JOIN job j ON (j.jobnr = tjobnr) "_
	&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft) "_
	&" LEFT JOIN aktiviteter a ON (a.id = TAktivitetId) "_
	&" WHERE "& tjobnrKri &" "& perintKri &" AND ("& typerSQL &") "& jobstSQL &""_
	&" ORDER BY "& orderByKri
	
	'Response.Write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 0, 1
	v = 0
	while not oRec.EOF
		strWeekNum = datepart("ww", oRec("Tdato"),2,2)
		id = 1
				
				if fordelpamedarb = 1 then
				
				if (lastmedarb <> oRec("Tmnr") OR lastjobnr <> oRec("Tjobnr")) then
					if v <> 0 then
					
					call fordelpamedarbejdere()%>
					
					<%
					end if
				end if
				
				end if
				
				if fordelpamedarb = 2 then
				
				if (lastaktid <> oRec("Taktivitetid") OR lastjobnr <> oRec("Tjobnr")) then
					if v <> 0 then
					
					call fordelpaakt()%>
					
					<%
					end if
				end if
				
				end if
				
				if lastjobnr <> oRec("Tjobnr") then
				%>
				<%if v <> 0 then
					call subtotaler%>
				<%
				timerTotpajob = 0
				end if
				
				if lastmedarb <> oRec("Tmnr") OR lastjobnr <> oRec("Tjobnr") OR lastaktid <> oRec("Taktivitetid") then
					if v <> 0 then%>
					<tr>
						<td height=100 bgcolor="#ffffff" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<%
					end if
				end if
				
				
				
				 tTop = 20
	            tLeft = 0
	            tWdth = 900
            	
            	
	             call tableDiv(tTop,tLeft,tWdth)
	             %>
                <table border="0" width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="page-break-after: always;">


          
				
				<tr>
					<td bgcolor="#ffffe1" colspan="8" height=60 style="padding:10px;">
					Job: <b><%=oRec("Tjobnavn")%> (<%=oRec("Tjobnr")%>)</b><br>
					Kontakt: <b><%=oRec("Tknavn")%>&nbsp;(<%=oRec("Tknr")%>)</b><br>
					<%
					
						'*** Jobansvarlige ***
						
						strSQL2 = "SELECT mnavn, mnr, mid, job.id, jobnavn, jobnr, jobans1, jobstatus, jobstartdato, jobslutdato, budgettimer, ikkebudgettimer, email FROM job LEFT JOIN medarbejdere ON (mid = jobans1) WHERE job.jobnr = '"& oRec("Tjobnr") &"'"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						jobstatus = oRec2("jobstatus")
						jobstartdato = oRec2("jobstartdato")
						jobslutdato = oRec2("jobslutdato")
						jobForkalkulerettimer = (oRec2("budgettimer") + oRec2("ikkebudgettimer"))
						jobForkalfakbare = oRec2("budgettimer")
						emailjobans1 = oRec2("email")
						strJobnavn = oRec2("jobnavn")
						strJobnr = oRec2("jobnr")
						end if
						oRec2.close
						
						
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2, email FROM job, medarbejdere WHERE job.jobnr = '"& oRec("Tjobnr") &"' AND mid = jobans2"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						emailjobans2 = oRec2("email")
						end if
						oRec2.close
						
						
						if showtimerkorsel = 1 then%>
						Jobansvarlig 1: <b><a href="mailto:<%=emailjobans1%>" class=vmenu><%=jobans1txt%></a></b>
						<%if len(trim(jobans2txt)) <> 0 then%>
						<br>
	   					Jobansvarlig 2: <b><a href="mailto:<%=emailjobans2%>" class=vmenu><%=jobans2txt%></a></b>
						<%end if
						jobans1txt = ""
						jobans2txt = ""
						%><br>
						<%select case jobstatus
						case 1
						thisjobstatus = "Aktiv"
						case 2
						thisjobstatus = "Passiv"
						case 0
						thisjobstatus = "Lukket"
						end select%>
						Status: <b><%=thisjobstatus%></b><br>
						Periode: <b><%=formatdatetime(jobstartdato, 1)%> - <%=formatdatetime(jobslutdato, 1)%></b>
						<br>
						Forkalkuleret timer: <b><%=formatnumber(jobForkalkulerettimer, 2)%></b>
						<br>Aftale: <b><%=oRec("aftnavn")%></b>
						<br>&nbsp;
						<%end if%>
					</td>
					
				</tr>
				
				
				
				<tr bgcolor="#5582D2">
					<td class='alt' style="width:90px;">
                        &nbsp;<b>Dato</b></td>
					<td class='alt' style="width:300px;"><b>Aktivitet</b></td>
					<td class='alt' style="width:90px;">&nbsp;
					<%if print <> "j" then %>
                     <b>Tilføj kommentar</b>
                       <%end if %></td>
					<td class='alt' align=right style="padding-right:3px; width:80px;"><b>Antal</b><br>Tid</td>
					<td class='alt' style="padding-right:5px; padding-left:15px; width:120px;"><b>Type</b></td>
					<td class='alt' align=right style="padding-right:3px; width:80px;"><b>Enheder</b></td>
					<td class='alt' align=right style="padding-right:3px; width:150px;"><b>Medarb.</b></td>
					<td class='alt' align=right>&nbsp;</td>
					
				</tr>
				
				<%end if%>
				<tr>
					<td bgcolor="#d6dff5" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<tr>
				<td valign="top" style="padding-top:3px; padding-left:3px; height:30px;">
				<%=formatdatetime(oRec("Tdato"), 2)%></td>
				
				<td valign="top" style="padding-top:3px;"><b>
				<%=oRec("Anavn")%></b>
				
				    <%
				    '**** Kommentar ***
				    
				    if len(oRec("timerkom")) <> 0 AND oRec("offentlig") <> 1 then 'tilgængelig for kunde%>
				
					<br /><%=oRec("timerkom")%>
				    <%end if%>
				    
				
					<%'**** Kundekommentarer ****%>
				
					<%
					strSQLin = "SELECT note FROM incidentnoter WHERE timerid = "& oRec("tid") &" ORDER BY id"
					oRec3.open strSQLin, oConn, 3 
					i = 0
					Timergodkendt = 0
					while not oRec3.EOF 
					
					
					
					Response.write "<br>"& oRec3("note") 
					
					
					
					i =  i + 1
					oRec3.movenext
					wend
					oRec3.close  
					
					
				%>
				</td>
				<td valign=bottom>
                    &nbsp;
				<%
				if oRec("godkendtstatus") = 0 AND print <> "j" then%>
				&nbsp;&nbsp;<a href="#" onClick=showkommentar('<%=oRec("tid")%>')><img src="../ill/notebook_add.png" alt="Tilføj kommentar.." border="0"></a>
				<%end if%>
				</td>
				
				<td align="right" style="padding-top:3px; padding-right:3px;" valign="top"><b><%=formatnumber(oRec("Timer"), 2)%></b>
				
				<%
				
				enheder = oRec("faktor") * oRec("timer")
				
				if len(enheder) <> 0 then
				enheder = enheder
				else
				enheder = 0
				end if
				%>
				
				
				<%select case oRec("tfaktim")
				case 1,2 
				enhang = "timer"
				    if oRec("tfaktim") = 1 then
				    aktFakbare = aktFakbare + oRec("Timer")
				    antalFakbare = antalFakbare + oRec("Timer")
				    medarbFakbare = medarbFakbare + oRec("Timer")
	                else
	                aktIkkeFakbare = aktIkkeFakbare + oRec("Timer")
				    antalIkkeFakbare = antalIkkeFakbare + oRec("Timer")
				    medarbIkkeFakbare = medarbIkkeFakbare + oRec("Timer")
				    end if
				    
				    aktEnh = aktEnh + enheder
				    antalEnh = antalEnh + enheder
				    medarbEnh = medarbEnh + enheder
				    
				case 5
				enhang = "km"
				aktKm = aktKm + oRec("Timer")
				antalKm = antalKm + oRec("Timer")
				medarbKm = medarbKm + oRec("Timer")
				end select
				%>
				<%=enhang%>
				
				<%
				if len(oRec("sttid")) <> 0 then
					if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
					Response.write "<font class=megetlillesort><br>"& left(formatdatetime(oRec("sttid"), 3), 5) & " - " & left(formatdatetime(oRec("sltid"), 3), 5)
					end if
				end if%>
				</td>
				
				<td style="padding-top:3px; padding-left:15px; padding-right:5px;" valign="top">
				
				<%call akttyper(oRec("tfaktim"), 4) %>
				
				 <%=akttypenavn %>
				</td>
				
				
				<td align="right" style="padding-top:3px; padding-right:3px;" valign="top"><b><%=formatnumber(enheder, 2)%></b></td>
				<%
				
				
				
				%>
				<td align=right style="padding-top:3px; padding-right:3px;" valign="top"><%=oRec("Tmnavn")%></td>
				<td align=right style="padding-top:3px;" valign="top">
			    &nbsp;</td>
				</tr>
				
				
				
			<%v = v + 1
			lastaktnavn = oRec("Anavn")
			lastmedarbnavn = oRec("Tmnavn")
			lastmedarb = oRec("Tmnr")
			lastjobnr = oRec("Tjobnr")
			lastjobnavn = oRec("Tjobnavn")
			lastSid = oRec("sid")
			lastaktid = oRec("Taktivitetid")
			
	oRec.movenext
	wend
	oRec.Close
	'Set oRec = Nothing

	if v = 0 then
	%>
	<tr>
		<td bgcolor="#d6dff5" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td bgcolor="#ffffff" colspan="8" height=60><font class=roed><b>Der forefindes ingen registreringer på de(t) valgte job i den valgte periode!</b></font></td>
	</tr>
	<%
	end if
	
	if v > 0  AND fordelpamedarb = 1 then
	
	call fordelpamedarbejdere()
	
	end if
	
if v > 0 then	
call subtotaler
end if
%>

	



<br><br>&nbsp;

<%end function
	

Set oRec = Nothing
end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->