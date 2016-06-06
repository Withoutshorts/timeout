<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	thisfile = "stat_fak" %>
	<SCRIPT LANGUAGE="JavaScript">
	
	function checkAll(field)
	{
	field.checked = true;
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	function uncheckAll(field)
	{
	field.checked = false;
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
	}
	
	function showFak() {
	document.all["faktura"].style.visibility = "visible";
	document.all["statknap"].style.visibility = "visible";
	}
	function showStat() {
	document.all["faktura"].style.visibility = "hidden";
	document.all["statknap"].style.visibility = "hidden";
	}
	
	//function NewWin_popupcal(url)    {
	//	window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
	//}
	
	
	//function hidefaknr(){
	//document.all["FM_fakint"].value = ""
	//}
	
	//function hidedatointerval(){
	//document.all["FM_usedatointerval"].checked = ""
	//}
	</script>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
	<%de = 1%>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top"><br><img src="../ill/header_fak.gif" alt="" border="0" width="279" height="34">
	<hr align="left" width="560" size="1" color="#000000" noshade>
<!-------------------------------Sideindhold------------------------------------->
<%
varYear = datepart("y",date)
%>
<table cellspacing="0" cellpadding="0" border="0">
<tr>
	<td valign="top" width="220">
	
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="#5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
			<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="279" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
			<td colspan=3 valign="top" class="alt">1. Vælg Job</td>
	</tr>
	<%
	if len(Request("filt")) <> 0 then
		filt = Request("filt")
		else
			if len(request.cookies("filt")) <> 0 then
			filt = request.cookies("filt")
			else
			filt = ""
			end if
	end if
		
	'*** Fra webblik ****
	if len(request("fakdenne")) <> 0 then
	fakturerdenne = request("fakdenne")	
	else
	fakturerdenne = 0
	end if
		
	Response.cookies("filt") = filt
	
	%>
	<form name=soeg id=soeg action="stat_fak.asp?menu=stat_fak&FM_kunde=<%=request("FM_kunde")%>&shokselector=1" method="post" name="soeg" id="soeg">
	<tr bgcolor="#EFF3FF">
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=3 style="padding-top:5px;"><b>Søg:</b><br>
		A) På kunde ved at bruge filteret nede til venstre.<br>
		B) Søg på jobnr eller jobnavn:<br>
		<%
		jnavn = request("jnavn")
		%>
		<input type="text" name="jnavn" size="25" value="<%=jnavn%>">
		<%
		
		
		select case filt 
		case "akt"
		chk1 = "CHECKED"
		chk2 =  ""
		
		case else
		chk1 =  ""
		chk2 =  "CHECKED"
		
		end select
		%>
		<br>
		(Job der er en del af en <u>Aftale</u> kan ikke faktureres.)<br>
		<input type="radio" name="filt" value="akt" <%=chk1%>> Vis kun aktive job.<br>
		<input type="radio" name="filt" value="alle" <%=chk2%>> Vis alle job.
		<img src="../ill/blank.gif" width="120" height="1" alt="" border="0">
		<input type="image" name="soeg" src="../ill/brug-filter.gif"><br>&nbsp;
		</td>
		<td valign="top" align="right" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr></form>
	<!--<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="3"><b>C) På status:</b>&nbsp;&nbsp;&nbsp;<a href="stat_fak.asp?menu=stat_fak&filt=aaben&shokselector=1&FM_kunde=<%=request("FM_kunde")%>"><img src="../ill/status_groen.gif" width="14" height="19" alt="Åbne" border="0"></a>&nbsp;&nbsp;&nbsp;<a href="stat_fak.asp?menu=stat_fak&filt=lukket&shokselector=1&FM_kunde=<%=request("FM_kunde")%>"><img src="../ill/status_roed.gif" width="14" height="19" alt="Lukkede" border="0"></a>&nbsp;&nbsp;&nbsp;<a href="stat_fak.asp?menu=stat_fak&filt=passiv&shokselector=1&FM_kunde=<%=request("FM_kunde")%>"><img src="../ill/status_graa_filter2.gif" width="18" height="20" alt="Passive" border="0"></a></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>-->
	
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
		<td colspan=3 valign=top><b>Marker de ønskede job/aftaler:</b><br>
		<form action="fak_osigt.asp?menu=kon" method="post" name="statselector" id="statselector">
		<img src="../ill/blank.gif" width="50" height="1" alt="" border="0"><a href="#" name="CheckAll" onClick="checkAll(document.statselector.FM_job), checkAll(document.statselector.FM_job)"><img src="../ill/alle.gif" border="0"></a>
		&nbsp;&nbsp;<a href="#" name="UnCheckAll" onClick="uncheckAll(document.statselector.FM_job), uncheckAll(document.statselector.FM_job)"><img src="../ill/ingen.gif" border="0"></a>
		</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
	</tr>
	</table>
	
	<div style="position:relative; height:235; overflow:auto; border-bottom:1px #003399 solid;">
	<table cellspacing="0" cellpadding="0" border="0">
	<%
	'** Vælger sql sætning efter sortering og filter ***
	
	select case filt
	case "akt"
	varFilt = " jobstatus = 1 "
	case "alle"
	varFilt = " jobstatus <> 999 "
	case else
	varFilt = " jobstatus = 1 "
	end select
	
	
	
	if len(jnavn) > 0 then
	varFilt = varFilt & " AND (job.jobnavn LIKE '"& jnavn &"%' OR job.jobnr = '"& jnavn &"')"
	else
	varFilt = varFilt
	end if
	
	if len(Request("FM_Kunde")) <> 0 then 
		if Request("FM_Kunde") <> "0" then
			varJobknrKri = "job.jobknr = "& Request("FM_Kunde") &" AND "
		else
			varJobknrKri = ""
		end if
	else
	varJobknrKri = ""
	end if
	
	antalJob = 0
	
	strSQL = "SELECT id, jobnr, jobnavn, jobknr, Kkundenavn, kid, fakturerbart, jobstatus, jobslutdato FROM job, kunder WHERE "& varJobknrKri &" "& varFilt &" AND kid = jobknr AND fakturerbart = 1 AND serviceaft = 0 GROUP BY jobnr ORDER BY Jobnavn"
	'Response.write strSQL
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	bgcolor = "#d2691e"
	jobtype = "E"
	antalE = 1
	imgikon = "<img src='../ill/blank.gif' width='1' height='1' alt='' border='0'>"
	%>
	<tr bgcolor="#eff3ff">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="22" alt="" border="0"></td>
		<td>
		
		<%if cint(fakturerdenne) = oRec("jobnr") then
		chkThis = "CHECKED"
		else
		chkThis = ""
		end if%>
		
		<input type="checkbox" name="FM_job" id="FM_job" value="<%=oRec("jobnr")%>" <%=chkThis%>>
		<input type="hidden" name="FM_job_alle" value="<%=oRec("jobnr")%>"></td>
		<td width=270><%=imgikon%>&nbsp;&nbsp;<%=oRec("Jobnr")%>&nbsp;&nbsp;<%=left(oRec("Jobnavn"), 14)%>&nbsp;&nbsp;<font class='lillegray'>(<%=left(oRec("Kkundenavn"), 16)%>)</font></td>
		<td>&nbsp;&nbsp;<%
			select case oRec("jobstatus")
			case 1
				if (oRec("jobslutdato") - date) < 0 then
				strStatImg = "<img src='../ill/status_slut.gif' width='16' height='19' alt='' border='0'>"
				else
				strStatImg = "<img src='../ill/status_groen.gif' width='14' height='19' alt='' border='0'>"
			end if
			case 2
			strStatImg = "<img src='../ill/status_graa.gif' width='14' height='20' alt='' border='0'>"
			case 0
			strStatImg = "<img src='../ill/status_roed.gif' width='14' height='19' alt='' border='0'>"
			end select
		
		Response.write strStatImg%></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="22" alt="" border="0"></td>
	</tr>
	<%
	antalJob = antalJob + 1
	oRec.movenext
	wend
	oRec.close
	
	%>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="3" valign="bottom"><img src="../ill/tabel_top.gif" width="279" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	</div>
	
	
	
	</td>
	<td valign="top" width="40">&nbsp;</td>
	<td valign="top" width="220">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="#5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
			<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="208" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
			<td colspan=2 class="alt">2. Kør Fakturerings oversigt</td>
	</tr>
	<tr bgcolor="#EFF3FF">
	<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td colspan="2" valign="top"><br><b>A) Periodeafgrænsning.</b><br>
	Finder fakturaer og timer til fakturering i den valgte periode.<br>  <input type="hidden" name="FM_usedatointerval" id="FM_usedatointerval" value="1">
	<!--#include file="inc/weekselector_s.asp"--><br><br>
	<br>
	<input type="hidden" name="FM_vis_betalte" value="0">
	<b>B)</b> Angiv et <b>faktura nr</b> / <b>interval</b> her, hvis du ønsker at finde bestemte fakturaer.<br>
	<input type="text" name="FM_fakint" size="12">&nbsp;<font class=megetlillesort>f.eks. 1045 ell. 1045-1067</font><br>
	
	<br />
    <input id="use_nr_or_interval" name="use_nr_or_interval" <%=useniCHK%> value=1 type="checkbox" /><b>Ignorer valgte kontakter og job.</b><br />
    Hvis der er angivet et faktura nr / interval (B) bruges dette, ellers bruges periode (A).<br><br>
    
	<font class=megetlillesort>(Den i pkt. A valgte periode-afgrænsning bruges stadigvæk til visning af timer til fakturering, samt som afgrænsning ved oprettelse af nye fakturaer.)<br><br>&nbsp;</td>
	<td valign="top" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=2 bgcolor="#ffffe1" style="padding:2px; border:1px #cccccc dashed;"><b>Auto-tilvælg print af medarbejdere?</b><br>
		Når jeg opretter en faktura, skal udspecificering af medarbejdere på print automatisk være slået til?<br>
		<%
		if request.cookies("tvmedarb") = "j" then
		chkA = "CHECKED"
		chkB = ""
		else
		chkB = "CHECKED"
		chkA = ""
		end if
		%>
		<input type="radio" name="FM_chkmed" value="1" <%=chkA%>> Ja 
		<input type="radio" name="FM_chkmed" value="0" <%=chkB%>> Nej</td>
		<td valign="top" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
		<td colspan="2" align=center><br>
		<input type="image" src="../ill/statpil.gif" alt="Kør Statistik!">
		</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom"><img src="../ill/tabel_top.gif" width="208" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
</td>
</tr>

</form>
</table>
<br><br>
</div>

<%
'now close and clean up
	
	oConn.close
	Set oRec = Nothing
	set oConn = Nothing 

end if

%>
<!--#include file="../inc/regular/footer_inc.asp"-->
