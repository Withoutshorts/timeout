<%response.buffer = true%>
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
	thisfile = "stat" %>
	<SCRIPT LANGUAGE="JavaScript">
	function updateseljob(jobnr) {
	offSet = document.getElementById("seljob").value;
	offSetL = document.getElementById("seljob").value.length;
	offSetJobnr = jobnr.length;
	pkt = offSet.indexOf("#"+jobnr+"#"); 
	
			if (pkt == -1) {
			seljobval = document.getElementById("seljob").value +",#"+jobnr+"#";
			document.getElementById("seljob").value = seljobval;
			} else {
			newstrSelRA = offSet.substring(0, pkt - 1)
			newstrSelRB = offSet.substring(pkt + offSetJobnr + 2, offSetL)
			seljobval = (newstrSelRA + newstrSelRB);
			document.getElementById("seljob").value = seljobval;
			}
			
		//submit knap
		if (document.getElementById("seljob").value.length > 3) {
			if (document.getElementById("selmed").value.length > 3){
				document.getElementById("subm").style.visibility = "visible"
				document.getElementById("subm_ghost").style.visibility = "hidden"
				}
		}else{
		document.getElementById("subm").style.visibility = "hidden"
		document.getElementById("subm_ghost").style.visibility = "visible"
		}
	}
	
	function updateselmed(mid) {
	stat = document.getElementById("selstat").value;
	offSet = document.getElementById("selmed").value;
	offSetL = document.getElementById("selmed").value.length;
	offSetMid = mid.length;
	pkt = offSet.indexOf("#"+mid+"#"); 
			
			if (pkt == -1) {
			selmedval = document.getElementById("selmed").value +",#"+mid+"#";
			document.getElementById("selmed").value = selmedval;
			} else {
			newstrSelRA = offSet.substring(0, pkt - 1)
			newstrSelRB = offSet.substring(pkt + offSetMid + 2, offSetL)
			selmedval = (newstrSelRA + newstrSelRB);
			document.getElementById("selmed").value = selmedval;
			}
		
		
		//submit knap
		if (document.getElementById("selstat").value == "stat_pies") {
			if (document.getElementById("selmed").value.length > 3) {
			document.getElementById("subm").style.visibility = "visible"
			document.getElementById("subm_ghost").style.visibility = "hidden"
			} else {
			document.getElementById("subm").style.visibility = "hidden"
			document.getElementById("subm_ghost").style.visibility = "visible"
			}
		}else{
			if (document.getElementById("selmed").value.length > 3) {
				if (document.getElementById("seljob").value.length > 3){
					document.getElementById("subm").style.visibility = "visible"
					document.getElementById("subm_ghost").style.visibility = "hidden"
				}
			}else{
			document.getElementById("subm").style.visibility = "hidden"
			document.getElementById("subm_ghost").style.visibility = "visible"
			}
		}
	
	}
	
		
	function checkAll(field, usejobormed)
	{
	if (usejobormed == 1){
	document.getElementById("seljob").value = document.getElementById("FM_job_all").value;
		if (document.getElementById("selmed").value.length > 3) {
		document.getElementById("subm").style.visibility = "visible"
		document.getElementById("subm_ghost").style.visibility = "hidden"	
		} else {
		document.getElementById("subm").style.visibility = "hidden"
		document.getElementById("subm_ghost").style.visibility = "visible"	
		}
	} else {
	document.getElementById("selmed").value = document.getElementById("FM_med_all").value;
		if (document.getElementById("seljob").value.length > 3) {
		document.getElementById("subm").style.visibility = "visible"
		document.getElementById("subm_ghost").style.visibility = "hidden"	
		} else {
		document.getElementById("subm").style.visibility = "hidden"
		document.getElementById("subm_ghost").style.visibility = "visible"	
		}
	}
	field.checked = true;
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	
	
	function uncheckAll(field, usejobormed){
	if (usejobormed == 1){
	document.getElementById("seljob").value =  "0#"
	} else {
	document.getElementById("selmed").value =  "0#"
	}
	field.checked = false;
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
		
	document.getElementById("subm").style.visibility = "hidden"
	document.getElementById("subm_ghost").style.visibility = "visible"	
	}
	
	function showFak() {
	document.all["faktura"].style.visibility = "visible";
	document.all["statknap"].style.visibility = "visible";
	}
	function showStat() {
	document.all["faktura"].style.visibility = "hidden";
	document.all["statknap"].style.visibility = "hidden";
	}
	
	
	 if (document.images){
			plus = new Image(200, 200);
			plus.src = "ill/plus.gif";
			minus = new Image(200, 200);
			minus.src = "ill/minus2.gif";
			}
	
			function expand (de){
				if (document.all("Menu" + de)){
					if (document.all("Menu" + de).style.display == "none"){
					document.all("Menu" + de).style.display = "";
					document.images["Menub" + de].src = minus.src;
				}else{
					document.all("Menu" + de).style.display = "none";
					document.images["Menub" + de].src = plus.src;
				}
			}
		}
	
	function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
	}
	
	function knapkundekri(){
	document.getElementById("medarbtop").style.visibility = "hidden"
	document.getElementById("medarbtop").style.display = "none"
	document.getElementById("medarb").style.visibility = "hidden"
	document.getElementById("medarb").style.display = "none"
	document.getElementById("progrp").style.visibility = "hidden"
	document.getElementById("progrp").style.display = "none"
	document.getElementById("jobkri").style.visibility = "visible"
	document.getElementById("jobkri").style.display = ""
	document.getElementById("job").style.visibility = "visible"
	document.getElementById("job").style.display = ""
	document.getElementById("antaljob").style.visibility = "visible"
	document.getElementById("antaljob").style.display = ""
	document.getElementById("jobtop").style.visibility = "visible"
	document.getElementById("jobtop").style.display = ""
	document.getElementById("periode").style.visibility = "hidden"
	document.getElementById("periode").style.display = "none"
	document.images("kjob").src = "../ill/knap_valgjob_on.gif";
	document.images("kmed").src = "../ill/knap_medarbkri_off.gif";
	document.images("kper").src = "../ill/knap_periode_off.gif";
	document.getElementById("subm").style.top = 558
	document.getElementById("subm").style.left = 700
	document.getElementById("subm_ghost").style.top = 558
	document.getElementById("subm_ghost").style.left = 700
	}
	
	function knapmedarbkri(){
	document.getElementById("medarbtop").style.visibility = "visible"
	document.getElementById("medarbtop").style.display = ""
	document.getElementById("medarb").style.visibility = "visible"
	document.getElementById("medarb").style.display = ""
	document.getElementById("progrp").style.visibility = "visible"
	document.getElementById("progrp").style.display = ""
	document.getElementById("jobkri").style.visibility = "hidden"
	document.getElementById("jobkri").style.display = "none"
	document.getElementById("job").style.visibility = "hidden"
	document.getElementById("job").style.display = "none"
	document.getElementById("antaljob").style.visibility = "hidden"
	document.getElementById("antaljob").style.display = "none"
	document.getElementById("jobtop").style.visibility = "hidden"
	document.getElementById("jobtop").style.display = "none"
	document.getElementById("periode").style.visibility = "hidden"
	document.getElementById("periode").style.display = "none"
	document.images("kjob").src = "../ill/knap_valgjob_off.gif";
	document.images("kmed").src = "../ill/knap_medarbkri_on.gif";
	document.images("kper").src = "../ill/knap_periode_off.gif";
	document.getElementById("subm").style.top = 430
	document.getElementById("subm").style.left = 700
	document.getElementById("subm_ghost").style.top = 430
	document.getElementById("subm_ghost").style.left = 700
	}
	
	
	function knapperiodekri(){
	document.getElementById("medarbtop").style.visibility = "hidden"
	document.getElementById("medarbtop").style.display = "none"
	document.getElementById("medarb").style.visibility = "hidden"
	document.getElementById("medarb").style.display = "none"
	document.getElementById("progrp").style.visibility = "hidden"
	document.getElementById("progrp").style.display = "none"
	document.getElementById("jobkri").style.visibility = "hidden"
	document.getElementById("jobkri").style.display = "none"
	document.getElementById("job").style.visibility = "hidden"
	document.getElementById("job").style.display = "none"
	document.getElementById("antaljob").style.visibility = "hidden"
	document.getElementById("antaljob").style.display = "none"
	document.getElementById("jobtop").style.visibility = "hidden"
	document.getElementById("jobtop").style.display = "none"
	document.getElementById("periode").style.visibility = "visible"
	document.getElementById("periode").style.display = ""
	document.images("kjob").src = "../ill/knap_valgjob_off.gif";
	document.images("kmed").src = "../ill/knap_medarbkri_off.gif";
	document.images("kper").src = "../ill/knap_periode_on.gif";
	document.getElementById("subm").style.top = 360
	document.getElementById("subm").style.left = 398
	document.getElementById("subm_ghost").style.top = 360
	document.getElementById("subm_ghost").style.left = 398
	}
	
	function renssog(){
	document.getElementById("jobsog").value = "";
	}
	</script>
	
	
	
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	
<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
<%de = 1%>
<table cellspacing="0" cellpadding="0" border="0" width="600">
<tr>
   <td valign="top"><br>
<img src="../ill/header_stat.gif" alt="" border="0" width="758" height="49"><br>
</td></tr></table>
</div>
<!-------------------------------Sideindhold------------------------------------->
<%

useJobsog = request("jobsog")
'Response.write useJobsog

varYear = datepart("y",date)

'*** Ignorer projektgrupper ****
if len(request("FM_ignorer")) <> 0 then
ign = 2
else
ign = 1 'request("FM_ignorer")
end if

if len(request("FM_kunde")) <> 0 AND len(useJobsog) = 0 then
useKunde = request("FM_kunde")
else
useKunde = 0
end if

if len(request("alle_eksterne")) <> 0 then
useAlle_eksterne = request("alle_eksterne")
else
useAlle_eksterne = 2 'eksterne
end if

if len(request("status")) <> 0 then
useStatus = request("status")
else
useStatus = 1
end if

if len(request("seljob")) <> 0 then
seljob = request("seljob")
else
seljob = "0#"
end if

if len(request("selmed")) <> 0 then
selmed = request("selmed")
else
selmed = "0#"
end if

if request("FM_viskunaktivemedarb") = "1" OR (request.cookies("visallemedarb") = "j" AND len(request("FM_viskunaktivemedarb")) = 0) then

'visallemed = request("FM_viskunaktivemedarb")
kunaktiveKri = " AND mansat <> 2 "
response.cookies("visallemedarb") = "j"
vamChk1 = "CHECKED"
vamChk2 = ""
		
else
kunaktiveKri = " "
response.cookies("visallemedarb") = "n"
vamChk1 = ""
vamChk2 = "CHECKED"

end if



if len(request("show")) <> 0 then
show = request("show")
else
show = "joblog_timberebal" '"joblog_z_b"
end if

select case show
case "joblog_status"
thisheader = "Jobstatus"
case "joblog_z_b"
thisheader = "Timefordeling"
case "joblog_korsel"
thisheader = "Kørsel"
case "joblog"
thisheader = "Joblog"
case "oms"
thisheader = "Omsætning"
case "word"
thisheader = "Eksporter"
case "stat_pies"
thisheader = "%vis timefordeling"
case "joblog_timberebal"
thisheader = "Timeforbrug, Omsætning og Ressourcetimer."
end select

varProgruppe = request("FM_progrupper")
'Response.write "varProgruppe" & varProgruppe
select case varProgruppe
case "", "10"
varProgruppe = "10"
case else
varProgruppe = varProgruppe
end select
'Response.write "varProgruppe" & varProgruppe
%>
<div style="position:absolute; top:73; left:335; z-index:1000;"><font class="stor-blaa">
>> <%=thisheader%></font>
</div>
<%

if len(request("showdiv")) <> 0 then
showdiv = request("showdiv")
else
showdiv = 2
end if

if show <> "stat_pies" then
showdiv = showdiv
else
showdiv = 2
end if
%>
<form action="<%=show%>.asp?menu=stat&func=tot" method="post" name="statselector" id="statselector">
<%
if showdiv <> 1 then
kjobshow = "off"
kmedshow = "on"
else
kjobshow = "on"
kmedshow = "off"
end if
%>

<%
select case show
case "joblog_status", "joblog_z_b", "joblog_korsel", "joblog", "oms", "word", "joblog_timberebal"
vz = "visible"
disp = ""
case else
vz = "hidden"
disp = "none"
end select%>
<div id="knapjobkri" name="knapjobkri" style="position:absolute; left:325; top:128; visibility:<%=vz%>; display:<%=disp%>;">
<a href="#" onClick="knapkundekri();"><img src="../ill/knap_valgjob_<%=kjobshow%>.gif" width="90" height="23" alt="" border="0" name=kjob id=kjob></a>
</div>

<%
select case show
case "joblog_status", "joblog_z_b", "joblog_korsel", "joblog", "oms", "stat_pies", "word", "joblog_timberebal"
vz = "visible"
disp = ""
case else
vz = "hidden"
disp = "none"
end select%>
<div id="knapmedarbkri" name="knapmedarbkri" style="position:absolute; left:204; top:128; visibility:<%=vz%>; display:<%=disp%>;">
<!-- stat.asp?menu=stat&showdiv=2 -->
<a href="#" onClick="knapmedarbkri();"><img src="../ill/knap_medarbkri_<%=kmedshow%>.gif" width="120" height="23" alt="" border="0" name=kmed id=kmed></a>
</div>

<%
select case show
case "joblog_z_b", "joblog", "joblog_korsel", "word", "joblog_timberebal"
vz = "visible"
disp = ""
case else
vz = "hidden"
disp = "none"
end select%>
<div id="knapperiode" name="knapperiode" style="position:absolute; left:416; top:128; visibility:<%=vz%>; display:<%=disp%>;">
<a href="#" onClick="knapperiodekri();"><img src="../ill/knap_periode_off.gif" width="120" height="23" alt="" border="0" name=kper id=kper></a>
</div>


<%
divjobLeft = 565
divjobTop = 172
if showdiv <> 1 then
divjobDisp = "none"
divjobVz = "hidden"
else
divjobDisp = ""
divjobVz = "visible"
end if
%>	
	
	<div id="jobtop" name="jobtop" style="position:absolute; left:<%=divjobLeft%>; top:<%=divjobTop%>; visibility:<%=divjobVz%>; display:<%=divjobDisp%>; z-index:500;"> 	
	<table cellspacing="0" cellpadding="0" border="0" width=350>
	<tr bgcolor="#EFF3FF">
		<td valign="top" colspan=5 style="border-top:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="5" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
			<td colspan=5 style="border-right:1px #003399 solid; padding-left:1;"><font size=3><b>(B)</b></font>&nbsp;&nbsp;<b>Vælg job.</b><br>
			Vælg de(t) ønskede job nedenfor. Søg på kunde og job kriterier under pkt. (A) til venstre.<br>
			<%if show <> "joblog_status" then%>
			<img src="../ill/blank.gif" width="105" height="11" alt="" border="0"><a href="#" name="CheckAll" onClick="checkAll(document.statselector.FM_jobE, 1), checkAll(document.statselector.FM_jobI, 1)"><img src="../ill/alle.gif" border="0"></a>
			&nbsp;&nbsp;<a href="#" name="UnCheckAll" onClick="uncheckAll(document.statselector.FM_jobE, 1), uncheckAll(document.statselector.FM_jobI, 1)"><img src="../ill/ingen.gif" border="0"></a>
			<%else%>
			<img src="../ill/blank.gif" width="105" height="24" alt="" border="0">
			<%end if%>
			</td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top" colspan=5 style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"></td>
	</tr>
	</table>
	</div>
	
	<div id="job" name="job" style="position:absolute; left:<%=divjobLeft+2%>; top:<%=divjobTop+87%>; width:348; height:288; visibility:<%=divjobVz%>; display:<%=divjobDisp%>; background-color:#FFFFFF; overflow:auto; border:1px #003399 solid; z-index:500;"> 	
	<table cellspacing="0" cellpadding="0" border="0" width=320>
	<%
	'** Vælger sql sætning efter sortering og filter ***
	
	if len(useJobsog) > 0 then
	varFilt = " (job.jobnavn LIKE '"& useJobsog &"%' OR job.jobnr = '"& useJobsog &"')" 
	
	if ign <> 1 then
	varFilt = varFilt & " AND (projektgruppe1 = "& varProgruppe &" OR projektgruppe2 = "& varProgruppe &" OR projektgruppe3 = "& varProgruppe &" OR projektgruppe4 = "& varProgruppe &" OR projektgruppe5 = "& varProgruppe &" OR projektgruppe6 = "& varProgruppe &" OR projektgruppe7 = "& varProgruppe &" OR projektgruppe8 = "& varProgruppe &" OR projektgruppe9 = "& varProgruppe &" OR projektgruppe10 = "& varProgruppe &") "
	else
	varFilt = varFilt
	end if
	
	varJobknrKri = ""
	alleeksKri = ""
	
	strSQL = "SELECT id, jobnr, jobnavn, jobknr, Kkundenavn, kid, fakturerbart, jobstatus, jobslutdato FROM job, kunder WHERE ("& varFilt &") AND kid = jobknr GROUP BY job.id ORDER BY kkundenavn, jobnavn, jobnr"
	else
	
	'*** Status
	select case useStatus
	case 1
	varFilt = " jobstatus = 1 "
	case 2
	varFilt = " jobstatus = 0 "
	case 3
	varFilt = " jobstatus = 2 "
	case else
	varFilt = " jobstatus <> 3 " 'tilbud
	end select
	
	'*** Kunde
	if len(Request("FM_Kunde")) <> 0 then 
		if Request("FM_Kunde") <> "0" then
			varJobknrKri = "job.jobknr = "& Request("FM_Kunde") &" AND "
		else
			varJobknrKri = ""
		end if
	else
	varJobknrKri = ""
	end if
	
	'*** Interne / Eksterne
	if cint(useAlle_eksterne) = 2 OR len(useAlle_eksterne) = 0 then
	alleeksKri = ""
	else
	alleeksKri = "OR ("& varJobknrKri &" "& varFilt &" AND fakturerbart = 0)"
	end if
	
	if ign <> 1 then
	varFilt = varFilt & " AND (projektgruppe1 = "& varProgruppe &" OR projektgruppe2 = "& varProgruppe &" OR projektgruppe3 = "& varProgruppe &" OR projektgruppe4 = "& varProgruppe &" OR projektgruppe5 = "& varProgruppe &" OR projektgruppe6 = "& varProgruppe &" OR projektgruppe7 = "& varProgruppe &" OR projektgruppe8 = "& varProgruppe &" OR projektgruppe9 = "& varProgruppe &" OR projektgruppe10 = "& varProgruppe &") "
	else
	varFilt = varFilt
	end if
	
	strSQL = "SELECT id, jobnr, jobnavn, jobknr, Kkundenavn, kid, fakturerbart, jobstatus, jobslutdato FROM job, kunder WHERE ("& varJobknrKri &" "& varFilt &" AND fakturerbart <> 0 AND kid = jobknr) "& alleeksKri &" AND kid = jobknr ORDER BY kkundenavn, Jobnavn, jobnr"
	end if
	
	jobId_all = "0#"
	antalJob = 0
	'Response.write strSQL
	oRec.open strSQL, oConn, 3
	
	while not oRec.EOF
	
	if oRec("fakturerbart") = "1" then
	bgcolor = "#d2691e"
	jobtype = "E"
	antalE = 1
	'imgikon = "<img src='../ill/eksternt_job_trans.gif' width='15' height='15' alt='' border='0'>"
	else
	bgcolor = "lightblue"
	jobtype = "I"
	antalI = 1
	'imgikon = "<img src='../ill/internt_job_trans.gif' width='15' height='15' alt='' border='0'>"
	end if
	
	if lastKundenavn <> oRec("Kkundenavn") then%>
	<tr bgcolor="#ffffff">
		<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=3 height=20 style="padding-left:5;"><b><%=oRec("Kkundenavn")%></b></td>
		<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<%end if%>
	<tr bgcolor="#ffffff">
		<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td style="padding-left:5;">
		<%
		if instr(seljob, "#"&oRec("jobnr")&"#") > 0 then
		thisChecked = "CHECKED"
		else
		thisChecked = ""
		end if
		
		if show <> "joblog_status" then
		useFtype = "checkbox"
		else
		useFtype = "radio"
		end if%>
		<input type="<%=useFtype%>" name="FM_job" id="FM_job<%=jobtype%>" value="<%=oRec("jobnr")%>" onClick="updateseljob('<%=oRec("jobnr")%>');" <%=thisChecked%>>
		<!-- ?? Bruges i joblog_z_b.asp --><input type="hidden" name="FM_job_alle" id="FM_job_alle" value="<%=oRec("jobnr")%>"></td>
		<td width="300" style="padding-left:5;"><%=oRec("Jobnr")%>&nbsp;&nbsp;<%=oRec("Jobnavn")%></td>
		<td>&nbsp;</td>
		<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<%
	jobId_all = jobId_all &",#"&oRec("jobnr")&"#"
	lastKundenavn = oRec("Kkundenavn")
	antalJob = antalJob + 1
	oRec.movenext
	wend
	oRec.close
	
	%>
	<%if antalE <> 1 then%>
	<input type="hidden" id="FM_jobE" value="0">
	<%end if%>
	<%if antalI <> 1 then%>
	<input type="hidden" id="FM_jobI" value="0">
	<%end if%>
	
	<%if antalJob = 0 then%>
	<tr><td style="padding:20;"><br><img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0">Ingen Job matcher de valgte Projektgruppe, Kunde eller Job kriterier.</td></tr>
	<%end if%>
	</table>
	<input type="hidden" name="FM_job_all" id="FM_job_all" value="<%=jobId_all%>">
	</div>
	
	<div id="antaljob" name="antaljob" style="position:absolute; left:<%=divjobLeft+2%>; top:<%=divjobTop+450%>; width:348; visibility:<%=divjobVz%>; display:<%=divjobDisp%>; z-index:500;"> 	
	<table><tr bgcolor="#ffffe1"><td style="padding:5; border:1px #003399 solid;">Der er <b><%=antalJob%></b> job på listen.<br> Antallet af job afhænger af de filter-kriterier der er valgt i <b>pkt. (A)</b>.<br>
	<%
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus <> 3"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneJob = antalEksterneJob
	else
	antalEksterneJob = 0
	end if
	%>
	<br>Der er ialt oprettet <b><%=antalEksterneJob%> Eksterne</b> job i systemet.
	
	
	<%
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneAktiveJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneAktiveJob = antalEksterneAktiveJob
	else
	antalEksterneAktiveJob = 0
	end if
	%>
	<br>Heraf er <b><%=antalEksterneAktiveJob%> Aktive</b>.
	
	<%
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 0"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneLukkedeJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneLukkedeJob = antalEksterneLukkedeJob
	else
	antalEksterneLukkedeJob = 0
	end if
	%>
	<br>Heraf er <b><%=antalEksterneLukkedeJob%> Lukkede</b>.
	
	<%
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 2"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksternePassiveJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksternePassiveJob = antalEksternePassiveJob
	else
	antalEksternePassiveJob = 0
	end if
	%>
	<br>Heraf er <b><%=antalEksternePassiveJob%> Passive</b>.
	
	<br>&nbsp;
	</td></tr>
	</table></div>


<%
divmedLeft = 565
divmedTop = 172
if showdiv <> 1 then
divmedDisp = ""
divmedVz = "visible"
else
divmedDisp = "none"
divmedVz = "hidden"
end if
%>
<div id="medarbtop" name="medarbtop" style="position:absolute; left:<%=divmedLeft%>; top:<%=divmedTop%>; visibility:<%=divmedVz%>; display:<%=divmedDisp%>; z-index:500;"> 	
<table cellspacing="0" cellpadding="0" border="0" width="350">
<tr bgcolor="#EFF3FF">
		<td valign="top" colspan=4 style="border-top:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="5" alt="" border="0"></td>
</tr>
<tr bgcolor="#EFF3FF">
		<td colspan=4 style="border-right:1px #003399 solid; padding-left:1;"><font size=3><b>(B)</b></font>&nbsp;<b>Vælg medarbejdere.</b></td>
</tr>
<tr bgcolor="#EFF3FF">
	<td valign="top"><img src="../ill/blank.gif" width="1" height="40" alt="" border="0"></td>
	<td colspan=2 valign=top style="padding-top:10; padding-left:75;"><a href="#" name="CheckAll" onClick="checkAll(document.statselector.FM_medarb, 2)"><img src="../ill/alle.gif" border="0"></a>
	&nbsp;&nbsp;<a href="#" name="UnCheckAll" onClick="uncheckAll(document.statselector.FM_medarb, 2)"><img src="../ill/ingen.gif" border="0"></a></td>
	<td valign="top" align="right" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="40" alt="" border="0"></td>
</tr>
<tr bgcolor="#EFF3FF">
		<td valign="top" colspan=4 style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"></td>
	</tr>
</table>
</div>
	
<div id="medarb" name="medarb" style="position:absolute; left:<%=divmedLeft+2%>; top:<%=divmedTop+72%>; width:348; height:169; visibility:<%=divmedVz%>; display:<%=divmedDisp%>; overflow:auto; border:1px #003399 solid; background-color:#FFFFFF; z-index:500;"> 	
<table cellspacing="0" cellpadding="0" border="0" width="320">
<tr bgcolor="#FFFFFF">
		<td valign="top" colspan=5><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"></td>
</tr>


<%
	
	antalMed = 0
	medId_all = "0#"
	
	strSQL = "SELECT Mid, Mnavn, Mnr, ProjektgruppeId, MedarbejderId, mansat FROM progrupperelationer, medarbejdere WHERE ProjektgruppeId = "& varProgruppe &" AND Mid = MedarbejderId "& kunaktiveKri &" ORDER BY Mnavn"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr bgcolor="#FFFFFF">
	<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<%
		if instr(selmed, "#"&oRec("mid")&"#") > 0 then
		thisChecked = "CHECKED"
		else
		thisChecked = ""
		end if
		%>
	<td width=50><input type="checkbox" name="FM_medarb" value="<%=oRec("Mid")%>" <%=thisChecked%> onClick="updateselmed('<%=oRec("mid")%>');"><input type="hidden" name="FM_medarb_alle" value="<%=oRec("Mid")%>"></td>
	<td width=220><%=oRec("Mnavn")%></td>
	
	<%
	if oRec("mansat") = "2" then
	Response.write "<td bgcolor=#c4c4c4 style='padding-left:3px; padding-right:3px;'>Deaktiveret!"
	Else
	Response.write "<td>&nbsp;"
	end if
	%>
	</td>
	<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<%
	medId_all = medId_all &",#"&oRec("mid")&"#"
	antalMed = antalMed + 1
	oRec.movenext
	wend
	oRec.close
	
	if antalMed <= 1 then%>
	<input type="hidden" id="FM_medarb" value="0">
	<%end if%>
	<input type="hidden" name="FM_med_all" id="FM_med_all" value="<%=medId_all%>">
	
	<%if antalMed = 0 then%>
	<tr><td style="padding:20;" colspan=4><br><img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0">Ingen medarbejdere matcher den valgte Projektgruppe.</td></tr>
	<%end if%>
	
	</table>
	</div>
	

	
<%
divperLeft = 200
divperTop = 150
divperDisp = "none"
divperVz = "hidden"
%>	
	
	
	<div id="periode" name="periode" style="position:absolute; left:<%=divperLeft%>; top:<%=divperTop%>; visibility:<%=divperVz%>; display:<%=divperDisp%>;"> 	
	<table cellspacing="0" cellpadding="0" border="0" width=350>
	<tr bgcolor="#5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
			<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="334" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
			<td colspan=2 class="alt"><b>Vælg periode</b></td>
	</tr>
	<tr bgcolor="#EFF3FF">
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="158" alt="" border="0"></td>
	<td colspan="2" valign="top"><br><b>Vælg Periode:</b><br><br><!--#include file="inc/weekselector_s.asp"--></td>
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="158" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top" colspan=4 style="border-bottom:1px #003399 solid; border-right:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"></td>
	</tr>
	</table>
	</div>
	
	
<%
if show <> "stat_pies" then
	if showdiv = 1 then
	divsubLeft = 700
	divsubTop = 558
	else
	divsubLeft = 700
	divsubTop = 430
	end if
divsubDisp = ""
divsubVz = "hidden"
else
divsubLeft = 700
divsubTop = 430
divsubDisp = ""
divsubVz = "visible" 
end if
%>	
	
	<input type="hidden" name="selstat" id="selstat" value="<%=show%>">
	<div id="subm" name="subm" style="position:absolute; left:<%=divsubLeft%>; top:<%=divsubTop%>; visibility:<%=divsubVz%>; display:<%=divsubDisp%>; z-index:100;"> 	
	<input type="image" src="../ill/statpil.gif" alt="Kør Statistik!">
	</form>
	</div>
	
	<div id="subm_ghost" name="subm_ghost" style="position:absolute; left:<%=divsubLeft%>; top:<%=divsubTop%>; visibility:visible; display:; z-index:50;"> 	
	<img src="../ill/kor_stat_ghost.gif" width="116" height="30" alt="" border="0">
	</div>
	<!-- slut main form -->
	
<%
if showdiv <> 1 then
divjobkriVz = "hidden"
divjobkriDisp = "none"
else
divjobkriVz = "visible"
divjobkriDisp = ""
end if
%>	

<div id="jobkri" name="jobkri" style="position:absolute; left:200; top:150; visibility:<%=divjobkriVz%>; display:<%=divjobkriDisp%>; z-index:100;">
<table cellspacing="0" cellpadding="0" border="0" width="366">
<form action="stat.asp?menu=stat&showdiv=1&show=<%=show%>&FM_progrupper=<%=varProgruppe%>" method="post" id=kundeogjob name=kundeogjob>
<input type="hidden" name="selmed" id="selmed" value="<%=selmed%>">
<input type="hidden" name="FM_viskunaktivemedarb" id="FM_viskunaktivemedarb" value="<%=visallemed%>">

<input type="hidden" name="osigt" id="osigt" value="1">
<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td valign="top" colspan=2><img src="../ill/tabel_top.gif" width="350" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
	<td class=alt colspan=2><b>Kunde og job kriterier.</b></td>
</tr>
<tr bgcolor="#EFF3FF">
	<td valign="top" rowspan=8 bgcolor="#EFF3FF" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td colspan=2 height=20 style="padding-top:5; padding-left:4;"><b><font size=3>(A)</font> Søg / Filtrer.</b><br>
	Inden der vælges job i pkt (B), kan du søge efter et specifikt job (1) eller filtrere på alle job (2). <br> &nbsp;</td>
	<td valign="top" rowspan=8 bgcolor="#EFF3FF" align="right" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
</tr>
<tr bgcolor="#EFF3FF"><!-- #EFF3FF -->
		<td colspan="2" style="padding-top:5; padding-left:4; border-left:1px #cccccc solid; border-right:1px #cccccc solid; border-bottom:1px #cccccc solid; border-top:1px #cccccc solid;"><b>1)</b> Søg på <b>jobnavn</b> eller <b>jobnr.</b><br>
		<input type="text" name="jobsog" id="jobsog" value="<%=useJobsog%>" style="font-size : 9px; width:228px;"><br><br>&nbsp;</td>
</tr>
<tr bgcolor="#EFF3FF">
	<td colspan=2 height=10>&nbsp;</td>
</tr>
<tr bgcolor="#EFF3FF">
		<td colspan=2 style="padding:4; border-top:1px #cccccc solid;  border-left:1px #cccccc solid;  border-right:1px #cccccc solid;">
		<%
		'**** kundekri baseret på progrp. ***
		if ign <> 1 then
		varFilt2 = " projektgruppe1 = "& varProgruppe &" OR projektgruppe2 = "& varProgruppe &" OR projektgruppe3 = "& varProgruppe &" OR projektgruppe4 = "& varProgruppe &" OR projektgruppe5 = "& varProgruppe &" OR projektgruppe6 = "& varProgruppe &" OR projektgruppe7 = "& varProgruppe &" OR projektgruppe8 = "& varProgruppe &" OR projektgruppe9 = "& varProgruppe &" OR projektgruppe10 = "& varProgruppe &""
		else
		varFilt2 = " id <> 0 "
		end if
		
		strSQL_KK = "SELECT id, jobnr, jobnavn, jobknr, Kkundenavn, kid, fakturerbart, jobstatus, jobslutdato FROM job, kunder WHERE ("& varJobknrKri &" "& varFilt2 &") AND kid = jobknr GROUP BY kkundenavn"
		
		'Response.write strSQL_KK
		oRec2.open strSQL_KK, oConn, 3
		strFM_Kunder_kri = "("
		while not oRec2.EOF
		'** Opbygger kundelisten til kunde og job kri baseret på den valgte projektgruppe. **
		strFM_Kunder_kri = strFM_Kunder_kri &" Kid = "& oRec2("kid")&" OR "
		oRec2.movenext
		wend
		oRec2.close
		
		strFM_Kunder_kri = strFM_Kunder_kri & " Kid = 0)"
		
		%>
		<b>2)</b> Eller vælg alle job tilhørende en bestemt <b>kunde</b>.<br>
		<select name="FM_kunde" size="1" style="font-size : 9px; width:185 px;" onChange="renssog();">
		<option value="0">Alle kunder</option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' AND "& strFM_Kunder_kri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(useKunde) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=left(oRec("Kkundenavn"), 30)%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select><br>&nbsp;</td>
		</tr>
		
		<tr bgcolor="#EFF3FF">
		<td colspan=2 style="padding:4; border-left:1px #cccccc solid;  border-right:1px #cccccc solid;">
		<!--
		<b>Filter kriterier.</b><br>
		
		<b>Interne / Eksterne job:</b><br>
		<
		if cint(useAlle_eksterne) = 1 then
		chk1 = "CHECKED"
		chk2 = ""
		else
		chk1 = ""
		chk2 = "CHECKED"
		end if
		%>
		Vis alle job <img src='../ill/internt_job_trans.gif' width='15' height='15' alt='' border='0'> <img src='../ill/eksternt_job_trans.gif' width='15' height='15' alt='' border='0'> 
		<input type="radio" name="alle_eksterne" value="1" <=chk1%>>&nbsp;&nbsp; Kun eksterne job <img src='../ill/eksternt_job_trans.gif' width='15' height='15' alt='' border='0'> <input type="radio" name="alle_eksterne" value="2" <=chk2%>>
		-->
		<input type="hidden" name="alle_eksterne" id="alle_eksterne" value="2">
		</td>
	</tr>
	
	<tr bgcolor="#EFF3FF">
	<%
		if cint(useStatus) = 0 then
		chk0 = "CHECKED"
		chk1 = ""
		chk2 = ""
		chk3 = ""
		end if
		
		if cint(useStatus) = 1 then
		chk1 = "CHECKED"
		chk0 = ""
		chk2 = ""
		chk3 = ""
		end if
		
		if cint(useStatus) = 2 then
		chk0 = ""
		chk1 = ""
		chk2 = "CHECKED"
		chk3 = ""
		end if
		
		if cint(useStatus) = 3 then
		chk0 = ""
		chk1 = ""
		chk2 = ""
		chk3 = "CHECKED"
		end if
		%>
		<td colspan=2 style="padding:4; border-left:1px #cccccc solid; border-right:1px #cccccc solid; border-bottom:1px #cccccc solid;"><b>Job status:</b><br>
		<input type="radio" name="status" value="0" <%=chk0%>>Vis alle.<br>
		<input type="radio" name="status" value="1" <%=chk1%>><img src='../ill/status_groen.gif' width='14' height='19' alt='' border='0'> Aktive.<br>
		<input type="radio" name="status" value="2" <%=chk2%>><img src='../ill/status_roed.gif' width='14' height='19' alt='' border='0'> Lukkede.<br>
		<input type="radio" name="status" value="3" <%=chk3%>><img src='../ill/status_graa.gif' width='14' height='20' alt='' border='0'> Passive.<br>
		</td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td colspan=2>
		<%if ign = 2 then
		chkign = "CHECKED"
		else
		chkign = ""
		end if%>
		<input type="checkbox" name="FM_ignorer" value="2" <%=chkign%>>Vis kun de job (og kunder) som den valgte projektgruppe (valgt under medarbejdere i step 1) har adgang til. <!--Ignorer de valgte projektgrupper's rettigheder og vis alle job.--><br>
		</td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td align=center height=31 colspan=2><input type="image" src="../ill/brug-filter.gif"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td valign="bottom" colspan=2><img src="../ill/tabel_top.gif" width=350" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</form>
</table>
</div>	
	
 
 
<%
divprgLeft = 200
divprgTop = 150
if showdiv = 1 then
divprgDisp = "none"
divprgVz = "hidden"
else
divprgDisp = ""
divprgVz = "visible"
end if
%>	

<div id="progrp" name="progrp" style="position:absolute; left:<%=divprgLeft%>; top:<%=divprgTop%>; visibility:<%=divprgVz%>; display:<%=divprgDisp%>; z-index:100;"> 	
<table cellspacing="0" cellpadding="0" border="0" width="366">
<form action="stat.asp?menu=stat&showdiv=2&show=<%=show%>" method="post" id="progrp" name="progrp">
<input type="hidden" name="seljob" id="seljob" value="<%=seljob%>">
<input type="hidden" name="FM_kunde" id="prgFM_kunde" value="<%=useKunde%>">
<input type="hidden" name="status" id="prgstatus" value="<%=useStatus%>">
<!--<input type="hidden" name="jobsog" id="prgjobsog" value="<=useJobsog%>">-->
<input type="hidden" name="alle_eksterne" id="prgalle_eksterne" value="<%=useAlle_eksterne%>">
<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td valign="top" colspan=2><img src="../ill/tabel_top.gif" width="350" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
	<td colspan=2 class=alt><b>Vælg projektgruppe.</b></td>
</tr>
<tr bgcolor="#EFF3FF">
		<td valign="top" style="border-left:1px #003399 solid;" height=40><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td height="100" valign="top" colspan=2 style="padding-top:5;"><font size=3><b>(A)</b></font><br>
		Hvis du ønsker at få vist medarbejdere fra en speciel projektgruppe kan denne vælges nedenfor. Herefter vælges medarbejdere (B).<br><br><select name="FM_progrupper" size="8" style="font-size : 11px; width:200;">
	<%
			strSQL = "SELECT projektgrupper.id AS id, navn FROM projektgrupper ORDER BY navn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request("FM_progrupper")) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	<td valign="top" align="right" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
	<td valign="top" style="border-left:1px #003399 solid;" height=20><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td colspan=2 valign=bottom style="padding-top:10;">
	<input type="radio" name="FM_viskunaktivemedarb" id="FM_viskunaktivemedarb" value="1" <%=vamChk1%>> Vis kun aktive medarbejdere.<br>
	<input type="radio" name="FM_viskunaktivemedarb" id="FM_viskunaktivemedarb" value="2" <%=vamChk2%>> Vis alle medarbejdere.<br><br>
	<br><br>
	<input type="image" src="../ill/brug-filter.gif" onClick="renssog();"></td>
	<td valign="top" align="right" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td valign="bottom" colspan=2><img src="../ill/tabel_top.gif" width="350" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
</table>
</form>
</div>	 
 

<%
'now close and clean up
	
	oConn.close
	Set oRec = Nothing
	set oConn = Nothing 

end if

%>
<!--#include file="../inc/regular/footer_inc.asp"-->
