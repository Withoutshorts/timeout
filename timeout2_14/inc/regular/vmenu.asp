<div id="vmenu_usr" style="position:absolute; left:10; top:60; visibility:visible;">
<table cellpadding="0" cellspacing="0" border="0">
<tr>
<%'************* TSA/CRM menu *********************
if menu = "crm" then
showmenu = "crm"
else
showmenu = request("mainmenu")
end if

select case showmenu
case "crm"
%>
<td><img src="../ill/v-menu_menu.gif" alt="" border="0"></td>
</tr>
</table>	
<table cellspacing="0" cellpadding="2" border="0" bgcolor="#EFF3FF" width="159">
<tr>
<td align="right"><%=varSelectedCRM_kal%></td>
<td><a href="crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal&status=0&id=0&emner=0" class='vmenu'>Kalender</a></td>
</tr>
<tr>
<td align="right"><%=varSelectedCRM_osigt%></td>
<td><a href="kunder.asp?menu=crm&shokselector=1&ketype=e&selpkt=osigt" class="vmenu">Kontakter</a></td>
</tr>
<!--
<tr>
<td align="right"><=varSelectedCRM_oprfirma%></td>
<td><a href='kunder.asp?menu=crm&ketype=e&func=opret&id=0&selpkt=oprfirma' class='vmenu'>Opret nyt firma&nbsp;<img src='../ill/pil_groen_lille.gif' width='17' height='17' alt='' border='0'></a></td>
</tr>-->
<td align="right"><%=varSelectedCRM_hist%></td>
<td><a href='crmhistorik.asp?menu=crm&ketype=e&func=hist&id=0&selpkt=hist' class='vmenu'>Aktions historik</a></td>
</tr>


<%if level <= 3 OR level = 6 then%>
<tr>
	<td align="right">&nbsp;</td>
	<td><a href='medarb.asp?menu=medarb' class='vmenualt'>Medarbejdere</a></td>
</tr>
<%end if%>
<%case else%>
<td><img src="../ill/v-menu_menu.gif" alt="" border="0"></td>
</tr>
</table>	

<%


	treg0206thisMidVmenu = 1
	strSQL = "SELECT timereg FROM medarbejdere WHERE mid = "& session("mid")
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	treg0206thisMidVmenu = oRec("timereg") 
	end if
	oRec.close 
	


if cint(treg0206thisMidVmenu) = 0 then
	tregLinkvmenu = "<a href=""timereg.asp?menu=timereg"" target=""_top"" class=""vmenu"">Timeregistrering</a>"
else
	tregLinkvmenu = "<a href=""timereg_2006_fs.asp"" target=""_top"" class=""vmenu"">Timeregistrering</a>"
end if

%>

<table cellspacing="0" cellpadding="2" border="0" bgcolor="#EFF3FF" width="159">
<tr>
	<td align="right"><%=varSelectedTimereg%></td>
	<td><%=tregLinkvmenu%></td>
</tr>

<%'if lto = "gp" OR lto = "bowe" OR (lto = "titoonic" AND session("mid") = 21) OR lto = "dencker" OR lto = "outz" OR lto = "perspektivait" OR (lto = "kringit" AND session("mid") = 11) OR (lto = "demo") OR (lto = "intranet") OR (lto = "execon" AND session("mid") = 36 OR session("mid") = 13) OR (lto = "external") OR (lto = "skybrud" AND session("mid") = 1) then%>
<!--<tr>
	<td align="right"><=varSelectedTimereg%></td>
	<td><a href="timereg_2006_fs.asp" class='rmenu'>Timereg. ver. 2006 BETA</a></td>
</tr>-->
<%'end if%>

<%if level <= 3 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedwebblik%></td>
	<td><a href="webblik_joblisten.asp" class='vmenu'>TimeOut idag</a></td>
</tr>
<%end if%>

<%if level <= 3 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedJobs%></td>
	<td><a href="jobs.asp?menu=job&shokselector=1&fromvemenu=j" class='vmenu'>Jobs</a></td>
</tr>
<%end if%>
<%if level <= 3 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedRes%></td>
	<td><a href="ressource_forside.asp?menu=res" class='vmenu'>Ressourcer</a></td>
</tr>
<%end if%>
<%if level <= 3 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedKund%></td>
	<td><a href="kunder.asp?menu=kund&visikkekunder=1" class='vmenu'>Kontakter</a></td>
</tr>
<%end if%>
<tr>
	<td align="right"><%=varSelectedMed%></td>
	<td><a href="medarb.asp?menu=medarb" class='vmenu'>Medarbejdere</a></td>
</tr>
<%if level <= 2 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedStat%></td>
	<td><a href="joblog_timetotaler.asp" class='vmenu'>Statistik</a></td>
</tr>
<%end if%>

<%if (level <= 2 OR level = 6) AND strE5 = "on" then%>
<!--<tr>
	<td align="right"><=varSelectedFak%></td>
	<td><a href="stat_fak.asp?menu=stat_fak&FM_kunde=0&shokselector=1" class='vmenu'>Fakturering</a></td>
</tr>-->
<tr>
	<td align="right"><%=varSelectedKon%></td>
	<td><a href="stat_fak.asp?menu=stat_fak&FM_kunde=0&shokselector=1" class='vmenu'>Bogføring og fakturering</a></td>
</tr>
<%end if%>

<%
if level <= 3 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedMat%></td>
	<td><a href="materialer.asp?menu=mat" class='vmenu'>Material. og Produktion.</a></td>
</tr>
<%end if
%>
<%end select%>
</table>
</div>



<%
'*** Er smiley aktiv?? ***
if global_inc = "j" then '** Er global_inc læst?
	call ersmileyaktiv()
else
	smilaktiv = 0
	strSQL = "SELECT smiley FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	smilaktiv = oRec("smiley") 
	end if
	oRec.close
end if



'************ TSA submenu **************************
select case menu 
case "medarb"
varTopGif = "medarb"
	if level <= 2 OR level = 6 then 
	Links = 2
	else
	Links = 0
	end if
Redim link(Links)
	
	link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='medarb.asp?menu=medarb' class='vmenu'>Medarbejdere</a></td>"
	if level <= 2 OR level = 6 then 
	link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='bgrupper.asp?menu=medarb' class='vmenu'>Brugergrupper</a></td>"
	link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='medarbtyper.asp?menu=medarb' class='vmenu'>Medarbejdertyper</a></td>"
	end if


case "kund"
varTopGif = "kund" 

if lto = "ravnit" OR lto = "intranet" then
Links = 5
link3kunder = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='signoff.asp?menu=kund' class='vmenu'>Sign-Off</a></td>"
else
Links = 4
link3kunder = "<td align='right'><img src='../ill/blank.gif' alt='' border='0'></td><td>&nbsp;</td>"
end if

Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='kunder.asp?menu=kund' class='vmenu'>Kontakter</a></td>"
link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='kunder.asp?menu=kund&func=opret' class='vmenu'>Opret ny Kontakt</a></td>"
link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='serviceaft_osigt.asp?menu=kund&id=0&func=osigtall' class='vmenu'>Aftaleoversigt</a></td>"
link(3) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='infobase.asp?menu=kund&id=0' class='vmenu'>Infobase</a></td>"
link(4) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='filer.asp?kundeid=0&jobid=0' class='vmenu'>Filarkiv</a></td>"


if lto = "ravnit" OR lto = "intranet" then
link(5) = link3kunder
end if



case "job"
varTopGif = "jobs" 
Links = 6
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='jobs.asp?menu=job&shokselector=1&fromvemenu=j' class='vmenu'>Joboversigt</a></td>"
link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='jobs.asp?menu=job&func=opret&id=0&int=1' class='vmenu'>Opret nyt job</a></td>"
link(2) = "<td colspan=2 style=""padding-left:5px;"">&nbsp;</td>"
link(3) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='projektgrupper.asp?menu=job' class='vmenu'>Projektgrupper</a></td>"
link(4) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='akt_gruppe.asp?menu=job&func=favorit' class='vmenu'>Stam-aktiviteter og Grp.</a></td>"
link(5) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='pipeline.asp?menu=job&FM_kunde=0&FM_progrupper=10' class='vmenu'>Pipeline</a></td>"
link(6) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='pg_allokering.asp?menu=job' class='vmenu'>Jobplanner (Årsoversigt)</a></td>"

'link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='jobs.asp?menu=job&func=opret&id=0&int=0' class='vmenu'>Opret nyt <img src='../ill/internt_job_trans.gif' width='15' height='15' alt='' border='0'> job&nbsp;</a></td>"

case "res"
varTopGif = "res" 
Links = 4
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='ressource_forside.asp?menu=res' class='vmenu'>Ressource Overblik</a></td>"
link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='jbpla_w.asp?menu=res' class='vmenu'>Ressource Kalender</a></td>"
link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='ressource_belaeg_jbpla.asp' class='vmenu'>Ressource Timer</a></td>"
link(3) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='ressource_belaeg.asp?menu=res' class='vmenu'>Ressource Belægning</a></td>"
link(4) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='jbpla_k.asp?menu=res' class='vmenu'>Job belastning</a></td>"



case "mat"
varTopGif = "mat" 
Links = 4
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='materialer.asp?menu=mat' class='vmenu'>Materialer</a></td>"
link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='materiale_grp.asp?menu=mat' class='vmenu'>Materialegrupper</a></td>"
link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='produktioner.asp?menu=mat' class='vmenu'>Standard produktioner</a></td>"
link(3) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='leverand.asp?menu=mat' class='vmenu'>Leverand. og Servicepart.</a></td>"


case "stat"
varTopGif = "stat" 

Links = 11
Redim link(Links)
'link(0) = "<td align='right'><img src='../ill/sum.gif' alt='Joblog' border='0' width='23' height='22'></td><td><a href='stat.asp?menu=stat&show=joblog_z_b' class='vmenu'>Timefordeling</a></td>"
link(0) = "<td align='right'><img src='../ill/sum.gif' alt='Joblog' border='0' width='23' height='22'></td><td><a href='joblog_timetotaler.asp' class='vmenu'>Timeforbrug - Grandtotal</a>&nbsp;<font color='gold'><b>NY!</b></font></td>"
link(1) = "<td align='right'><img src='../ill/status.gif' alt='Status' border='0' width='23' height='22'></td><td><a href='stat.asp?menu=stat&show=joblog_status' class='vmenu'>Nøgletal og Status</a></td>"
link(2) = "<td align='right'><img src='../ill/joblog.gif' alt='Joblog' border='0' width='23' height='22'></td><td><a href='stat.asp?menu=stat&show=joblog' class='vmenu'>Joblog</a></td>"
link(3) = "<td align='right'><img src='../ill/korsel.gif' alt='Kørsel' border='0' width='23' height='22'></td><td><a href='stat.asp?menu=stat&show=joblog_korsel' class='vmenu'>Kørsel</a></td>"
link(4) = "<td align='right'><img src='../ill/oms.gif' width='23' height='22' alt='' border='0'></td><td><a href='oms.asp?menu=stat' class='vmenu'>Omsætning</a></td>"
link(5) = "<td align='right'><img src='../ill/pie_ikon.gif' width='23' height='22' alt='' border='0'></td><td><a href='stat.asp?menu=stat&show=stat_pies' class='vmenu'>%-vis timefordeling</a></td>"
link(6) = "<td align='right'><img src='../ill/word.gif' width='23' height='22' alt='' border='0'></td><td><a href='stat.asp?menu=stat&show=word' class='vmenu'>Eksport (Word, Excel)</a></td>"
link(7) = "<td align='right'><img src='../ill/ikon_fomr.gif' width='23' height='22' alt='' border='0'></td><td><a href='fomr.asp?menu=stat&func=stat' class='vmenu'>%-fordeling på<br>Forretningsområder</a></td>"

link(8) = "<td align='right'><img src='../ill/ikon_stempelur.gif' width='23' height='22' alt='' border='0'></td><td><a href='stempelur.asp?menu=stat&func=stat' class='vmenu'>Logind historik (Stempelur)</a></td>"


if cint(smilaktiv) = 1 then
	link(9) = "<td align='right'><img src='../ill/ikon_smil.gif' width='23' height='22' alt='' border='0'></td><td><a href='smileystatus.asp?menu=stat' class='vmenu'>Smileys</a></td>"
else
	link(9) = ""
end if


link(10) = "<td align='right'><img src='../ill/ikon_mat.gif' width='23' height='22' alt='' border='0'></td><td><a href='materiale_stat.asp?menu=stat' class='vmenu'>Materialeforbrug.</a></td>"




'case "stat_fak"
'varTopGif = "fak" 
'Links = 0
'Redim link(Links)
'	if strE5 = "on" then
'	link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='stat_fak.asp?menu=stat_fak&FM_kunde=0&shokselector=1' class='vmenu'>Fakturerings Kørsel</font></a></td>"
'	else
'	link(0) = "<td align='right'>&nbsp;</td>"
'	end if

case "kon", "stat_fak"
varTopGif = "fak" 
Links = 10
Redim link(Links)
	if strE5 = "on" then
	link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='posteringer.asp?menu=kon&id=0' class='vmenu'>Kassekladde</a></td>"
	link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='kontoplan.asp?menu=kon' class='vmenu'>Kontoplan</a></td>"
	link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='momskoder.asp?menu=kon' class='vmenu'>Momskoder</a></td>"
	link(3) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='nogletalskoder.asp?menu=kon' class='vmenu'>Nøgletalskoder</a></td>"
	link(4) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='posteringer.asp?menu=kon&id=0&kontonr=-2&kid=0' class='vmenu'>U-moms</a></td>"
	link(5) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='posteringer.asp?menu=kon&id=0&kontonr=-1&kid=0' class='vmenu'>I-moms</a></td>"
	
	link(6) = "<td align='right'><br><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><br><a href='stat_fak.asp?menu=stat_fak&FM_kunde=0&shokselector=1' class='vmenu'>Fakturering Job</a></td>"
	'link(7) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href=Javascript:NewWin('joblog_afstem.asp?menu=stat&show=joblog_afstem') target='_self' class='vmenu'>Afstemningsoversigt</a></td>"
	link(7) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='joblog_afstem.asp?menu=kon&show=joblog_afstem' class='vmenu'>Afstemnings o.sigt (Job)</a></td>"
	link(8) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='fak_serviceaft_osigt.asp?menu=stat_fak&func=osigtall' class=vmenu>Fakturering Aftaler</a></td>"
	link(9) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='fak_serviceaft_saldo.asp?menu=kon&aftid=-1&kundeid=0' class=vmenu>Afstemning (Aftaler)</a></td>"
	
	link(10) = "<td align='right' valign=top><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='#' class=vmenualt>Søg i fakturaer</a></td>"
	
	
	
	
	else
	link(0) = "<td align='right'>&nbsp;</td>"
	link(1) = "<td align='right'>&nbsp;</td>"
	link(2) = "<td align='right'>&nbsp;</td>"
	link(3) = "<td align='right'>&nbsp;</td>"
	link(4) = "<td align='right'>&nbsp;</td>"
	link(5) = "<td align='right'>&nbsp;</td>"
	link(6) = "<td align='right'>&nbsp;</td>"
	link(7) = "<td align='right'>&nbsp;</td>"
	link(8) = "<td align='right'>&nbsp;</td>"
	link(9) = "<td align='right'>&nbsp;</td>"
	end if	
	
case "crm"
varTopGif = "blank" 
Links = 0
Redim link(Links)
case else
varTopGif = "timereg" 
Links = 4
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='timereg.asp?menu=tiemreg' class='vmenu'>Timeregistrering</a></td>"
link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='joblog.asp?menu=timereg&FM_medarb="&session("mid")&"&FM_job=0&selmedarb="&session("mid")&"' class='vmenu'>Medarbejder Log</a></td>"
link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='joblog_korsel.asp?menu=timereg&FM_medarb="&session("mid")&"&FM_job=0&selmedarb="&session("mid")&"' class='vmenu'>Medarbejder Kørsel</a></td>"
link(3) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='timereg.asp?func=showguide&FM_use_me="&usemrn&"' class='vmenu'>Guiden dine aktive job&nbsp;<font color='darkred'>!</font></a></td>"
link(4) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href=""#"" class=vmenu onClick=""showrestimer()"">Din uge</a></td>"
'link(5) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='stempelur.asp?menu=stat&func=stat&kunmedarb=1' class='vmenu'>Logind historik (Stempelur)</a></td>"


%>
<%end select%>


<div id="vmenudiv" style="position:absolute; left:10; top:270; width:130; height:300; visibility:visible;">
<%if menu = "crm" then
Response.write "<br><br>"
end if%>
<table cellspacing="0" cellpadding="0" border="0" width="159">
<tr>
	<td colspan="2"><img src="../ill/<%=varTopGif%>_top_blaa.gif" alt="" border="0"></td>
</tr>
</table>
<!--</table>-->	
<table cellspacing="0" cellpadding="2" border="0" bgcolor="#EFF3FF" width="159">
<%for intCounter = 0 to Links%>
<tr bgcolor="EFF3FF">
<%=link(intCounter)%>
</tr>
<%next%>

</table>
<br>
<%
'** slut submenu ****


'**************************************************
'* Sætter variable på kunder DD i 
'* vemnu alt efter hvilken fil man er inde på. 
'**************************************************

select case thisfile
'case "timereg"
'ketypeKri = " ketype <> 'e' "
'actionLine = "timereg.asp?menu=timereg"
'name1 = "searchstring"
'imgsearch = "../ill/v-menu_soeg.gif"
'writethis = "(Dine aktive job)"
'hiddenfield1 = "<input type='hidden' name='FM_use_me' value='"& request("FM_use_me") &"'>"
'	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
'	hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_dag' value='"&request("FM_start_dag")&"'>"
'	hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_mrd' value='"&request("FM_start_mrd")&"'>"
'	hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_aar' value='"&request("FM_start_aar")&"'>"
'	else
'	hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_dag' value='"&request("strdag")&"'>"
'	hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_mrd' value='"&request("strmrd")&"'>"
'	hiddenfield1 = hiddenfield1 & "<input type='hidden' name='FM_start_aar' value='"&request("straar")&"'>"
'	end if
case "jobs"
ketypeKri = " ketype <> 'e' "
actionLine = "jobs.asp?menu=job&sort=kunkunde&shokselector=1"
name1 = "FM_kunde"
imgsearch = "../ill/v-menu_soeg.gif"
writethis = ""
hiddenfield1 = ""
'case "stat"
'ketypeKri = " ketype <> 'e' "
'actionLine = "stat.asp?menu=stat&sort=kunkunde&shokselector=1"
'name1 = "FM_kunde"
'imgsearch = "../ill/v-menu_soeg.gif"
'writethis = ""
'hiddenfield1 = ""
case "stat_fak"
ketypeKri = " ketype <> 'e' "
actionLine = "stat_fak.asp?menu=stat_fak&sort=kunkunde&shokselector=1"
name1 = "FM_kunde"
imgsearch = "../ill/v-menu_soeg.gif"
writethis = ""
hiddenfield1 = ""
case "crmkalender"
ketypeKri = " ketype <> 'k' "
actionLine = "crmkalender.asp?menu=crm&selpkt=kal"
actionline = actionline &"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&shownumofdays="&shownumofdays&"&FM_start_aar="&strAar&""
name1 = "id"
imgsearch = "../ill/v-menu_soeg_f.gif"
writethis = ""
hiddenfield1 = ""
case "infobase"
ketypeKri = " ketype <> 'e' "
actionLine = "infobase.asp?menu=kund"
name1 = "id"
imgsearch = "../ill/v-menu_infobase.gif"
writethis = "(Global)"
hiddenfield1 = ""
case "crmhistorik"
ketypeKri = " ketype <> 'k' "
actionLine = "crmhistorik.asp?menu=crm&func=hist&selpkt=hist"
name1 = "id"
imgsearch = "../ill/v-menu_soeg_f.gif"
writethis = ""
hiddenfield1 = ""
'case "kunder"

'	if menu <> "crm" then
'	ketypeKri = " ketype <> 'e' "
'	actionLine = "kunder.asp?menu=kund" 'midlertidig (bruges ikke endnu) 
'	hiddenfield1 = ""
'	imgsearch = "../ill/v-menu_soeg_f.gif"
'	writethis = ""
'	name1 = "searchstring"
'	else
'	ketypeKri = " ketype <> 'k' "
'	actionLine = "kunder.asp?menu=crm&selpkt=osigt&ketype=e" 
'	name1 = "id"
'	imgsearch = "../ill/v-menu_soeg_f.gif"
'	writethis = ""
'	hiddenfield1 = ""
'	end if

end select




select case thisfile
case "crmkalender", "jobs", "stat_fak", "infobase" '"stat", "timereg", , "crmhistorik"
if (thisfile = "jobs") AND request("shokselector") <> 1 OR thisfile = "crmhistorik" AND (func = "opret" OR func = "red") then 'thisfile = "stat" OR 

else %>
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="120">
	<form action="<%=actionLine%>" method="post">
	<%=hiddenfield1%>
	<tr>
		<td><img src="<%=imgsearch%>" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td height="30" valign="middle" align="center" style="border-right: 1px solid #003399; border-left: 1px solid #003399;">
		<table>
		<tr>
		<td><select name="<%=name1%>" size="1" style="font-size : 9px; width:125 px;">
		<option value="0">Alle <%=writethis%></option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(request(name1)) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=left(oRec("Kkundenavn"), 18)%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select></td>
		<td>&nbsp;</td>
		</tr>
	</table>
	
	
	
</td></tr>
<%
'**** Medarbejder filter (kun CRM) ********
if level <= 2 AND menu = "crm" AND (thisfile <> "infobase") then
%><tr bgcolor="#064CB9">
		<td height="20">&nbsp;&nbsp;<font class="stor-hvid">Vis for medarbejder:</td>
	</tr>
<tr bgcolor="#EFF3FF">
<td height="30" valign="middle" style="border-right: 1px solid #003399; border-left: 1px solid #003399;">	
<table>
<tr>
<td><img src="../ill/blank.gif" alt="" border="0" width="10" height="1">
<select name="medarb" size="1" style="font-size : 11px; width:125 px;">
<option value="0">Alle</option>
<%
		strSQL = "SELECT mnavn, mid FROM medarbejdere ORDER BY mnavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
			if cint(medarb) = cint(oRec("mid")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
		
		%>
		<option value="<%=oRec("mid")%>" <%=isSelected%>><%=left(oRec("mnavn"), 15)%></option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
</select></td>
<td>&nbsp;</td>
</tr>
</table></td>
</tr>
<%end if


'**** Filter efter emne, status og medarbejder (Kun crmhistorik)  ****
if thisfile = "crmhistorik" AND func = "hist" then

%>
<tr bgcolor="#064CB9">
	<td height="20">&nbsp;&nbsp;<font class="stor-hvid">Vis emne:</td>
</tr>
<tr bgcolor="#EFF3FF">
	<td height="30" valign="middle" style="border-right: 1px solid #003399; border-left: 1px solid #003399;">	
	<table>
	<tr>
	<td><img src="../ill/blank.gif" alt="" border="0" width="10" height="1">
	<select name="emner" size="1" style="font-size : 11px; width:125 px;">
	<option value="0">Alle</option>
	<%
			strSQL = "SELECT navn, id FROM crmemne ORDER BY navn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request("emner")) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	<td>&nbsp;</td>
	</tr>
	</table></td>
	</tr>

<tr bgcolor="#064CB9">
	<td height="20">&nbsp;&nbsp;<font class="stor-hvid">Vis status:</td>
</tr>
<tr bgcolor="#EFF3FF">
	<td height="30" valign="middle" style="border-right: 1px solid #003399; border-left: 1px solid #003399;">	
	<table>
	<tr>
	<td><img src="../ill/blank.gif" alt="" border="0" width="10" height="1">
	<select name="status" size="1" style="font-size : 11px; width:125 px;">
	<option value="0">Alle</option>
	<%
			strSQL = "SELECT navn, id FROM crmstatus ORDER BY navn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request("status")) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	<td>&nbsp;</td>
	</tr>
	</table></td>
	</tr>

<%end if


'************************************************************************************
%>
<tr bgcolor="#EFF3FF"><td align="center" style="padding-top:2; border-right: 1px solid #003399; border-left: 1px solid #003399; border-bottom: 1px solid #003399;"><input type="image" src="../ill/brug-filter.gif" border="0">
	</td></form></tr>
</table>
<%
end if 'jobs, stat shockselector
end select%>

</div>

