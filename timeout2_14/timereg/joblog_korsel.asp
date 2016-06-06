<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	if request("FM_medarb") = "" OR request("FM_job") = "" then%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 40
		call showError(errortype)
		else
	%>
	<%if request("print") <> "j" then%>
	<!--#include file="../inc/regular/header_inc.asp"-->	
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%else%>
	<html>
	<head>
		<title>timeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<%end if%>
	<!--#include file="inc/convertDate.asp"-->
	
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	
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
	'linket = request("linket")
	FM_medarb = request("FM_medarb")
	
	selmedarb = FM_medarb 'request("selmedarb")
	if len(selmedarb) = "0" then
	selmedarb = 0
	else
	selmedarb = selmedarb
	end if
	
	'selaktid = request("selaktid")
	
	FM_job = request("FM_job")
	intJobnr = FM_job 'request("jobnr")
	
	'if len(Request("mrd")) <> "0" then
	'	strReqMrd = Request("mrd")
	'else
	'strReqMrd = month(now)
	'end if
	
	'**** finder ud af om der er valgt et år ***
	'if len(Request("year")) <> "0" then
		'if Request("year") = "-1" then
		'strReqAar = "0"
		'else
		'strReqAar = Request("year")
		'end if	
	'else
	'strReqAar = right(year(now), 2)
	'end if
	
	if len( request("FM_start_dag")) <> 0 then
	strMrd = request("FM_start_mrd")
	strDag = request("FM_start_dag")
	strAar = right(request("FM_start_aar"),2) 
	strDag_slut = request("FM_slut_dag")
	strMrd_slut = request("FM_slut_mrd")
	strAar_slut = right(request("FM_slut_aar"),2)
	else
	strMrd_temp = dateadd("m", -1, date)
	strMrd = month(strMrd_temp)
	strDag_temp = dateadd("d", -30, date)
	strDag = day(strDag_temp)
	if strMrd <> 12 then
	strAar= right(year(now), 2)
	else
	strAar_temp = dateadd("yyyy", -1, date)
	strAar = right(year(strAar_temp),2)
	end if
	
	strDag_slut = day(now)
	strMrd_slut = month(now)
	strAar_slut = right(year(now), 2)
	end if
	
	
	'** Indsætter cookie **
	Response.Cookies("st_dag") = strDag
	Response.Cookies("st_dag").Expires = date + 40
	
	Response.Cookies("st_md") = strMrd
	Response.Cookies("st_md").Expires = date + 40
	
	Response.Cookies("st_aar") = strAar
	Response.Cookies("st_aar").Expires = date + 40
	
	
	Response.Cookies("sl_dag") = strDag_slut
	Response.Cookies("sl_dag").Expires = date + 40
	
	Response.Cookies("sl_md") = strMrd_slut
	Response.Cookies("sl_md").Expires = date + 40
	
	Response.Cookies("sl_aar") = strAar_slut
	Response.Cookies("sl_aar").Expires = date + 40
	
	
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	'thisjoblog = "j"
	thisfile = "joblog_korsel"
	if request("print") <> "j" then
	%>
	<!--include file="inc/joblog_z_mrdk.asp"-->
	<!-- mrd knapper -->
	<%end if
		'*** Hvis ikke det er medarbejder log vises stat submenu *******
		if menu <> "timereg" then%>
		<!--include file="inc/stat_submenu.asp"-->
			<%if request("print") <> "j" then%>
			<div style="position:absolute; left:680; top:81;">
			<a href="javascript:NewWin('joblog_korsel.asp?menu=stat&print=j&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>')" target="_self" class='rmenu'>&nbsp;Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
			</div>
			<%else%>
			<table cellspacing="0" cellpadding="0" border="0" width="880">
				<tr>
					<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
					<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
				</tr>
			</table>
			<%end if%>
		<%end if
	if request("print") <> "j" then
	pleft = 190
	ptop = 55
	else
	pleft = 20
	ptop = 10
	end if
	%>	
	<div id="sindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible; z-index:100;">
	<br><br>
	<%if request("print") = "j" then%>
	<br><font class="pageheader"><b>Joblog Kørsel</b> <%=formatdatetime(strDag&"/"&strMrd&"/"&strAar,1)%> - <%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut,1)%> </font><br>
	<%else%>
	<font class="pageheader"><b>Joblog Kørsel</b></font><br>
	<%end if%>
	
<!-------------------------------Sideindhold------------------------------------->
<%if request("print") <> "j" then%>
<br>
<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="625">
<tr bgcolor="#5582D2">
	<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
	<td colspan=4 valign="top"><img src="../ill/tabel_top.gif" width="609" height="1" alt="" border="0"></td>
	<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
	<td colspan=4 valign="top" class="alt">&nbsp;<b>Job og periode:</b></td>
</tr>	
<tr><form action="joblog_korsel.asp?menu=<%=menu%>" method="post">
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
	<td><input type="hidden" name="FM_medarb" value="<%=FM_medarb%>">
	<input type="hidden" name="FM_job" value="<%=FM_job%>"></td>
	<!--#include file="inc/weekselector_b.asp"-->
	<td><input type="image" src="../ill/statpil.gif" vspace=5></td>
	<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
</tr>
</form>
<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=4 valign="bottom"><img src="../ill/tabel_top.gif" width="609" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
</tr>
</table>
<%end if%>

<br><br>
<table border="0" width=625 cellpadding="0" cellspacing="0" bgcolor="#ffffff">
<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan=6 valign="top"><img src="../ill/tabel_top.gif" width="608" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="31" alt="" border="0"></td>
</tr>
<tr height="20" bgcolor="#5582D2">
	<td class='alt' width="100">&nbsp;&nbsp;Uge&nbsp;&nbsp;&nbsp;Dato</td>
	<td class='alt' width="300">&nbsp;&nbsp;Aktivitet</td>
	<td class='alt' width="40">Km</td>
	<td width="120" class='alt'>Medarb.</td>
	<td align="center" class='alt' width="70">Taste dato</td>
	<td align="center" width="5">&nbsp;</td>
</tr>
<%
'*** Finder de medarbejdere og job der er valgt ***
if menu <> "timereg" then
Dim intMedArbVal 
Dim b
Dim intJobnrKriValues 
Dim i
			
			ugeAflsMidKri = " u.mid = 0 "
			medarbKri = "( Tmnr = 0 "
			intMedArbVal = Split(selmedarb, ", ")
			For b = 0 to Ubound(intMedArbVal)
				
				if selmedarb <> "0" then
				medarbKri = medarbKri & " OR Tmnr = " & intMedArbVal(b)
				ugeAflsMidKri = ugeAflsMidKri & " OR u.mid = " & intMedArbVal(b)
				 
				else
				medarbKri = "( Tmnr <> 0 "
				ugeAflsMidKri = " u.mid <> 0 "
				end if
				
			next
			medarbKri = medarbKri & ") "
			ugeAflsMidKri = "("& ugeAflsMidKri &")"
			
			jobKri = "( Tjobnr = '0' "
			intJobnrKriValues = Split(intJobnr, ", ")
				For i = 0 to Ubound(intJobnrKriValues)
				if intJobnr <> "0" then
				jobKri = jobKri & " OR Tjobnr = '" & intJobnrKriValues(i) & "'"
				else
				jobKri = "( Tjobnr <> '0' "
				end if
				next
			jobKri = jobKri & ") "
			
			
else
medarbKri = "(Tmnr = " & selmedarb &") "
jobKri = "( Tjobnr <> '0' ) "
ugeAflsMidKri = " (u.mid = "& selmedarb &")"
end if


'if len(request("selaktid")) <> "0" then
'selaktidKri = "AND TAktivitetId = " & selaktid &""
'else
'selaktidKri = ""
'end if

'& selaktidKri &"

    
    
         '*** Smiley ***
		call ersmileyaktiv()
        call smileyAfslutSettings()
	
		'*** Afsluttede uger *****
		dim medabAflugeId
		redim medabAflugeId(200)
		strSQL2 = "SELECT u.status, u.afsluttet, WEEK(u.uge, 1) AS ugenr, u.id, u.mid FROM ugestatus u WHERE "& ugeAflsMidKri &" AND uge BETWEEN '"& strStartdato &"' AND '"& strSlutdato &"' GROUP BY u.mid, uge"
		'Response.write strSQL2
		'Response.flush
		oRec2.open strSQL2, oConn, 3
		while not oRec2.EOF
			medabAflugeId(oRec2("mid")) = medabAflugeId(oRec2("mid")) & "#"& oRec2("ugenr") &"#,"
		oRec2.movenext
		wend
		oRec2.close 

	
	totkmallealle = 0
	ekspTxt = ""
	lastjobnavn = ""
	lastmedarb = 0
	
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, "_
	&" TAktivitetNavn AS Anavn, "_
	&" TAktivitetId, Tknavn, Tknr, Tmnr, "_
	&" Tmnavn, Timer, Tid, "_
	&" godkendtstatus, Tfaktim, TimePris, Timerkom, "_
	&" m.mnr, m.init, mnavn, m.mid "_
	&" FROM timer "_
	&" LEFT JOIN medarbejdere m ON (mid = tmnr) "_
	&" WHERE "& medarbKri &" AND "& jobKri &" AND (tfaktim = 5) AND "_
	&" (Tdato >= '"& strStartDato &"' AND Tdato <= '"& strSlutDato &"')"_
	&" ORDER BY Tknr, Tjobnavn, Tjobnr, Tmnavn, Tdato DESC"
	'&" ORDER BY Tknavn, Tjobnavn, Tjobnr, Tmnr, Tdato DESC "
	
	'Response.write strSQL
	
	
	'strSQL = "SELECT * FROM timer WHERE tid = 0"
	oRec.open strSQL, oConn, 0, 1
	
	v = 0
	while not oRec.EOF
		strWeekNum = datepart("ww", oRec("tdato"), 2, 2)
		id = 1
		%>
		<!--include file="inc/dato2.asp"-->
		<%		
		'Response.write lastmedarb &"<>" & oRec("Tmnr")
				if lastmedarb <> oRec("Tmnr") OR lastjobnr <> oRec("Tjobnr") then
					if v <> 0 then%>
					<tr>
						<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<tr>
						<td bgcolor="#ffffe1" align=right colspan="8" height=20 style="border-left:1 #003399 solid; border-right:1 #003399 solid; padding-right:10;"><%=lastmedarbnavn%>:&nbsp; <b><%=formatnumber(medarbtimer, 2)%></b> km.</td>
					</tr>
					<%
					medarbtimer = 0
					end if
				end if
				
				if lastjobnr <> oRec("Tjobnr") then
				%>
					    <%if v <> 0 then%>
					    <tr>
						    <td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					    </tr>
					    <tr>
						    <td bgcolor="#ffdfdf" align=right colspan="8" height=20 style="border-left:1 #003399 solid; border-right:1 #003399 solid; padding:3px 10px 0px 0px;">Ialt på <b><%=lastjobnavn%></b> i den valgte periode: <b><%=formatnumber(timerTotpajob, 2)%></b> km.<br /><br />&nbsp;</td>
					    </tr>
					    <%
					    timerTotpajob = 0
				        end if%>
				
				<tr>
					<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<tr>
					<td bgcolor="#EFF3FF" colspan="8" height=20 style="border-left:1 #003399 solid; border-right:1 #003399 solid; padding:10px 10px 0px 10px;"><b><%=oRec("Tjobnavn")%> (<%=oRec("Tjobnr")%>)</b>&nbsp;&nbsp;<%=oRec("Tknavn")%><br>
					
					<%
						'*** Jobansvarlige ***
						
						strSQL2 = "SELECT mnavn, mnr, mid, job.id, jobans1 FROM job, medarbejdere WHERE job.jobnr = '"& oRec("Tjobnr") &"' AND mid = jobans1"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2 FROM job, medarbejdere WHERE job.jobnr = '"& oRec("Tjobnr") &"' AND mid = jobans2"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						
						%>
						<font class=megetlillesilver>
						Jobansvarlig 1: <%=jobans1txt%><br>
	   					Jobansvarlig 2: <%=jobans2txt%></font>
						<%
						jobans1txt = ""
						jobans2txt = ""
						%>
					</td>
				</tr>
				<%end if%>
				
				<tr>
					<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
				<td valign="top" style="border-left:1 #003399 solid;" height="20">&nbsp;</td>
				<td height="18">&nbsp;<%=strWeekNum%>&nbsp;&nbsp;&nbsp;<%=formatdatetime(oRec("Tdato"), 2)%>&nbsp;</td>
				<td>&nbsp;
				
				<%
				 '*** Er uge alfsuttet af medarb, er smiley og autogk slået til
                erugeafsluttet = instr(medabAflugeId(oRec("mid")), "#"&datepart("ww", oRec("Tdato"), 2, 2)&"#")
                
                
                'if (cint(erugeafsluttet) <> 0 AND autogk = 1 AND smilaktiv = 1) OR _
                '(smilaktiv = 1 AND autolukvdato = 1 AND _
                '(day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) = year(now) AND DatePart("m", oRec("Tdato")) < month(now)) OR _
                '(DatePart("yyyy", oRec("Tdato")) < year(now))) then

                        
                   


                 strMrd_sm = datepart("m", oRec("Tdato"), 2, 2)
                strAar_sm = datepart("yyyy", oRec("Tdato"), 2, 2)
                strWeek = datepart("ww", oRec("Tdato"), 2, 2)
                strAar = datepart("yyyy", oRec("Tdato"), 2, 2)

                if cint(SmiWeekOrMonth) = 0 then
                usePeriod = strWeek
                useYear = strAar
                else
                usePeriod = strMrd_sm
                useYear = strAar_sm
                end if

                
                call erugeAfslutte(useYear, usePeriod, oRec("mid"))
		        
		        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		        'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		        'Response.Write "tjkDag: "& tjkDag & "<br>"
		        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		             strSQL2 = "SELECT risiko FROM job WHERE jobnr = '"& oRec("Tjobnr") &"'"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						risiko = oRec2("risiko")
						end if
						oRec2.close
                
                call lonKorsel_lukketPer(oRec("Tdato"), risiko)
		         
                'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                 if ( ( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 ) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) = year(now) AND DatePart("m", oRec("Tdato")) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) < year(now) AND DatePart("m", oRec("Tdato")) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("Tdato")) < year(now) AND DatePart("m", oRec("Tdato")) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("Tdato")) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
               
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
				
				if oRec("godkendtstatus") = 0 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 0 AND (cdate(lastfakdato) <  cdate(oRec("Tdato"))) then 'AND level = 1%>
				
	            <a href="javascript:popUp('rediger_tastede_dage_2006.asp?id=<%=oRec("Tid")%>','600','500','250','120');" target="_self"; class=vmenu>Kørsel</a>
				
				<%else %>
				Kørsel
				<%end if %>
				</td>
				<td align="center"><b><%=oRec("Timer")%></b></td>
				<%
				medarbtimer = medarbtimer + oRec("Timer")
				timerTotpajob = timerTotpajob + oRec("Timer")
				totkmallealle = totkmallealle + oRec("Timer")
				%>
				<td><%=oRec("mnavn") & " (" & oRec("mnr") &")" %>
				<%if len(oRec("init")) <> 0 then%>
				<%= " - "& oRec("init") %>
				<%end if%></td>
				<td align="center"><font class="medlillesilver"><%=convertDate(oRec("TasteDato"))%></font></td>
				<td>&nbsp;</td>
				<td valign="top" style="border-right:1 #003399 solid;" height="20">&nbsp;</td></tr>
				    
				    
				    <%if len(oRec("timerkom")) <> 0 then%>
				    <tr>
					    <td valign="top" style="border-left:1 #003399 solid;" height="20">&nbsp;</td>
					    <td>&nbsp;</td>
					    <td valign="top" width=315 style="padding-left:10;">
					    <%="<font class='lillegray'>"& oRec("timerkom")&"</font>"%><br>&nbsp;</td>
					    <td colspan="4">&nbsp;</td>
					    <td valign="top" style="border-right:1 #003399 solid;" height="20">&nbsp;</td></tr>
				    </tr>
				    <%end if
			
			
			
			
			v = v + 1
			lastmedarbnavn = oRec("mnavn")
			lastmedarb = oRec("Tmnr")
			lastjobnr = oRec("Tjobnr")
			
			
			ekspTxt = ekspTxt & formatdatetime(oRec("Tdato"), 2) &"; "& oRec("Tknavn") & " ("&oRec("Tknr")&"); "& oRec("Tjobnavn") & " ("&oRec("Tjobnr")&"); Kørsel; "& oRec("Timer") & "; " & oRec("Tmnavn") & chr(013)
			
			
			
	oRec.movenext
	wend
	oRec.Close
	'Set oRec = Nothing
	
	
		%>		
	

<tr>
		<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFe1" align=right colspan="8" height=20 style="border-left:1 #003399 solid; border-right:1 #003399 solid; padding-right:10;"><%=lastmedarbnavn%>:&nbsp; <b><%=formatnumber(medarbtimer, 2)%></b> km.</td>
	</tr>
<tr>
	<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
</tr>
<tr>
	<td bgcolor="#ffdfdf" align=right colspan="8" height=20 style="border-left:1 #003399 solid; border-right:1 #003399 solid; padding:3px 10px 0px 0px;">Ialt på <b><%=lastjobnavn%></b> i den valgte periode: <b><%=formatnumber(timerTotpajob, 2)%></b> km.<br /><br />&nbsp;</td>
</tr>
<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="6" valign="bottom"><img src="../ill/tabel_top.gif" width="609" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
</table>
<br><br>
Der er på de ovenstående job og medarbejdere, i den valgte periode kørt ialt: <u><b><%=formatnumber(totkmallealle, 2)%></b></u> Km.

<%if request("print") <> "j" then


	%><br><br>
	
	<form action="joblog_korsel_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>">
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=ekspTxt%>">
			<input type="hidden" name="txt20" id="txt20" value="">
			
			<input type="submit" value="Eksporter">
			
			</form>
<%end if%>

<br>
<br>&nbsp;
</div>
<br>
<br>
<%end if 'validering %>
<%end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->