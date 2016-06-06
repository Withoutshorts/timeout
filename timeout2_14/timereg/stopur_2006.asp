	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
	<%
	'*********** Sætter sidens global variable ***********************************
	
	'* Hvis der vælges en anden bruger
	'* den den der er logget på.
	 
	 '*** Mid !***
	if len(Request("FM_use_me")) <> 0 then
	usemrn = Request("FM_use_me")
	else
	usemrn = session("Mid")
	end if
	
	thisfile = "stopur"
	func = request("func")
	
	level = session("rettigheder")
	
	
	
	if func = "db" then
	jobid = request("jobid")
	
	'idag = day(now)&"/"&month(now)&"/"&year(now)
	tid = time()
	idag = year(now)&"/"&month(now)&"/"&day(now)
	
	fundet = 0
	
	
	if len(request("sletlog")) <> 0 then
	slet = 1
	else
	slet = 0
	end if
	
	if slet <> 0 then
		
		strSQL2 = "DELETE FROM stopur WHERE jobid = "& jobid &" AND medid = "& usemrn
		oConn.execute(strSQL2)
		
		Response.redirect "stopur_2006.asp?FM_use_me="&usemrn
	
	end if 'slet
	
	if len(request("slet")) <> 0 then
	ryd = 1
	else
	ryd = 0
	end if
	
	
	if ryd <> 1 then
		strSQL = "SELECT id, sttid, sltid FROM stopur WHERE jobid = "& jobid &" AND medid = "& usemrn
		oRec.open strSQL, oConn, 3
		'Response.write strSQL &"<br>"
		while not oRec.EOF
		fundet = 1
			
			
					strSQL2 = "UPDATE stopur SET sltid = '"& idag &" "& tid &"' WHERE id = "& oRec("id")
					oConn.execute(strSQL2)
					
			
		oRec.movenext
		wend
		oRec.close
	
	end if
	
			if fundet = 0 then
			
			strSQL = "INSERT INTO stopur (jobid, medid, sttid) VALUES "_
			&" ("& jobid &", "& usemrn &", '"& idag &" "& tid &"')"
			
			'Response.write strSQL
			oConn.execute(strSQL)
			
			end if
			
	
	
	
	
	Response.redirect "stopur_2006.asp?FM_use_me="&usemrn
	
	end if
	
	
	
	
	'************************************************************************************************
	'*** Viser Liste (Stopur) *******
	'************************************************************************************************
	
	'**** Funktioner ****
	public sttid, sltid
	function formattider(tid1, tid2)
	
	if len(trim(tid1)) <> 0 then
		if left(formatdatetime(tid1, 3), 5) <> "00:00" then
		sttid = left(formatdatetime(tid1, 3), 5)
		else
		sttid = ""
		end if
	else
	sttid = ""
	end if
	
	if len(trim(tid2)) <> 0 then
		if left(formatdatetime(tid2, 3), 5) <> "00:00" then
		sltid = left(formatdatetime(tid2, 3), 5)
		else
		sltid = ""
		end if
	else
	sltid = ""
	end if
	
	end function
	
	
	sub stopur
	
	%>
	<td align=center style="padding:2px; border-top:1px silver dashed;">
		<input type="hidden" name="jobid" id="jobid" value="<%=oRec2("id")%>">
		<input type="text" name="sttid" value="<%=sttid%>" style="width:30px; height:15px; border:1px limegreen solid; font-size:10px; font-family:arial;"> &nbsp;&nbsp;
		<input type="text" name="sltid" value="<%=sltid%>" style="width:30px; height:15px; border:1px darkred solid; font-size:10px; font-family:arial;"> = 
		<input type="text" name="ialttid" value="<%=timerthis%>" style="width:30px; height:15px; border:1px #000000 solid; font-size:10px; font-family:arial;">
		</td>
		<td style="padding:2px; border-top:1px silver dashed;">
		 <input type="checkbox" name="slet" id="slet" value="1">
		 <input type="submit" value="Start/Stop" style="height:15px; font-size:8px; font-family:arial;">
		  </td>
	<%
	end sub%>
	
	
	
<%

function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
end function

function SQLBless2(s)
		dim tmp2
		tmp2 = s
		tmp2 = replace(tmp2, ".", ",")
		SQLBless2 = tmp2
end function

'*******************************************************************************
%>
	<div id="sindhold" style="position:absolute; left:20; top:0; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	
<%

  

 
 
 '************************************************************************************************** 
  'Opbygger timereg SQL states
  '**************************************************************************************************
 	
				'*** Tilvalgte job i guiden dine aktive job ****
				varUseJob = "("
				
						'******************************************************
						'**** Henter job fra Guiden Dine aktive job ***********
						strSQL3 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE medarb = "& usemrn &""
						oRec3.open strSQL3, oConn, 3
						
						while not oRec3.EOF 
						varUseJob = varUseJob & " j.id = "& oRec3("jobid") & " OR "
						oRec3.movenext
						wend 
						
						oRec3.close 
						
						
				
				if varUseJob = "(" then
				varUseJob = " j.id = 0 AND "
				else
				varUseJob_len = len(varUseJob)
				varUseJob_left = varUseJob_len - 3
				varUseJob_use = left(varUseJob, varUseJob_left) 
				varUseJob = varUseJob_use & ") AND "
				end if
				
				
				
				'**** Henter projektgrupper ***
				call hentbgrppamedarb(usemrn)
				
	
	'**************************************************************************************************
	
	
	%>
	<div name="kontakterogjob" id="kontakterogjob" style="position:absolute; left:20px; top:10px; display:; width:500px; visibility:visible; background-color:#ffffff; padding:5px; border:1px #5582d2 solid; border-right:2px #5582d2 solid; border-bottom:2px #5582d2 solid;">
	<table cellspacing=0 cellpadding=0 border=0>
	
	<tr bgcolor="#5582d2">
		<td width=200 class=alt style="padding-left:3px;"><b>Kunder</b> (Job)</td>
		<td class=alt width=120 style="padding-left:5px;"><b>Start / Stop tid</b></td>
		<td width=80 class=alt style="padding-left:5px;"><b>Ryd / Opd.</b></td>
	</tr>
	<%
	
	'************** Job med adgang til (ved ignorer projektgrupper) ******************************************
	
	if ignorerProgrp = 1 then
	
	strJobmedadgang = "#0#"
	
		strSQL = "SELECT j.id AS id FROM job j "_
		&" WHERE "& varUseJob &" j.jobstatus = 1 AND j.fakturerbart = 1 "& strPgrpSQLkri & " ORDER BY j.id" 
		
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
		strJobmedadgang = strJobmedadgang &",#"& oRec("id")&"#"
		
		oRec.movenext
		wend
		
		oRec.close
	
	end if
	
	
	'************** Main SQL Call ******************************************
	strSQL2 = "SELECT j.id AS id, jobnr, jobnavn, jobknr, kkundenavn, kkundenr, "_
	&" kid, jobans1, jobans2 FROM job j, kunder k"_
	&" WHERE "& varUseJob &" j.jobstatus = 1 AND j.fakturerbart = 1 AND k.Kid = j.jobknr "& strPgrpSQLkri & " GROUP BY j.id ORDER BY k.kkundenavn, j.jobnavn, j.id, j.jobnavn" 
	
	'Response.write strSQL2
	'Response.flush
	'Response.write "<hr>"
	
	oRec2.Open strSQL2, oConn, 3
	
	lastKid = 0
	x = 0
	while not oRec2.EOF 
	
	editok = 0
	bgThis = "ffffff"
	sttid = ""
	sltid = ""
	timerthis = ""
	
	
	
	if lastKid <> oRec2("kid") then%>
	<tr bgcolor="#eff3ff">
		<td colspan=3 style="padding:2px;"><b><%=oRec2("kkundenavn")%> (<%=oRec2("kkundenr")%>)</b></td>
	</tr>
	<%end if%>
	
	<form action="stopur_2006.asp?func=db&medid=<%=usemrn%>" method=POST>
	<tr bgcolor="<%=bgThis%>">
		
		<td style="padding:2px; border-top:1px silver dashed;" class=lille>
		<%=oRec2("jobnavn")%> (<%=oRec2("jobnr")%>)</td>
		
		
		<%
		strStopUrLog = ""
		
		strSQLstopur = "SELECT s.sttid, s.sltid FROM stopur s WHERE s.jobid = "& oRec2("id") &" AND s.medid = "& usemrn &" ORDER BY id DESC"
		'Response.write strSQLstopur
		'Response.flush 
		oRec3.open strSQLstopur, oConn, 3
		s = 0
		while not oRec3.EOF
		
		call formattider(oRec3("sttid"), oRec3("sltid"))
		
		if len(trim(sttid)) <> 0 AND len(trim(sltid)) <> 0 then
			timerthis = ""
			totalmin = datediff("n", sttid, sltid)
			call timerogminutberegning(totalmin)
			timerthis = thoursTot&"."&tminTot
		end if
		
		if s = 0 then
			call stopur
		end if
		
		strStopUrLog = strStopUrLog & sttid &" - "& sltid & " = " & timerthis & vbcrlf 
		
		s = s + 1
		oRec3.movenext
		wend
		
		
		
		oRec3.close
		
		if s = 0 then
			call stopur
		end if%>
		
	
	</tr>
	<%if s <> 0 then%>
	<tr><td colspan=3 style="padding:5px; border-bottom:1px silver dashed;"><textarea cols="40" rows="4" name="FM_log_<%=oRec2("id")%>" id="FM_log_<%=oRec2("id")%>" id="FM_log_<%=oRec2("id")%>"><%=strStopUrLog%></textarea>&nbsp; <input type="checkbox" name="sletlog" id="sletlog" value="1"><input type="submit" value="Slet log" style="height:15px; font-size:8px; font-family:arial;"></td></tr>
	<%end if%>
	</form>
	
	
	<%
	x = x + 1
	lastKid = oRec2("kid")
	oRec2.movenext
	wend
	oRec2.close
	
	%>
	
	</table>
	
	
	

</div>


<%
end if 'user session
%>
<!--#include file="../inc/regular/footer_inc.asp"-->





