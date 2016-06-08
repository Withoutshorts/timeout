		
		<div id="divopgaveknap" name="divopgaveknap" style="position:absolute; top:262; left:2; z-index:100;">
		<a href="#" onclick="showopgave()" class='alt'><img src="../ill/knap_opglisten_on.gif" width="90" height="23" alt="" border="0" id="knapopgliste"></a></td>
		</div>
		
		<div id="divmilepalknap" name="divmilepalknap" style="position:absolute; top:262; left:93; z-index:100;">
		<a href="#" onclick="showmilepal()" class='alt'><img src="../ill/knap_milepal_off.gif" width="61" height="23" alt="" border="0" id="knapmilepal"></a>
		</div>
		
		<div id="divtimerknap" name="divtimerknap" style="position:absolute; top:262; left:155; z-index:100;">
		<a href="#" onclick="showtimer()" class='alt'><img src="../ill/knap_timer_off.gif" width="42" height="23" alt="" border="0" id="knaptimer"></a>
		</div>


<DIV ID="divmilepal" name="divmilepal" style="position: absolute; display:none; visibility:hidden; top:282; z-index:2000; width:<%=wdt%>px;">
<table cellpadding="0" cellspacing="0" border="0" width="200">
	<tr>
		<td colspan=3 bgcolor="#003399" height=1><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td style="border-left: 1px solid #003399; border-right: 1px solid #003399; border-bottom: 1px solid #003399; padding-left:3px; padding-right:3; padding-top:3; padding-bottom:5px;"><b>Milepæle de næste 7 dage:</b><br>
		<table cellpadding="0" cellspacing="0" border="0" width="190">
		<%
		function jobnavn()
				if level <= 3 OR level = 6 then
				Response.write "<tr bgcolor=#ffffff><td colspan=3><b><a href='milepale.asp?menu=job&jid="&oRec3("thisjobid")&"&func=list&rdir=treg' class=vmenu>"& oRec3("jobnavn") & "</a></b></td></tr>"
				else
				Response.write "<tr bgcolor=#ffffff><td colspan=3><b>"& oRec3("jobnavn") & "</b></td></tr>"
				end if
		end function
		
		
		datoKri = dateadd("d", 7, daynow)
		strSQL3 = "SELECT timereg_usejob.id, medarb, jobid, jobstartdato, jobslutdato, jobnavn, job.id AS thisjobid, milepale.navn AS navn, milepal_dato, milepale_typer.navn AS typenavn, ikon, milepale.type FROM timereg_usejob LEFT JOIN job ON (job.id = jobid AND jobstatus = 1) LEFT JOIN milepale ON (jid = job.id) LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) WHERE medarb = "& usemrn &" ORDER BY job.id, milepal_dato"
		'Response.write strSQL3
		oRec3.open strSQL3, oConn, 3
		
		lastjid = 0
		while not oRec3.EOF 
		
		if lastjid <> oRec3("jobid") then
		used = "n"
			if len(oRec3("jobstartdato")) <> 0 AND len(oRec3("jobslutdato")) <> 0 then
				if cdate(oRec3("jobstartdato")) >= cdate(daynow) AND cdate(oRec3("jobstartdato")) <= cdate(datoKri) then
					if used <> "j" then
					call jobnavn()
					used = "j"
					end if
				Response.write "<tr bgcolor=#ffffff><td width=130 valign=top><img src='../ill/mp_gron.gif' width='10' height='10' alt='' border='0'>&nbsp;Job startdato</td><td>&nbsp;<font class='megetlillesilver' width=50>"& oRec3("jobstartdato") &"</td><td>&nbsp;</td></tr>"
				end if
				if cdate(oRec3("jobslutdato")) >= cdate(daynow) AND cdate(oRec3("jobslutdato")) <= cdate(datoKri) then
					if used <> "j" then
					call jobnavn()
					used = "j"
					end if
				Response.write "<tr bgcolor=#ffffff><td valign=top><img src='../ill/mp_end.gif' width='10' height='10' alt='' border='0'>&nbsp;Job slutdato</td><td>&nbsp;<font class='megetlillesilver'>"& oRec3("jobslutdato") &"</td><td>&nbsp;</td></tr>"
				end if
			end if
		end if
		
		if len(oRec3("milepal_dato")) <> 0 then
			if cdate(oRec3("milepal_dato")) >= cdate(daynow) AND cdate(oRec3("milepal_dato")) <= cdate(datoKri) then
				if used <> "j" then
				call jobnavn()
				used = "j"
				end if
			Response.write "<tr bgcolor=#ffffff><td valign=top><img src='../ill/mp_"& oRec3("ikon") &".gif' width='10' height='10' alt='' border='0'>&nbsp;"& oRec3("typenavn") &"</td><td>&nbsp;<font class='megetlillesilver'>"& oRec3("milepal_dato") &"</td><td>&nbsp;</td></tr>"
			end if
		end if
		
		lastjid = oRec3("jobid")
		oRec3.movenext
		wend 
		
		oRec3.close 
		
		if used = "n" then
		Response.write "<br>(ingen)"
		end if%>
		
		</table></td>
	</tr>
</table>
</div>

<DIV ID="divopgave" name="divopgave" style="position: absolute; display: ; top:282; z-index:500; width:<%=wdt%>px;">
<table cellpadding="0" cellspacing="0" border="0" width="200">
	<tr>
		<td colspan=3 bgcolor="#003399" height=1><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td width="199" style="border-left: 1px solid #003399; padding-left:3; padding-top:3;" valign="top" colspan="2"><b>Opgavelisten:</b><br>
		<%
		cdmd = month(now())
		cddag = day(now())
		cdaar = year(now())
		
		cdidag = cdaar&"/"&cdmd&"/"&cddag
		
				'*** tæller antal af todo's **
				'*** + antal siden sidste login ***
				strSQL = "SELECT todo.id AS id, sortorder, medarbid, todoid, todo.dato, todo.slutdato AS slutdato, kundeid, jobstatus, job.id AS jid, startdato, slutdato, udfort FROM todo, todo_rel LEFT JOIN job ON (job.id = kundeid) WHERE (todo.sortorder > 0 AND todo.dato > '2004-1-10') AND (todo_rel.todoid = todo.id AND todo_rel.medarbid = "& usemrn &") ORDER BY todo.id"
				
				'Response.write strSQL
				'Response.flush
				
				oRec.open strSQL, oConn, 3
				x = 0
				nye = 0
				gluge = 0
				gamle = 0
				
				'**** Finder datoer på ugens 7 dage **************
				select case weekday(useDate)
				case 1
				todosondag = dateadd("d", 0 , useDate)
				todomandag = dateadd("d", 1 , useDate)
				todotirsdag = dateadd("d", 2 , useDate)
				todoonsdag = dateadd("d", 3 , useDate)
				todotorsdag = dateadd("d", 4 , useDate)
				todofredag = dateadd("d", 5 , useDate)
				todolordag = dateadd("d", 6 , useDate)
				case 2
				todomandag = dateadd("d", 0 , useDate)
				todosondag = dateadd("d", -1 , useDate)
				todotirsdag = dateadd("d", 1 , useDate)
				todoonsdag = dateadd("d", 2 , useDate)
				todotorsdag = dateadd("d", 3 , useDate)
				todofredag = dateadd("d", 4 , useDate)
				todolordag = dateadd("d", 5 , useDate)
				case 3
				todotirsdag = dateadd("d", 0 , useDate)
				todosondag = dateadd("d", -2 , useDate)
				todomandag = dateadd("d", -1 , useDate)
				todoonsdag = dateadd("d", 1 , useDate)
				todotorsdag = dateadd("d", 2 , useDate)
				todofredag = dateadd("d", 3 , useDate)
				todolordag = dateadd("d", 4 , useDate)
				case 4
				todoonsdag = dateadd("d", 0 , useDate)
				todosondag = dateadd("d", -3 , useDate)
				todomandag = dateadd("d", -2 , useDate)
				todotirsdag = dateadd("d", -1 , useDate)
				todotorsdag = dateadd("d", 1 , useDate)
				todofredag = dateadd("d", 2 , useDate)
				todolordag = dateadd("d", 3 , useDate)
				case 5
				todotorsdag = dateadd("d", 0 , useDate)
				todosondag = dateadd("d", -4 , useDate)
				todomandag = dateadd("d", -3 , useDate)
				todotirsdag = dateadd("d", -2 , useDate)
				todoonsdag = dateadd("d", -1 , useDate)
				todofredag = dateadd("d", 1 , useDate)
				todolordag = dateadd("d", 2 , useDate)
				case 6
				todofredag = dateadd("d", 0 , useDate)
				todosondag = dateadd("d", -5 , useDate)
				todomandag = dateadd("d", -4 , useDate)
				todotirsdag = dateadd("d", -3 , useDate)
				todoonsdag = dateadd("d", -2 , useDate)
				todotorsdag = dateadd("d", -1 , useDate)
				todolordag = dateadd("d", 1 , useDate)
				case 7
				todolordag = dateadd("d", 0 , useDate)
				todosondag = dateadd("d", -6 , useDate)
				todomandag = dateadd("d", -5 , useDate)  
				todotirsdag = dateadd("d", -4 , useDate)
				todoonsdag = dateadd("d", -3 , useDate)
				todotorsdag = dateadd("d", -2 , useDate)
				todofredag = dateadd("d", -1 , useDate)
				end select
				
				
				lowerKri =  convertDateYMD(todosondag)
				upperKri = convertDateYMD(todolordag)
				while not oRec.EOF 
					
					if oRec("jobstatus") = "1" then 'job = aktiv
						x = x + 1
						
						if (cdate(oRec("slutdato")) > cdate(lowerKri)) AND (cdate(oRec("slutdato")) <= cdate(upperKri)) then
						nye = nye + 1
						
							if cint(oRec("udfort")) = 4 then
							gluge = gluge + 1
							end if
						end if
						
						if cint(oRec("udfort")) = 4 then
						gamle = gamle + 1
						end if
					end if
				oRec.movenext
				wend
				oRec.close
		
		totalpaliste = x
		nyepaliste = nye
		gamlepaaliste = gamle
		gamleuge = gluge
		'*** Sætter lokal dato/kr format. *****
		Session.LCID = 1030
		%>
		<%if level <= 3 OR level = 6 then%>
		<a href="javascript:NewWin_todo('todo_add.asp?func=opr&wk=<%=todosondag%>');" target="_self" class='vmenu'>Ny opgave <img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"></a><br>
		<%end if%>
		<a href="javascript:NewWin_todo('todo_add.asp?func=list&wk=<%=todosondag%>');" target="_self" class='vmenu'>Opgavelisten <img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"></a><br>
			<br><b>Du har:</b> (<%=session("user")%>)<br>
			Total:&nbsp;<font color="#999999"><%=totalpaliste%> opgave(r),</font> <font color="darkred"><%=gamlepaaliste%> afsluttet.</font><br>
			Uge <%=datepart("ww", useDate, 2, 2)%>:&nbsp;<font color="Goldenrod"><%=nyepaliste%> opgave(r),</font> <font color="darkred"><%=gamleuge%> afsluttet.</font>
			<br><br>
			<u><b>Dagens opgaver d. <%=daynow%>:</b></u><br>
			<font color="#5582d2">
			<%
			strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobknr, jobstatus, todo.editor, todo.id, kundeid, emne, todo.beskrivelse, udfort, klassificering, kid, kkundenavn, sortorder, startdato, slutdato, medarbid, todoid FROM todo, todo_rel LEFT JOIN job ON (job.id = kundeid) LEFT JOIN kunder ON (kid = jobknr) WHERE kundeid <> -1 AND (todoid = todo.id AND medarbid = "& session("Mid") &") AND udfort = 1 AND slutdato = '"& cdidag &"' AND emne <> '' AND jobstatus = 1 GROUP BY sortorder ORDER BY sortorder "
			oRec.open strSQL, oConn, 3 
			fu = 0
			while Not oRec.EOF 
			
			if level <= 3 OR level = 6 then
			Response.write "<a href=""javascript:NewWin_todo('todo_add.asp?func=red&id="&oRec("id")&"');"" target=""_self"" class=""vmenu"">" &oRec("kkundenavn") &"</a><br>" &  oRec("jobnavn") &" ("& oRec("jobnr") &")<br> " & oRec("emne") & "<br><br>"
			else
			Response.write "<b>" &oRec("kkundenavn") &"</b><br>" &  oRec("jobnavn") &" ("& oRec("jobnr") &")<br> " & oRec("emne") & "<br><br>"
			end if
			
			fu = 1
			oRec.movenext
			wend
			oRec.close
			
			if fu = 0 then
			Response.write "(Ingen)"
			end if
			%>
			
		<br>&nbsp;</td>
		<td style="border-right: 1px solid #003399;">&nbsp;</td>
	</tr>
		<tr>
		<td colspan="3" style="border-top: 1px solid #003399;"><img src="../ill/blank.gif" alt="" border="0"></td>
		</tr>
</table>
</div>
	
	


