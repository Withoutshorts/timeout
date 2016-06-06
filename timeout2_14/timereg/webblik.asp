<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<%

func = request("func")

thisfile = "webblik.asp"

if len(request("FM_parent")) <> 0 then
parent = request("FM_parent")
else
parent = 0
end if

if len(request("oldone")) <> 0 then
oldones = request("oldone")
else
oldones = 0
end if

if len(request("lastid")) <> 0 then
lastId = request("lastid")
else
lastId = 0
end if

function todolist(klikket)

if cint(lastId) = oRec("id") then
bgthis = "#ffff99"
else
bgthis = "#ffffff"
end if
%>
<tr>
		<%if antalsubs < 1 then
		strImgLI = "<img src='../ill/webblik_li_8.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<%if antalsubs >= 1 AND antalsubs <= 5 then
		strImgLI = "<img src='../ill/webblik_li_10.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<%if antalsubs > 5 AND antalsubs <= 15 then
		strImgLI = "<img src='../ill/webblik_li_12.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<%if antalsubs > 15 AND antalsubs <= 30 then
		strImgLI = "<img src='../ill/webblik_li_14.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<%if antalsubs > 30 then
		strImgLI = "<img src='../ill/webblik_li_16.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<td valign=top style="padding-top:3px; padding-bottom:3px; border-bottom:1px #cccccc dashed;"><%=strImgLI%></td>
		<td bgcolor="<%=bgthis%>" width=250 valign=top style="padding-top:3px; padding-bottom:3px; border-bottom:1px #cccccc dashed;">
		
		<% if oRec("afsluttet") = 1 then%>
		<s>
		<%end if%>
		
		<a href="webblik.asp?FM_parent=<%=oRec("id")%>&todolevel=<%=oRec("level")%>" class="<%=acls%>"><%=oRec("navn")%></a>&nbsp;<font class=lillegray>(<%=antalsubs%>)</font>
		
		<%if oRec("delt") <> 0 then%>
		<font class=megetlillesort>(delt)</font> 
		<%end if%>
		
		<% if oRec("afsluttet") = 1 then%>
		</s>
		<%end if%>
		</td>
		
		<%'*** Andre medarbejdere med adgang ***
			medarbemails = " "
			strSQL2 = "SELECT medarbid, email, mid FROM todo_rel_new "_
			&" LEFT JOIN medarbejdere ON (mid = medarbid) WHERE todoid = " & oRec("id") &" GROUP BY mid ORDER BY mid" ' & " AND medarbid <> " & session("mid")
			'Response.write strSQL2
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF  
				medarbemails = medarbemails & oRec2("mid") & ","
			oRec2.movenext
			wend
			oRec2.close 
		
		len_medarbemails = len(medarbemails)
		left_medarbemails = left(medarbemails, len_medarbemails - 1)
		medarbemails = trim(left_medarbemails)
		%>
		<td width=30 bgcolor="<%=bgthis%>" align=right valign=top style="padding-top:3px; padding-bottom:3px; padding-right:5px; border-bottom:1px #cccccc dashed;">&nbsp;<a href="#" onClick="edittodo('<%=oRec("id")%>','<%=oRec("navn")%>','<%=medarbemails%>','<%=formatdatetime(oRec("dato"), 0)%>', '<%=oRec("afsluttet")%>', <%=parent%>)"><img src="../ill/blyant.gif" width="12" height="11" alt="" border="0"></a></td>
		<td width=90 bgcolor="<%=bgthis%>" valign=top style="padding-top:3px; padding-bottom:3px; border-bottom:1px #cccccc dashed;">&nbsp;<a href="webblik_oprtodo.asp?func=slet&id=<%=oRec("id")%>&FM_parent=<%=parent%>&oldone=<%=oldones%>" class="red">[x]</a>&nbsp;&nbsp;
		<%if klikket <> 0 AND oRec("afsluttet") <> 1 then%>
		<a href="webblik_oprtodo.asp?func=op&id=<%=oRec("id")%>&FM_parent=<%=parent%>&oldone=<%=oldones%>"><img src="../ill/wb_pil_op.gif" width="10" height="10" alt="" border="0"></a>
		<a href="webblik_oprtodo.asp?func=ned&id=<%=oRec("id")%>&FM_parent=<%=parent%>&oldone=<%=oldones%>"><img src="../ill/wb_pil_ned.gif" width="10" height="10" alt="" border="0"></a>&nbsp; <font class=megetlillesilver>(<%=oRec("sortorder")%>)</font>
		<%end if%>
		</td>
	</tr>
	<%
	
	
end function


sub periodeForm%>
	<form method="post" action="webblik.asp?menu=webblik&FM_usedatokri=1">
		<!--#include file="inc/weekselector_s.asp"-->&nbsp;
		<%if level = 1 then
		if request("FM_ignorer_pg") = "1" then
		chkThis = "CHECKED"
		else
		chkThis = ""
		end if
		%>
		<br><input type="checkbox" name="FM_ignorer_pg" value="1" <%=chkThis%>> Vis job/milepæle uanset tilknytning til dine projektgrupper.
		<%end if%>
		<br><input type="submit" value="Vis periode.">
	</form>
<%end sub


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	%>
	
	<script>
	function showtodo(val, todorelmedarb){
	document.getElementById("addtodo").style.visibility = "visible"
	document.getElementById("addtodo").style.display = ""
	//document.getElementById("FM_todo_rel").value = todorelmedarb
	
	var medarbCHK = todorelmedarb;
	var medarbCHK_array = new Array();
	medarbCHK_array = medarbCHK.split(",");
	
	for (i = 0; i < document.oprredtodo.FM_todo_rel.length; i++){
		document.oprredtodo.FM_todo_rel[i].checked = false;
		for (v = 0; v < medarbCHK_array.length; v++){
			if (document.oprredtodo.FM_todo_rel[i].value == medarbCHK_array[v]){
			document.oprredtodo.FM_todo_rel[i].checked = true;
			}
		}
	}
	
	document.getElementById("FM_parent").value = val;
	document.getElementById("FM_todo").focus()
	}
	
	function hidetodo(){
	document.getElementById("addtodo").style.visibility = "hidden"
	document.getElementById("addtodo").style.display = "none"
	document.getElementById("FM_todo").value = ""
	}
	
	function edittodo(val, todo, todorelmedarb, dato, afl, valparent){
	if (afl == 1) {
	document.getElementById("FM_afsluttet").checked = true
	}
	
	document.getElementById("addtodo").style.visibility = "visible"
	document.getElementById("addtodo").style.display = ""
	document.getElementById("id").value = val;
	//document.getElementById("FM_todo_rel").value = todorelmedarb
	
	var medarbCHK = todorelmedarb;
	var medarbCHK_array = new Array();
	medarbCHK_array = medarbCHK.split(",");
	
	for (i = 0; i < document.oprredtodo.FM_todo_rel.length; i++){
		document.oprredtodo.FM_todo_rel[i].checked = false;
		for (v = 0; v < medarbCHK_array.length; v++){
			if (document.oprredtodo.FM_todo_rel[i].value == medarbCHK_array[v]){
			document.oprredtodo.FM_todo_rel[i].checked = true;
			}
		}
	}
	
	document.getElementById("FM_todo").value = todo;
	document.getElementById("redi").value = "Sidst redigeret: " + dato;
	document.getElementById("submittodo").value = "Opdater ToDo"
	document.getElementById("FM_parent").value = valparent;
	}
	
	</script>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	<div id="sindhold" style="position:absolute; left:0px; top:42px; visibility:visible;">
	
	<%call tsamainmenu(2)%><br><br>
	<table>
	<tr>
		<td rowspan=2><img src="../ill/header_to_idag_ikon.gif" width="58" height="48" alt="" border="0"></td>
		<td><img src="../ill/header_to_idag.gif" width="898" height="34" alt="" border="0"></td>
	</tr>
	<tr><td valign=top>&nbsp;&nbsp;&nbsp;&nbsp;<%=weekdayname(weekday(day(now) &"/"& month(now) &"/"& year(now)))%> d.<%=formatdatetime(day(now) &"/"& month(now) &"/"& year(now), 1)%></td></tr>
	</table>
	
	<table cellspacing=20 cellpadding=0 border=0 width=950>
	<tr>
		<td width=400 valign=top style="border:2px #cccccc solid; padding:10px;" bgcolor="#ffffff"><h4>ToDo listen</h4>
	
	<a href="Javascript:history.back();" class=todo_lille><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0"> Tilbage</a>&nbsp;&nbsp;
	|&nbsp;&nbsp;<a href="webblik.asp?FM_parent=0" class=todo_lille>ToDo oversigt</a><br><br>
	<table cellspacing=0 cellpadding=0 border=0 widht=400>
	<tr><form method=post action="webblik.asp?FM_parent=<%=parent%>">
		<td colspan=4 style="border-bottom:1px #cccccc solid; padding:2px;">
		<%
		todolevel = request("todolevel")
		if todolevel < 3 then
		%>
		<a href="#" onClick="showtodo('<%=parent%>','<%=session("mid")%>')" class=todo_stor>Opret ny ToDo her&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		<%else%>
		<br><b>Der kan maks oprettes 3 niveauer ToDo's!</b>
		<%end if
		
		
		if len(request("FM_afsluttede")) <> 0 then
			visafl = request("FM_afsluttede")
			if visafl = "1" then
			visaflCHK0 = "CHECKED"
			visaflCHK1 = ""
			else
			visafl = 0
			visaflCHK0 = ""
			visaflCHK1 = "CHECKED"
			end if
			Response.cookies("webblik_visafl") = visafl
			Response.cookies("webblik_visafl").expires = date + 40
		else
			if request.cookies("webblik_visafl") = "1" then
			visafl = 1
			visaflCHK0 = "CHECKED"
			visaflCHK1 = ""
			else
			visafl = 0
			visaflCHK0 = ""
			visaflCHK1 = "CHECKED"
			end if
		end if
		
		%>
		
		
		<br><b>Vis afsluttede ToDo's?</b>
		<input type="radio" name="FM_afsluttede" id="FM_afsluttede" value="1" <%=visaflCHK0%>> Ja &nbsp;&nbsp;
		<input type="radio" name="FM_afsluttede" id="FM_afsluttede" value="0" <%=visaflCHK1%>> Nej &nbsp;&nbsp;
		<input type="submit" value="Vis" style="width:30; font-size:10px;"><br><br>
		
		
	</td>
	</form>
	</tr>
	<%
	'**** Todo ****************'
	if oldones <> "1" then
	sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
	datointervalslut = dateadd("d", -7, sqlDatostart)  
	sqlDatoslut = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	else
	sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
	sqlDatoslut = "2005/1/1"
	end if
	
	
	if cint(parent) = 0 then
	
	datoSQL = " AND dato BETWEEN '"& sqlDatoslut &"' AND '"& sqlDatostart &"'"
	
	else
		
		
		'**** klikket Todo SQL **************
		strSQL = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, t.afsluttet, t.sortorder"_
		&" FROM todo_new t WHERE t.id = "& parent 
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
					
					'*** Opdaterer dato på redigering på den viste todo ***
					sqlToDoVsitDato = year(now) & "/" & month(now) & "/" & day(now)
					strSQL = "UPDATE todo_new SET dato = '"& sqlToDoVsitDato &"' WHERE id = "& oRec("id")
					oConn.execute(strSQL)
					
					antalsubs = 0
			
					strSQL2 = "SELECT count(parent) AS antalsubs FROM todo_new WHERE parent = " & oRec("id") &" GROUP BY parent"
					oRec2.open strSQL2, oConn, 3 
					if not oRec2.EOF then 
					
					antalsubs = oRec2("antalsubs")
					
					end if
					oRec2.close 
					
				acls = "todo_stor"
				
				call todolist(0)
		
		
		end if
		oRec.close 
		
		
	datoSQL = ""
	end if
	
	
	if visafl = 1 then
	visaflSQL = " AND t.afsluttet <> 99 "
	else
	visaflSQL = " AND t.afsluttet <> 1 "
	end if
	
	'**** Main Todo SQL **************
	strSQL = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, t.afsluttet, t.sortorder, trn.medarbid, trn.todoid "_
	&" FROM todo_rel_new trn "_
	&" RIGHT JOIN todo_new t ON (t.id = trn.todoid "& visaflSQL &" AND (parent = "& parent &" "& datoSQL &"))"_
	&" WHERE medarbid = "& session("mid") &" GROUP BY t.id ORDER BY t.afsluttet, t.sortorder"
	
	
	'Response.write strSQL
	
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
			
			antalsubs = 0
			
			strSQL2 = "SELECT count(parent) AS antalsubs FROM todo_new WHERE parent = " & oRec("id") &" GROUP BY parent"
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then 
			
			antalsubs = oRec2("antalsubs")
			
			end if
			oRec2.close 
			
	
			
		acls = "todo_mellem"
		
		
	call todolist(1)
	
	oRec.movenext
	wend
	oRec.close %>
	
	
	<tr>
		<td colspan=4 style="border-bottom:1px #cccccc solid; padding:2px;">
	<%
	'*** Mere end 7 dage gamle ToDo's ***********************
	if cint(parent) = 0 then
	
			if request("FM_visalle") = "1" then
			Response.cookies("visallegamle") = "j"
			Response.cookies("visallegamle").expires = date + 10
			fmallechk0 = ""
			fmallechk1 = "CHECKED"
			else
			Response.cookies("visallegamle") = "n"
			Response.cookies("visallegamle").expires = date + 10
			fmallechk0 = "CHECKED"
			fmallechk1 = ""
			end if%>
			<br><br><br><form method=post action="webblik.asp?FM_parent=<%=parent%>">
			<b>Mere end en uge gamle ToDo's:</b><br>
			Vis kun maks 3 md. gamle  <input type="radio" name="FM_visalle" id="FM_visalle" value="0" <%=fmallechk0%>>&nbsp;&nbsp;
			Vis alle:<input type="radio" name="FM_visalle" id="FM_visalle" value="1" <%=fmallechk1%>>
			<input type="submit" value="Vis" style="width:30; font-size:10px;">
			</form></td>
			</tr><tr>
			<td colspan=4><div style="position:relative; height:400px; overflow:auto;">
			<%
			
			sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
			datointervalslut = dateadd("d", -8, sqlDatostart)  
			sqlDatostart = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
			
			if request.cookies("visallegamle") = "j" then
			sqlDatoslut = "2002/1/1"
			else
			sqlDatoslut_temp = dateadd("d", -90, day(now)&"/"& month(now)&"/"&year(now)) 
			sqlDatoslut = year(sqlDatoslut_temp)&"/"& month(sqlDatoslut_temp)&"/"&day(sqlDatoslut_temp) 
			end if
			
				
			strSQL3 = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, trn.medarbid, trn.todoid, t.afsluttet "_
			&" FROM todo_rel_new trn "_
			&" RIGHT JOIN todo_new t ON (t.id = trn.todoid "& visaflSQL &" AND (parent = "& parent &" "& hentklikketSQL &")"_
			&" AND dato BETWEEN '"& sqlDatoslut &"' AND '"& sqlDatostart &"') "_
			&" WHERE medarbid = "& session("mid") &" GROUP BY t.id ORDER BY t.navn" 
			
			'Response.write "<br><br>" & strSQL
			'Response.flush
			
			oRec3.open strSQL3, oConn, 3 
			while not oRec3.EOF 
				
			%>
			<a href="webblik.asp?FM_parent=<%=oRec3("id")%>&todolevel=<%=oRec3("level")%>&oldone=1" class="todo_lille">
			<% if oRec3("afsluttet") = 1 then%>
			<s>
			<%end if%>
			
			<%=oRec3("navn")%>
			
			<% if oRec3("afsluttet") = 1 then%>
			</s>
			<%end if%>
			</a>
			 
			 
			 &nbsp;<a href="webblik_oprtodo.asp?func=slet&id=<%=oRec3("id")%>&FM_parent=<%=parent%>&oldone=<%=oldones%>" class="red">[x]</a>&nbsp;&nbsp; 
			<br>
			<%
			oRec3.movenext
			wend
			oRec3.close 
	
	end if
	%>
	</div>
	</td></tr>
	</table>
	</td>

	
	
	
	
	<%
	'*******************************************************
	'**** Igangværende Job *********************************
	'*******************************************************
	%>
	
	
	<td valign=top width=450 style="border:2px orange solid; padding:10px;" bgcolor="#ffffff"><h4>Igangværende Job</h4>
	Aktive job, som du er tilknyttet via dine projektgrupper, og som har en <b>slutdato</b> i det valgte datointerval.
	<h5> (Maks 25 job vises på listen.)</h5>
	<%call periodeform%>
	<table cellspacing=0 cellpadding=0 border=0 width=450>
	<tr>
		 <td colspan=3 align=right style="padding:2px;"><a href="jobs.asp?menu=job&func=opret&id=0&int=1&rdir=webblik" class=todo_stor>Opret nyt job &nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a></td>
	</tr>
	<%
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut 'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	
	Response.write "<b>Periode:</b> " & formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)
	
	if request("FM_ignorer_pg") <> "1" then
		call hentbgrppamedarb(session("mid"))
	else
		strPgrpSQLkri = ""
	end if
	
	strSQL = "SELECT j.id, jobnavn, jobnr, jobknr, kkundenavn, "_
	&" jobans2, kkundenr, jobslutdato, j.beskrivelse, jobans1, mnavn, "_
	&" ikkebudgettimer, budgettimer, jobtpris, sum(r.timer) AS restimer "_
	&" FROM job j LEFT JOIN kunder ON (kid = jobknr) "_
	&" LEFT JOIN medarbejdere m ON (m.mid = jobans1)"_
	&" LEFT JOIN ressourcer r ON (r.jobid = j.id)"_
	&" WHERE fakturerbart = 1 AND jobstatus = 1 AND jobslutdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' "& strPgrpSQLkri &""_
	&" GROUP BY j.id ORDER BY jobslutdato LIMIT 0, 25"
	
	'Response.write strSQL
	'Response.flush
	
	c = 0
	budgetIalt = 0
	budgettimerIalt = 0
	restimerIalt = 0
	
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	bg = right(c, 1)
	select case bg
	case 0,2,4,6,8
	bgthis = "#ffffff"
	case else
	bgthis = "#eff3ff"
	end select
	
	'** Rettigheder til at redigere **
	if level = 1 then
	editok = 1
	else
			if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
			editok = 1
			end if
	end if
	%>
	
	<tr bgcolor="<%=bgthis%>">
		<td width=130 valign=top style="padding:3px; border-top:1px #cccccc dashed;"><b><%=left(oRec("kkundenavn"), 30)%></b></td>
		<td width=280 valign=top style="padding:3px; border-top:1px #cccccc dashed;" class=lille>
		<%if editok = 1 then%>
		<a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&rdir=webblik"><%=left(oRec("jobnavn"), 30)%></a>&nbsp;&nbsp;(<%=oRec("jobnr")%>)&nbsp;&nbsp;
		&nbsp;&nbsp;<a href="jobs.asp?menu=job&func=slet&id=<%=oRec("id")%>&jobnr_sog=a&filt=aaben&fm_kunde_sog=0&rdir=webblik" class=red>[x]</a><%else%>
		<b><%=left(oRec("jobnavn"), 30)%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</b>
		<%end if%>
		<br>
		<font color="red"><%=formatdatetime(oRec("jobslutdato"), 1)%></font>&nbsp;|&nbsp;
		<font color="#c4c4c4"><i><%=oRec("mnavn")%></i></font>
		</td>
		<td valign=top>&nbsp;</td>
		
	</tr>
	<%if len(oRec("beskrivelse")) <> 0 then%>
	<tr bgcolor="<%=bgthis%>">
		<td>&nbsp;</td>
		<td colspan=2 style="padding-top:3px;"><font size=1 color="#5582d2"><%=oRec("beskrivelse")%></font></td>
	</tr>
	<%end if
	%>
	<tr bgcolor="<%=bgthis%>">
		<td height=8 colspan=3><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<%
	
	budgetIalt = budgetIalt + oRec("jobtpris")
	budgettimerIalt = budgettimerIalt + (oRec("budgettimer") + oRec("ikkebudgettimer"))
	
	if len(oRec("restimer")) <> 0 then
	restimerIalt = restimerIalt + oRec("restimer")
	else
	restimerIalt = restimerIalt
	end if
	
	c = c + 1
	oRec.movenext
	wend
	oRec.close 
	
	%></table>
	<br><br>
	Budgetteret omsætning ialt: <b><%=formatcurrency(budgetIalt, 2)%></b><br>
	Forkalkuleret timer ialt: (fakturerbare og ikke fakturerbare timer) <b><%=formatnumber(budgettimerIalt, 2)%></b> <br>
	Ressourcetimer tildelt på medarbejdere ialt: <b><%=formatnumber(restimerIalt, 2)%></b> <br>
	
	
	</td>
	</tr>
	
	
	
	
	
	<%
	'**********************************************************
	'**************** Milepæle *******************************'
	'**********************************************************
	%>
	<tr>
	<td valign=top style="border:2px #5582d2 solid; padding:10px;" bgcolor="#ffffff"><h4>Milepæle</h4>
	<%call periodeForm%>
	<font class=megetlillesort>(Job start- og slut -datoer er ikke vist)</font>
	<br>
	<%
	
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
		
	
	Response.write "<b>Periode:</b> " & formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)
	%>
	<table cellspacing=0 cellpadding=0 border=0 width=400>
	<tr>
		 <td colspan=3 align=right style="border-bottom:1px #999999 dashed; padding:2px;">
			<a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=0&rdir=webblik','650','500','250','120');" target="_self" class=todo_stor>Opret Ny milepæl&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		</td>
	</tr>
	<tr><td style="padding-top:0px;">
	
		
		<table cellspacing=0 cellpadding=0 border=0 width=100%>
		<%
		if func = "sletmilepal" then
		'*** Her slettes en milepal ***
		milepalid = request("milepalid")
		strSQLmil = "DELETE FROM milepale WHERE id = "& milepalid &""
		'response.write strSQLmil
		oConn.execute(strSQLmil)
		end if
		
		
		
		'datointervalstart = dateadd("d", -5, year(now)&"/"& month(now)&"/"&day(now)) 
		'sqlDatostart = year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
		'datointervalslut = dateadd("d", 65, sqlDatostart)  
		'sqlDatoslut = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
		
		lastDato = ""
		strSQL = "SELECT m.navn, m.milepal_dato, m.jid, m.beskrivelse, j.jobnavn, j.jobnr, m.id AS mileid, "_
		&" mt.id, mt.navn AS typenavn, k.kkundenavn, k.kkundenr, j.id AS jobid "_
		&" FROM milepale m "_
		&" LEFT JOIN  job j ON (j.id = m.jid "& strPgrpSQLkri &") "_
		&" LEFT JOIN  kunder k ON (k.kid = j.jobknr) "_
		&" LEFT JOIN milepale_typer mt ON (mt.id = m.type) "_
		&" WHERE milepal_dato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' ORDER BY m.milepal_dato"
		
		'Response.write strSQL
		'Response.flush
		
		oRec.open strSQL, oConn, 3 
		x = 0
		while not oRec.EOF 
		
		if len(trim(oRec("jobnavn"))) <> 0 then
		
			if lastDato <> oRec("milepal_dato") then%>
			<tr><td valign=bottom style="padding-top:10px; border-bottom:1px #999999 dashed; padding:5px;" bgcolor="#d6dff5"><font class=stor-blaa><%=formatdatetime(oRec("milepal_dato"), 2)%></font></td></tr>
			<%end if%>
			
			<tr><td style="border-bottom:1px #999999 dashed; padding:5px;" class=lille>
			<%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>)<br>
			<a href="jobs.asp?menu=job&func=red&id=<%=oRec("jobid")%>&int=1&rdir=webblik" class=todo_mellem><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>)</a><br>
			<font color="red"><%=oRec("typenavn")%>:</font> <a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec("mileid")%>&jid=<%=oRec("jid")%>&rdir=webblik','650','500','250','120');" target="_self" class=rmenu><%=oRec("navn")%></a>&nbsp;&nbsp;
			<a href="webblik.asp?func=sletmilepal&milepalid=<%=oRec("mileid")%>" class=red>[x]</a><br>
			<%if len(oRec("beskrivelse")) <> 0 then%>
			<font color="#999999"><%=oRec("beskrivelse")%></font>
			<%end if%>
			</td></tr>
			
			<%
			lastDato = oRec("milepal_dato")
			x = x + 1
		
		end if
		
		oRec.movenext
		wend
		oRec.close 
		%>
		</table>
	</td></tr></table>
	</td>
	
	
	
	<%
	'**********************************************************
	'**************** Job til fakturering *********************
	'**********************************************************
	%>
	
	
	<td style="border:2px forestgreen solid; padding:10px;" bgcolor="#ffffff" valign=top><h4>Job til fakturering</h4>
	Job <b>uden fakturaer</b>, men med registrerede <b>timer på, indenfor seneste 30 dage.</b><br>
	
	<%
	
	sqlDatoslut = year(now)&"/"& month(now)&"/"&day(now) 
	datointervalslut = dateadd("d", -30, sqlDatoslut)  
	sqlDatostart = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	
	Response.write "<br><b>Periode:</b> " & formatdatetime(day(datointervalslut)&"/"& month(datointervalslut) &"/"& year(datointervalslut), 1) & " - " & formatdatetime(day(now)&"/"& month(now)&"/"&year(now) , 1)
	
	%>
	<table cellspacing=0 cellpadding=0 border=0>
	<tr>
		 <td colspan=2 align=right style="padding:2px;">&nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT j.id, jobnavn, jobnr, jobknr, kkundenavn, jobans2, kkundenr, jobslutdato, "_
	&" j.beskrivelse, jobans1, mnavn, tjobnr, sum(t.timer) AS timer, f.faknr, f.fid, f.fakdato, kid, j.serviceaft, s.navn AS aftalenavn FROM timer t, job j "_
	&" LEFT JOIN kunder ON (kid = jobknr) "_
	&" LEFT JOIN medarbejdere m ON (mid = jobans1)  "_
	&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft)"_
	& "LEFT JOIN fakturaer f ON (f.jobid = j.id AND f.fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"')"_ 
	&" WHERE (tfaktim = 1 AND tdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"') AND (jobnr = tjobnr "& strPgrpSQLkri &") "_
	&" GROUP BY tjobnr ORDER BY jobslutdato, f.fakdato DESC" ' LIMIT 0, 10" 
	
	'Response.write strSQL
	'Response.flush
	
	
	c = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	if oRec("serviceaft") <> 0 then
	bgthis = "#eff3ff"
	else
		'bg = right(c, 1)
		'select case bg
		'case 0,2,4,6,8
		bgthis = "#ffffff"
		'case else
		'bgthis = "WhiteSmoke"
		'end select
	end if
	
	
	if len(oRec("faknr")) > 1 then
	
	else
	%>
	
	<tr bgcolor="<%=bgthis%>">
		
		<td width=300 valign=top style="padding:3px; border-top:1px #2c962d dashed;" class=lille><%=left(oRec("kkundenavn"), 30)%>&nbsp;(<%=oRec("kkundenr")%>)<br>
		<a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&rdir=webblik" class=todo_mellem><%=left(oRec("jobnavn"), 30)%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</a>
		<br>
		<%if oRec("serviceaft") <> 0 then%>
		Aftale: <%=oRec("aftalenavn")%><br>
		<%end if%>
		<font color="red"><%=formatdatetime(oRec("jobslutdato"), 1)%></font> | <font color="#c4c4c4"><i><%=oRec("mnavn")%></i></font>
		</td>
		<td width=150 valign=top  style="padding:3px; border-top:1px #2c962d dashed;">
		<%if oRec("serviceaft") = 0 then%>
		<a href="stat_fak.asp?menu=stat_fak&FM_kunde=<%=oRec("kid")%>&shokselector=1&fakdenne=<%=oRec("jobnr")%>" class=todo_mellem>Opret ny faktura &nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		<%else%>
		<a href="fak_serviceaft_osigt.asp?menu=stat_fak&func=osigtall&id=<%=oRec("kid")%>" class=todo_mellem>Opret ny faktura &nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		<%end if%>
		<br>(Timer: <%=oRec("timer")%>)
		</td>
		
	</tr>
	<tr bgcolor="<%=bgthis%>">
		<td height=8 colspan=2><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<%
	c = c + 1
	
	end if
	oRec.movenext
	wend
	oRec.close 
	
	%></table>
	
	
	</td>
	</tr>
	</table>
	</div>
	
	
	
	
	
	
	<div id=addtodo name=addtodo style="position:absolute; left:90; top:120; visibility:hidden; display:none; width:400; height:200; background-color:#ffffe1; border:2px #cccccc solid; padding:10px;">
	<form action="webblik_oprtodo.asp?func=oprettodo" name=oprredtodo id=oprredtodo method="post">
	<input type="hidden" name="FM_parent" id="FM_parent" value="0">
	<input type="hidden" name="oldone" id="oldone" value="<%=oldones%>">
	<input type="hidden" name="id" id="id" value="0">
	<b>ToDo</b>:<br><input type="text" name="FM_todo" id="FM_todo" value="" style="width:350px;">
	<br><input type="text" name="redi" id="redi" value="" style="border:0px; background-color:#ffffe1; width:150px; font-size:9px;">
	<br><input type="checkbox" name="FM_afsluttet" id="FM_afsluttet" value="1"> <b>ToDo afsluttet.</b> <br>
	
	<br>
	<b>Del denne todo med:</b> Vælg flere ved at holde [ctrl] nede.<br>
	<div style="position:relative; height:120; overflow:auto; background-color:#ffffff; border:1px #cccccc solid; padding:10px;">
	<%
	strSQL = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE mansat <> '2' ORDER BY mnavn"
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	%>
	<input type="checkbox" name="FM_todo_rel" id="FM_todo_rel" value="<%=oRec("mid")%>"> <%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>)<br>
	<%
	oRec.movenext
	wend
	oRec.close 
	%>
	</div>
	
	<!--<input type="textarea" name="FM_todo_rel" id="FM_todo_rel" value="" style="width:350px; height:50px;">-->
	<br>
	<!--<font class=megetlillesort>Du vil altid selv have adgang til denne Todo uanset om email fremgår her.
	<br>Ved oprettelse, vil nye underliggende ToDo's nedarve delte emailadresser</font>
	<br>-->
	<input type="checkbox" name="FM_opdater_u" id="FM_opdater_u" value="1"> <b>Opdater underliggende ToDo's til at følge denne ToDo's deling.</b>
	<br><br>
	<input type="submit" name="Tilføj ToDo" id="submittodo" value="Tilføj ToDo"><br>
	<br><br><a href="#" onClick="hidetodo()">[Luk vindue]</a>
	</form>
	</div>
	
	
	
	
	
	
	
	
	
	
	
	
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->