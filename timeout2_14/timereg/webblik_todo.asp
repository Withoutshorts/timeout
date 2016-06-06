<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->

<%



if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")
case "FM_sortOrder"
Call AjaxUpdate("todo_new","sortorder","")
End select
Response.End
end if

func = request("func")
print = request("print")

thisfile = "webblik_todo.asp"

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

if len(trim(request("nomenu"))) <> 0 then
nomenu = request("nomenu")
else
nomenu = 0
end if
			
			





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
	
	function edittodo(val, todo, todorelmedarb, dato, afl, valparent, forvafsl, tododato_dag, tododato_mrd, tododato_aar, pblic ){
	if (afl == 1) {
	document.getElementById("FM_afsluttet").checked = true
	}
	
	if (forvafsl == 1) {
	document.getElementById("forvafsl").checked = true
}

    if (pblic == 1) {
    document.getElementById("FM_public").checked = true
    }
	
	document.getElementById("FM_start_dag").value = tododato_dag
	document.getElementById("FM_start_mrd").value = tododato_mrd
	document.getElementById("FM_start_aar").value = tododato_aar
	
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
	
	<% 
	if len(request("medid")) <> 0 then
	medid = request("medid")
	response.cookies("webblik")("medid") = medid
	else
	    if len(request.cookies("webblik")("medid")) <> 0 then
	    medid = request.cookies("webblik")("medid")
	    else
	    medid = session("mid")
	    end if
	end if
	
	if medid = 0 then
	useMid = session("mid")
	else
	useMid = medid
	end if

   
	%>
	
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <script src="inc/todo_jav.js"></script>

	
    <% 
    tm = 0
    if print <> "j" AND nomenu <> 1 AND tm = 9999 then %>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(2)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call webbliktopmenu()
	%>
	</div>
	<% end if %>


	
	<div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
	
	
	<%
	oimg = "ikon_todo_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "ToDo's"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	%>
	
	
	<% if print <> "j" then %>
	<table cellspacing=0 cellpadding=0 border=0 width=750>
	
	<tr>
		<td valign=top style="padding:10px;">
	<!--
	<a href="Javascript:history.back();" class=vmenu>
        <img src="../ill/pil_venstre_16.png" border=0 /> Tilbage</a>&nbsp;&nbsp;
	&nbsp;&nbsp;
    -->
   </td></tr>
	</table>
	<% end if %>

	
	
	<table cellspacing=0 cellpadding=0 border=0 width=750 id="incidentlist">
	<form method="post" action="webblik_todo.asp?FM_parent=<%=parent%>&nomenu=<%=nomenu%>">
	<tr>
	    <td colspan=2 style="padding-top:20px;"><h4>Seneste 31 dage:</h4>
        <% if print <> "j" then %>
         <a href="webblik_todo.asp?FM_parent=0&nomenu=<%=nomenu %>" class=vmenu><img src="../ill/folder_up.png" border="0" /> Til øverste niveau</a>
         <%end if %>
	
       
        </td>
	    <td colspan=3 align=right style="padding:4px;">
	    <%
	    todolevel = request("todolevel")
		if todolevel <= 3 then
		if print <> "j" then
		%>
		<a href="#" onClick="showtodo('<%=parent%>','<%=useMid%>')" class=todo_stor>Tilføj ny ToDo <img src="../ill/add2.png" alt="Tilføj ny ToDo" border="0"></a>
		<% end if
		else%>
		<br><b>Der kan maks oprettes 3 niveauer ToDo's!</b>
		<%end if %>
		&nbsp;
	    </td>
	</tr>
	<tr>
		<td colspan=3 bgcolor="#8caae6" style="border-left:1px #5582d2 solid; border-top:1px #5582d2 solid;  padding:5px 10px 10px 10px;">
		<%
		    
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
			
			Response.cookies("webblik")("visafl") = visafl
		
		else
			if request.cookies("webblik")("visafl") = "1" then
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
		
		
    <%
	
	
	
	
	strSQL = "SELECT mnavn, mnr, init, mid FROM medarbejdere WHERE mansat <> '2' AND mansat <> '3' ORDER BY mnavn"%>
	
	<b>Medarbejder:</b> &nbsp;
	<% if print <> "j" then %>
	<select id="medid" name="medid" style="font-size:9px; width:300px;" onchange="submit();">
	<option value=0>Alle</option>
	
	<%
	oRec.open strSQL, oConn, 3
    while not oRec.EOF
    
    if cint(medid) = cint(oRec("mid")) then
    mCHK = "SELECTED"
    else
    mCHK = ""
    end if

    call meStamdata(oRec("mid"))
    %>
    <option value="<%=oRec("mid") %>" <%=mCHK %>><%=meTxt %></option>
    <%
    oRec.movenext
    wend
    Response.Write "</select>"
    else
    if medid = 0 then
    strSQLslut = "<> 0"
    else
    strSQLslut = "= " & cint(medid)
    end if
    strSQL = "SELECT mnavn, mnr, init FROM medarbejdere WHERE mid " & strSQLslut
    oRec.open strSQL, oConn, 3
    if medid = 0 then
    Response.Write "Alle"
    else
    Response.Write(oRec("mnavn"))
    end if
    end if
    oRec.close

    if len(trim(request("FM_sorter"))) <> 0 then
    sort = request("FM_sorter")
    else
    sort = 0
    end if

    if sort = 1 then
    sort1CHK = "CHECKED"
    SQLorby = "t.id DESC"
    else
    sort0CHK = "CHECKED"
    SQLorby = "t.sortorder"
    end if


	 if print <> "j" then
	 %>
 
		<br /><br /><b>Vis afsluttede ToDo's?</b>
		<input type="radio" name="FM_afsluttede" id="FM_afsluttede" value="1" <%=visaflCHK0%> onclick="submit();"> Ja &nbsp;&nbsp;
		<input type="radio" name="FM_afsluttede" id="FM_afsluttede" value="0" <%=visaflCHK1%> onclick="submit();"> Nej &nbsp;&nbsp;

    <br /><br />
    <b>Sorter efter:</b>
    

        <input type="radio" name="FM_sorter" id="Radio1" value="1" <%=sort1CHK%> onclick="submit();"> Senest oprettede &nbsp;&nbsp;
		<input type="radio" name="FM_sorter" id="Radio2" value="0" <%=sort0CHK%> onclick="submit();"> Prioitet &nbsp;&nbsp;<br />

        </td>
 <td colspan=2 bgcolor="#8caae6" style="border-top:1px #5582d2 solid; border-right:1px #5582d2 solid;  padding:10px 10px 10px 10px;">

		<input type="submit" value="Vis ToDo's >> "><br>
		
		<% end if %>

        <br /> 
       Træk i en ToDo for at sortere i rækkefølgen.
	</td>
	
	</tr>
	<!--</table>
	
	
	<table cellspacing=0 cellpadding=0 border=0 width=750 bgcolor="#ffffff">-->
	<%
	'**** Todo ****************'
	if oldones <> "1" then
	sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
	datointervalslut = dateadd("d", -31, sqlDatostart)  
	sqlDatoslut = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	else
	sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
	sqlDatoslut = "2005/1/1"
	end if
	
	
	if visafl = 1 then
	visaflSQL = " AND t.afsluttet <> 99 "
	else
	visaflSQL = " AND t.afsluttet <> 1 "
	end if
	
	if cint(parent) = 0 then
	
	datoSQL = " AND dato BETWEEN '"& sqlDatoslut &"' AND '"& sqlDatostart &"'"
	
	else
		
		
		'**** klikket Todo SQL **************
		strSQL = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, t.afsluttet, t.sortorder, t.tododato, t.forvafsl, t.public "_
		&" FROM todo_new t WHERE t.id = "& parent 
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
					
					'*** Opdaterer dato på redigering på den viste todo ***
					sqlToDoVsitDato = year(now) & "/" & month(now) & "/" & day(now)
					strSQL = "UPDATE todo_new SET dato = '"& sqlToDoVsitDato &"' WHERE id = "& oRec("id")
					oConn.execute(strSQL)
					
					antalsubs = 0
			
					strSQL2 = "SELECT count(t.parent) AS antalsubs FROM todo_new t WHERE t.parent = " & oRec("id") &" "& visaflSQL &" GROUP BY t.parent"
					
					'Response.Write strSQL2
					'Response.flush
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
	
	
	if medid <> 0 then
	medidSQLkri = " medarbid = "& medid &" AND " 
	else
	medidSQLkri = " medarbid <> 0 AND " 
	end if
	
	
	'**** Main Todo SQL **************
	strSQL = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, "_
	&" t.afsluttet, t.sortorder, trn.medarbid, trn.todoid, t.tododato, t.forvafsl, t.public "_
	&" FROM todo_rel_new trn "_
	&" LEFT JOIN todo_new t ON (t.id = trn.todoid "& visaflSQL &" AND parent = "& parent &" "& datoSQL &")"_
	&" WHERE ("& medidSQLkri &" trn.todoid = t.id) OR t.public = 1 GROUP BY t.id ORDER BY "& SQLorby
	
	'definer variabler til csv eksportering
	ekspTxt = "id;dato;tododato;beskrivelse;medarbejdere;medarbejderid;afsluttet;forventet afsluttet;sortorder;"
	ekspTxt = ekspTxt &"xx99123sy#z"
	
	'Response.write strSQL
	'Response.Flush
	x = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
		
			antalsubs = 0
			
			strSQL2 = "SELECT count(parent) AS antalsubs FROM todo_new t WHERE parent = " & oRec("id") &" "& visaflSQL &" GROUP BY parent"
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then 
			    antalsubs = oRec2("antalsubs")
			end if
			oRec2.close 
			
	
			
		acls = "todo_mellem"
		
		
	call todolist(1)
	
	x = x + 1
	
	oRec.movenext
	wend
	oRec.close %>
	<tr><td bgcolor="#ffffff" colspan=5 style="border-left:1px #8cAAe6 solid; border-right:1px #8cAAe6 solid; border-bottom:1px #8cAAe6 solid; padding:10px 10px 10px 10px;">
		Der er ialt: <b><%=x %></b> ToDo's på listen.</td></tr>
	</table>
	</form>
	
	
	<br /><br />
	<h4>Gamle ToDo's:</h4>
	<table cellspacing=0 cellpadding=0 border=0 width=750 bgcolor="#ffffff">
	
	<tr>
		<td bgcolor="#8caae6" style="border:1px #5582d2 solid; padding:10px 10px 10px 10px;">
		
	<%
	'*** Mere end 1 md gamle ToDo's ***********************
	if cint(parent) = 0 then
	
			if request("FM_visalle") = "1" then
			Response.cookies("webblik")("visallegamle") = "j"
			fmallechk0 = ""
			fmallechk1 = "CHECKED"
			visalle="1"
			else
			Response.cookies("webblik")("visallegamle") = "n"
			fmallechk0 = "CHECKED"
			fmallechk1 = ""
			end if
			
			Response.cookies("webblik").expires = date + 40
			%>
			<b>Mere end en 31 dage gamle ToDo's:</b> (maks 200)<br>
			<% if print <> "j" then %>
			<form method=post action="webblik_todo.asp?FM_parent=<%=parent%>&nomenu=1">
			Vis kun maks 4 md. gamle  <input type="radio" name="FM_visalle" id="FM_visalle" value="0" <%=fmallechk0%>>&nbsp;&nbsp;
			Vis alle:<input type="radio" name="FM_visalle" id="FM_visalle" value="1" <%=fmallechk1%>>
			
			
			<br />Vis afsluttede ToDo's?
		    <input type="radio" name="FM_afsluttede" id="FM_afsluttede" value="1" <%=visaflCHK0%>> Ja &nbsp;&nbsp;
		    <input type="radio" name="FM_afsluttede" id="FM_afsluttede" value="0" <%=visaflCHK1%>> Nej &nbsp;&nbsp;
		    <input type="submit" value="Vis" style="width:30; font-size:10px;">
		
			<% end if %>
			</td>
			</tr>
			</form>
			<tr>
			<td style="padding:20px 20px 10px 10px; border:1px #8cAAe6 solid; border-top:0px;">
			
			<!--<div style="position:relative; height:200px; overflow:auto;">-->
			<%
			
			sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
			datointervalslut = dateadd("d", -32, sqlDatostart)  
			sqlDatostart = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
			
			if request.cookies("webblik")("visallegamle") = "j" then
			sqlDatoslut = "2002/1/1"
			else
			sqlDatoslut_temp = dateadd("d", -120, day(now)&"/"& month(now)&"/"&year(now)) 
			sqlDatoslut = year(sqlDatoslut_temp)&"/"& month(sqlDatoslut_temp)&"/"&day(sqlDatoslut_temp) 
			end if
			
			
			strSQL3 = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, "_
			&" trn.medarbid, trn.todoid, t.afsluttet, t.tododato, t.forvafsl, t.sortorder, t.public"_
			&" FROM todo_rel_new trn "_
			&" LEFT JOIN todo_new t ON (t.id = trn.todoid "& visaflSQL &" AND (parent = "& parent &" "& hentklikketSQL &")"_
			&" AND dato BETWEEN '"& sqlDatoslut &"' AND '"& sqlDatostart &"') "_
			&" WHERE ("& medidSQLkri &" trn.todoid = t.id) OR t.public = 1 GROUP BY t.id ORDER BY t.dato DESC LIMIT 200" 
			
			'Response.write "<br><br>" & strSQL3
			'Response.flush
			
			oRec3.open strSQL3, oConn, 3 
			while not oRec3.EOF 
				
			%>
			<a href="webblik_todo.asp?FM_parent=<%=oRec3("id")%>&todolevel=<%=oRec3("level")%>&oldone=1&nomenu=1" class="todo_lille">
			<% if oRec3("afsluttet") = 1 then%>
			<s>
			<%end if%>
			
			<%=formatdatetime(oRec3("dato"), 2) &" - "&oRec3("navn")%>
			
			
			<%if oRec3("forvafsl") <> 0 then %>
		    &nbsp;<font class=megetlillesort>Forv. udført: <%=formatdatetime(oRec3("tododato"), 2) %></font>
		    <%end if%>
			
			<% if oRec3("afsluttet") = 1 then%>
			</s>
			<%end if%>
			</a>
			 
			 
			 &nbsp;<a href="webblik_oprtodo.asp?func=slet&id=<%=oRec3("id")%>&FM_parent=<%=parent%>&oldone=<%=oldones%>&nomenu=1" class="red">[x]</a>&nbsp;&nbsp; 
			
			        
			        
			 <%'*** Andre medarbejdere med adgang ***
			
			medarbNavn = ""
			strSQL2 = "SELECT medarbid, email, mid, mnavn FROM todo_rel_new "_
			&" LEFT JOIN medarbejdere ON (mid = medarbid) WHERE todoid = " & oRec3("id") &" GROUP BY mid ORDER BY mid" ' & " AND medarbid <> " & session("mid")
			'Response.write strSQL2
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF  
				medarbNavn = medarbNavn & oRec2("mnavn") & ", "
			oRec2.movenext
			wend
			oRec2.close 
		
		
		
		    len_medarbNavn = len(medarbNavn)
		    left_medarbNavn = left(medarbNavn, len_medarbNavn - 2)
		    medarbNavn = trim(left_medarbNavn)
		    %>
    		
	        <font class=megetlillesilver> - <%=left(medarbNavn, 50) %></font><br />
			        
			      
			<%

            if len(trim(oRec3("tododato"))) <> 0 then
            tdDato = oRec3("tododato")
            else
            tdDato = ""
            end if

            
            htmlparseCSV(oRec3("navn"))
            todoTxt = htmlparseCSVTxt
            todoTxt = replace(todoTxt, "''", "")
            todoTxt = replace(todoTxt, "'", "")

			ekspTxt = ekspTxt &Replace(oRec3("todoid"),";",":")
			ekspTxt = ekspTxt &Replace(oRec3("dato"),";",":")
			ekspTxt = ekspTxt &";"&Replace(tdDato,";",":")
			ekspTxt = ekspTxt &";"& todoTxt &";"&Replace(medarbNavn,";",":")&";"
			ekspTxt = ekspTxt &Replace(oRec3("medarbid"),";",":")&";"&Replace(oRec3("afsluttet"),";",":")&";"
			ekspTxt = ekspTxt &Replace(oRec3("forvafsl"),";",":")&";"&Replace(oRec3("sortorder"),";",":")&";"
            ekspTxt = ekspTxt &"xx99123sy#z"
			

            htmlparseCSV(ekspTxt)
            ekspTxt = htmlparseCSVTxt
			
			oRec3.movenext
			wend
			oRec3.close 
	
	end if
	%>
	<!--</div>-->
	&nbsp;</td></tr>
	</table>
    <br /><br />&nbsp;


	<%if print <> "j" then

ptop = 130
pleft = 800
pwdt = 120

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="webblik_todo_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=ekspTxt%>">
            <input type="hidden" name="txt20" id="txt20" value="">

<tr>
    <td align=center><input type="image" src="../ill/export1.png">
    </td>
    </td><td>.csv fil eksport</td>
    </tr>
    <tr>
    <td align=center><a href="webblik_todo.asp?print=j&FM_visalle=<%=visalle %>&FM_parent=<%=parent%>&oldone=<%=oldones%>&lastid=<%=lastId%>" target="_blank"  class='vmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
</td><td>Print version</td>
   </tr>
   
	</form>
   </table>
</div>
	<%
	else
	Response.Write("<script language=""JavaScript"">window.print();</script>")
	end if %>

	
	
	
	<div id=addtodo name=addtodo style="position:absolute; left:180px; top:10px; visibility:hidden; display:none; width:400px; height:500px; background-color:#ffffe1; border-left:1px #cccccc solid; border-top:1px #cccccc solid; border-right:2px #999999 solid; border-bottom:2px #999999 solid; padding:10px;">
	<form action="webblik_oprtodo.asp?func=oprettodo&nomenu=<%=nomenu%>" name=oprredtodo id=oprredtodo method="post">
	<input type="hidden" name="FM_parent" id="FM_parent" value="0">
	<input type="hidden" name="oldone" id="oldone" value="<%=oldones%>">
	<input type="hidden" name="id" id="id" value="0">
	
	<%
	oimg = "ikon_nytodo_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "ToDo"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	%>
	
	
        <textarea id="FM_todo" name="FM_todo" cols="45" rows="3"></textarea>
	<br><input type="text" name="redi" id="redi" value="" style="border:0px; background-color:#ffffe1; width:150px; font-size:9px;">
	
	
	<br>
	<b>Del denne todo med:</b> (Vælg <a href="#" id="a_sel_all" class=vmenu>alle</a> / <a href="#" id="a_sel_none" class=vmenu>ingen</a>)<br>
	<div style="position:relative; height:120; overflow:auto; background-color:#ffffff; border:1px #cccccc solid; padding:10px;">
	<%
	strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> '2' AND mansat <> '3' ORDER BY mnavn"
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 

    call meStamdata(oRec("mid"))
	%>
	<input type="checkbox" name="FM_todo_rel" id="FM_todo_rel" class="cls_todo_rel" value="<%=oRec("mid")%>"> <%=meTxt %><br>
	<%
	oRec.movenext
	wend
	oRec.close 
	%>
	</div>
	
	<!--<input type="textarea" name="FM_todo_rel" id="FM_todo_rel" value="" style="width:350px; height:50px;">-->
	<br>
    <input type="checkbox" name="FM_public" id="FM_public" value="1" <%=pubCHK %>> <b>Offentlig</b> (vis på timereg. siden for alle medarb.)<br />
    <input type="checkbox" name="FM_afsluttet" id="FM_afsluttet" value="1"> <b>ToDo afsluttet.</b> <br>
	
	<!--<font class=megetlillesort>Du vil altid selv have adgang til denne Todo uanset om email fremgår her.
	<br>Ved oprettelse, vil nye underliggende ToDo's nedarve delte emailadresser</font>
	<br>-->
	<input type="checkbox" name="FM_opdater_u" id="FM_opdater_u" value="1"> <b>Opdater underliggende ToDo's til at følge denne ToDo's deling.</b>
	<br>
	
	<%if cint(forvafsl) = 1 then
	forvafslCHK = "CHECKED"
	else
	forvafslCHK = ""
	end if %>
	
    <input id="forvafsl" name="forvafsl" type="checkbox" value="1" <%=forvafslCHK%> > <b>Angiv forventet udførelsesdato.</b>
	<table cellspacing="0" cellpadding="0" border="0">
	<tr>
	
	<%
	
	if forvafsl = 1 then
	strDag = day(tododato)
	strMrd = month(tododato)
	strAar = year(tododato)
	else
	strDag = day(now())
	strMrd = month(now())
	strAar = right(year(now()), 4)
	end if  
	
	%>
	<!--#include file="inc/dato2_b.asp"-->
		<td style="padding-left:10;"><br>Dato:&nbsp;</td>
		<td><br>
		<select name="FM_start_dag" id="FM_start_dag">
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
		<option value="31">31</option></select>
		
		<select name="FM_start_mrd" id="FM_start_mrd">
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
		
		
		<select name="FM_start_aar" id="FM_start_aar">
		<option value="<%=strAar%>"><%=strAar%></option>
	    <option value="2006">2006</option>
		<option value="2007">2007</option>
		<option value="2006">2006</option>
		<option value="2007">2007</option>
		<option value="2008">2008</option>
		<option value="2009">2009</option>
		<option value="2010">2010</option>
		<option value="2011">2011</option>
		<option value="2012">2012</option></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		</td>
	</tr>
	</table>
	<br /><br />
	<input type="submit" name="Tilføj ToDo" id="submittodo" value="Tilføj ToDo"><br>
	<br><br><a href="#" onClick="hidetodo()">[Luk vindue]</a>
	</form>
	</div>
	
	
	
	
	
	
	
	
	
	
	
	
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->