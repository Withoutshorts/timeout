<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	level = session("rettigheder")
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	'Response.write "request(wk)" & request("wk")
	
	if len(request("wk")) <> 0 then
	useWK = datepart("ww", request("wk"), 2, 3)
	sqlweekdato = request("wk") 'year(request("wk"))&"/"&month(request("wk"))&"/"&day(request("wk"))
	else
	useWK = datepart("ww", now, 2, 3)
	sqlweekdato = month(now)&"/"&day(now)&"/"&year(now) 'year(now)&"/"&month(now)&"/"&day(now)
	end if
	
	if len(request("FM_visstatus")) <> 0 then
	intVisstatus = request("FM_visstatus")
	else
	intVisstatus = 1 '**vis kun aktive
	end if
	
	
	if len(request("FM_viskunde")) > 0 then
	intViskunde = request("FM_viskunde")
	else
	intViskunde = 0
	end if
	
	if len(request("FM_vismedarb")) <> 0 then
	intVismedarb = request("FM_vismedarb")
	else
	intVismedarb = 0
	end if
	
	if len(request("FM_periodesel")) <> 0 then
	persel = request("FM_periodesel")
	else
	persel = 6
	end if
	
	select case func
	case "slet"
	
	
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<div id="sindhold" style="position:absolute; left:210; top:180; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF">
	<tr>
	    <td style="border:1px #8CAAE6 solid; padding-right:20; padding-left:20; padding-top:20; padding-bottom:20;"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0"><br>
		Du er ved at <b>slette</b> en opgave.
		<br>Er du sikker på at du vil slette denne opgave?<br><br>
		<a href="todo_add.asp?func=sletja&&FM_vismedarb=<%=intVismedarb%>&FM_viskunde=<%=intViskunde%>&id=<%=id%>&FM_visstatus=<%=intVisstatus%>&FM_periodesel=<%=persel%>&wk=<%=request("wk")%>">Ja, slet opgave!</a>
		<img src="ill/blank.gif" width="50" height="1" alt="" border="0"><a href="javascript:history.back()">Nej, Annuler!</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletja"
		
		strSQL = "DELETE FROM todo WHERE id = "& id &""
		oConn.execute(strSQL)
		Response.redirect "todo_add.asp?func=list&FM_viskunde="&intViskunde&"&FM_vismedarb="&intVismedarb&"&FM_visstatus="&intVisstatus&"&FM_periodesel="&persel&"&wk="&request("wk")
	
	case "changeudfort"
		
		thisid = request("FM_udfort_id")
		status = request("FM_udfort")
		
		strSQL = "UPDATE todo SET udfort = "& status &" WHERE id = "& thisid &""
		oConn.execute(strSQL)
		Response.redirect "todo_add.asp?func=list&FM_viskunde="&intViskunde&"&FM_vismedarb="&intVismedarb&"&id="&thisid&"&FM_visstatus="&intVisstatus&"&FM_periodesel="&persel&"&wk="&request("wk")
		
	case "moveup"
	
		lastID = request("lastid")
		lastsoID = request("lastso")
		thissoID = request("thisso")
		
		if len(lastID) > 0 then
		'** opdaterer forrrige **
		oConn.execute("UPDATE todo SET sortorder = "& thissoID &" WHERE id = "& lastID &"")
		'** opdaterer denne **
		oConn.execute("UPDATE todo SET sortorder = "& lastsoID &" WHERE id = "& id &"")
		end if
		
		Response.redirect "todo_add.asp?func=list&FM_viskunde="&intViskunde&"&FM_vismedarb="&intVismedarb&"&id="&id&"&FM_visstatus="&intVisstatus&"&FM_periodesel="&persel&"&wk="&request("wk")
	
	case "movedown"
		
		oprsortorder = request("thisso")
		
		strSQL = request("useSQL")
		oRec.open strSQL, oConn, 3
		usenext = "n"
		'stop = "n"
		while not oRec.EOF 'AND stop <> "j" 
			
			if usenext = "j" then
			nextsortorder = oRec("sortorder")
			nextID = oRec("id")
			'stop = "j"
			end if
			
			if cint(oRec("id")) = cint(id) then
			usenext = "j"
			else
			usenext = "n"
			end if
			
		oRec.movenext
		wend
		oRec.close
		
		if len(nextID) > 0 then
			'** opdaterer forrrige **
			oConn.execute("UPDATE todo SET sortorder = "& oprsortorder &" WHERE id = "& nextID &"")
			'** opdaterer denne **
			oConn.execute("UPDATE todo SET sortorder = "& nextsortorder &" WHERE id = "& id &"")
		end if
		
		Response.redirect "todo_add.asp?func=list&FM_viskunde="&intViskunde&"&FM_vismedarb="&intVismedarb&"&id="&id&"&FM_visstatus="&intVisstatus&"&FM_periodesel="&persel&"&wk="&request("wk")
	
	case "dbopr", "dbred"
		
		'*** validering ***
		
		if len(request("FM_emne")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<%
		errortype = 43
		call showError(errortype)
		
		else
		
			'*tjekker om startdag eksisterer ** 
			if Request("FM_start_dag") > 28 then 
			select case Request("FM_start_mrd")
			case "2"
			strStartDay = 28
			case "4", "6", "9", "11"
			strStartDay = 30
			case else
			strStartDay = Request("FM_start_dag")
			end select
			else
			strStartDay = Request("FM_start_dag")
			end if
			
			
			'*tjekker om slutdag eksisterer ** 
			if Request("FM_start_dag_2") > 28 then 
			select case Request("FM_start_mrd_2")
			case "2"
			strSlutDay = 28
			case "4", "6", "9", "11"
			strSlutDay = 30
			case else
			strSlutDay = Request("FM_start_dag_2")
			end select
			else
			strSlutDay = Request("FM_start_dag_2")
			end if
	
		
			dtStartdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&strStartDay
			dtSlutdato = request("FM_start_aar_2")&"/"&request("FM_start_mrd_2")&"/"&strSlutDay
		
			if cdate(dtStartdato) > cdate(dtSlutdato) then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			<%
			errortype = 44
			call showError(errortype)
		
			else	
					'** er der valgt medarbejdre **
					usedpg = request("FM_usedpg")
					t = 0
					erdervalgtmedarb = "0,"
					Dim validMyPGArray
					validMyPGArray = split(usedpg, ", ")
					for t = 0 to Ubound(validMyPGArray)
						erdervalgtmedarb = erdervalgtmedarb & request("pg_"&validMyPGArray(t)&"")
					next
					
					if len(erdervalgtmedarb) <= 2 then
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
					<%
					errortype = 45
					call showError(errortype)
					
					else
		
						
						function SQLBless2(s)
						dim tmp
						tmp = s
						tmp = replace(tmp, "'", "''")
						SQLBless2 = tmp
						end function
						
						dtdato = year(now)&"/"&month(now)&"/"&day(now)
						strEmne = SQLBless2(request("FM_emne"))
						strTekst = SQLBless2(request("FM_tekst"))
						intStatus = request("FM_status")
						intKundeid = request("FM_kunde")
						intPrio = request("FM_priotet")
						intOprsom = request("FM_nummer")
						
						
						if func = "dbred" then
						
							strSQL = "UPDATE todo SET "_
							& " editor = '"& session("user") &"', "_
							& " dato = '"& dtdato &"', "_
							& " startdato = '"& dtStartdato &"', "_
							& " slutdato = '"& dtSlutdato &"', "_
							& " emne = '"& strEmne &"', "_ 
							& " beskrivelse = '"& strTekst &"', "_
							& " mid = "& session("mid") &", "_
							& " udfort = "& intStatus &", "_
							& " klassificering = "& intPrio &", "_
							& " kundeid = "& intKundeid &" WHERE id = "& id
							
							oConn.execute(strSQL)
							 
							thistodoid = id
							
								'*** sletter tidligere relationer ***
								strSQL3 = "DELETE FROM todo_rel WHERE todoid =" & id
								oConn.execute(strSQL3)
								
						 
						else
							'*** Henter sortorder ***
							if intOprsom = "0" then 'nederst
							
								strSQL1 = "SELECT sortorder FROM todo ORDER BY sortorder DESC"
								
								oRec.open strSQL1, oConn, 3
								if not oRec.EOF then 
								usenewsortorder = oRec("sortorder") + 1
								else
								usenewsortorder = 1
								end if
								oRec.close
							
							else
								
								strSQL1 = "SELECT id, sortorder FROM todo WHERE sortorder >= "& intOprsom &" ORDER BY sortorder"
								
								oRec.open strSQL1, oConn, 3
								while not oRec.EOF
									thisneworder = oRec("sortorder") + 1 
									oConn.execute("UPDATE todo SET sortorder = "& thisneworder &" WHERE id = " & oRec("id") &"")
								oRec.movenext
								wend
								oRec.close
								
								usenewsortorder = intOprsom
							
							end if
							
							'*** opretter ny opgave *** 
							strSQL = "INSERT INTO todo (editor, dato, startdato, slutdato, emne, beskrivelse, mid, udfort, sortorder, klassificering, kundeid)"_
							& " VALUES ('"& session("user") &"', '"& dtdato &"', "_
							& "'"& dtStartdato &"', '"& dtSlutdato &"', "_
							& "'"& strEmne &"', '"& strTekst &"', "& session("mid") &", "& intStatus &", "& usenewsortorder &", "_
							& ""& intPrio &", "& intKundeid &")"
							
							oConn.execute(strSQL)
							
							'*** Henter ToDo Id ***
							strSQL4 = "SELECT id FROM todo ORDER BY id DESC"
							oRec.open strSQL4, oConn, 3
							if not oRec.EOF then 
							thistodoid = oRec("id") 
							else
							thistodoid = 0
							end if
							oRec.close
						
						end if
						
						
						'** Opret medarbrelationer **
						t = 0
						m = 0
						Dim MyPGArray
						Dim MyMedarbArray 
						medarbiswritten = "#,"
						MyPGArray = split(usedpg, ", ")
						for t = 0 to Ubound(MyPGArray)
						'Response.write MyPGArray(t) & "<br>"
						'Response.write request("pg_"&MyPGArray(t)&"") &"<br><br>"
						
								MyMedarbArray = split(request("pg_"&MyPGArray(t)&""), ", ")
								for m = 0 to Ubound(MyMedarbArray)
									if instr(medarbiswritten, "#"&MyMedarbArray(m)&"#") = 0 then
									strSQL2 = "INSERT INTO todo_rel (medarbid, todoid) VALUES ("& MyMedarbArray(m) &", "& thistodoid &")"
									oConn.execute(strSQL2)
									medarbiswritten = medarbiswritten &"#"&MyMedarbArray(m)&"#," & "(" & thistodoid & ")"
									'Response.write "medarbiswritten" & medarbiswritten & "<br>"
									else
									medarbiswritten = medarbiswritten
									end if 
								next
						next
						
				Response.redirect "todo_add.asp?func=list&FM_viskunde="&intViskunde&"&FM_vismedarb="&intVismedarb&"&id="&thistodoid&"&FM_visstatus="&intVisstatus&"&FM_periodesel="&persel&"&wk="&request("wk")
				
				end if 'validering
			end if 'validering
		end if 'validering
	
	case "opr", "red", "list"
	%>
	<html>
		<head>
		<title>TimeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css">
		<script LANGUAGE="javascript">
		function chkthis(field)
		{
		usefield = document.all["pg_"+field+""]
		usefield.checked = true;
		for (i = 0; i < usefield.length; i++)
		  usefield[i].checked = true ;
		}
		
		function unchkthis(field)
		{
		usefield = document.all["pg_"+field+""]
		usefield.checked = false;
		for (i = 0; i < usefield.length; i++)
			usefield[i].checked = false ;
		}
		
		function NewWin_popupcal(url) {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
			
			
			function showudfort(thisid){
			document.getElementById("udfort").style.display = "";
			document.getElementById("udfort").style.visibility = "visible";
			document.getElementById("udfort").style.left = window.event.x - 150;
			document.getElementById("udfort").style.top = window.event.y - 10;
			document.getElementById("FM_udfort_id").value = thisid
			}
			
			function closeudfort(){
			document.getElementById("udfort").style.display = "none";
			document.agetElementById("udfort").style.visibility = "hidden";
			}
			
		
			function showtodotekst(thisid){
		
			if (document.getElementById("lastvalue").value > 0){
			closethis = document.getElementById("lastvalue").value;
			document.getElementById("b_"+closethis+"").style.display = "none";
			}
			document.getElementById("b_"+thisid+"").style.display = "";
			document.getElementById("b_"+thisid+"").style.visibility = "visible";
			document.getElementById("b_"+thisid+"").style.left = window.event.x - 150;
			document.getElementById("b_"+thisid+"").style.top = window.event.y - 10;
			document.getElementById("lastvalue").value = thisid
			}
			
			function closetodotekst(thisid){
			document.getElementById("b_"+thisid+"").style.display = "none";
			document.getElementById("b_"+thisid+"").style.visibility = "hidden";
			}
	
		</script>
		</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<!--#include file="inc/convertDate.asp"-->
	<%
	if func = "opr" OR func = "red" then
		
			if func = "red" then
				dbfunc = "dbred"
				submitknap = "opdaterpil"
				
				strSQL = "SELECT editor, dato, startdato, slutdato, emne, beskrivelse, mid, udfort, sortorder, klassificering, kundeid FROM todo WHERE id = "& id
				
				oRec.open strSQL,oConn, 3
				while not oRec.EOF 
					
					dtStartdato = oRec("startdato")
					useopfdato = oRec("slutdato")
					strDag = day(dtStartdato)
					strMrd = month(dtStartdato)
					strAar = year(dtStartdato)
					stremne = oRec("emne")
					strtekst = oRec("beskrivelse")
					intStatus = oRec("udfort")
					intPrio = oRec("klassificering")
					intKundeID = oRec("kundeid")
					
				oRec.movenext
				wend
				oRec.close
			
			else
				dbfunc = "dbopr"
				submitknap = "opretpil"
				intKundeID = intViskunde
				
				if len(request("strdag")) <> 0 then
				strDag = request("strdag")
				strMrd = request("strmrd")
				strAar = request("straar")
				else
				strDag = day(now)
				strMrd = month(now)
				strAar = year(now)
				end if
				
				useopfdato = dateadd("d", 0, strDag&"/"&strMrd&"/"&strAar)
			end if
			 
			%>
			<!--#include file="inc/dato2_b.asp"-->
			<!-------------------------------Sideindhold------------------------------------->
			<div id="sindhold" style="position:absolute; left:10; top:10; visibility:visible;">
			<table border=0 cellpadding=0 cellspacing=0>
			<tr>
			<td valign="top"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
			<td valign="bottom">Opret eller rediger en ny opgave til <b>opgavelisten</b> her.<br>
			</td></tr>
			</table>
			<br>
			<table cellspacing="1" cellpadding="0" border="0" bgcolor="#d6dff5">
			<form action="todo_add.asp?func=<%=dbfunc%>&id=<%=id%>&FM_viskunde=<%=intViskunde%>&FM_vismedarb=<%=intVismedarb%>&FM_visstatus=<%=intVisstatus%>&FM_periodesel=<%=persel%>" method="post" name="pgmedarbsel">
			<input type="hidden" name="wk" value="<%=request("wk")%>">
			<tr>
				<td>Kunde&nbsp;|&nbsp;Job:</td><td><select name="FM_kunde">
				<%
				strSQL = "SELECT job.id, jobnavn, jobnr, jobknr, kid, kkundenavn, kkundenr FROM job, kunder WHERE jobstatus = 1 AND kid = jobknr ORDER BY kkundenavn, jobnavn"
				
				oRec.open strSQL, oConn, 3
				foundone = "n"
				while not oRec.EOF
				
				if cint(intKundeID) = cint(oRec("id")) then
				ksel = "SELECTED"
				foundone = "j"
				else
				ksel = ""
				foundone = foundone
				end if%>
				
				<option value="<%=oRec("id")%>" <%=ksel%>><%
				if len(oRec("kkundenavn")) > 25 then
				Response.write left(oRec("kkundenavn"), 25)&".."
				else
				Response.write oRec("kkundenavn")
				end if
				
				if len(oRec("jobnavn")) > 35 then
				Response.write "&nbsp;&nbsp;|&nbsp;&nbsp;"& left(oRec("jobnavn"), 35)&".."
				else
				Response.write "&nbsp;&nbsp;|&nbsp;&nbsp;"& oRec("jobnavn") 
				end if
				
				Response.write "&nbsp;&nbsp;&nbsp;(" & oRec("jobnr") & ")"%>
				</option>
				
				
				<%oRec.movenext
				wend
				oRec.close
				
				%>
				<!--if foundone = "j" then%>
				<option value="0">Ikke job relateret (Alle)</option>
				<else%>
				<option value="0" SELECTED>Ikke job relateret (Alle)</option>
				<end if%>-->
			</select>&nbsp;(kun aktive job.)</td>
			</tr>
			<tr>
				<td>Start dato:</td><td>
				<select name="FM_start_dag">
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
				<option value="<%=strAar%>"><%=strAar%></option>
				<option value="2002">2002</option>
				<option value="2003">2003</option>
			   	<option value="2004">2004</option>
			   	<option value="2005">2005</option>
				<option value="2006">2006</option>
				<option value="2007">2007</option></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
				</td>
			</tr>
			<tr>
				<td>Slut dato:</td><td>
				<%
						useopfdag = day(useopfdato)
						useopfmrd = month(useopfdato)
						useopfmrdNavn = left(monthname(useopfmrd), 3)
						useopfaar = year(useopfdato) 
				%>
						<select name="FM_start_dag_2">
						<option value="<%=useopfdag%>"><%=useopfdag%></option> 
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
						
						<select name="FM_start_mrd_2">
						<option value="<%=useopfmrd%>"><%=useopfmrdNavn%></option>
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
						
						<select name="FM_start_aar_2">
						<option value="<%=useopfaar%>"><%=useopfaar%></option>
						<option value="2002">2002</option>
						<option value="2003">2003</option>
					   	<option value="2004">2004</option>
					   	<option value="2005">2005</option>
						<option value="2006">2006</option>
						<option value="2007">2007</option>
						</select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=2')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
					</td>
			</tr>
			<tr>
				<td>Emne: </td><td><input type="text" name="FM_emne" value="<%=stremne%>" size="30">&nbsp;&nbsp;(evt. aktivitet)</td>
			</tr>
			<tr>
				<td valign=top>Tekst:</td><td><textarea cols="40" rows="4" name="FM_tekst" id="FM_tekst"><%=strtekst%></textarea></td>
			</tr>
			<%if func = "opr" then%>
			<tr>
				<td colspan=2>Opret som nummer: <select name="FM_nummer">
				<option value="0" SELECTED>Nederst</option>
				<%
				strSQL = "SELECT id, sortorder FROM todo WHERE sortorder > 1 ORDER BY sortorder DESC"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF 
				%>
				<option value="<%=oRec("sortorder")%>"><%=oRec("sortorder")%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
				<option value="1">1</option>
		    </select> på opgavelisten.</td>
			</tr>
			<%end if%>
			<tr>
				<td>Prioitet:</td><td><select name="FM_priotet">
				<%
				if func = "red" then
					select case intPrio
					case 1
					pr1 = "SELECTED"
					pr2 = ""
					pr3 = ""
					case 2
					pr1 = ""
					pr2 = "SELECTED"
					pr3 = ""
					case 3
					pr1 = ""
					pr2 = ""
					pr3 = "SELECTED"
					end select
				else
				pr1 = ""
				pr2 = "SELECTED"
				pr3 = ""
				end if
				%>
				<option value="1" <%=pr1%>>Høj</option>
				<option value="2" <%=pr2%>>Middel</option>
				<option value="3" <%=pr3%>>Lav</option>
			</select></td>
			</tr>
			<tr>
				<td>Status:</td><td><select name="FM_status">
				<%if func = "red" then
					select case intStatus
					case 1
					st1 = "SELECTED"
					st2 = ""
					st3 = ""
					st4 = ""
					case 2
					st1 = ""
					st2 = "SELECTED"
					st3 = ""
					st4 = ""
					case 3
					st1 = ""
					st2 = ""
					st3 = "SELECTED"
					st4 = ""
					case 4
					st1 = ""
					st2 = ""
					st3 = ""
					st4 = "SELECTED"
					end select
				else
				st1 = "SELECTED"
				st2 = ""
				st3 = ""
				st4 = ""
				end if%>
				<option value="1" <%=st1%>>Igang</option>
				<option value="2" <%=st2%>>Venter</option>
				<option value="3" <%=st3%>>Afventer anden opgave</option>
				<option value="4" <%=st4%>>Afsluttet</option>
			</select></td>
			</tr>
			<%if func = "red" then%>
			<tr>
				<td colspan=2><br><b>Medarbejdere med adgang til denne opgave</b>:<br><%
						strcheckmedarb = "0#"
						strSQL2 = "SELECT medarbid, mnavn, mid FROM todo_rel, medarbejdere WHERE mid = medarbid AND todoid = "& id &" ORDER BY mnavn"
						oRec2.open strSQL2, oConn, 3
							while not oRec2.EOF 
							%>
							<input type="checkbox" CHECKED name="pg_0" id="pg_0" value="<%=oRec2("mid")%>">&nbsp;&nbsp;<%=oRec2("mnavn")%><br>
							<%
							oRec2.movenext
							wend
						oRec2.close
					%>
					<input type="hidden" name="FM_usedpg" value="0">
					</td>
			</tr>
			<%	
			end if
			%>
			<tr>
				<td colspan=2><br><b>Tilføj de medarbejdere der skal have adgang til denne opgave:</b><br>
				&nbsp;</td>
			</tr>
			</table>
			<table cellspacing="1" cellpadding="0" border="0" bgcolor="#d6dff5">
			<tr>
			<%
				
				strSQL = "SELECT id, navn FROM projektgrupper WHERE id = 10 ORDER BY navn"
				oRec.open strSQL, oConn, 0, 1
				x = 1
				while not oRec.EOF
				
				select case x
				case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80 'Antal projektgrupper
				%>
				</tr><tr>
				<%
				end select
				%> 
				
					<td width="200" valign="top"><!--<b><=oRec("navn")%></b><br>-->
					<a href="#" onClick="chkthis('<%=oRec("id")%>');" class="vmenuglobal"><img src="../ill/alle.gif" width="43" height="24" alt="" border="0"></a>&nbsp;&nbsp;<a href="#" onClick="unchkthis('<%=oRec("id")%>');" class="vmenualt"><img src="../ill/ingen.gif" width="43" height="24" alt="" border="0"></a><br>
					<%strSQL3 = "SELECT medarbejderid, projektgruppeid, mid, mnavn FROM progrupperelationer LEFT JOIN medarbejdere ON (mid = progrupperelationer.medarbejderid) WHERE projektgruppeid = "& oRec("id") &" AND mnavn <> '' ORDER BY mnavn"
					oRec3.open strSQL3, oConn, 0, 1
					while not oRec3.EOF
					%>
					<input type="checkbox" name="pg_<%=oRec("id")%>" id="pg_<%=oRec("id")%>" value="<%=oRec3("mid")%>">&nbsp;&nbsp;<%=oRec3("mnavn")%><br>
					<%
					oRec3.movenext
					wend
					oRec3.close
					%><br>
					</td>
				<input type="hidden" name="FM_usedpg" value="<%=oRec("id")%>">
				<%
				x = x + 1
				oRec.movenext
				wend
				oRec.close
				%>
			
			</tr>
			<tr>
				<td colspan=3 align=center>
				<input type="image" src="../ill/<%=submitknap%>.gif"></td>
			</tr>
			</form>
			</table>
		<%end if 'opr/red
		
		
		if func = "list" then
		%>
		<!-------------------------------Sideindhold------------------------------------->
			<div id="sindhold" style="position:absolute; left:10; top:10; visibility:visible;">
			<table border=0 cellpadding=0 cellspacing=0>
			<tr>
			<td valign="top"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
			<td width=330><b>Opgavelisten.</b><br>
			Vælg f.eks "alle" kunder, samt den ønskede medarbejder for at se hver enkelt medarbejders opgaveliste.<br>
			Der bliver <u>kun</u> vist opgaver fra <u>aktive</u> job.
			</td><td valign="bottom">
			<%if level <= 3 OR level = 6 then%>
			<a href="todo_add.asp?func=opr&FM_vismedarb=<%=intVismedarb%>&FM_viskunde=<%=intViskunde%>&FM_visstatus=<%=intVisstatus%>&FM_periodesel=<%=persel%>&wk=<%=sqlweekdato%>" class="vmenu">Tilføj ny <b>Opgave</b> <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
			<%end if%>
			</td></tr>
			</table>
			<br>
		
		
		<!-- filter -->
		<table cellspacing="0" cellpadding="0" border="0" bgcolor="#fffff1" width="600">
		<tr bgcolor="#5582D2">
				<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="22" alt="" border="0"></td>
				<td colspan="2" valign="top"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
				<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="22" alt="" border="0"></td>
			</tr>
			<tr bgcolor="#5582D2">
				<td colspan="2" style="padding-left:3px;" class="alt"><b>Filter:</b></td>
			</tr>
			<form action="todo_add.asp?func=list&id=<%=id%>" method="post" name="viskundesel">
			<tr>
				<td valign="top" rowspan=2><img src="../ill/tabel_top.gif" width="1" height="175" alt="" border="0"></td>
				<td valign=top style="padding-top:2;" width="450"><b>Kunde&nbsp;|&nbsp;Job (aktive) filter:</b><br><select name="FM_viskunde" style="font-family: arial,helvetica,sans-serif; font-size: 10px;">
				<%
				
				strSQL = "SELECT job.id, jobnavn, jobnr, jobknr, kid, kkundenavn, kkundenr FROM job, kunder WHERE jobstatus = 1 AND kid = jobknr ORDER BY kkundenavn, jobnavn"
				
				oRec.open strSQL, oConn, 3
				foundone = "n"
				while not oRec.EOF
					if cint(request("FM_viskunde")) = cint(oRec("id")) then
					ksel = "SELECTED"
					foundone = "j"
					else
					ksel = ""
					foundone = foundone
					end if
				%>
				<option value="<%=oRec("id")%>" <%=ksel%>><%
				if len(oRec("kkundenavn")) > 25 then
				Response.write left(oRec("kkundenavn"), 25)&".."
				else
				Response.write oRec("kkundenavn")
				end if
				
				if len(oRec("jobnavn")) > 35 then
				Response.write "&nbsp;&nbsp;|&nbsp;&nbsp;"& left(oRec("jobnavn"), 35)&".."
				else
				Response.write "&nbsp;&nbsp;|&nbsp;&nbsp;"& oRec("jobnavn") 
				end if
				
				Response.write "&nbsp;&nbsp;&nbsp;(" & oRec("jobnr") & ")"%>
				</option>
				<%oRec.movenext
				wend
				oRec.close
				
				if foundone = "n" then%>
				<option value="0" SELECTED>Alle</option>
				<%else%>
				<option value="0">Alle</option>
				<%end if%>
			</select>
			</td><td valign=top style="padding-top:2;"><b>Medarbejder filter:</b><br>
			<select name="FM_vismedarb" style="font-family: arial,helvetica,sans-serif; font-size: 10px;">
				<%
				
				strSQL = "SELECT mid, mnavn FROM medarbejdere ORDER BY mnavn"
				oRec.open strSQL, oConn, 3
				foundone = "n"
				while not oRec.EOF
					if cint(intVismedarb) = cint(oRec("mid")) then
					msel = "SELECTED"
					foundone = "j"
					else
					msel = ""
					foundone = foundone
					end if
				
					if level <= 3 OR level = 6 then%>
					<option value="<%=oRec("mid")%>" <%=msel%>><%=oRec("mnavn")%></option>
					<%
					else
						if cint(session("mid")) = cint(oRec("mid")) then
						%>
						<option value="<%=oRec("mid")%>" SELECTED><%=oRec("mnavn")%></option>
						<%
						end if
					end if
					
				oRec.movenext
				wend
				oRec.close
				
				if level <= 3 OR level = 6 then
					if foundone = "n" then%>
					<option value="0" SELECTED>Alle</option>
					<%else%>
					<option value="0">Alle</option>
					<%end if
				end if%>
			</select>
			</td>
			<td valign="top" rowspan=2 align=right><img src="../ill/tabel_top.gif" width="1" height="175" alt="" border="0"></td>
		</tr>
		<tr>
			<td valign="top" style="padding-top:2;">
			<b>Periode / dato filter:</b><br>
			<%if cint(persel) = 1 then
			psel1 = "CHECKED"
			end if%> 
			<input type="radio" name="FM_periodesel" value="1" <%=psel1%>>&nbsp;<b>Aktive</b> (startdato før dagsdato).<br>
			<%if cint(persel) = 6 then
			psel6 = "CHECKED"
			end if%> 
			<input type="radio" name="FM_periodesel" value="6" <%=psel6%>>&nbsp;<b>Uge 
			<select name="wk">
			<option value="<%=sqlweekdato%>" SELECTED><%=useWK%></option>
			<%
			for w = 1 to 52
			if w > useWK then
			thisWkLowerKri = (7 * (w-useWK))
			end if
			if w = useWK then
			thisWkLowerKri = (0)
			end if
			if w < useWK then
			thisWkLowerKri = -(7 * (useWK-w))
			end if
			%>
			<option value="<%=dateadd("d", thisWkLowerKri, sqlweekdato)%>"><%=w%></option>
			<%
			next
			%>
			</select>'s opgaver.</b> (Vi er i uge <%=datepart("ww", now, 2 ,3)%> nu.)
			<br>
			<%if cint(persel) = 2 then
			psel2 = "CHECKED"
			end if%>
			<input type="radio" name="FM_periodesel" value="2" <%=psel2%>>&nbsp;<b>Dagens opgaver</b> (slutdato er idag).<br>
			<%if cint(persel) = 3 then
			psel3 = "CHECKED"
			end if%>
			<input type="radio" name="FM_periodesel" value="3" <%=psel3%>>&nbsp;<b>Fremtidige.</b><br>
			<%if cint(persel) = 4 then
			psel4 = "CHECKED"
			end if%>
			<input type="radio" name="FM_periodesel" value="4" <%=psel4%>>&nbsp;<b>Slutdato overskreddet.</b><br>
			<%if cint(persel) = 5 then
			psel5 = "CHECKED"
			end if%>
			<input type="radio" name="FM_periodesel" value="5" <%=psel5%>>&nbsp;<b>Vis alle.</b><br></td>
			<td valign="bottom">
			<b>Status filter:</b><br>
			<%if cint(intVisstatus) = 1 then
			stchk1 = "CHECKED"
			end if%>
			<input type="radio" name="FM_visstatus" value="1" <%=stchk1%>> Vis kun igangværende.<br> 
			<%if cint(intVisstatus) <> 1 then
			stchk2 = "CHECKED"
			end if%>
			<input type="radio" name="FM_visstatus" value="2" <%=stchk2%>> Vis alle. <br>
			<br><input type="image" src="../ill/brug-filter.gif"><br>&nbsp;
			</td>
			</tr>
			<tr>
				<td colspan=4 height="1" bgcolor="#003399" style="border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			</tr>
			<input type="hidden" name="lastvalue" id="lastvalue" value="0">
			</form>
			</table>
			<br>
			
			
			<table cellspacing="0" cellpadding="0" border="0" bgcolor="#d6dff5" width="600">
				<tr bgcolor="#5582D2">
					<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="22" alt="" border="0"></td>
					<td colspan="5" valign="top"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
					<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="22" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#5582D2">
					<td width="130" style="padding-left:3px;" class="alt">Job / Kunde / Opr. af</td>
					<td width="55" style="padding-left:3px;" class="alt">Prioitet</td>
					<td width="335" style="padding-left:8px;" class="alt">Emne / Beskrivelse</td>
					<td width="30" style="padding-left:3px;" class="alt">Status</td>
					<td width="50" height=20 style="padding-left:3px;" class="alt">&nbsp;</td>
				</tr>
				<%
				'**** Kunde kriterie *****
				if len(request("FM_viskunde")) = 0 OR request("FM_viskunde") = 0 then
				kundekri = " kundeid <> -1 "
				else
				kundekri = " kundeid = " & request("FM_viskunde") &""
				end if
				
				'**** Medarbejder kriterie ****************
				'Har medarbejerderen adgang til denne todo?
				'******************************************
				
				if level <= 3 OR level = 6 then 
					if cint(intVismedarb) = 0 then
					medarbKri = "(todoid = todo.id AND medarbid <> 0)"
					else
					medarbKri = "(todoid = todo.id AND medarbid = "& intVismedarb &")"
					end if
				else
					medarbKri = "(todoid = todo.id AND medarbid = "& session("Mid") &")"
				end if
					
				
				'**** Status kriterie *****
				if cint(intVisstatus) = 1 then
				statKri = " udfort = 1"
				else
				statKri = " udfort <> 10" 'alle 
				end if
				
				dagsdato = year(now)&"/"&month(now)&"/"&day(now)
				
				 
				'*** dato kriterie ***
				select case cint(persel)
				case 1 '(startdato før dagsdato, også dem der er over deadline)
				perKri = " startdato <= '"&dagsdato&"'"
				case 2 '(slut dato = dags dato)
				perKri = " slutdato = '"&dagsdato&"'"
				case 3 '(start dato nyere end dags dato).
				perKri = " startdato > '"&dagsdato&"'"
				case 4 '(slut dato er før dags dato)
				perKri = " slutdato < '"&dagsdato&"'"
				case 5 '(vis alle)
				perKri = " startdato > '2001-1-1'"
				case 6 '(denne uges opgaver)
				
				lowerKri = sqlweekdato
				upperKri = dateadd("d", 6, sqlweekdato)
				perKri = " (slutdato >= '"& convertDateYMD(lowerKri) &"' AND slutdato <= '"& convertDateYMD(upperKri) &"') "
				'Response.write perKri
				end select 
				
				
				strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobknr, jobstatus, todo.editor, todo.id, kundeid, emne, todo.beskrivelse, udfort, klassificering, kid, kkundenavn, sortorder, startdato, slutdato, medarbid, todoid FROM todo, todo_rel "_
				&" LEFT JOIN job ON (job.id = kundeid) "_
				&" LEFT JOIN kunder ON (kid = jobknr) "_
				&" WHERE "& kundekri &" AND "& medarbKri &" AND "& statKri &" AND "& perKri &" AND emne <> '' AND jobstatus = 1 "_
				&" GROUP BY todo.id ORDER BY sortorder"
				'Response.write strSQL
				oRec.open strSQL, oConn, 3
				x = 0
				
				'oRec.movefirst
				lastToDoID = 0
				lastsoID = 1
				
				while not oRec.EOF
				
				
				'if oRec("jobstatus") = "1" then 'job = aktiv
					select case oRec("klassificering")
					case 1
					showprio = "<b>Høj</b>"
					usepfont = "#000000"
					case 2
					showprio = "Mid"
					usepfont = "#000000"
					case 3
					showprio = "Lav"
					usepfont = "#999999"
					end select
					
					select case oRec("udfort")
					case 1
					showudf = "Igang"
					useudfont = "LimeGreen"
					case 2
					showudf = "Venter"
					useudfont = "#CCCCCC"
					case 3
					showudf = "Afv.opg."
					useudfont = "#999999"
					case 4
					showudf = "Afsluttet"
					useudfont = "darkred"
					end select
					
					if len(id) <> 0 then
					id = id
					else
					id = cint(0)
					end if
					
					if len(oRec("id")) <> 0 then
						if cint(oRec("id")) = cint(id) then
						bgthis = "#FFFFF1"
						else
						bgthis = "#ffffff"
						end if
					else
						bgthis = "#FFFFF1"
					end if
					%>
					<tr bgcolor="<%=bgthis%>">
						<td valign="top" rowspan=2 style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
						<td style="padding-left:3px;"><b><%
						if len(oRec("jobnavn")) > 18 then
						Response.write left(oRec("jobnavn"), 18) &".."
						else
						Response.write oRec("jobnavn")
						end if%></b></td>
						<td style="padding-left:3px;" align=right><a href="todo_add.asp?func=moveup&id=<%=oRec("id")%>&thisso=<%=oRec("sortorder")%>&lastid=<%=lastToDoID%>&lastso=<%=lastsoID%>&FM_vismedarb=<%=intVismedarb%>&FM_viskunde=<%=intViskunde%>&FM_visstatus=<%=intVisstatus%>&FM_periodesel=<%=persel%>&wk=<%=request("wk")%>"><img src="../ill/pil_moveup.gif" width="10" height="12" alt="Flyt op på listen" border="0"></a>&nbsp;<a href="todo_add.asp?func=movedown&id=<%=oRec("id")%>&thisso=<%=oRec("sortorder")%>&FM_vismedarb=<%=intVismedarb%>&FM_viskunde=<%=intViskunde%>&FM_visstatus=<%=intVisstatus%>&FM_periodesel=<%=persel%>&useSQL=<%=strSQL%>&wk=<%=request("wk")%>"><img src="../ill/pil_movedown.gif" width="10" height="12" alt="" border="0"></a>&nbsp;&nbsp;<font color="<%=usepfont%>"><%=showprio%></font></td>
						<td style="padding-left:8px;"><a href="todo_add.asp?func=red&id=<%=oRec("id")%>&FM_vismedarb=<%=intVismedarb%>&FM_viskunde=<%=intViskunde%>&FM_visstatus=<%=intVisstatus%>&FM_periodesel=<%=persel%>&wk=<%=request("wk")%>" class="vmenu">
						<%
						'Response.write oRec("sortorder") & "#&nbsp;"
						if len(oRec("emne")) > 50 then 
							if oRec("udfort") = 4 then
							Response.write "<s>"& left(oRec("emne"), 50) & "...</s>"
							else
							Response.write left(oRec("emne"), 50) & "..."
							end if
						else
							if oRec("udfort") = 4 then
							Response.write "<s>"&oRec("emne")&"</s>"
							else
							Response.write oRec("emne")
							end if
						end if
						%></a></td>
						<td style="padding-left:3px;"><a href="#" onClick="showudfort('<%=oRec("id")%>');"><font color="<%=useudfont%>" size=1><u><%=showudf%></u></font></a></td>
						<td style="padding-left:3px; padding-top:2px;">&nbsp;<%if level <= 3 OR level = 6 then%>
						<a href="todo_add.asp?func=slet&&FM_vismedarb=<%=intVismedarb%>&FM_viskunde=<%=intViskunde%>&id=<%=oRec("id")%>&FM_visstatus=<%=intVisstatus%>&FM_periodesel=<%=persel%>&wk=<%=request("wk")%>"><img src="../ill/slet_crmkal.gif" width="12" height="12" alt="Slet opgave" border="0"></a>
						<%end if%></td>
						<td valign="top" align=right rowspan=2 style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					</tr>
					<tr bgcolor="<%=bgthis%>">
						<td style="padding-left:3px;" colspan="2" valign="top">
						<%if len(oRec("kkundenavn")) > 18 then 
						Response.write left(oRec("kkundenavn"), 18) & ".."
						else
						Response.write oRec("kkundenavn")
						end if
						%><br><font color=#999999><%=oRec("editor")%></td>
						<td colspan=3 style="padding-left:8px;" valign="top"><a href="#" onClick="showtodotekst('<%=oRec("id")%>')" class="kalblue"><%
						
						if len(oRec("beskrivelse")) > 50 then
							if oRec("udfort") = 4 then
							Response.write "<s>"& left(oRec("beskrivelse"), 50) & "</s>..."
							else
							Response.write left(oRec("beskrivelse"), 50) & "..."
							end if
						else
							if oRec("udfort") = 4 then
							Response.write "<s>"& oRec("beskrivelse") & "</s>"
							else
							Response.write oRec("beskrivelse")
							end if
						end if
						%></a>
						<div id="b_<%=oRec("id")%>" name="b_<%=oRec("id")%>" style="position:absolute; width:400; height:250; overflow:auto; left:300; top:250; visibility:hidden; display:none; background-color:#FFFFFF; filter:alpha(opacity=90); border-left:1 #990200 solid; border-top:1 #990200 solid; border-bottom:1 #990200 solid; border-right:1 #990200 solid; padding-left:5px; padding-top:5px; padding-right:10px; padding-bottom:2px;">
						<a href='#' onClick=closetodotekst(<%=oRec("id")%>)>Luk vindue</a><br><b>Beskrivelse:</b>
						<form><textarea cols="45" rows="8" name="FM_thistxt_<%=oRec("id")%>" id="FM_thistxt_<%=oRec("id")%>" id="FM_thistxt_<%=oRec("id")%>" style="border:0;"><%=oRec("beskrivelse")%></textarea><br>
						<font size=1>Teksten kan ikke opdateres her.</font></form>
						</div>
						<br><font color=#999999 size=1><%=oRec("startDato")%> til <%=oRec("slutDato")%>
						<%
						strDag = day(now)
						strMrd = month(now)
						strAar = year(now)
						
						if cdate(strDag&"/"&strMrd&"/"&strAar) > oRec("slutDato") then
						Response.write "&nbsp;<font color=darkred size=2><b>!</b></font>"
						end if
						%>
						</td>
					</tr>
					<tr>
						<td colspan=7 height="1" style="border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					</tr>
					<%
					x = 1
					
					lastToDoID = oRec("id")
					lastsoID = oRec("sortorder")
				'end if 'jobstatus = 1 (aktiv)
				
				oRec.movenext
				wend
				oRec.close
				%>
				<tr>
					<td colspan=7 height="10" bgcolor="#5582D2" style="border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
				</tr>
			</table>
			
		<%
		end if 'list%>
		<br><br><br>&nbsp;
		</div>
		
		<div id=udfort name=udfort style="position:absolute; width:150; height:150; overflow:none; left:300; top:250; visibility:hidden; display:none; background-color:#FFFFE1; filter:alpha(opacity=90); border-left:1 #990200 solid; border-top:1 #990200 solid; border-bottom:1 #990200 solid; border-right:1 #990200 solid; padding-left:5px; padding-top:5px; padding-right:10px;">
		<a href='#' onClick=closeudfort()>Luk vindue</a><br><b>Vælg status:</b>
		<form name=udf id=udf action="todo_add.asp?func=changeudfort&FM_vismedarb=<%=intVismedarb%>&FM_viskunde=<%=intViskunde%>&FM_visstatus=<%=intVisstatus%>&FM_periodesel=<%=persel%>&wk=<%=request("wk")%>" method="post">
		<input type="radio" name="FM_udfort" value="1" CHECKED>&nbsp;Igang<br>
		<input type="radio" name="FM_udfort" value="2">&nbsp;Venter<br>
		<input type="radio" name="FM_udfort" value="3">&nbsp;Afvent. anden opg.<br>
		<input type="radio" name="FM_udfort" value="4">&nbsp;Afsluttet<br>
		<input type="hidden" id="FM_udfort_id" name="FM_udfort_id" value="0">
		<input type="image" src="../ill/opdaterpil.gif">
		</form>
		</div>
		
		
		<%
		end select
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
