<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
	'** Faste filter kri ***'
	thisfile = "medarbtyper_grp"
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
	
	
	
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<h4>Medarbejdertype - grupper</h4>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> et medarbejdertype gruppe. Er dette korrekt?<br />
        Du vil samtidig slette alle relationer til denne gruppe</td>
	</tr>
	<tr>
	   <td><a href="medarbtyper_grp.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM medarbtyper_grp WHERE id = "& id &"")
    

	Response.redirect "medarbtyper_grp.asp?menu=tok&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
        useleftdiv = "c"
		errortype = 8
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
        if len(trim(request("FM_opencalc"))) <> 0 then
        opencalc = 0 '** Lukket MAKS 1
                
                   '*** Nulstiller
                   oConn.execute("UPDATE medarbtyper_grp SET opencalc = 1 WHERE id <> 0")
        else
        opencalc = 1 '** Åben (default)
        end if

     

		if func = "dbopr" then
		oConn.execute("INSERT INTO medarbtyper_grp (navn, editor, dato, opencalc) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& opencalc &")")
		else
		oConn.execute("UPDATE medarbtyper_grp SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', opencalc = "& opencalc &" WHERE id = "&id&"")
		end if

        
		
		Response.redirect "medarbtyper_grp.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	opencalc = 1

	else
	strSQL = "SELECT navn, editor, dato, opencalc FROM medarbtyper_grp WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	opencalc = oRec("opencalc")
    end if
	oRec.close
	
    if opencalc = 0 then
    opencalcCHK = "CHECKED"
    else
    opencalcCHK = ""
    end if

	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:82px; visibility:visible;">
	<h4>Medarbejdertype - grupper</h4>

        <form action="medarbtyper_grp.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
            
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	
    	
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td>Navn:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30"></td>
	</tr>
    <tr>
		<td>Lukket for kalkulation:<br />
        (under jobbudget, maks. 1 grp.)</td>
		<td><input type="CHECKBOX" name="FM_opencalc" value="1" <%=opencalcCHK %>></td>
	</tr>
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="100" height="1" alt="" border="0"><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	
	</table>
            </form>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>

	
	
	

	<%case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">

        <%	tTop = 0
	tLeft = 0
	tWdth = 800
	
	
	call tableDiv(tTop,tLeft,tWdth) %>

	<h4>Medarbejdertype - grupper</h4>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top">
	
	Sortér efter <a href="medarbtyper_grp.asp?menu=tok&sort=navn">Navn</a> eller <a href="medarbtyper_grp.asp?menu=tok&sort=nr">Id nr.</a>
	

          <span style="position:relative; left:600px; top:-20px;"><% 
                nWdt = 180
                nTxt = "Opret nyt gruppe"
                nLnk = "medarbtyper_grp.asp?menu=tok&func=opret"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %> 
	        </span>


	
	</td>
	</tr>
	</table>
	
	<table cellspacing="0" cellpadding="4" border="0" width="100%">

	<tr bgcolor="#5582D2">
		<td width="50" height=20 class=alt><b>Id</b></td>
		<td class=alt><b>Gruppe</b></td>
        <td class=alt>Bruges af antal medarbejdertyper</td>
		<td class=alt>Åben for kalkulation</td>
        <td>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")

    strSQL = "SELECT mt.id AS id, navn, COUNT(m.id) AS antal, opencalc FROM medarbtyper_grp AS mt "

	strSQL = strSQL & " LEFT JOIN medarbejdertyper AS m ON (m.mgruppe = mt.id) GROUP BY mt.id "

    if sort = "navn" then
	strSQL = strSQL & "ORDER BY navn"
	else
	strSQL = strSQL & "ORDER BY m.id"
	end if

    'Response.Write strSQL
    'Response.flush

	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	
	<tr bgcolor="#ffffff" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'#ffffff');">
		
		<td height=20 style="border-bottom:1px #CCCCCC Solid;"><%=oRec("id")%></td>
		<td style="border-bottom:1px #CCCCCC Solid;"><a href="medarbtyper_grp.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
        <td style="border-bottom:1px #CCCCCC Solid;"><%=oRec("antal")%></td>
       
        <td style="border-bottom:1px #CCCCCC Solid;"><%=oRec("opencalc") %></td>
		<td style="border-bottom:1px #CCCCCC Solid;">
        <%if cint(oRec("antal")) = 0 AND oRec("id") <> 1 then %>
        <a href="medarbtyper_grp.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a>
        <%else %>
        &nbsp;
        <%end if %></td>
		
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>	
	<tr bgcolor="#5582D2">
		<td colspan=5 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			</tr>
	</table>
	
</div>
</div> <!-- tablediv -->
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
