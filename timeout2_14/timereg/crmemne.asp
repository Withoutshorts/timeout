<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->


<!--#include file="../inc/regular/global_func.asp"-->

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%


 




if len(session("user")) = 0 then
	

	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***

    call menu_2014()

	%>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	<h4>Aktions - Emner</h4>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> et emne. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="crmemne.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM crmemne WHERE id = "& id &"")
	Response.redirect "crmemne.asp?menu=tok&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		
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
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO crmemne (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
		else
		oConn.execute("UPDATE crmemne SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "crmemne.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato FROM crmemne WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if

    call menu_2014()
	%>
	
   

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	<h4>CRM - Emner</h4>
        <form action="crmemne.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<table cellspacing="0" cellpadding="4" border="0" width="400">
	
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
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:350px;"></td>
	</tr>
	<tr>
		<td colspan="2" align="right"><br><br><input type="submit" value="<%=varSubVal%> >>"></td>
	</tr>
	
	</table>
    </form>

	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="16" height="17" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case else
        
        
    call menu_2014()    
    %>
	
	


   

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; width:600px; height:600; visibility:visible;">
	<h4>CRM - Emner</h4>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top">
	Sortér efter <a href="crmemne.asp?menu=tok&sort=navn">navn</a> eller <a href="crmemne.asp?menu=tok&sort=nr">nr</a>
	<img src="../ill/blank.gif" width="100" height="1" alt="" border="0">
	<a href="crmemne.asp?menu=tok&func=opret">Opret nyt emne <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
	<td height="30"><b>Id</b></td>
	<td width="60"><b>Emne</b></td>
	<td width="130">&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	strSQL = "SELECT id, navn FROM crmemne ORDER BY navn"
	else
	strSQL = "SELECT id, navn FROM crmemne ORDER BY id"
	end if
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
		<td><%=oRec("id")%></td>
		<td height="20"><a href="crmemne.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		<td><a href="crmemne.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>	
	</table>
	</TD>
	</TR></TABLE>
	
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
