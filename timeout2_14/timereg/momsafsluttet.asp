<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	
	select case func 
	case "opd"
	
	
	editdato = year(now) &"/"& month(now) &"/"& day(now)
	
	 thisID = -1
	 thisKomm = ""
	 thisafslutdato = ""
	 thisafslutdatoSQL = ""
	 
	
	
	
	antalids = split(request("FM_id"),",")
	
	for x = 0 to UBOUND(antalids)
	
	
	errType = 0
	    
	    thisID = trim(antalids(x))
	    thisKomm = replace(request("FM_kom_"&trim(thisID)&""),"'","''")
	    
	    if len(trim(thisKomm)) > 255 then
	    errType = 1
	            
	        %>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
			<%
			errortype = 97
			call showError(errortype)
			
			Response.end
	            
	    end if
	    
	    thisafslutdato = trim(request("FM_afslutdato_dag_"&thisID&"")) &"-"& trim(request("FM_afslutdato_md_"&thisID&"")) &"-"& trim(request("FM_afslutdato_aar_"&thisID&""))
	    thisafslutdatoSQL = trim(request("FM_afslutdato_aar_"& thisID &"")) &"/"& trim(request("FM_afslutdato_md_"& thisID &"")) &"/"& trim(request("FM_afslutdato_dag_"& thisID &""))
	    
	    if isDate(thisafslutdato) = false then
	    errType = 2
	    
	        %>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
			<%
			errortype = 98
			call showError(errortype)
			
			Response.end
	    
	    end if
	    
	    
	    
	if thisID = 0 then
	strSQL = "INSERT INTO momsafsluttet (editor, dato, kommentar, afslutdato) VALUES "_
	&"('"& session("user") &"', '"& editdato &"', '"& thisKomm &"','"& thisafslutdatoSQL &"' )"
	else
	strSQL = "UPDATE momsafsluttet SET editor = '"& session("user") &"', "_
	&" dato = '"& editdato &"', kommentar = '"& thisKomm &"', afslutdato = '"& thisafslutdatoSQL &"' WHERE id = "& thisID
	end if
	
	if errType = 0 AND trim(len(thiskomm)) <> 0 then
    'Response.Write strSQL &"<br>"
	oConn.execute(strSQL)
	end if
	
	if trim(len(thiskomm)) = 0 then
	oConn.execute("DELETE FROM momsafsluttet WHERE id = "& thisID)
	end if
	
	next
	
	       
			
    Response.Redirect "momsafsluttet.asp"
	
	case else 
	
	
	
	%>
	
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
	<h4>Momsperiode(r) afsluttet</h4>
	Der kan ikke foretages posteringer på konti i en afsluttet periode.<br />
	<br />
	<table cellspacing=0 cellpadding=2 border=0 width=500>
	<form action="momsafsluttet.asp?func=opd" method="post">
	<tr bgcolor="#5582d2">
	    <td class=alt colspan=2 style="border-left:1px #003399 solid; border-top:1px #003399 solid; border-right:1px #003399 solid;"><b>Momsperioder afsluttet:</b></td>
	</tr>
	
	<tr bgcolor="#5582d2">
	<td class=alt style="border-left:1px #003399 solid; border-bottom:1px #003399 solid;"><b>Kommentar til periode</b></td>
	<td class=alt style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><b>Momsperiode afsluttet dato</b></td>
	</tr>
	
	
	
	<%strSQL = "SELECT id, dato, afslutdato, kommentar FROM momsafsluttet WHERE id <> 0 ORDER BY afslutdato DESC" 
	oRec.open strSQL, oConn, 3
	
	
	while not oRec.EOF 
	
	
	%>
	
	<input id="FM_id" name="FM_id" type="hidden" value="<%=trim(oRec("id"))%>" />
	
	<tr bgcolor="#ffffff">
	
	<td style="border-left:1px #5582d2 solid; border-bottom:1px #5582d2 solid;">
        <input id="FM_kom_<%=oRec("id")%>" name="FM_kom_<%=oRec("id")%>" type="text" value="<%=oRec("kommentar") %>" /></td>
	<td style="border-right:1px #5582d2 solid; border-bottom:1px #5582d2 solid;">
        <input id="FM_afslutdato_dag_<%=oRec("id")%>" name="FM_afslutdato_dag_<%=oRec("id")%>" type="text" size=2 value="<%=datepart("d",oRec("afslutdato"))%>" /> -
	    <input id="FM_afslutdato_md_<%=oRec("id")%>" name="FM_afslutdato_md_<%=oRec("id")%>" type="text" size=2 value="<%=datepart("m", oRec("afslutdato"))%>" /> -
	    <input id="FM_afslutdato_aar_<%=oRec("id")%>" name="FM_afslutdato_aar_<%=oRec("id")%>" type="text" size=4 value="<%=datepart("yyyy", oRec("afslutdato"))%>" /> dd-mm-åååå</td>
	</tr>
	<%
	
	oRec.movenext
	wend
	oRec.close
	%>
	
	
	</table>
	<br />
	<br />
	<br />
	<table cellspacing=0 cellpadding=4 border=0 width=500>
	<input id="FM_id" name="FM_id" type="hidden" value="0" />
	
	<tr bgcolor="#cccccc">
	    <td style="border-left:1px #5582d2 solid; border-top:1px #5582d2 solid; border-right:1px #5582d2 solid;" colspan=2><br /><b>Tilføj afslutning af momsperiode:</b> (for at tilføje, skal kommentar feltet være udfyldt)</td>
	</tr>
	
	<tr bgcolor="#cccccc">
	
	<td style="border-left:1px #5582d2 solid; border-bottom:1px #5582d2 solid;">
        <input id="FM_kom_0" name="FM_kom_0" type="text" /></td>
	<td style="border-bottom:1px #5582d2 solid; border-right:1px #5582d2 solid;">
        <input id="FM_afslutdato_dag_0" name="FM_afslutdato_dag_0" type="text" size=2 value="<%=day(now)%>" /> -
	    <input id="FM_afslutdato_md_0" name="FM_afslutdato_md_0" type="text" size=2 value="<%=month(now)%>" /> -
	    <input id="FM_afslutdato_aar_0" name="FM_afslutdato_aar_0" type="text" size=4 value="<%=year(now)%>" /> dd-mm-åååå
	    </td>
	
	</tr>
	<tr>
	    <td colspan=2 align=right style="padding-right:20px;"><input id="Submit1" type="submit" value="Tilføj / opdater" /></td>
	</tr>
	
	</table>
	<br />
        For at slette en afslutning, skal kommentar feltet ryddes før submit.
	
	</form>
<%

    end select
    
    
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
