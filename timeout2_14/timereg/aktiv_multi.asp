<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
	
	%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call jobtopmenu()
	%>
	</div>
	
	<div id="sindhold" style="position:absolute; left:20px; top:132px; visibility:visible; z-index:100;">
   
	
	<!-------------------------------Sideindhold------------------------------------->
	<h4>Multi opdater aktiviteter</h4>
	
	
	
	<%
	if len(trim(request("FM_aktnavn"))) <> 0 then
	aktnavn = request("FM_aktnavn")
	else
	aktnavn = ""
	end if
	 %>
	
	<form action=aktiv_multi.asp method=post>
        Vælg aktivitet:&nbsp;<input id="FM_aktnavn" name="FM_aktnavn" type="text" value="<%=aktnavn %>" style="width:200px;" />
        <input id="Submit1" type="submit" value="submit" />
	</form>
	
	
	
	
	<%
	if len(trim(aktnavn)) <> 0 then
	%>
	<table>
	<tr>
	    <td>Id</td>
	<td><b>Aktiviteter</b></td><td>
        Kontakt</td>
        <td>Jobnavn (nr)</td>
        </tr>
	
    <%
    strSQL = "SELECT a.id AS aid, a.navn AS anavn, a.job, jobnavn, jobnr, k.kkundenavn, k.kkundenr FROM aktiviteter a "_
    &" LEFT JOIN job j ON (j.id = a.job) "_
    &" LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.navn LIKE '"& aktnavn &"%'"
    
    Response.Write strSQL
    Response.flush
    
    oRec.open strSQL, oConn, 3
    while not oRec.EOF 
    %>
    <tr><td><%=oRec("aid") %></td><td><%=oRec("anavn") %></td><td><%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>)</td><td><%=oRec("jobnavn") %> (<%=oRec("jobnr") %>)</td></tr>
    <%
    oRec.movenext
    wend
    
    
    %>
    </table>
    
    
    <%
    
    end if
    %>
     </div>

<%
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
