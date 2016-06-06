<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/kontakter_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

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
<SCRIPT language=javascript src="inc/serviceaft_osigt_func.js"></script>





<%

if len(request("id")) <> 0 then
id = request("id")
else
id = 0
end if

intKid = id

thisfile = "serviceaft_osigt.asp"
func = request("func")
FM_soeg = request("FM_soeg")

visalle = request("visalle")

if len(request("FM_usedatokri")) <> 0 then
fmudato = request("FM_usedatokri")
response.cookies("so_usefmudato") = fmudato
else
	if len(request.cookies("so_usefmudato")) <> 0 then
	fmudato = request.cookies("so_usefmudato")
	else
	fmudato = 0
	response.cookies("so_usefmudato") = fmudato
	end if
end if


filterKri = request("filter_per")
filterStatus = request("status")

select case func 

case "sletsaft"
	saftid = request("saftid")
	'*** Her spørges om det er ok at der slettes en saftale ***
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:100; visibility:visible;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en aftale. Er dette korrekt?<br />
	    Du vil samtidig slette alle faktuaer oprettet på denne aftale.</td>
	</tr>
	<tr>
	   <td><a href="serviceaft_osigt.asp?menu=kun&func=sletsaftok&saftid=<%=saftid%>&id=<%=id%>&FM_soeg=<%=FM_soeg%>&visalle=<%=visalle%>&FM_usedatokri=<%=fmudato%>&filter_per=<%=filterKri%>&status=<%=filterStatus%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
case "sletsaftok"
saftid = request("saftid")
'*** Her slettes en aftale ***
oConn.execute("DELETE FROM serviceaft WHERE id = "& saftid &"")

    '*** Sletter fakturaer på aftale ****
	strSQLfak = "SELECT fid FROM fakturaer WHERE jobid = 0 AND aftaleid = "& id &""
	oRec.open strSQLfak, oConn, 3
	
	while not oRec.EOF 
	        
	        oConn.execute("DELETE FROM faktura_det WHERE fakid = "& oRec("fid") &"")
	        oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& oRec("fid") &"")
	        oConn.execute("DELETE FROM fak_mat_spec WHERE matfakid = "& oRec("fid") &"")
	        
	oRec.movenext
	wend
	oRec.close
	
	
	oConn.execute("DELETE FROM fakturaer WHERE jobid = "& id &"")

'*** Opdaterer job ****
oConn.execute("UPDATE job SET serviceaft = 0 WHERE serviceaft = "& saftid &"")

Response.redirect "serviceaft_osigt.asp?menu=kund&id="&id&"&FM_soeg="&FM_soeg&"&func="&visalle&"&brugdatokri="&fmudato&"&filter_per="&filterKri&"&status="&filterStatus

	
case else	
'****************************** Aftaler. ***********************************'

        
        call menu_2014()
        
        if func <> "osigtall" then
		call kundemenu("seraft")

        topPx = 92
        opTopPX = 0

        else
        opTopPX = -10
        topPx = 102
	end if%>
	
	
	
	<div id="seraft" name="seraft" Style="position:relative; left:20px; top:<%=topPx%>px; display:; visibility:visible;">
	
	<!--#include file="inc/serviceaft_osigt_inc.asp"-->
	<br><br>
	</div>
	
	<br><br><br><br><br><br><br><br>&nbsp;
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
