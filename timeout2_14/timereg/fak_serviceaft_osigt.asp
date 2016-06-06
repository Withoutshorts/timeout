<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->


<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
%>

<SCRIPT language=javascript src="inc/serviceaft_osigt_func.js"></script>

<!--#include file="../inc/regular/header_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->



<%

if len(request("id")) <> 0 then
id = request("id")
else
id = 0
end if

thisfile = "fak_serviceaft_osigt.asp"
func = request("func")
FM_soeg = request("FM_soeg")



if len(request("FM_usedatokri")) <> 0 then
fmudato = request("FM_usedatokri")
response.cookies("fso_usefmudato") = fmudato
response.cookies("fso_usefmudato").Expires = date + 20
else
	if len(request.cookies("fso_usefmudato")) <> 0 then
	fmudato = request.cookies("fso_usefmudato")
	else
	fmudato = 0
	response.cookies("fso_usefmudato") = fmudato
	response.cookies("fso_usefmudato").Expires = date + 20
	end if
end if


select case func 

case "sletsaft"
	saftid = request("saftid")
	'*** Her spørges om det er ok at der slettes en saftale ***
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:100; visibility:visible;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en aftale. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="serviceaft_osigt.asp?menu=kun&func=sletsaftok&saftid=<%=saftid%>&id=<%=id%>&FM_soeg=<%=FM_soeg%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
'*** Her slettes en serviceaftale ***
oConn.execute("DELETE FROM serviceaft WHERE id = "& saftid &"")

'*** Opdaterer job ****
oConn.execute("UPDATE job SET serviceaft = 0 WHERE serviceaft = "& saftid &"")

Response.redirect "serviceaft_osigt.asp?menu=kund&id="&id&"&FM_soeg="&FM_soeg

	
case else	
'****************************** Aftaler. ***********************************'
%>	
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:58; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="550">
	<tr>
    <td valign="top"><br>
	<img src="../ill/header_fak_aft.gif" alt="" border="0" width="800" height="33"></td>
	</tr></table>
	</div>
	
	<div id=seraft name=seraft Style="position:absolute; left:190; top:122; display:; visibility:visible;">
	<!--#include file="inc/fak_serviceaft_osigt_inc.asp"-->
	<br><br>
	<%if s > 0 then%>
	<h3>Funktioner:</h3>
	
	<form action="fak_serviceaft_eksport.asp" method="post">
	<input type="hidden" name="FM_aftaler" id="FM_aftaler" value="<%=aftalIds%>">
	<b>Eksportér ovenstående liste af aftaler som txt. fil i XAL format.</b><br>
	<textarea cols="60" rows="15" name="FM_eksportdata"><%=ekspTxt%></textarea><br>
	<br><input type="submit" Value="Eksporter data"><br>
	
	</form>
	<%end if%>
	</div>
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
