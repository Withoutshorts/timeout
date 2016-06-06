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
	
	func = request("func")
	id = request("id") 
	
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:170; top:80; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <%
	select case func
	case "med"
	strSQL = "SELECT id, navn FROM brugergrupper WHERE id=" & id
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	gruppeNavn = oRec("Navn")
	end if
	oRec.close
	%>
	<td valign="top"><font class="pageheader">Brugergrupper <img src="../ill/pil.gif" width="11" height="11" border="0" alt="">  Medlemmer</font>
	<br>
	<br><br><br>
	Medlemmer i <b><%=gruppeNavn%></b> gruppen:<br>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="80%">
	<tr>
	<td width="40"><b>Navn</b></td>
	<td>&nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT Mnavn, Mid FROM medarbejdere WHERE Brugergruppe = "&id&" ORDER BY Mnavn"
	oRec.open strSQL, oConn, 3
	
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'lightBlue');" onmouseout="mOut(this,'#fff8dc');">
	<td height="20"><a href="medarb_red.asp?menu=medarb&func=red&id=<%=oRec("Mid")%>"><%=oRec("Mnavn")%> </a></td>
	<td>&nbsp;</td>
	</tr>
	<%
	oRec.movenext
	wend
	%>	
	</table>
	<%
	case else
	%>
	<td valign="top"><font class="pageheader">Søgeresultater:</font>
	<br>
	<br>
	Sortér efter <a href="sogeresultater.asp?menu=xx&sort=navn">navn</a> eller <a href="bgrupper.asp?menu=medarb&sort=nr">rettigheds niveau</a>
	<br><br>
	Søgefunktionen er ikke tilsluttet systemet. Kontakt outZource.
	<%end select%>

<br><br>
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="16" height="17" alt="" border="0"></a>
<br>
<br>
</TD>
</TR>
</TABLE>
</div>

<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
