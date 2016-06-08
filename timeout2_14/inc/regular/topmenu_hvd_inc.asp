<script>
function showabout(){
document.all["about"].style.visibility = "visible";
}
function hideabout(){
document.all["about"].style.visibility = "hidden";
}
function NewWin(url)    {
window.open(url, 'Help', 'width=700,height=580,scrollbars=yes,toolbar=no');    
}
function NewWin_cal(url)    {
window.open(url, 'Calender', 'width=300,height=380,scrollbars=yes,toolbar=no');    
}
function NewWin_popupaktion(url)    {
window.open(url, 'Calender', 'width=410, height=650,scrollbars=yes,toolbar=no');    
}
function NewWin_help(url)    {
window.open(url, 'Help', 'width=650,height=580,scrollbars=yes,toolbar=no');    
}
function NewWin_large(url)    {
window.open(url, 'Help', 'width=980,height=600,scrollbars=yes,toolbar=no');    
}
function NewWin_todo(url)    {
window.open(url, 'Opgavelisten', 'width=650,height=650,scrollbars=yes,toolbar=no');    
}
</script>

<div name="about" id="about" style="position:absolute; top:64; left:100; visibility:hidden; background-color:#FFFFE1; filter:alpha(opacity=90); border-left:1 #990200 solid; border-top:1 #990200 solid; border-bottom:1 #990200 solid; border-right:1 #990200 solid; padding-left : 2px;padding-right : 2px; z-index:200;">
<table height="115" width="275" cellpadding=0 cellspacing=0 border=0>
<%
strSQL = "SELECT * FROM dbversion, licens"
oRec.open strSQL, oConn, 3
if not oRec.EOF then
licenshaver = oRec("licens")
%>
<tr><td></td><td align="right"><a href="#" onclick="hideabout()";><font size="1">Luk dette vindue</font></a></td></tr>
<tr><td width="75">Firma:</td><td> <%=oRec("licens")%></td></tr>
<tr><td>Serienr:</td><td> <%=oRec("key")%></td></tr>
<tr><td>DB version:</td><td> <%=oRec("dbversion")%></td></tr>
<tr><td>Licenstype: </td><td><%=oRec("licenstype")%></td></tr>
<tr><td>Brugere:</td><td> <%=oRec("klienter")%></td></tr>
<tr><td>Support: </td><td><a href="mailto:support@outzource.dk">support@outzource.dk</a></td></tr>
<%
'** tjek om der er tilvalgt CRM ***
len_licensType = instr(oRec("key"), "B")
	if len_licensType > 0 then
	licensType = "CRM"
	end if
end if
oRec.close
%>
</table>
<br>
</div>


<table cellspacing="0" cellpadding="0" border="0" bgcolor="#003399" width=90%>
<tr>
<td><a href="../login.asp"><img src="../ill/logo_topbar.gif" alt="" border="0"></a></td>
<td align="right">
<table cellpadding="2" cellspacing="0">
<tr>
<td valign="top"><img src="../ill/keys.gif" alt="" border="0"></td>
<td>
<font style="color:#ffffff;font-family:arial,sans-serif;font-size:11px;">Du er <strong>logget ind</strong> som:
	&nbsp;<%=session("user")%>&nbsp;&nbsp;<font color="#FFCC00">(<%=licenshaver%>)</font> 
	<%
	strSQL = "SELECT lastlogin FROM medarbejdere WHERE Mnavn ='" & session("user") &"'"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		strDato = oRec("lastlogin")
		end if
		oRec.close
	%>
	&nbsp;&nbsp;<br>Dato:&nbsp;<%=strDato%> - Sidste login:&nbsp;<%=session("strLastlogin")%>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="../sesaba.asp" STYLE="font-family : arial, sans-serif; font-size : 11px; color : ffffff;"><strong>Log ud</strong></a> &nbsp;&nbsp;</font>
</td>
</tr>
</table>


</td>
</tr>
</table>
<br>

