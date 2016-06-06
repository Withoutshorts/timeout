

<!--#include file="inc/convertDate.asp"--> 
<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
</script>

	<table cellspacing="0" cellpadding="0" border="0" width="200>
	<tr>
    <td valign="top"><br>
	&nbsp;Seneste 5 indtastninger
	 <!-------------------------------Sideindhold------------------------------------->
<br>
<table cellspacing="1" cellpadding="2" border="1" bgcolor="#ffffff" bordercolor="#cd853f">
<tr>
	<td><font class='tdblack11'><b>Dato</b></td>
	<td><font class='tdblack11'><b>Jobnavn</b></td>
	<td><font class='tdblack11'><b>Aktivitet</b></td>
	<td><font class='tdblack11'><b>Timer</b></td>
</tr>
<%
	
	oRec.Open "SELECT Tdato, Tjobnr, Tjobnavn, aktiviteter.navn AS Anavn, Tknavn, Timer, Tid, Tfaktim FROM timer, aktiviteter WHERE Tmnavn = '"&Session("user")&"' AND Tid > 0 AND aktiviteter.id = TAktivitetId ORDER BY Tdato DESC", oConn, 3
	'objRec.Open "SELECT Tdato, Tjobnr, Tjobnavn, Tknavn, Timer, Tid, Tfaktim FROM timer WHERE Tmnavn = 'Skarlsen' AND Tdato BETWEEN #"&strDatoKri&"# AND #"&strDatoDag&"# ORDER BY Tdato DESC", oConn, 3
	
	x = 0
	While not oRec.EOF And x < 5
	
	%>
	<tr onmouseover="mOvr('gift',this,'lightBlue');" onmouseout="mOut(this,'#ffffff');">
	<%
	Response.write "<td align=center><font class='tdblack11'>"& oRec("Tdato") &"</td>"
	Response.write "<td align=center><font class='tdblack11'>"& left(oRec("Tjobnavn"), 12) &"</td>"
	Response.write "<td align=center><font class='tdblack11'>"& left(oRec("Anavn"), 12) &"</td>"
	Response.write "<td align=center><font class='tdblack11'>"& oRec("Timer") &"</td></tr>"
	
	x = x + 1	
	oRec.movenext
	wend	  
	oRec.Close
	
%>
</table>

</TD>
</TR>
</TABLE>

