<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
%>
				<tr bgcolor="silver">
					<td valign="top" bgcolor="#ffffff"><img src="ill/sort_pil.gif" width="12" height="9" border="0" alt=""></td>
					<td bgcolor="#ffffff"><img src="ill/blank.gif" width="5" height="1" border="0" alt=""><a href="javascript:NewWinView('view.asp?id=<%=oRec("Aid")%>&type=Acc2')" target="_self"><img src="ill/upload/<=oRec("Billedenavn")%>" alt="" border="0"></a></td>
					<td valign="top" bgcolor="#ffffff" width="200"><font class="netscape"><b><%=oRec("navn_dk")%></b><br>
					Item no. <%=oRec("Varenr") & "<br>"%></td>
					<td align="right" valign="top" bgcolor="#ffffff"><b><%=oRec("Pris_dk") &",-</b></td><td align='right' valign='top' bgcolor='#ffffff'><b>" & oRec("Pris_eng")%>,-</b></td>
					<td align="right" valign="top" bgcolor="#ffffff"><img src="ill/blank.gif" width="10" height="1" border="0" alt=""><a href="inc/insert.asp?id=<%=oRec("Aid")%>&insptob=open&protype=acc"><font class="basket"><img src="ill/sort_pil.gif" width="12" height="9" alt="" border="0">Add to  <img src="ill/vogn_lille.gif" width="24" height="14" alt="" border="0"></a></td>
				</tr>
				<tr>
					<td colspan="6" bgcolor="silver"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<%


</body>
</html>
