<%

function todolist(klikket)

if cint(lastId) = oRec("id") then
bgthis = "#FFFF99"
else
bgthis = "#ffffff"
end if
%>
<tr>

		<%if antalsubs < 1 then
		strImgLI = "<img src='../ill/webblik_li_8.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<%if antalsubs >= 1 AND antalsubs <= 5 then
		strImgLI = "<img src='../ill/webblik_li_10.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<%if antalsubs > 5 AND antalsubs <= 15 then
		strImgLI = "<img src='../ill/webblik_li_12.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<%if antalsubs > 15 AND antalsubs <= 30 then
		strImgLI = "<img src='../ill/webblik_li_14.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<%if antalsubs > 30 then
		strImgLI = "<img src='../ill/webblik_li_16.gif' width='21' height='14' alt='' border='0'>" 
		end if%>
		
		<td valign=top bgcolor="<%=bgthis%>" class=lille style="padding:5px 5px 5px 5px; border-left:1px #8cAAe6 solid; width:80px; border-bottom:1px #cccccc dashed;">
        <input type="hidden" name="SortOrder" value="<%=oRec("sortorder")%>" />
	    <input type="hidden" name="rowId" value="<%=oRec("id")%>" />
        <%=strImgLI%> <br />(id: <%=oRec("id") %>)</td>
		<td bgcolor="<%=bgthis%>" valign=top style="padding:5px; width:480px; border-bottom:1px #cccccc dashed;">
		
		<% if oRec("afsluttet") = 1 then%>
		<s>
		<%end if%>
		
        <% if print <> "j" then %>
		<a href="#" onClick="edittodo('<%=oRec("id")%>','<%=replace(oRec("navn"), "'","")%>','<%=medarbemails%>','<%=formatdatetime(oRec("dato"), 0)%>', '<%=oRec("afsluttet")%>', '<%=parent%>', '<%=oRec("forvafsl") %>', '<%=day(oRec("tododato")) %>', '<%=month(oRec("tododato")) %>', '<%=year(oRec("tododato")) %>', '<%=oRec("public")%>')" class="<%=acls%>">
        <%=oRec("navn")%></a>
        &nbsp;<a href="webblik_todo.asp?FM_parent=<%=oRec("id")%>&todolevel=<%=oRec("level")%>&nomenu=<%=nomenu %>" class="<%=acls%>"><span style="color:#999999;">(<%=antalsubs%>)</span></a>
        <%else %>
        <%=oRec("navn")%> (<%=antalsubs%>)
        <%end if %>

        <br /><span style="font-size:9px; color:#8caae6;"><%=oRec("dato") %></span>
		
		
		
		<%'*** Andre medarbejdere med adgang ***
			medarbemails = " "
			medarbNavn = ""
			strSQL2 = "SELECT medarbid, email, mid, mnavn FROM todo_rel_new "_
			&" LEFT JOIN medarbejdere ON (mid = medarbid) WHERE todoid = " & oRec("id") &" GROUP BY mid ORDER BY mid" ' & " AND medarbid <> " & session("mid")
			'Response.write strSQL2
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF  
				medarbemails = medarbemails & oRec2("mid") & ","
				medarbNavn = medarbNavn & oRec2("mnavn") & ", "
			oRec2.movenext
			wend
			oRec2.close 
		
		len_medarbemails = len(medarbemails)
		left_medarbemails = left(medarbemails, len_medarbemails - 1)
		medarbemails = trim(left_medarbemails)
		
		len_medarbNavn = len(medarbNavn)
		left_medarbNavn = left(medarbNavn, len_medarbNavn - 2)
		medarbNavn = trim(left_medarbNavn)
		
		if IsDate(oRec("tododato")) then
		dato = formatdatetime(oRec("tododato"),2)
		else
		dato = oRec("tododato")
		end if 
		ekspTxt = ekspTxt &oRec("id")&";"&oRec("dato")&";"&dato&";"&Replace(Replace(oRec("navn"),"'","''"),";",":")&";"&Replace(medarbNavn,";",":")&";"&Replace(medarbemails,";",":")&";"&oRec("afsluttet")&";"&oRec("forvafsl")&";"&Replace(oRec("sortorder"),";",":")&";"
        ekspTxt = ekspTxt &"xx99123sy#z"
		
		%>
		
		<%if oRec("delt") <> 0 then%>
		<font class=megetlillesort>(delt)</font>&nbsp;
		<%end if%>

        <%if oRec("public") = 1 then%>
		<font class=megetlillesort>(offentlig)</font>&nbsp; 
		<%end if%>
		
		<% if oRec("afsluttet") = 1 then%>
		</s>
		<%end if%>
		
		<br /><font class=megetlillesilver>
        <%if len(medarbNavn) > 50 then %>
        <%=left(medarbNavn, 50) %>..
        <%else %>
        <%=medarbNavn %>
        <%end if %></font>
		
		</td>
        <td bgcolor="<%=bgthis%>" class=lille valign=top style="padding:5px; border-bottom:1px #cccccc dashed; white-space:nowrap;"><%if oRec("forvafsl") <> 0 then %>
		Forv. udført: <b><%=formatdatetime(oRec("tododato"), 2) %></b>
		<%end if%></td>
		
		<!--
		<form action="webblik_oprtodo.asp?func=op&id=<%=oRec("id")%>&FM_parent=<%=parent%>&oldone=<%=oldones%>" method="post" id="<%=oRec("id")%>">
        -->
        <td bgcolor="<%=bgthis%>" align=right valign=top style="padding:5px; border-bottom:1px #cccccc dashed;">&nbsp;
        <% if print <> "j" then %>
        <a href="#" onClick="edittodo('<%=oRec("id")%>','<%=replace(oRec("navn"), "'","")%>','<%=medarbemails%>','<%=formatdatetime(oRec("dato"), 0)%>', '<%=oRec("afsluttet")%>', '<%=parent%>', '<%=oRec("forvafsl") %>', '<%=day(oRec("tododato")) %>', '<%=month(oRec("tododato")) %>', '<%=year(oRec("tododato")) %>', '<%=oRec("public")%>')">
		<img src="../ill/edit.gif" alt="Rediger ToDo" border="0" /></a><% end if %></td>
		<td bgcolor="<%=bgthis%>" valign=top style="padding:5px; border-bottom:1px #cccccc dashed; border-right:1px #8cAAe6 solid;">&nbsp;
		<% if print <> "j" then %>
		<a href="webblik_oprtodo.asp?func=slet&id=<%=oRec("id")%>&FM_parent=<%=parent%>&oldone=<%=oldones%>"><img src="../ill/slet_16.gif" alt="Slet" border="0" /></a>&nbsp;&nbsp;
		<%if klikket <> 0 AND oRec("afsluttet") <> 1 then%>

		 <!--   <input id="FM_sortorder" name="FM_sortorder" value="<%=oRec("sortorder")%>" style="width:40px; font-family:verdana; font-size:8px;" type="text" />
            <input id="Submit_<%=oRec("id")%>" type="submit" value="<|>" style="font-family:verdana; font-size:8px;" />
            -->
		
		
		<%end if%>
		<% end if %>
		</td>
		
		
		
	</tr>
	<%
	
	
end function
%>