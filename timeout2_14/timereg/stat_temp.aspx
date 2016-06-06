<%@ Page Language="VB" %>

<!--
	<form action="joblog_timetotaler.asp?FM_jobsog=<%=jobSogValPrint%>&FM_job=<%=jobid%>&FM_aftaler=<%=aftaleid%>" method="post">
	
	<input type="hidden" name="FM_kunde" id="FM_kunde" value="<%=kundeid%>">
	<input type="hidden" name="FM_kundeans" id="FM_kundeans" value="<%=kundeans%>">
	<input type="hidden" name="FM_jobans" id="FM_jobans" value="<%=jobans%>">
	<input type="hidden" name="FM_jobans2" id="FM_jobans2" value="<%=jobans%>">
	<input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
	-->
	<tr>
	<td><input type="radio" name="FM_radio_projgrp_medarb" value="1" <%=projgrp_medarb1%>>&nbsp;<b>Medarbejder(e):</b>&nbsp;<br>
	<select name="FM_medarb" style="font-size : 11px; width:205px;">
	
	    <%
	    select case level
	    case 1,2,6
	    %>
	    <option value="0">Alle</option>
	    <%
	    medarbSelKri = " mansat <> 2 "
	    case else
	    medarbSelKri = " mid = " & selmedarb
	    end select
	
	
	strSQL = "SELECT Mid, Mnavn, Mnr, mansat, init FROM medarbejdere WHERE "& medarbSelKri &" ORDER BY Mnavn"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
		
		if cint(selmedarb) = oRec("mid") then
		thisChecked = "SELECTED"
		else
		thisChecked = ""
		end if
		%>
		<option value="<%=oRec("mid")%>" <%=thisChecked%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>) - <%=oRec("init")%></option>
	<%
	oRec.movenext
	wend
	oRec.close
	%></select>
	</td>
	
	<td>
	
	&nbsp;
	    
	    <% 
	    select case level
	    case 1,2,6
	    %>
	    
	
	<input type="radio" name="FM_radio_projgrp_medarb" value="2" <%=projgrp_medarb2%> onclick="setshowprogrpval();">&nbsp;<b>Projektgruppe(r):</b><br>
	<%
	strSQL = "SELECT p.id AS id, navn, count(medarbejderid) AS antal FROM projektgrupper p "_
	&" LEFT JOIN progrupperelationer ON (ProjektgruppeId = p.id) "_
	&" GROUP BY p.id ORDER BY navn"
	oRec.open strSQL, oConn, 3
	
	'Response.write strSQL
	'Response.flush
	%>
	
	<select name="FM_progrupper" style="font-size : 11px; width:205px;" onchange="setshowprogrpval();">
	<%
			
			
			while not oRec.EOF
			
			if cint(progrp) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			
			if cint(oRec("antal")) > 0 then%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
			<%end if
			
			oRec.movenext
			wend
			oRec.close
			%>
	</select>
	
	<%
	    end select
	    %>
	
	</td>
	<td valign=middle align=right>&nbsp;<input type="submit" value=" Vælg medarb. >> "></td>
	</tr>
	</table>
	
	
	
	<%if progrp_medarb = 2 AND request("showprogrpdiv") = "1" then %>
	<div id="progrpdiv" style="position:absolute; left:810px; top:120px; width:200px; height:250px; overflow:auto; background-color:#ffffff; border:1px #c4c4c4 solid; padding:10px; font-size:9px; z-index:20000;" onclick="closeprogrpdiv();" onmouseover="curserType('progrpdiv');">
	
	    <b>Medarbejdere i valgt projektgruppe:</b><br />
	    <%=medarbigrp %><br />
	    <a href="#" onclick="closeprogrpdiv();" class=red>[Klik her for at lukke]</a>
	    </div>
	<%end if %>
	
	<input id="showprogrpdiv" name="showprogrpdiv" value="0" type="hidden" />
	
	<table cellspacing=0 cellpadding=5 border=0 width=100% bgcolor="#FFFFFF">