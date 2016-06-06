<table border=1 cellspacing=0 cellpadding=0 bgcolor="#ffffff" width="600">
		<form id=filter name=filter method=post action="#">
		
		<tr><td style="border-left:1px #8caae6 solid; border-top:1px #8caae6 solid; border-right:1px #8caae6 solid; padding-top:8px;">
		
		&nbsp;&nbsp;<b>Vælg Job:</b>&nbsp;<select name="id" size="1" style="font-size : 9px; width:250px;" onChange="visjob()">
		<option value="0"> Alle - eller vælg det ønskede job..</option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid, jobnavn, jobnr, id, "_
				&" jobans1, jobans2 FROM job LEFT JOIN kunder ON (kid = jobknr) WHERE jobstatus = 1 ORDER BY jobnavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(id) = cint(oRec("id")) then
				
				'*** Jobansvarlige ***
				jobans1 = oRec("jobans1")
				jobans2 = oRec("jobans2")
				thisJobnr = oRec("jobnr")
				
				
				isSelected = "SELECTED"
					
					'*** rediger rettigheder ***
					if level = 1 then
					editok = 1
					else
							if cint(session("mid")) = jobans1 OR cint(session("mid")) = jobans2 OR (cint(jobans1) = 0 AND cint(jobans2) = 0) then
							editok = 1
							end if
					end if
				
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) --  <%=oRec("Kkundenavn")%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select>
		</td></tr>
		</table>