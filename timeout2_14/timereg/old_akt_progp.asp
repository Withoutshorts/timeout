<tr>
						<td valign="top" style="padding-left:40;">Projektgruppe 1:</td>
						<td><select name="FM_projektgruppe_1">
						<option value="1">Ingen</option>
							<%
							strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
							oRec.open strSQL, oConn, 3
							gp1 = 1
							While not oRec.EOF 
							projgId_1 = oRec("id")
							projgNavn_1 = oRec("navn")
							
							if trim(strProj_1) = trim(projgId_1) then
							varSelected_1 = "SELECTED"
							gp1 = strProj_1
							else
							varSelected_1 = ""
							gp1 = gp1
							end if
							
							if varLastName <> projgId_1 then
							%>
							<option value="<%=projgId_1%>" <%=varSelected_1%>><%=projgNavn_1%></option>
							<%
							end if
							varLastName = projgId_1
							oRec.movenext
							wend
							varLastName = ""
							oRec.close%>
				</select></td>
					</tr>
					
					
					<tr>
						<td valign="top" style="padding-left:40;">Projektgruppe 2:</td>
						<td><select name="FM_projektgruppe_2">
						<option value="1">Ingen</option>
						<%
							strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
							oRec.open strSQL, oConn, 3
							gp2 = 1
							While not oRec.EOF 
							projgId_2 = oRec("id")
							projgNavn_2 = oRec("navn")
							
							if trim(strProj_2) = trim(projgId_2) then
							varSelected_2 = "SELECTED"
							gp2 = strProj_2
							else
							varSelected_2 = ""
							gp2 = gp2
							end if
							
							if varLastName <> projgId_2 then
							%>
							<option value="<%=projgId_2%>" <%=varSelected_2%>><%=projgNavn_2%></option>
							<%
							end if
							
							varLastName = projgId_2
							oRec.movenext
							wend
							varLastName = ""
							oRec.close%>
				</select></td>
					</tr>
					<tr>
						<td valign="top" style="padding-left:40;">Projektgruppe 3:</td>
						<td><select name="FM_projektgruppe_3">
						<option value="1">Ingen</option>
						<%
							strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
							oRec.open strSQL, oConn, 3
							gp3 = 1
							While not oRec.EOF 
							projgId_3 = oRec("id")
							projgNavn_3 = oRec("navn")
							
							if trim(strProj_3) = trim(projgId_3) then
							varSelected_3 = "SELECTED"
							gp3 = strProj_3
							else
							varSelected_3 = ""
							gp3 = gp3
							end if
							
							if varLastName <> projgId_3 then
							%>
							<option value="<%=projgId_3%>" <%=varSelected_3%>><%=projgNavn_3%></option>
							<%
							end if
							
							varLastName = projgId_3
							oRec.movenext
							wend
							varLastName = ""
							oRec.close%>
				</select></td>
					</tr>
						<tr>
						<td valign="top" style="padding-left:40;">Projektgruppe 4:</td>
						<td><select name="FM_projektgruppe_4">
						<option value="1">Ingen</option>
						<%
							strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
							oRec.open strSQL, oConn, 3
							gp4 = 1
							While not oRec.EOF 
							projgId_4 = oRec("id")
							projgNavn_4 = oRec("navn")
							
							if trim(strProj_4) = trim(projgId_4) then
							varSelected_4 = "SELECTED"
							gp4 = strProj_4
							else
							varSelected_4 = ""
							gp4 = gp4
							end if
							
							if varLastName <> projgId_4 then
							%>
							<option value="<%=projgId_4%>" <%=varSelected_4%>><%=projgNavn_4%></option>
							<%
							end if
							
							varLastName = projgId_4
							oRec.movenext
							wend
							varLastName = ""
							oRec.close%>
				</select></td>
					</tr>
						<tr>
						<td valign="top" style="padding-left:40;">Projektgruppe 5:</td>
						<td><select name="FM_projektgruppe_5">
						<option value="1">Ingen</option>
						<%
							strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
							oRec.open strSQL, oConn, 3
							gp5 = 1
							While not oRec.EOF 
							projgId_5 = oRec("id")
							projgNavn_5 = oRec("navn")
							
							if trim(strProj_5) = trim(projgId_5) then
							varSelected_5 = "SELECTED"
							gp5 = strProj_5
							else
							varSelected_5 = ""
							gp5 = gp5
							end if
							
							if varLastName <> projgId_5 then
							%>
							<option value="<%=projgId_5%>" <%=varSelected_5%>><%=projgNavn_5%></option>
							<%
							end if
							
							varLastName = projgId_5
							oRec.movenext
							wend
							varLastName = ""
							oRec.close%>
					</select></td>
					</tr>
