<%
	'****************************************************************************
	'Tildel timer på medarbejdere på hver enkelt aktivitet.
	'****************************************************************************
	
	function dagepaaakt()
		for d = 0 to mthDaysUse - 1
		getnyDato = dateadd("d", d, jobStartKri)
		showday = d + 1
		
		select case weekday(getnyDato)
		case 1, 7
		dagebgcolor = "#cccccc"
		case else
		dagebgcolor = "#ffffff"
		end select 
		
		%>	
		<td valign="top" bgcolor="<%=dagebgcolor%>" width=18 align=center><font class=megetlillesort><%=showday%></td>
		<%
		next
	end function
	
	
	
	dim resdato
	dim restimer
	
	t = 0
	redim resdato(t)
	redim restimer(t)
	
	for x = 0 to x - 1
	%>
	<div style="position:absolute; left:10; top:150; visibility:hidden; display:none; width:650; height:350; border:1px #003399 solid; overflow:auto; padding:5px; background-color:#d6dff5;" name="d_<%=arrjobid(x)%>" id="d_<%=arrjobid(x)%>">
	<a href="#" onClick="hide(<%=arrjobid(x)%>)">Luk vindue</a><br>
		<b><%=arrjobnavn(x)%> (<%=arrjobnr(x)%>)</b><br>
		<%=arrjstdato(x)%> til <%=arrslutdato(x)%><br>
		
		<%
		jobdatediff = datediff("d", arrjstdato(x), arrslutdato(x))
		%>
		<table border="0" cellpadding="0" celspacing="0">
			
		<%
			strSQL2 = "SELECT medarbejdere.mid, medarbejdere.mnavn, medarbejdere.mnr, aktiviteter.id, navn, aktstartdato, "_
			&" aktslutdato, projektgruppe1, projektgruppe2, "_
			&" projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, "_
			&" projektgruppe8, projektgruppe9, projektgruppe10, projektgruppeid, medarbejderid "_
			&" FROM aktiviteter LEFT JOIN "_
			&" progrupperelationer ON (projektgruppeid = projektgruppe1 "_
			&" OR projektgruppeid = projektgruppe2 "_
			&" OR projektgruppeid = projektgruppe3 OR projektgruppeid = projektgruppe4 "_
			&" OR projektgruppeid = projektgruppe5 OR projektgruppeid = projektgruppe6 "_
			&" OR projektgruppeid = projektgruppe7 OR projektgruppeid = projektgruppe8 "_
			&" OR projektgruppeid = projektgruppe9 OR projektgruppeid = projektgruppe10) "_
			&" LEFT JOIN medarbejdere ON (medarbejdere.mid = medarbejderid) "_
			&" WHERE job = " & arrjobid(x) &" AND aktstatus = 1 AND fakturerbar <> 2 ORDER BY navn, mnavn"
			
			
			'Response.write strSQL2
			'Response.flush
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF 
					
				
					
				
				if oRec2("id") <> lastaktid then%>
				<tr><td valign=right colspan=31><br><b><%=oRec2("navn")%></b></td></tr>
				<tr>
						<%
						call dagepaaakt()
						%>
				</tr>
				<%end if%>
				
			<tr><td valign=right colspan=31><%=oRec2("mnavn")%> (<%=oRec2("mnr")%>) <br>
			<%
			if len(oRec2("mid")) <> 0 then
			usemid = oRec2("mid")
			else
			usemid = 0
			end if
			
			if len(oRec2("id")) <> 0 then
			useaktid = oRec2("id")
			else
			useaktid = 0
			end if
			
			'*** res timer på datoer ***
			strSQL3 = "SELECT ressourcer.timer AS restimer, ressourcer.dato AS resdato FROM "_
			&" ressourcer WHERE ressourcer.mid = "&usemid&" AND aktid = "&useaktid
			oRec3.open strSQL3, oConn, 3 
			t = 1
			while not oRec3.EOF 
			Redim preserve resdato(t)
			Redim preserve restimer(t)
			
			resdato(t) = oRec3("resdato")
			restimer(t) = oRec3("restimer")
			
			t = t + 1
			oRec3.movenext
			wend
			oRec3.close 
			
			
			%>
			</td></tr>
			 
			<tr>
							
						<%for d = 0 to mthDaysUse - 1
						sql2day = d + 1
						akttimer_medarb = 0
						
						'*** Sammenligner timer på diverse datoer ***
						for t = 1 to t - 1
						sql2resdato = sql2day&"/"&monththis&"/"&yearthis
							'Response.write cdate(resdato(t)) &"="& cdate(sql2resdato) & "<br>"
							if cdate(resdato(t)) = cdate(sql2resdato) then
							akttimer_medarb = restimer(t)
							end if
						next
						
						if akttimer_medarb <> 0 then
						akttimer_medarb = akttimer_medarb
						else
						akttimer_medarb = ""
						end if
						%>	
						<td valign="top" width=18><input type="text" name="F_<%=sql2day %>" id="F_<%=sql2day %>" value="<%=akttimer_medarb%>" style="width:18px; font-size:8;"></td>
						<%
						next%>
			</tr>
			<%
			lastaktid = oRec2("id") 
			oRec2.movenext
			wend
			oRec2.close 
		%></table>
		
		</div>
	<%	
	next
	%>
	
	<!--- hidden 0-div -->
	<input type="hidden" name="FM_lastdiv" id="FM_lastdiv" value="0">
	<div style="position:absolute; left:0; top:0; visibility:hidden; display:none;" name="d_0" id="d_0">
	</div>