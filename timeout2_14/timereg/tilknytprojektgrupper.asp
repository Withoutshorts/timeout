<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->	

<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/convertDate.asp"--> 

	
	<%
	
	id = request("id")
	medarbid = request("medid")
	func = request("func")

   
	
	if func = "db" then
	
	'*** Henter værdier ***
	strProjektgr1 = request("FM_projektgruppe_1")
	strProjektgr2 = request("FM_projektgruppe_2")
	strProjektgr3 = request("FM_projektgruppe_3")
	strProjektgr4 = request("FM_projektgruppe_4")
	strProjektgr5 = request("FM_projektgruppe_5")
	strProjektgr6 = request("FM_projektgruppe_6")
	strProjektgr7 = request("FM_projektgruppe_7")
	strProjektgr8 = request("FM_projektgruppe_8")
	strProjektgr9 = request("FM_projektgruppe_9")
	strProjektgr10 = request("FM_projektgruppe_10")
	
	
	
	'*** Opdaterer job ***
	strSQLupd = "UPDATE job SET "_
	&" projektgruppe1 = "& strProjektgr1 &", "_ 
	&" projektgruppe2 = "& strProjektgr2 &", "_ 
	&" projektgruppe3 = "& strProjektgr3 &", "_ 
	&" projektgruppe4 = "& strProjektgr4 &", "_
	&" projektgruppe5 = "& strProjektgr5 &", "_
	&" projektgruppe6 = "& strProjektgr6 &", "_ 
	&" projektgruppe7 = "& strProjektgr7 &", "_ 
	&" projektgruppe8 = "& strProjektgr8 &", "_ 
	&" projektgruppe9 = "& strProjektgr9 &", "_
	&" projektgruppe10 = "& strProjektgr10 &" "_
	&" WHERE id = "& id &""
	
	'Response.Write strSQLupd
	'Response.flush
	
	oConn.execute(strSQLupd)
	
	
	'**** Opdaterer aktiviteter *** 
	strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& id 
	
	oRec2.open strSQL, oConn, 3
		if not oRec2.EOF then
		
		oConn.execute("UPDATE aktiviteter SET"_
		&" projektgruppe1 = "& oRec2("projektgruppe1") &" , projektgruppe2 = "& oRec2("projektgruppe2") &", "_
		&" projektgruppe3 = "& oRec2("projektgruppe3") &", "_
		&" projektgruppe4 = "& oRec2("projektgruppe4") &", "_
		&" projektgruppe5 = "& oRec2("projektgruppe5") &", "_
		&" projektgruppe6 = "& oRec2("projektgruppe6") &", "_
		&" projektgruppe7 = "& oRec2("projektgruppe7") &", "_
		&" projektgruppe8 = "& oRec2("projektgruppe8") &", "_
		&" projektgruppe9 = "& oRec2("projektgruppe9") &", "_
		&" projektgruppe10 = "& oRec2("projektgruppe10") &" WHERE job = "& id &"")
		
		end if
	oRec2.close
	
	
	
	'*** tilføjer job i timereg_usejob (Vis guide) ****
	strSQL = "SELECT DISTINCT(MedarbejderId) FROM progrupperelationer WHERE ("_
	&" ProjektgruppeId = "& strProjektgr1 &""_
	&" OR ProjektgruppeId =" & strProjektgr2 &""_
	&" OR ProjektgruppeId =" & strProjektgr3 &""_
	&" OR ProjektgruppeId =" & strProjektgr4 &""_
	&" OR ProjektgruppeId =" & strProjektgr5 &""_
	&" OR ProjektgruppeId =" & strProjektgr6 &""_
	&" OR ProjektgruppeId =" & strProjektgr7 &""_
	&" OR ProjektgruppeId =" & strProjektgr8 &""_
	&" OR ProjektgruppeId =" & strProjektgr9 &""_
	&" OR ProjektgruppeId =" & strProjektgr10 &""_
	&") GROUP BY MedarbejderId"
	oRec3.open strSQL, oConn, 3
	while not oRec3.EOF

        
       
        
     
        
		strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& oRec3("MedarbejderId") &", "& id &")"
        
		oConn.execute(strSQL3)
		
    oRec3.movenext
	wend
	oRec3.close
	

	
	'window.opener.top.frames["a"].location.reload();

	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	//Response.Write("<script language=""JavaScript"">window.opener.top.frames['a'].location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
	
	
	else
	
	%>
	<div id="sindhold" style="position:absolute; left:40; top:20; width:600; height:600; visibility:visible;">
	<%
	
	strSQL = "SELECT id, projektgruppe1, projektgruppe2, "_
	&" projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, "_
	&" projektgruppe8, projektgruppe9, projektgruppe10 "_
	&" FROM job WHERE id=" & id 
	
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strProj_1 = oRec("projektgruppe1")
	strProj_2 = oRec("projektgruppe2")
	strProj_3 = oRec("projektgruppe3")
	strProj_4 = oRec("projektgruppe4")
	strProj_5 = oRec("projektgruppe5")
	strProj_6 = oRec("projektgruppe6")
	strProj_7 = oRec("projektgruppe7")
	strProj_8 = oRec("projektgruppe8")
	strProj_9 = oRec("projektgruppe9")
	strProj_10 = oRec("projektgruppe10")
	
	end if
	oRec.close
	
	%>
	
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<form action="tilknytprojektgrupper.asp?func=db&id=<%=id%>&medid=<%=medarbid%>" method="POST">
	<tr>
    <td valign="top">
	<h4>Tilknyt ny Projektgruppe til job: <span style="font-size:10px;">(og underliggende aktiviteter)</span></h4>
	<br>

    <div style="background-color:#ffffff; width:400px; border:1px #999999 solid; padding:10px; color:#000000;">
    <b>Den valgte medarbejder er med i følgende projektgrupper:</b><br />
	<%
		
		
		strSQL = "SELECT pg.navn AS pgnavn "_
		&" FROM medarbejdere, progrupperelationer "_
		&" LEFT JOIN projektgrupper pg ON (pg.id = ProjektgruppeId) WHERE medarbejdere.Mid="& medarbid &"  AND MedarbejderId = Mid GROUP BY ProjektgruppeId" 
		oRec.Open strSQL, oConn, 3
		
			While Not oRec.EOF
			
			Response.write oRec("pgnavn") &"<br>"
			
			oRec.MoveNext
			Wend
		
		oRec.Close
		
		
	%>
	<br>
    </div>

    <br />
	
    <span style="color:red; border:1px #999999 solid; padding:4px; background-color:#FFFFFF;">Vær opmærksom på at du overskriver alle projektgruppe indstillinger på både job og aktiviteter!</span>
	<br /><br /><b>Tilknyt projektgrupper til jobbet:</b>
	<br>

				
							<%
							p = 0
							for p = 1 to 10
							varSelected = ""
							select case p
							case 1
							strProj = strProj_1
							case 2
							strProj = strProj_2
							case 3
							strProj = strProj_3
							case 4
							strProj = strProj_4
							case 5
							strProj = strProj_5
							case 6
							strProj = strProj_6
							case 7
							strProj = strProj_7
							case 8
							strProj = strProj_8
							case 9
							strProj = strProj_9
							case 10
							strProj = strProj_10
							end select
							
							if strProj <> 1 then
							varDisab = "DISABLED"
							else
							varDisab = ""
							end if
							
							       
							%>
							Projektgruppe <%=p%>:&nbsp;
								<select name="FM_projektgruppe_<%=p%>" style="width:300px;" <%=varDisab%>>
								<%
									strSQL = "SELECT id, navn FROM projektgrupper ORDER BY navn"
									oRec.open strSQL, oConn, 3
									gp = 1
									While not oRec.EOF 
									projgId = oRec("id")
									projgNavn = oRec("navn")
									
									if cint(strProj) = cint(projgId) then
									varSelected = "SELECTED"
									gp = strProj
									else
									varSelected = ""
									'gp = gp
									end if
									
									
									
									
									%>
									<option value="<%=projgId%>" <%=varSelected%>><%=projgNavn%> </option>
									<%
									oRec.movenext
									wend
									oRec.close%>
						</select><br>
						
						<%
							
							 if varDisab = "DISABLED" then
									%>
                                    <input name="FM_projektgruppe_<%=p%>" value="<%=gp%>" type="hidden" />
									<%
									end if
							
							
							select case p
							case 1
							gp1 = gp
							case 2
							gp2 = gp
							case 3
							gp3 = gp
							case 4
							gp4 = gp
							case 5
							gp5 = gp
							case 6
							gp6 = gp
							case 7
							gp7 = gp
							case 8
							gp8 = gp
							case 9
							gp9 = gp
							case 10
							gp10 = gp
							end select
							
							next
							
							%>
							<br><br>
							<!--<input type="checkbox" name="FM_sendmail" value="1" checked>Send mail til jobansvarlige om at du har tilknyttet en ny projektgruppe.
							--><input type="submit" value="Opdater">
							</td>
						</tr>
						</form>
				</table>
		</div>
		<%end if%>
		
		
		<!--#include file="../inc/regular/footer_inc.asp"-->