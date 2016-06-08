



<%
	Set oRec = server.createobject("adodb.recordset")
	if cbool(oRec.state) = true then
	oRec.close
	end if
	set oRec = nothing

    Set oRec2 = server.createobject("adodb.recordset")
	if cbool(oRec2.state) = true then
	oRec2.close
	end if
	set oRec2 = nothing

    Set oRec3 = server.createobject("adodb.recordset")
	if cbool(oRec3.state) = true then
	oRec3.close
	end if
	set oRec3 = nothing

    Set oRec4 = server.createobject("adodb.recordset")
	if cbool(oRec4.state) = true then
	oRec4.close
	end if
	set oRec4 = nothing

    Set oRec5 = server.createobject("adodb.recordset")
	if cbool(oRec5.state) = true then
	oRec5.close
	end if
	set oRec5 = nothing

    Set oRec6 = server.createobject("adodb.recordset")
	if cbool(oRec6.state) = true then
	oRec6.close
	end if
	set oRec6 = nothing




'* afslutter load tid ***
timeB = now
loadtime = datediff("s",timeA, timeB)

'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
'Response.write "<br><br><br><br><br><br><br><br><br>&nbsp;&nbsp;&nbsp;<font class=megetlillesilver>Sidens loadtid: "& loadtime &" sekund(er)." &"<br>"
'Response.flush

'if media <> "print" AND media <> "export" AND thisfile = "xtimereg_akt_2006" then%>
<!--
<div id="loadtid" style="position:absolute; z-index:1000000000; top:100px; left:1250px; width:200px;"><font class=megetlillesilver>Sidens loadtid: <%=loadtime%> sekund(er)</div>
    -->
    
<%'end if


if cdbl(loadtime) > 30 AND thisfile = "timereg_akt_2006" AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\timereg_akt_2006.asp" then
					  	            
                                    'Sender notifikations mail
		                            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            Mailer.CharSet = 2
		                            Mailer.FromName = "TimeOut performance service "
		                            Mailer.FromAddress = "timeout@outzource.dk"
		                            Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            Mailer.AddRecipient "TimeOut Support", "support@outzource.dk"
		                           
            						
                             
            						    
						            Mailer.Subject = "Loadtid alert "& lto &", "& thisfile
		                            strBody = "Hej TimeOut Support" & vbCrLf & vbCrLf
                                    strBody = strBody & session("user") & " fra "& lto &" har haft en loadtid på "& thisfile &" på: "& loadtime & " sek."& vbCrLf & vbCrLf
                                    strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            strBody = strBody & session("user") & vbCrLf & vbCrLf
            		                
            		
            		
		                            Mailer.BodyText = strBody
            		
		                            Mailer.sendmail()
		                            Set Mailer = Nothing

                                    'Response.Write "mail afsendt - " & request.form("jq_lukjob")
				            
end if
            	


'** Stat ***'

statTime = time
statDato = year(now)&"/"&month(now)&"/"&day(now)
stTempLen = len(cstr(request.servervariables("PATH_TRANSLATED")))
stTempInstr = instr(cstr(request.servervariables("PATH_TRANSLATED")), "imereg")
statSide = right(cstr(request.servervariables("PATH_TRANSLATED")), (stTempLen - (stTempInstr + 6)))

yearnow = year(now)
monthnow = month(now)

'** IKKE localhost
'strConnStat = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" '62.182.173.226
'Set oConnStat = Server.CreateObject("ADODB.Connection")
'oConnStat.open strConnStat
'statSQL = "INSERT INTO admin_stat_"& yearnow &"_"& monthnow &" (dato, tpunkt, bruger, lto, side) VALUES ('"&statDato&"', '"&statTime&"', '"& session("user") &"', '"&lto&"', '"& statSide &"')"
'oConnStat.execute(statSQL)
'oConnStat.close
'Set oConnStat = nothing	
%>


</body>
</html>
