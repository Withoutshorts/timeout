



<%
	Set oRec = server.createobject("adodb.recordset")
	if cbool(oRec.state) = true then
	oRec.close
	end if
	set oRec = nothing


'* afslutter load tid ***
timeB = now
loadtime = datediff("s",timeA, timeB)
Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write "<br><br><br><br><br><br><br><br><br>&nbsp;&nbsp;&nbsp;<font class=megetlillesilver>Sidens loadtid: "& loadtime &" sekund(er)." &"</font><br>"
Response.flush

'** Stat ***

statTime = time
statDato = year(now)&"/"&month(now)&"/"&day(now)
stTempLen = len(cstr(request.servervariables("PATH_TRANSLATED")))
stTempInstr = instr(cstr(request.servervariables("PATH_TRANSLATED")), "imereg")
statSide = right(cstr(request.servervariables("PATH_TRANSLATED")), (stTempLen - (stTempInstr + 6)))

'strConnStat = "driver={MySQL};server=81.19.249.41; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_admin;"
'Set oConnStat = Server.CreateObject("ADODB.Connection")
'oConnStat.open strConnStat
'statSQL = "INSERT INTO stat (dato, tpunkt, bruger, lto, side) VALUES ('"&statDato&"', '"&statTime&"', '"& session("user") &"', '"&lto&"', '"& statSide &"')"
'oConnStat.execute(statSQL)
'oConnStat.close
'Set oConnStat = nothing	
%>


<div style="position:absolute; z-index:20000; background-color:pink; left:20px; top:2000; width:100px; display:; visibility:hidden;">
<%
for each x in Request.Cookies
  
  if Request.Cookies(x).HasKeys then
    for each y in Request.Cookies(x)
      response.write(x & ":" & y & "=" & Request.Cookies(x)(y))
      response.write("<br />")
    next
  else
    Response.Write(x & "=" & Request.Cookies(x) & "<br />")
  end if
  'response.write "<br>"
next
%>
<hr>
<%
'dim i
'dim j
j=Session.Contents.Count
Response.Write("Session variables: " & j)
For i=1 to j
  Response.Write(Session.Contents(i) & "<br />")
Next
%>
<hr>
<%
For Each i in Session.StaticObjects
  Response.Write(i & "<br />")
Next
%>
</div>
<p>&nbsp;</p>
</body>
</html>
