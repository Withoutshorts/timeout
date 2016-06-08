<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Migration</title>
</head>

<body>
<%
'strConnect_old = "DBQ=e:\Selskab\regnskab\timesag\timesag.mdb;DRIVER={Microsoft Access Driver (*.mdb)};"

'strConnect = "timeout21_outzource"

'strConnect = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=d:\sites\outzource.dk\timeout\db\timeOut.mdb;"
Set oConn_old = Server.CreateObject("ADODB.Connection")
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oRec3 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn_old.open strConnect_old
oConn.open strConnect

strSQL = "SELECT * FROM timer"
oRec.open strSQL, oConn_old, 3

while not oRec.EOF

if oRec("Tfaktim") = true then
varFaktim = 1
else
varFaktim = 0
end if

Select case oRec("Tjobnr")
case "100"
varKnr = 1
varJobnr = 1
varAid = 20
varAnavn = "Redesign"
tfaktim = 1
case "106"
varKnr = 2
varJobnr = 3
varAid = 22
varAnavn = "Estimat af timeforbrug"
tfaktim = 1
case "11"
varKnr = 5
varJobnr = 5
varAid = 31
varAnavn = "Udvikling"
tfaktim = 0
case "104"
varKnr = 1
varJobnr = 2
varAid = 21
varAnavn = "Redesign"
tfaktim = 1
case "5"
varKnr = 5
varJobnr = 7
varAid = 0
varAnavn = "Intern"
tfaktim = 0
case else
varKnr = 5
varJobnr = 6
varAid = 0
varAnavn = "Intern"
tfaktim = 0
end select

 oConn.Execute ("INSERT INTO timer (Tdato, Tmnavn, Tmnr, Tjobnavn, Tjobnr, Tknavn, Tknr, Timer, Tfaktim, Timerkom, TAktivitetId, TAktivitetNavn, Taar, TimePris, Tastedato)"_
 &" VALUES ("_
 &"'"& oRec("Tdato") &"', "_ 
 &"'Søren Karlsen', "_
 &"1, "_ 
 &"'"& oRec("Tjobnavn") &"', "_ 
 &""& varJobnr &", "_
 &"'"& oRec("Tknavn")&"', "_ 
 &""& varKnr &", "_ 
 &""& oRec("Timer")&", "_
 &""& tfaktim &", "_ 
 &"'"& oRec("Timerkom")&"', "_ 
 &""&varAid&", "_
 &"'"& varAnavn &"', "_
 &"2002, "_
 &"550, "_
 &"'3/27/2002')")

Response.write oRec("Tdato") &": " & oRec("Tid") & "ok!" & "<br>" 
 
oRec.movenext
wend
%><br>
<br>

done!
</body>
</html>
