<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>

<%

lto = "cst"
strLicenskey = "2.2009-0105-TO100" 'CST 
strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_cst;"
'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wwf_BAK;"

'**CST sever ***'
'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=93.161.131.214; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cst;"
	    

'strConnThis = "mySQL_timeOut_intranet"
'lto = "intranet - local"



Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oRec3 = Server.CreateObject ("ADODB.Recordset")
Set oRec4 = Server.CreateObject ("ADODB.Recordset")
Set oRec5 = Server.CreateObject ("ADODB.Recordset")
Set oRec6 = Server.CreateObject ("ADODB.Recordset")

Set oCmd = Server.CreateObject ("ADODB.Command")
oConn.open strConnThis


%>


<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<script language="javascript">

//function clwin(){
//window.location.reload()
//}

//function setTimer() {
//       setTimeout ("clwin()", 120000); //30000 millisekunder = 30 sek. 
//}
</script>
	<title>TimeOut Stempelur Service</title>
</head>

<body> <!-- onload="setTimer();" -->
<h4>TimeOut Stempelur Service</h4>
Denne side må ikke lukkes ned. <br><br>

<br>
Overfører terminal loginds...

<%

thisfile = "stempelur_terminal_ns"


'Finder dags dato
tidspunkt = (DatePart("h", Now) + 1) & ":00:00"
tidspunkt30 = (DatePart("h", Now) + 1) & ":59:59"
datenow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
usedate = Year(Now) & "/" & Month(Now) & "/" & Day(Now)

'session("afsendelser") = session("afsendelser") & "<br>Dato:" & datenow & " kl:" & tidspunkt &" - "& tidspunkt30 

Response.write "<br>Dato: " & now &"<hr>"


'oRec.close
'oConn.close
'Response.Write strConnect_aktiveDB & "<br>"
'Finder Terminal loginds.
strSQLtm = "SELECT lht_id, lht_mid, lht_logind, lht_type FROM login_historik_terminal "_
&" WHERE lht_status <> 1 AND lht_mid <> 0 ORDER BY lht_logind"

'Response.Write strSQLtm
'Response.flush

c = 0
oRec6.open strSQLtm, oConn, 3
while not oRec6.EOF 

Response.Write "<hr>"& oRec6("lht_id") & " Dato & tid: "& oRec6("lht_logind") & "<br>"
Response.flush

call logindStatus(oRec6("lht_mid"), 1, 2, oRec6("lht_logind"))

strSQLup = "UPDATE login_historik_terminal SET lht_status = 1, fo_logud_kode = "& fo_logud &" WHERE lht_id = "& oRec6("lht_id")
oConn.execute(strSQLup)


'strSQLup = "UPDATE login_historik_terminal SET lht_status = 1, fo_logud_kode = "& fo_logud &" WHERE lht_id = "& oRec6("lht_id")
'oConn.execute(strSQLup)

c = c + 1
oRec6.movenext
wend
oRec6.close
oConn.close


		
%>
<br><br>
Der er overført: <b><%=c%></b><br><br>


<!--#include file="../inc/regular/footer_inc.asp"-->

<!--
<script language="javascript">
    setTimeout('close_terminal()', 10000);
    function close_terminal() {
        window.close()
    }
</script>
    -->