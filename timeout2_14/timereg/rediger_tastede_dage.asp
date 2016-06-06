<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_inc.asp"-->	
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/vmenu.asp"-->
<!--#include file="../inc/regular/rmenu.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<div id="sindhold" style="position:absolute; left:190; top:80; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top">
	<!-------------------------------Sideindhold------------------------------------->
	<img src="../ill/header_red_reg.gif" alt="" border="0"><hr align="left" width="450" size="1" color="#000000" noshade>
	

Dags dato:&nbsp;<%=convertDate(date)%><br>

<%
intJoblog = Request("joblog")
StrTid = Request.queryString("id")

id = StrTid
'strReqMrd = Request("mrd")
medarb = request("medarb")
intjobnr = request("jobnr")
eks = request("eks")
'strReqAar = Request("year")
lastFakdag = request("lastFakdag")
selmedarb = request("selmedarb")
selaktid = request("selaktid")
strJob = request("FM_job")
strMedarb = request("FM_medarb")
'strFMAar = request("FM_aar")

strMrd = request("FM_start_mrd")
strDag = request("FM_start_dag")
strAar = right(request("FM_start_aar"),2) 
strDag_slut = request("FM_slut_dag")
strMrd_slut = request("FM_slut_mrd")
strAar_slut = right(request("FM_slut_aar"),2)


if intJoblog = "1" then
strUser = Request("medarb")
else
strUser = Session("user")
end if

%>
Indtaset af: <%= strUser %><br>
<%
	oRec.Open "SELECT * FROM timer WHERE Tmnavn = '"& strUser &"' AND Tid =" & StrTid & "", oConn, 3
	
	if not oRec.EOF then 
		  StrTid = oRec("Tid")
		  StrTdato = oRec("Tdato")
		  StrTimer = oRec("Timer")
		  StrTimerkom = oRec("Timerkom")
		  StrTjobnr = oRec("Tjobnr")
		  StrTjobnavn = oRec("Tjobnavn")
		  StrTknavn = oRec("Tknavn") 
		  StrUdato = "12/12/2007" 
		  intAktId = oRec("TAktivitetId")	
		  strAktnavn = oRec("Taktivitetnavn")
		  intOff = oRec("offentlig")	
	end if
		  
	'now close and clean up
	oRec.Close
	Set oRec = Nothing
	
%>

<form action="db_tastede_dage.asp?menu=<%=menu%>&jobnr=<%=intJobnr%>&eks=<%=eks%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=strJob%>&FM_medarb=<%=strMedarb%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" method="POST">
<input type="Hidden" name="joblog" value="<%=intJoblog%>">
<input type="Hidden" name="id" value="<%=StrTid%>">
<input type="Hidden" name="FM_jobnr" value="<%=StrTjobnr%>">
<input type="Hidden" name="FM_AktId" value="<%=intAktId%>">
<table cellspacing="2" cellpadding="2" border="0">
<tr>
    <td>Jobnavn:&nbsp;&nbsp;<%=StrTjobnavn %><br></td>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Jobnr:&nbsp;&nbsp;<%= StrTjobnr %></td>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Kundenavn:&nbsp;&nbsp;<%=StrTknavn %></td>
    <td>&nbsp;</td>
</tr>
<tr><td>Aktivitet:&nbsp;&nbsp;<b><%=strAktnavn%></b></td><td>&nbsp;</td></tr>
<tr>
    <td>Dato:&nbsp;&nbsp;
	<!--#include file="inc/dato2.asp"-->
	<select name="dag">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>&nbsp;
		
		<select name="mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
		
		
		<select name="aar">
		<option value="<%=strAar%>">
		<%if id <> 0 then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		<option value="02">2002</option>
		<option value="03">2003</option>
	   	<option value="04">2004</option>
	   	<option value="05">2005</option>
		<option value="06">2006</option>
		<option value="07">2007</option></select></td>

 		<td>&nbsp;</td>
</tr>
<tr>
    <td>
	<%if intjoblog = 1 then%>
	Timer:
	<%else%>
	Km:
	<%end if%>&nbsp;&nbsp;<input type="Text" name="Timer" Value="<%=StrTimer%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br></td>
    <td>&nbsp;</td>
</tr>
<tr>
    <td valign="top">Kommentar:&nbsp;&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr>
    <td align="left"><textarea name="Timerkom" cols="60" rows="6" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=StrTimerkom%></textarea></td>
	 <td>&nbsp;</td>
</tr>
<tr>
    <td align="left">Skjul kommentar på kundeside:
	<select name="FM_off" id="FM_off" style="font-size:9px;">
	<%
	if intOff = 0 then
	nejsel = "SELECTED"
	jasel = ""
	else
	nejsel = ""
	jasel = "SELECTED"
	end if
	%>
	<option value="0" <%=nejsel%>>Nej</option>
	<option value="1" <%=jasel%>>Ja</option>
	</select></td>
	 <td>&nbsp;</td>
</tr>
<tr>
<td>
<input type="image" src="../ill/opdaterpil.gif"></td>
<td>&nbsp;</td>
</tr>
</table>
</form>
</TD>
</TR>
</TABLE>
<!--#include file="../inc/regular/footer_inc.asp"-->
