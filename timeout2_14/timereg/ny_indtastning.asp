<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_inc.asp"-->
<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
</script>


<!--#include file="inc/dato1.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/vmenu.asp"-->
<!--#include file="../inc/regular/rmenu.asp"-->
	
	<div id="sindhold" style="position:absolute; left:170; top:80; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top">
 <!-------------------------------Sideindhold------------------------------------->
<br>

<font class="pageheader">Indtastning af timeforbrug</font>
<% 

	oRec.Open "SELECT * FROM Medarbejdere WHERE Mnavn='" & Session("User") & "'", oConn, 3
	
	
	While Not oRec.EOF
		strMnr = oRec("Mnr")
		strMnavn = oRec("Mnavn")
		strMansat = oRec("Mansat")
		strMkat = oRec("Mkat")  
		oRec.MoveNext
	Wend 
	
	oRec.Close
	
	Session("user")= strMnavn 
	
%>

<table cellspacing="2" cellpadding="2" border="0">
<tr>
	<form action="naeste.asp" method="POST"> 
	<input type="Hidden" name="FMnr" value="<%=strMnr%>">
	<!--<input type="Hidden" name="FDato" value="strdato">--> 
    <td>Indtastninger for dato:</td>
    <td>&nbsp; 
	<select name="dag">
		<option value="<%=session("strDag")%>"><%= session("strDag")%></option> 
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
		<option value="<%=session("strMrd")%>"><%=strMrdNavn%></option>
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
		<option value="<%=session("strAar")%>"><%=session("strAar")%></option>
		<option value="99">1999</option>
	   	<option value="00">2000</option>
	   	<option value="01">2001</option></select>
	</td>
</tr>
<tr>
    <td>Medarbejder:</td>
    <td>&nbsp;<b><%=strMnavn%></b></td>
</tr>
<tr>
    <td>Medarbejder nr:</td>
    <td>&nbsp;<%=strMnr%></td>
</tr>
<tr>
    <td>Funktion:</td>
    <td>&nbsp;<%=strMkat%></td>
</tr>
</table>




<table cellspacing="2" cellpadding="4" border="0">
<tr>
	<td colspan="3">Jobnr&nbsp;&nbsp;&nbsp;Jobnavn&nbsp;&nbsp;&nbsp;&nbsp;Aktivitet
	<br>
	<select name="Fjobnr" size="5" style="width:300;">
	<% 
	oRec.Open "SELECT job.id, jobnr, jobnavn, Jobknavn, aktiviteter.id AS Aid, aktiviteter.navn, aktiviteter.job  FROM job, aktiviteter WHERE jobstatus = 1 AND aktiviteter.job = job.id ORDER BY Jobnr DESC", oConn, 3 
	
	While Not oRec.EOF
	
		strJobnr = oRec("Jobnr")
		strJobnavn = oRec("Jobnavn")
		strknavn = oRec("Jobknavn")
		strknavntr = Left(strknavn, 10)
		
		%><option value='<%=strJobnr &","&oRec("Aid")%>'><%=oRec("Jobnr")%>&nbsp;&nbsp;&nbsp;<%=oRec("Jobnavn")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=oRec("navn")%></option>
	 	<%
	oRec.MoveNext
	Wend 
	
	oRec.Close
	%>
	</select></td>
 
</tr>
<tr> 

   <td colspan="3">Timer:&nbsp;&nbsp;<input type="Text" name="FTimer" maxlength="4" value="" size="4" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><font Size="1">&nbsp;&nbsp;(antal arbejdstimer angivet med 1 decimal)</font></td>

</tr>
<tr>
    <td valign="top">Kommentar:<br>
    <textarea name="FTimerkom" cols="35" rows="5" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></textarea></td>
	<td></td>
</tr>
<tr>
    <td align="right" colspan="3">&nbsp;</td>
<tr>
    <td align="center" colspan="3"><input type="submit" value="Indlæs registrering" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid; cursor:hand; width:116;"></td>
</tr>
</table>
</form></td>
</TR>
</TABLE>
</div>
<div id="onthefly" style="position:absolute; left:500; top:80; width:260; height:600; visibility:visible;">
<!--#include file="inc/se_data_onthefly.asp"-->
</div>
<!--#include file="../inc/regular/footer_inc.asp"-->