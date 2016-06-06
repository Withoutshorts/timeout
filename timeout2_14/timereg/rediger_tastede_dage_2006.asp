<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/regular/header_hvd_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->		
<!--#include file="inc/convertDate.asp"--> 
<div id="sindhold" style="position:absolute; left:20px; top:5px; width:60%; height:600px; visibility:visible;">
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top">
	<!-------------------------------Sideindhold------------------------------------->
	<br /><h3><%=tsa_txt_181 %></h3>

<%=tsa_txt_180 %>:&nbsp;<%=formatdatetime(date, 1)%><br>

<%

call erkmDialog() 

id = Request("id")

jobid = 0
intAktId = 0
inttknr = 0


	strSQL = "SELECT tdato, timer, timerkom, tjobnr, tjobnavn, "_
	&" tknavn, TAktivitetId, Taktivitetnavn, tmnavn, tknr, k.kkundenr, offentlig, timepris, tmnr,"_
	&" t.editor, tastedato, sttid, sltid, t.valuta, j.id AS jid, bopal FROM timer t "_
	&" LEFT JOIN job j ON (j.jobnr = t.tjobnr) "_ 
	&" LEFT JOIN kunder k ON (k.kid = tknr) WHERE Tid =" & id 
	
	'Response.write strSQL
	'Response.flush
	
	oRec.Open strSQL, oConn, 3
	
	if not oRec.EOF then 
		
		  StrTdato = oRec("Tdato")
		  StrTimer = oRec("Timer")
		  StrTimerkom = oRec("Timerkom")
		  jobnr = oRec("tjobnr")
		  jobid = oRec("jid")
		  StrTjobnavn = oRec("Tjobnavn")
		  StrTknavn = oRec("Tknavn") 
		  StrUdato = "12/12/2007" 
		  intAktId = oRec("TAktivitetId")	
		  strAktnavn = oRec("Taktivitetnavn")
		  intOff = oRec("offentlig")	
		  strUser = oRec("tmnavn")
		  inttknr = oRec("kkundenr")
		  timepris = oRec("timepris")
		  medid = oRec("tmnr")
		  editor = oRec("editor")
		  tastedato = oRec("tastedato") 
		  
		 	 if len(trim(oRec("sttid"))) <> 0 then
				if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
				sttid = left(formatdatetime(oRec("sttid"), 3), 5)
				else
				sttid = ""
				end if
			else
			sttid = ""
			end if
	
			if len(trim(oRec("sltid"))) <> 0 then
				if left(formatdatetime(oRec("sltid"), 3), 5) <> "00:00" then
				sltid = left(formatdatetime(oRec("sltid"), 3), 5)
				else
				sltid = ""
				end if
			else
			sltid = ""
			end if
			
			intValuta = oRec("valuta")
			
			bopal = oRec("bopal")
		  
	end if
	'now close and clean up
	oRec.Close
	Set oRec = Nothing	
%>




<%=tsa_txt_182 %>: <b><%=editor%></b> <b><%=formatdatetime(tastedato, 1)%></b><br><br>
<table cellspacing="2" cellpadding="2" border="0" width=500>
    <tr><td align="right" colspan="2"><a href="db_tastede_dage_2006.asp?func=slet&id=<%=id %>" id="slet" style="color:red;">[X - <%=tsa_txt_221 %>]</a></td></tr>
<form action="db_tastede_dage_2006.asp" method="POST">
<input type="Hidden" name="id" value="<%=id%>">
<input type="Hidden" name="medid" value="<%=medid%>">
<input type="Hidden" name="jobnr" value="<%=jobnr%>">


<tr>
    <td width=90><b><%=tsa_txt_022 %>:</b></td><td><%=left(StrTknavn, 30) %> (<%=inttknr%>)</td>
</tr>
<tr>
    <td><b><%=tsa_txt_067 %>:</b></td><td><%=left(StrTjobnavn, 30) %> (<%= jobnr %>)</td>
 </tr>

 
	<tr>
		<td><b><%=tsa_txt_068 %>:</b></td><td>


		
		<%
        
        if session("rettigheder") = 1 then 
        strSQLakt = "SELECT navn, id FROM aktiviteter WHERE job = "& jobid &" ORDER BY fase, sortorder, navn"
        'Response.write "<br><br><br><br><br><br><br><br><br><br>"& strSQLakt
        'Response.flush
		oRec3.open strSQLakt, oConn, 3
		
         %>
        <select id="Select2" name="intAktId" style="width:200px;">
        <%
		
      
		while not oRec3.EOF 
		
		if cdbl(intAktId) = oRec3("id") then
	    aSel = "SELECTED"
        else
		aSel = ""
		end if%>
		 <option <%=aSel %> value="<%=oRec3("id") %>"><%=oRec3("navn") %></option>
		<%
		
		oRec3.movenext
		wend 
		oRec3.close
		 %>
		 
		  </select>

          <%else 
          
          strSQLakt = "SELECT navn, id FROM aktiviteter WHERE id = "& intAktId
          oRec3.open strSQLakt, oConn, 3
        aNavn = "(ikke fundet)"
		if not oRec3.EOF then 
	    aNavn = oRec3("navn")
	    end if 
		oRec3.close
          
          Response.write aNavn

          %>
<input type="hidden" name="intAktId" value="<%=intAktId %>" />
<%end if %>

		</td>
	</tr>


<tr>
	<td><b><%=tsa_txt_101 %>:</b></td><td><%= strUser %></td>
</tr>
<tr>
    <td><b><%=tsa_txt_183 %>:</b></td><td>
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
		<option value="07">2007</option>
		<option value="08">2008</option>
		<option value="09">2009</option>
		<option value="10">2010</option>
		<option value="11">2011</option>
		<option value="12">2012</option>
		<option value="13">2013</option>
		<option value="14">2014</option>
        <option value="15">2015</option>
        <option value="16">2016</option>
        <option value="17">2017</option>
        <option value="18">2018</option>
        <option value="19">2019</option></select></td>

 		
</tr>
<tr>
    <td><b><%=tsa_txt_070 %>:</b></td><td><input type="Text" name="Timer" Value="<%=StrTimer%>" style="border: 1px #86B5E4 solid; width:45px;">
	&nbsp;&nbsp;<b><%=tsa_txt_184 %>:</b>&nbsp;<input type="text" name="FM_sttid" value="<%=sttid%>" size=2> - 
	<input type="text" size=2 name="FM_sltid" value="<%=sltid%>"> (<%=tsa_txt_185 %>)</td>
  </tr>
  
  <%if session("rettigheder") <= 2 OR session("rettigheder") = 6 then %>
  <tr>
    <td valign=top style="padding-top:5px;"><b><%=tsa_txt_186 %>:</b></td><td><input type="Text" name="timepris" Value="<%=timepris%>"  style="border: 1px #86B5E4 solid; width:80px;">
	 &nbsp;<select name="FM_valuta" id="Select1" style="width:70px;">
		    <option value="<%=intValuta %>"><%=tsa_txt_229 %></option>
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(oRec3("id")) = cint(intValuta) then
		    valGrpCHK = "SELECTED"
		    else
		    valGrpCHK = ""
		    end if
    		
    		
		    %>
		    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode")%></option>
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>

            <br />
              <input type="checkbox" value="1" name="FM_opdater_timepriser" /> Opdater timepris på denne aktivitet for valgte medarbejder<br /><br />
        Opdater timepriser fra d. <input type="text" name="FM_opdatertpfra" value="<%=formatdatetime(now,2) %>" style="font-size:9px; width:60px;" /> til dd.<br />
        (også lukkede uger og hvis der foreligger faktura)
	
	<br>
        &nbsp;
	</td>
  </tr> 
  <%else %>
  <input type="hidden" name="timepris" value="<%=timepris%>" />
  <input type="hidden" name="FM_valuta" value="<%=intValuta %>" />
  <%end if %>


 <tr>
    <td valign="top" colspan=2><b><%=tsa_txt_051 %>:</b></td>
</tr>
</table>
<table cellspacing="2" cellpadding="2" border="0">
<tr>
    <td align="left" colspan=2><textarea name="Timerkom" cols="60" rows="5" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=StrTimerkom%></textarea></td>
</tr>
<tr>
    <td align="left" colspan=2><%=tsa_txt_053 %>:
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
	<option value="0" <%=nejsel%>><%=tsa_txt_054 %></option>
	<option value="1" <%=jasel%>><%=tsa_txt_055 %></option>
	</select></td>
</tr>
<%


if kmDialogOnOff = 1 then %>
  <tr>
    <td colspan=2>
    
    <%if cint(bopal) = 1 then 
    bpCHK = "CHECKED"
    else
    bpCHK = ""
    end if%>
    
        <input id="Checkbox1" name="FM_bopal" value="1" type="checkbox" <%=bpCHK %> /> <%=tsa_txt_287 %> (<%=tsa_txt_288%>)
	</td>
  </tr>
  <%else %>
<input id="name="FM_bopal"" name="FM_bopal" value="0" type="hidden" />
  <%end if %>
<tr>
	<td colspan=2 align=center>
        <input id="Submit1" type="submit" value="<%=tsa_txt_085 %>" /></td>
</tr>
</table>
</form>
</TD>
</TR>
</TABLE>
</div>
<!--#include file="../inc/regular/footer_inc.asp"-->
