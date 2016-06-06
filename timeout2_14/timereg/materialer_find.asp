<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	
	
	func = request("func")
	
	if len(request("FM_varenr")) <> 0 then
	varenr = request("FM_varenr")
	showvarenr = varenr
	else
	varenr = 0
	showvarenr = ""
	end if 
	%>
	
	<script type="text/javascript">
     function updateordre(id){
     
    var nextn = 0;
	nextn = (window.opener.document.getElementById("lastn").value / 1) - 1
	
	//alert(nextn)
	for (m=0;m<9;m++){
	window.opener.document.getElementById("td_"+nextn+"_"+m+"").style.display = ""
	window.opener.document.getElementById("td_"+nextn+"_"+m+"").style.visibility = "visible"
	}
	
	
	var thisid = id;
     
     thisvarenavn = document.getElementById("FM_navn_"+thisid+"").value;
     thismatgrp = document.getElementById("FM_matgrp_"+thisid+"").value;
     thisvarenr = document.getElementById("FM_varenr_"+thisid+"").value;
     thisbetegnelse = document.getElementById("FM_betegnelse_"+thisid+"").value;
     thisantal = document.getElementById("FM_antal_"+thisid+"").value;
     
    ordrelinVal = window.opener.document.getElementById("lastn").value;
     window.opener.document.getElementById("FM_navn_"+ordrelinVal+"").value = thisvarenavn;
      window.opener.document.getElementById("FM_varenr_"+ordrelinVal+"").value = thisvarenr;
       window.opener.document.getElementById("FM_gruppe_"+ordrelinVal+"").value = thismatgrp;
        window.opener.document.getElementById("FM_betegn_"+ordrelinVal+"").value =  thisbetegnelse;
        window.opener.document.getElementById("FM_antal_"+ordrelinVal+"").value =  thisantal;
        
        window.opener.document.getElementById("lastn").value = nextn
       
   }
    </script>
	
	
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<div id=side style="position:absolute; left:20px; top:20px;">
	<h4>Find materiale:</h4>
	
	<form action="materialer_find.asp?func=find" method="post">
        Søg på navn ell. varenr: <input id="FM_varenr" name="FM_varenr" type="text" size=10 value="<%=showvarenr%>" />
        <input id="Submit1" type="submit" value="Find >>" />
	</form>
	
	
	<%
	if func = "find" then 
	%>
	
	
	<table cellspacing=0 cellpadding=0 border=0 width=100%><tr>
	<form>
	<td>Gruppenavn / nr.</td>
	<td>Varenavn</td>
	<td>Varenr</td>
	<td>Betegnelse</td>
	<td>På lager</td>
    <td>Min. lager</td>
	<td>
        &nbsp;</td>
	</tr>
	<%
	
	strSQL = "SELECT m.id, m.navn, m.varenr, m.betegnelse, m.matgrp, m.antal, "_
	&" m.minlager, mg.navn AS gruppenavn, mg.nummer AS gruppenr "_
	&" FROM materialer m LEFT JOIN materiale_grp mg ON (mg.id = m.matgrp)  "_
	&" WHERE m.varenr = '"& varenr &"' OR m.navn LIKE '"& varenr &"%'"
	
	'Response.Write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	%>
	
	<tr>
	<td><input id="FM_matgrp_<%=oRec("id")%>" size=10 type="text" value="<%=oRec("gruppenavn") %>"/></td>
	<td><input id="FM_navn_<%=oRec("id")%>" size=10 type="text" value="<%=oRec("navn") %>" /></td>
	<td><input id="FM_varenr_<%=oRec("id")%>" size=6 type="text" value="<%=oRec("varenr") %>"/></td>
	<td><input id="FM_betegnelse_<%=oRec("id")%>" size=10 type="text" value="<%=oRec("betegnelse") %>"/></td>
	<td><%=oRec("antal") %></td>
	<td><%=oRec("minlager") %></td>
	
	<%
	bestilantal = (oRec("antal") - oRec("minlager"))
	if bestilantal < 0 then
	bestilantal = -(bestilantal)
	else
	bestilantal = 0
	end if
	%>
	
	<td><input id="FM_antal_<%=oRec("id")%>" size=2 type="text" value="<%=bestilantal %>"/></td>
	<td>
        &nbsp;<input id="bt_<%=oRec("id")%>" type="button" value=" >> " onclick="updateordre(<%=oRec("id")%>)" /></td>
	</tr>
	
	
	<%
	
	oRec.movenext
	wend
	oRec.close
	%>
	</table>
	</form>
	
	<%
	
	end if %>
	
	</div>
	
	

<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
