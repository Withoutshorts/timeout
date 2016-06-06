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
	
	
	id = request("pic")
	
	if len(trim(request("matid"))) <> 0 then
	matid = request("matid")
	else
	matid = 0
	end if
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<body>
	<!-------------------------------Sideindhold------------------------------------->
	
	<%
		strSQL = "select id, filnavn FROM filer WHERE id = " & id
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
			strPicNavn = "../inc/upload/"&lto&"/"&oRec("filnavn")
			strPicNavn_only = oRec("filnavn")
		end if
		oRec.close 
		
		
	if matid <> 0 then
	strSQL = "SELECT m.id, m.navn, m.varenr, m.antal, mg.navn AS gnavn, m.matgrp, "_
	&" m.enhed, m.pic, m.minlager, mg.av, m.indkobspris, m.salgspris, m.betegnelse, m.lokation "_
	&" FROM materialer m LEFT JOIN materiale_grp mg "_
	&" ON (mg.id = m.matgrp) WHERE m.id = "& matid &""
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
			
			matnavn = oRec("navn")
			matbetegn = oRec("betegnelse")
			matvarenr = oRec("varenr")
			matlok = oRec("lokation")
			matgrp = oRec("matgrp")
			matenhed = oRec("enhed")
			
			
		end if
		oRec.close 
	
	end if
		%>
		
	<div id="toptable" style="position:absolute; left:20px; top:50px; visibility:visible; width:520px;">
	<img src="<%=strPicNavn%>" width="150" height="100" alt="" border="0"> <br>
	<%=strPicNavn_only%>
	
	<%if matid <> 0 then %>
	<br /><br />
	<b>Navn:</b> <%=matnavn%><br />
	<b>Varenr:</b> <%=matvarenr %><br />
	<b>Gruppe:</b> <%=matgrp %><br />
	<b>Lokation:</b> <%=matlok %><br /><br />
	<b>Betegnelse:</b><br />
	<%=matbetegn%><br /><br />
	
	
	
	<%end if %>
	
	<br><br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="Luk Vindue" border="0"></a>
	<br><br>&nbsp;
	</div>
	
	
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
