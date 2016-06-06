<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	rdir = request("rdir")
	menu = "sdsk"
	func = request("func")
	
	thisfile = "sdsk_tilfoj"
	'Response.write "func: " & func
	
	if func = "dbopr" then
	id = 0
	else
		if len(request("id")) <> 0 then
		id = request("id")
		else
		id = 0
		end if
	end if
	
	'**** Kundelogin ?? ***'
	kview = request("usekview")
	
	'*** Til topmenu ***
	if len(request("FM_kontakt")) <> 0 then
	kontaktId = request("FM_kontakt")
	else
	kontaktId = 0
	end if
	
	if len(request("lastedit")) <> 0 then
	lastedit = request("lastedit")
	else
	lastedit = 0
	end if
	
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless2 = tmp
		end function
	
	
	select case func	
	case "dbopr", "dbred"
				
				
		if len(trim(request("FM_besk"))) <> 0 then
		strBesk = SQLBless(request("FM_besk"))
		else
		strBesk = ""
		end if
		
		call htmlreplace(strBesk)
		strBesk = htmlparseTxt
		
		datoSQL =  year(now)&"/"&month(now)&"/"&day(now)
		editor = session("user")
		editor2 = editor
		tid = time
		
		
		if len(request("FM_public")) <> 0 then
		pblic = 1
		else
		pblic = 0
		end if
		
		
		if len(request("FM_sdskrel")) <> 0 then
		intSdsk_rel = request("FM_sdskrel")
		else
		intSdsk_rel = 0
		end if
		
		
		if len(request("FM_losning")) <> 0 then
		intLosning = request("FM_losning")
		else
		intLosning = 0
		end if
		
		
		
		if func = "dbopr" then
		
		strSQL = "INSERT INTO sdsk_rel (besk, dato, editor, sdsktidspunkt, public, sdsk_rel, losning, sdskdato) "_
		&" VALUES ('"& strBesk &"', '"& datoSQL &"', '"& editor &"','" & tid & "', "& pblic &", "& intSdsk_rel &", "& intLosning &", '"& datoSQL &"')"
		
		'Response.write strSQL
		'Response.flush
		
		oConn.execute(strSQL)
					
		else
			
			
			strSQL = "UPDATE sdsk_rel SET besk = '"& strBesk &"', dato2 = '"& datoSQL &"', editor2 = '"& editor2 &"', "_
			&" public = "& pblic &", losning = "& intLosning &" WHERE id = " & id
			
			oConn.execute(strSQL)
			
	    end if
	    
	   
		if rdir <> "redsdsk" then	
		Response.Write("<script language=""JavaScript"">opener.location.href = 'sdsk.asp?lastedit="+lastedit+"&usekview="+kview+"&FM_kontakt="+kontaktId+"';</script>")
		else
		Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
		''Response.Write("<script language=""JavaScript"">opener.location.href = 'sdsk.asp?func=red&id="+id+"';</script>")
		end if
		
		
		
		Response.Write("<script language=""JavaScript"">window.close();</script>")	
		
		
		'Response.redirect "sdsk_tilfoj.asp?func=red&lastedit="&lastedit&"&sdskrelid="&intSdsk_rel&"&usekview="&kview&"&FM_kontakt="&kontaktId&"&id="&id
		''Response.end
	case "red","opr"
			
			%>
			<!--#include file="../inc/regular/header_hvd_inc.asp"-->
			<%
			
			if len(request("sdskrelid")) <> 0 then
			sdskrelid = request("sdskrelid")
			else
			sdskrelid = 0
			end if 
	
	
			if func = "red" then
				
				strSQL = "SELECT besk, dato, editor, sdsktidspunkt, public, losning "_
				&" FROM sdsk_rel WHERE id = "&id
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
				
					strBesk = oRec("besk")
					editor = oRec("editor")
					dtTidspunkt = oRec("sdsktidspunkt")
					intPublic = oRec("public") 
					dtDato = oRec("dato")
					intLosning = oRec("losning")
				
				end if
				oRec.close
			
				nFunc = "dbred"
				
				
			else
				
				strBesk = ""
				nFunc = "dbopr" 
				intPublic = 0
				
			
			end if
	
	
	
	if func = "red" then
	sVal = "Opdater"
	else
	sVal = "Opret"
	end if%>
	
	<body onLoad="self.focus()">
	
	<div id="sideindhold" name="sideindhold" style="position:absolute; left:20px; top:20px;">
	<h3><%=sVal%> Incident</h3>
	
	<a href="#" class=red>Luk og opdater Incidentliste [x]</a><br />
	<br />
	
	<%if func = "red" then%>
	Sidst redigeret af: <b><%=editor%></b> d. <%=formatdatetime(dtDato, 1)%>
	<%end if%>
	<table cellspacing=2 cellpadding=2 border=0>
	<form method=post action="sdsk_tilfoj.asp?func=<%=nFunc%>&id=<%=id%>&lastedit=<%=lastedit%>">
	<input type="hidden" name="FM_sdskrel" id="FM_sdskrel" value="<%=sdskrelid%>">
	<input type="hidden" name="usekview" id="usekview" value="<%=kview%>">
	<input type="hidden" name="FM_kontakt" id="FM_kontakt" value="<%=kontaktId%>">
    <input id="Hidden1" name="rdir" value="<%=rdir %>" type="hidden" />
	<tr>
		<td colspan=2><b>Beskrivelse</b><br>
		
		
		 <%
		                dim content
	                    content = strBesk
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_besk"
			            editorK.Text = content
			            editorK.FilesPath = "CuteEditor_Files"
			            editorK.AutoConfigure = "Minimal"
            			
			            editorK.Width = 500
			            editorK.Height = 280
			            editorK.Draw()
		                %>
		
		
		
		<!--<textarea cols="57" rows="7" name="FM_besk" id="FM_besk" id="FM_besk"><strBesk%></textarea>-->
		</td>
	</tr>
	<tr>
		<%
		if intPublic = 1 OR (func = "opr" AND kview = "j") then
		pubChecked = "CHECKED"
		else
		pubChecked = ""
		end if
		%>
		<td colspan=2><input type="checkbox" name="FM_public" id="FM_public" value="1" <%=pubChecked%>>Offentlig. (Incident tilgængelig for ekstern tilknyttet kontakt.)</td>
	</tr>
	<tr>
		<%
		if intLosning = 1 then
		losChecked = "CHECKED"
		else
		losChecked = ""
		end if
		%>
		<td colspan=2><input type="checkbox" name="FM_losning" id="FM_losning" value="1" <%=losChecked%>>Godkendt som løsning.</td>
	</tr>
	
	
	<tr>
		<td colspan=2 align=center><br><input type="submit" value="<%=sVal%> Incident"></td>
	</form>
	</tr>
	</table>
	</div>

<%
end select
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
