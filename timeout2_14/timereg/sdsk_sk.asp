<% response.buffer = true %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/sdsk_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/sdsk_inc.asp"-->
<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<%
'section for ajax calls
if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")
case "FM_ajaxstatus"
Call AjaxUpdate("sdsk","status","ændringen i status blev gemt")
	strMailID = Request.Form("id")
	intStatus = Request.Form("value")
	
			Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			' Sætter Charsettet til ISO-8859-1 
			Mailer.CharSet = 2
			' Afsenderens navn 
			Mailer.FromName = "TimeOut ServiceDesk"
			' Afsenderens e-mail 
			Mailer.FromAddress = "timeout_no_reply@outzource.dk"
			Mailer.RemoteHost = "webmail.abusiness.dk"
			Mailer.ContentType = "text/html"
		    
		    
		    '** Henter Afsender ***
			strSQL = "SELECT mnavn, mnr, init, email FROM medarbejdere WHERE mid = " & session("mid")
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
				
			afsNavn = oRec("mnavn") & "("& oRec("mnr") &") - " & oRec("init")
			afsEmail = oRec("email")
				
			oRec.movenext
			wend
				
			oRec.close
			
		            bcc = ""
		            call eml_inciAns()
		            call eml_inciCreator()
		            call eml_proGrp()
								call eml_Kpers(kpers)
					
					if instr(kpers2email, "@") <> 0 then
		            Mailer.AddBCC ""&kpers2&"",""& kpers2email  
		            bcc = bcc & "Kontaktperson: "&kpers2&", "& kpers2email & "<br>"
	                end if
	                
					call eml_kogjobAns()
				
				'** Henter statusnavn ***
				strSQL = "SELECT navn FROM sdsk_status WHERE id = " & intStatus
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				statusNavn = oRec("navn") 
				
				oRec.movenext
				wend
				oRec.close
				
			
		    ' Mailens emne
			Mailer.Subject = "Opdateret Incident (ID: "& StrMailID &") - "& strEmne 
			' Modtagerens navn og e-mail
			Mailer.AddRecipient ansvNavn, ansvEmail
			
			''*** Mail til Kundeansvarlig ***
			'if kundeans1 <> intAns AND kundeans1 <> 0 then
			'	Mailer.AddCC ""&modKansNavn&"",""& modKansEmail 
		    '		
		    '		if instr(modKansEmail, "@") <> 0 AND instr(bcc, modKansEmail) = 0 then
		    '	    bcc = bcc & ""&oRec("mnavn")&","& oRec("email") & "<br>"
		    '       end if
				 
		    '	end if
			
			' Selve teksten
					Mailer.BodyText = "Status på Incident: (ID: "& strMailID &") er blevet opdateret." & vbCrLf & vbCrLf _ 
					& "==============================================================" & vbCrLf _
					& "Status er nu: " & statusNavn & vbCrLf _
					& "Modtagere af denne mail:" & vbCrLf & replace(bcc, "<br>", " -- ") & vbCrLf & vbCrLf _
					& "Med venlig hilsen"& vbCrLf & afsNavn & ", "& afsEmail & vbCrLf 
					'& "Incident ansvarlig: " & ansvNavn & ", " & ansvEmail & vbCrLf & vbCrLf _ bcc
					'& "Adressen til TimeOut er: https://outzource.dk/"&lto&"" & vbCrLf & vbCrLf _ 
				
					If Mailer.SendMail Then
    				
					Else
    					Response.Write "Fejl...<br>" & Mailer.Response
  					End if	
case "FM_sortOrder"
Call	AjaxUpdate("sdsk","sortorder","")
case "FM_ajaxtype"
Call AjaxUpdate("sdsk","type","ændringen i kategori blev gemt")
case "FM_ajaxprio"
Call AjaxUpdate("sdsk","prioitet","ændringen i type blev gemt")
case "FM_ajaxprio_typ"
Call AjaxUpdate("sdsk","priotype","ændringen i prioritet blev gemt")
case "FM_CustDesc"
	strSQL = "Update kunder set beskrivelse = '" & request.Form("value") & "' where Kid = " & request.Form("id")
	'response.Write strSQL
	oConn.execute(strSQL)
case "FN_getCustDesc"

strSQL = "SELECT beskrivelse, Kid FROM kunder WHERE Kid = " & Request.Form("cust")
				'Response.write strSQL
				'Response.flush
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				Response.Write oRec("beskrivelse") 
				oRec.movenext
				wend
				oRec.close
End select
Response.End
end if


	
 %>
	<script>
	var newwindow = '';

	function popitup(url) {
	
	url = url +"&kundeid="+ document.getElementById("FM_kontakt").value
	
		if (!newwindow.closed && newwindow.location) {
			newwindow.location.href = url;
		}
		else {
			newwindow=window.open(url,'name','height=500,width=350,left=100,top=100');
			if (!newwindow.opener) newwindow.opener = self;
		}
		if (window.focus) {newwindow.focus()}
		return false;
	}
	
	function renssog(){
	document.getElementById("sog").value = ""
	}
	
	
	function BreakItUp()
	{
	  //Set the limit for field size.
	  var FormLimit = 302399 
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm.BigTextArea.value
	
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm.appendChild(objTEXTAREA)
	
	      TempVar = TempVar.substr(FormLimit)
	    }
	  }
	}
	
	function visduedate(){
	//alert(document.getElementById("FM_useduedate").checked)
	if (document.getElementById("FM_useduedate").checked == true) {
	document.getElementById("FM_useduedate_dag").disabled = false
    document.getElementById("FM_useduedate_md").disabled = false
	document.getElementById("FM_useduedate_aar").disabled = false
	document.getElementById("FM_duetime").disabled = false
	//document.getElementById("FM_esttid").disabled = false
	} else {
	document.getElementById("FM_useduedate_dag").disabled = true
    document.getElementById("FM_useduedate_md").disabled = true
	document.getElementById("FM_useduedate_aar").disabled = true
	document.getElementById("FM_duetime").disabled = true
	//document.getElementById("FM_esttid").disabled = true
	}
	}
	
	//function setEmneFocus(){
	//alert("her")
	//document.getElementById("sdsk_emne").focus() 
	//}
	
	</script>

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
	menu = "sdsk"
	func = request("func")
	
	if func = "opr1" AND request("FM_kontakt") = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%
	errortype = 77
	call showError(errortype)
	
	Response.end 
	end if 
	
	
	
	thisfile = "sdsk"
	
	if func = "dbopr" then
	id = 0
	else
		if len(request("id")) <> 0 then
		id = request("id")
		else
		id = 0
		end if
	end if
	
	print = request("print")
	
	dato = day(now)&"/"&month(now)&"/"&year(now)
	tid = time
	
	if len(request("lastedit")) <> 0 then
	lastedit = request("lastedit")
	else
	lastedit = 0
	end if
	
	if len(Request("bcc")) <> 0 then
	bcc = Request("bcc")
	showbcc = 1
	else
	showbcc = 0
	bcc = ""
	end if
	
	'**** Kundelogin ?? ***'
	kview = request("usekview")
	
	
	'*** Til topmenu ***
	if len(request("FM_kontakt")) <> 0 then
	kontaktId = request("FM_kontakt")
	else
	kontaktId = 0
	end if
	
	
	call datocookie
	stDatoSQL = strAar &"/"& strMrd &"/"& strDag
	slDatoSQL = strAar_slut &"/"& strMrd_slut &"/"& strDag_slut
	
	
	
	
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
	
	
	
	if func = "slet" then
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:40; visibility:visible;">
	
	<br><br><br>
	<div id="slet" style="position:relative; left:0; top:10; visibility:visible; border:1px red dashed; background-color:#ffff99; width:300px; padding:10px;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> et Incident. Er dette korrekt?<br>
		Du vil samtidig slette hele loggen på dette Incident.&nbsp;</td>
	</tr>
	<tr>
	   <td><a href="sdsk.asp?func=sletok&id=<%=id%>">Ja, slet Incident og tilhørende log >> </a><br><br><a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	</div>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	Response.end
	end if
	
	if func = "sletok" then
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM sdsk WHERE id = "& id &"")
	oConn.execute("DELETE FROM sdsk_rel WHERE sdsk_rel = "& id &"")
	Response.redirect "sdsk.asp"
	end if
	
		
	if func = "dbopr" OR func = "dbred" then
	
	
				if request("FM_kontakt") = "0" then
					antalErr = 1
					errortype = 69
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"--><%
					call showError(errortype)
					response.end
				end if
		
			
				if len(trim(request("FM_emne"))) = 0 then
					antalErr = 1
					errortype = 68
							 
							if kview = "j" then
							%>
							<!--#include file="../inc/regular/header_hvd_inc.asp"-->
							<%
							else
							%>
							<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
							<%
							end if
							%>
					<!--#include file="../inc/regular/topmenu_inc.asp"--><%
					call showError(errortype)
					response.end
				end if
				
				
				
				if request("FM_type") = "0" then
					antalErr = 1
					errortype = 70
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"--><%
					call showError(errortype)
					response.end
				end if
				
				if request("FM_prio") = "0" then
					antalErr = 1
					errortype = 71
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"--><%
					call showError(errortype)
					response.end
				end if
				
				if request("FM_status") = "0" then
					antalErr = 1
					errortype = 72
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"--><%
					call showError(errortype)
					response.end
				end if
				
			
				
				
				
				
		
		if request("FM_useduedate") = "1" then
		
		
		duedate = request("FM_start_aar_due") & "-" & request("FM_start_mrd_due") &"-"& request("FM_start_dag_due") &" "& request("FM_duetime")&":00"
		
		'Response.Write "duedate: " & duedate
		'Response.end
		
		if len(trim(request("FM_esttid"))) <> 0 then
		dblesttid = SQLBless2(request("FM_esttid"))
		else
		dblesttid = 0
		end if
		
		
		useduedate = 1
		
		        
		         if isdate(duedate) = false then
					antalErr = 1
					errortype = 121
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"--><%
					call showError(errortype)
					response.end
				end if
		        
		
		else
		
		duedate = "2001-1-1"
		useduedate = 0
		dblesttid = 0
		
		end if
		
		
		       
		
		
		
		
				
	
		strEmne = replace(request("FM_emne"), "'", Chr(34))
		
		
		
		if len(trim(request("FM_besk"))) <> 0 then
		strBesk = replace(request("FM_besk"), "'", Chr(34))
		'strBesk = replace(strBesk, "\", "#Backslash#")
		else
		strBesk = ""
		end if
		
		call htmlreplace(strBesk)
		strBesk = htmlparseTxt
		
		
		intType = request("FM_type")
		intPrio = request("FM_prio")
		intPriotype = request("FM_prio_typ")
		intAns = request("FM_ans")
		intKontakt = request("FM_kontakt")
		intStatus = request("FM_status")
		
		datoSQL =  year(now)&"/"&month(now)&"/"&day(now)
		
		editor = session("user")
		if kview = "j" then
		editor = editor &", "& session("kontaktpersEmail")
		end if
		
		if len(request("FM_forny")) <> 0 then
		forny = 1
		else
		forny = 0
		end if
		
		if len(request("FM_public")) <> 0 then
		pblic = 1
		else
		pblic = 0
		end if
		
		if len(request("FM_jobid")) <> 0 then
		jobid = request("FM_jobid")
		else
		jobid = 0
		end if
		
		if InStr(jobid, ",") then
		jobid = split(jobid,",")(1)
		end if
		
		strSogeord_1 = SQLBless(request("FM_sogeord_1"))
		strSogeord_2 = SQLBless(request("FM_sogeord_2"))
		strSogeord_3 = SQLBless(request("FM_sogeord_3"))
		strSogeord_4 = SQLBless(request("FM_sogeord_4"))
		
		creator = request("FM_creator")
		
		kpers = request("FM_kpers")
		kpers2 = replace(request("FM_kpers2"), "'", "''")
		kpers2email = replace(request("FM_kpers2email"), "'", "''")
		
		if func = "dbopr" then
		
		strSQL = "INSERT INTO sdsk (emne, besk, type, prioitet, priotype, kundeid, jobid, ansvarlig, esttid, dato, "_
		&" editor, tidspunkt, public, status, sogeord1, sogeord2, sogeord3, sogeord4, "_
		&" creator, kpers, kpers2, kpers2email, duedate, useduedate) "_
		&" VALUES ('"& strEmne &"', '"& strBesk &"', "& intType &", "& intPrio &", "& intPriotype &", "& intKontakt &", "& jobid &","_
		&" "& intAns &", "& dblesttid &", '"& datoSQL &"', '"& editor &"','" & time & "', "_
		&" "& pblic &", "& intStatus &", '"& strSogeord_1 &"', '"& strSogeord_2 &"', '"& strSogeord_3 &"', '"& strSogeord_4 &"',"_
		&" '"& creator &"', "& kpers &", '"& kpers2 &"', '"& kpers2email &"', '"& duedate &"', "& useduedate &")"
		
		
		'Response.write strSQL
		'Response.write flush
			
		oConn.execute(strSQL)
					
					
				strSQL = "SELECT id FROM sdsk ORDER by id DESC LIMIT 0, 1"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				lastedit = oRec("id")
				oRec.movenext
				wend
				
				oRec.close
				
			
			
			
			'***** Oprettter Mail object ***
			if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\sdsk.asp" then
			
			
			Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			' Sætter Charsettet til ISO-8859-1 
			Mailer.CharSet = 2
			' Afsenderens navn 
			Mailer.FromName = "TimeOut ServiceDesk"
			' Afsenderens e-mail 
			Mailer.FromAddress = "timeout_no_reply@outzource.dk"
			Mailer.RemoteHost = "webmail.abusiness.dk"
			Mailer.ContentType = "text/html"
		    
		    
		    '** Henter Afsender ***
			strSQL = "SELECT mnavn, mnr, init, email FROM medarbejdere WHERE mid = " & session("mid")
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
				
			afsNavn = oRec("mnavn") & "("& oRec("mnr") &") - " & oRec("init")
			afsEmail = oRec("email")
				
			oRec.movenext
			wend
				
			oRec.close
			
			
			
			
		            bcc = ""
		            call eml_inciAns()
		            call eml_inciCreator()
		            call eml_proGrp()
					call eml_Kpers(kpers)
					
					if instr(kpers2email, "@") <> 0 then
		            Mailer.AddBCC ""&kpers2&"",""& kpers2email  
		            bcc = bcc & "Kontaktperson: "&kpers2&", "& kpers2email & "<br>"
	                end if
	                
					call eml_kogjobAns()
				
				
				
				
				'** Henter statusnavn ***
				strSQL = "SELECT navn FROM sdsk_status WHERE id = " & intStatus
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				statusNavn = oRec("navn") 
				
				oRec.movenext
				wend
				oRec.close
				
				
				
			
		    ' Mailens emne
			Mailer.Subject = "Ny Incident (ID: "& lastedit &") - "& strEmne 
			' Modtagerens navn og e-mail
			Mailer.AddRecipient ansvNavn, ansvEmail
			
			''*** Mail til Kundeansvarlig ***
			'if kundeans1 <> intAns AND kundeans1 <> 0 then
			'	Mailer.AddCC ""&modKansNavn&"",""& modKansEmail 
		    '		
		    '		if instr(modKansEmail, "@") <> 0 AND instr(bcc, modKansEmail) = 0 then
		    '	    bcc = bcc & ""&oRec("mnavn")&","& oRec("email") & "<br>"
		    '       end if
				 
		    '	end if
			
			' Selve teksten
					Mailer.BodyText = "Ny Incident: (ID: "& lastedit &") er oprettet." & vbCrLf & vbCrLf _ 
					& "==============================================================" & vbCrLf _
					& "Kontakt: " & vbCrLf _
					& knavn & " ("& knr &") " & vbCrLf _
					& adresse & vbCrLf _
					& postnr & ", " & by  & vbCrLf _
					& "Tlf:" & telefon  & vbCrLf _
					& "==============================================================" & vbCrLf & vbCrLf & vbCrLf  _
					& "Status: "& vbCrLf & statusNavn & vbCrLf & vbCrLf _
					& "Emne og beskrivelse: "& vbCrLf & strEmne & vbCrLf _
					& strBesk & vbCrLf & vbCrLf & vbCrLf _ 
					& "Modtagere af denne mail:" & vbCrLf & replace(bcc, "<br>", " -- ") & vbCrLf & vbCrLf _
					& "Med venlig hilsen"& vbCrLf & afsNavn & ", "& afsEmail & vbCrLf 
					'& "Incident ansvarlig: " & ansvNavn & ", " & ansvEmail & vbCrLf & vbCrLf _ bcc
					'& "Adressen til TimeOut er: https://outzource.dk/"&lto&"" & vbCrLf & vbCrLf _ 
				
					If Mailer.SendMail Then
    				
					Else
    					Response.Write "Fejl...<br>" & Mailer.Response
  					End if	
				
			
			end if ''** Mail
			
			
			
			'*** Skal step 3 springes over?? **'
			if len(request("FM_tilknytjob")) <> 0 then
	        tilknytjob = 1
	        else
	        tilknytjob = 0
	        end if
	
	        response.cookies("sdsk")("tilknytjob") = tilknytjob 
	        response.cookies("sdsk").expires = date + 60
			
			
			if kview = "j" then
			Response.redirect "sdsk.asp?lastedit="&lastedit&"&usekview=j&FM_kontakt="&intKontakt
			else
			Response.redirect "sdsk.asp?lastedit="&lastedit&"&func=dbopr2&FM_kontakt="&intKontakt&"&bcc="&bcc&"&FM_tilknytjob="&tilknytjob
			end if
		
		
		else
		
		
		
		
		
			
			if forny = 1 then
			dagogtidSQL = ", dato = '"& datoSQL &"', tidspunkt = '"& time & "', rsptid_a_id = 0, rsptid_a_tid = 0, rsptid_b_id = 0, rsptid_b_tid = 0, dato2 = '"& datoSQL &"'" 
			else
			dagogtidSQL = ", dato2 = '"& datoSQL &"'"
			end if
			
			if forny <> 1 then
			if request("FM_opr_status") <> intStatus then
				
				'*** Gem rsp tid A ***
				strSQL = "SELECT id, rsptid_a FROM sdsk_status WHERE id = "& intStatus
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
							
							if oRec("rsptid_a") = 1 then
							
							'*** Timer til Rsp **
							strSQL2 = "SELECT dato, tidspunkt, rsptid_a_id FROM sdsk WHERE id = "& id & " AND rsptid_a_id = 0"
							oRec2.open strSQL2, oConn, 3
							if not oRec2.EOF then
							 	
								
							 	oprettet = oRec2("dato") &" "& formatdatetime(oRec2("tidspunkt"), 3)
								nu = dato &" "& tid
								responsTid = datediff("h", oprettet, nu)
								
									strSQL3 = "UPDATE sdsk SET rsptid_a_id = " & intStatus & ", rsptid_a_tid = "& responsTid &" WHERE id =" & id
									oConn.execute(strSQL3) 
							 
							end if
							oRec2.close
							
							
							end if
							
				end if
				oRec.close
				
				'*** Gem rsp tid B ***
				strSQL = "SELECT id, rsptid_b FROM sdsk_status WHERE id = "& intStatus
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
							
							if oRec("rsptid_b") = 1 then
							
							'*** Timer til Rsp **
							strSQL2 = "SELECT dato, tidspunkt, rsptid_b_id FROM sdsk WHERE id = "& id & " AND rsptid_b_id = 0"
							oRec2.open strSQL2, oConn, 3
							if not oRec2.EOF then
							 
							 	oprettet = oRec2("dato") &" "& formatdatetime(oRec2("tidspunkt"), 3)
								nu = dato &" "& tid
								responsTid = datediff("h", oprettet, nu)
									
									strSQL3 = "UPDATE sdsk SET rsptid_b_id = " & intStatus & ", rsptid_b_tid = "& responsTid &" WHERE id =" & id
									oConn.execute(strSQL3) 
							 
							end if
							oRec2.close
							
							end if
				
				end if
				oRec.close
				
				
				
			end if
			end if
			
			
			'************************************************************************************
			'*** Adviser projektgruppe ***
			'************************************************************************************
			if request("FM_opr_status") <> intStatus then
				
				
				'**Henter afsender ***
				if kview <> "j" then
				
				strSQL = "SELECT mnavn, mnr, init, email FROM medarbejdere WHERE mid = " & session("mid")
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
				
				afsNavn = oRec("mnavn") & "("& oRec("mnr") &") - " & oRec("init")
				afsEmail = oRec("email")
				
				end if
				
				oRec.close
				else
				
				afsNavn = session("user")
				afsEmail = session("kontaktpersEmail")
				
				end if
				
				
				'** Henter navn på ny status ***
				strSQL = "SELECT navn FROM sdsk_status WHERE id = " & intStatus
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				statusNavn = oRec("navn") 
				
				oRec.movenext
				wend
				
				oRec.close
				
				
			
			
			'***** Oprettter Mail object ***
			if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\sdsk.asp" then
					
			
			
			Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			' Sætter Charsettet til ISO-8859-1 
			Mailer.CharSet = 2
			' Afsenderens navn 
			Mailer.FromName = "TimeOut ServiceDesk"
			' Afsenderens e-mail 
			Mailer.FromAddress = "timeout_no_reply@outzource.dk"
			Mailer.RemoteHost = "webmail.abusiness.dk"
			Mailer.ContentType = "text/html"
			
			' Mailens emne
			Mailer.Subject = "Incident (ID: "& id &") - "& strEmne &" - "& statusNavn
			' Modtagerens navn og e-mail
			Mailer.AddRecipient afsNavn, afsEmail
			
			        bcc = ""
			        
					call eml_inciAns()
					call eml_inciCreator()
					call eml_proGrp()
					call eml_Kpers(kpers)
					
					if instr(kpers2email, "@") <> 0 then
		            Mailer.AddBCC ""&kpers2&"",""& kpers2email  
		            bcc = bcc & "Kontaktperson: "&kpers2&", "& kpers2email & "<br>"
	                end if
					
					call eml_kogjobAns()
					
					
					' Selve teksten
					Mailer.BodyText = "Vedr. Incident: (ID: "& id &") "& vbCrLf & vbCrLf _
					& "==============================================================" & vbCrLf _
					& "Kontakt: " & vbCrLf _
					& knavn & " ("& knr &") " & vbCrLf _
					& adresse & vbCrLf _
					& postnr & ", " & by  & vbCrLf _
					& "Tlf:" & telefon  & vbCrLf _
					& "==============================================================" & vbCrLf & vbCrLf & vbCrLf  _
					& "Emne og beskrivelse: " & vbCrLf & strEmne &"."& vbCrLf _
					& strBesk & vbCrLf & vbCrLf & vbCrLf _
					& "Denne Incident har skiftet status til: "& statusNavn &""  & vbCrLf & vbCrLf _ 
					& "Modtagere af denne mail:" & vbCrLf & replace(bcc, "<br>", " -- ") & vbCrLf & vbCrLf _
					& "Med venlig hilsen" & vbCrLf & afsNavn & ", "& afsEmail  &  vbCrLf 
				
					If Mailer.SendMail Then
    				
					Else
    					Response.Write "Fejl...<br>" & Mailer.Response
  					End if
					
			
			end if
			end if
			'*** Adviser via email ***********
			
			
			
			
			
			'**** Opdater Incident ***********
			strSQL = "UPDATE sdsk SET emne = '"& strEmne &"', besk = '"& strBesk &"', type = "& intType &","_
			&" prioitet = "& intPrio &", priotype = "& intPriotype &", kundeid = "& intKontakt &", jobid = "& jobid &", ansvarlig = "& intAns &", esttid = "& dblesttid &","_
			&" editor2 = '"& editor &"', public = "& pblic &", status = "& intStatus &" "& dagogtidSQL &", "_
			&" sogeord1 = '"& strSogeord_1 &"', sogeord2 = '"& strSogeord_2 &"', sogeord3 = '"& strSogeord_3 &"', sogeord4 = '"& strSogeord_4 &"',"_
			&" creator = "& creator &", kpers = "& kpers &", kpers2 = '"& kpers2 &"', "_
			&" kpers2email = '"& kpers2email &"', duedate = '"& duedate &"', useduedate = "& useduedate &" WHERE id = "&id
			
			'Response.Write strSQL
			oConn.execute(strSQL)
			
			lastedit = id
		
			
			'*** Luk aktiviteter ***
			intLuk = 0
			strSQL = "SELECT luk FROM sdsk_status WHERE id = "& intStatus
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			intLuk = oRec("luk")
			end if
			oRec.close
			
			if cint(intLuk) = 1 then 
			strSQL = " UPDATE aktiviteter SET aktstatus = 0 WHERE incidentid = "& id
			oConn.execute(strSQL)
			end if
			'************************
			
		if kview = "j" then
		Response.redirect "sdsk.asp?lastedit="&lastedit&"&usekview=j&FM_kontakt="&intKontakt
		else
		Response.redirect "sdsk.asp?lastedit="&lastedit&"&bcc="&bcc
		'Response.write bcc
		end if
		
		
		
		
		end if
					
		
		
	
	end if
	
	
	
	if func = "dbopr3" then
			
			
			'*** Tilknytter Incident til job og opretter aktivitet på det valgte job ***
			if len(request("FM_jobid")) <> 0 then
			jobid = request("FM_jobid")
			else
			jobid = 0
			end if	
			
			if len(request("FM_fomr")) <> 0 then
			fomr = request("FM_fomr")
			else
			fomr = 0
			end if
			
			
			
			strSQL = "UPDATE sdsk SET jobid = "& jobid &" WHERE id = "&id
			oConn.execute(strSQL)
			
			lastedit = id
			
			response.Cookies("sdsk")("opretakt") = request("FM_opretakt")
			response.cookies("sdsk").expires = date + 10
			
			if request("FM_opretakt") = "1" then
			
				'** Opretter aktivitet på det valgte job ***'
				'*** Henter nødvendig data. ***'
				'** Projektgrupper, start og slut dato mm. nedarves fra job ***'
				'** Faktor hentes fra SDSK **'
				
				strSQL = "SELECT s.emne, s.besk, s.prioitet, "_
				&" s.esttid, s.editor, sp.faktor "_
				&" FROM sdsk s "_
				&" LEFT JOIN sdsk_prioitet sp ON (sp.id = s.prioitet) WHERE s.id = "&id
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
				
				strEmne = SQLBless(oRec("emne"))
				strDatoSQL = session("dato") 'year(now)&"/"& month(now) &"/"& day(now)
				strEditor = oRec("editor")
				
				if len(trim(oRec("esttid"))) <> 0 then
				dblEsttid = SQLBless2(oRec("esttid"))
				else
				dblEsttid = 0
				end if
				
				if len(trim(oRec("faktor"))) <> 0 then
				cdblFaktor = SQLBless2(oRec("faktor"))
				else
				cdblFaktor = 0
				end if 
				
				if len(trim(oRec("besk"))) <> 0 then
				strBesk = SQLBless(oRec("besk"))
				else
				strBesk = ""
				end if
				
				
				end if
				oRec.close
							
				
				projektgruppe1 = 0
				projektgruppe2 = 0 
				projektgruppe3 = 0 
				projektgruppe4 = 0 
				projektgruppe5 = 0 
				projektgruppe6 = 0 
				projektgruppe7 = 0
				projektgruppe8 = 0
				projektgruppe9 = 0
				projektgruppe10 = 0 
				jobstartdato = strDatoSQL
				jobslutdato = strDatoSQL				
								
				strSQLpgrp = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, "_
				&" projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, jobstartdato, jobslutdato FROM job WHERE id = "& jobid 
				
				'Response.write strSQLpgrp & "<br><br>"
				
				oRec2.open strSQLpgrp, oConn, 3
					if not oRec2.EOF then
					
					projektgruppe1 = oRec2("projektgruppe1") 
					projektgruppe2 = oRec2("projektgruppe2") 
					projektgruppe3 = oRec2("projektgruppe3") 
					projektgruppe4 = oRec2("projektgruppe4") 
					projektgruppe5 = oRec2("projektgruppe5") 
					projektgruppe6 = oRec2("projektgruppe6") 
					projektgruppe7 = oRec2("projektgruppe7") 
					projektgruppe8 = oRec2("projektgruppe8") 
					projektgruppe9 = oRec2("projektgruppe9") 
					projektgruppe10 = oRec2("projektgruppe10") 
					jobstartdato = oRec2("jobstartdato")
					jobslutdato = oRec2("jobslutdato")
					
					
					end if
				oRec2.close
							
				
			'for t = 0 to 1
			
			'if t = 0 then
			fakbar = 1
			navn_ext = ""
			'else
			'fakbar = 0
			'navn_ext = "" '(ej fakturerbar)
			'end if
			
			'*** Opretter aktivitet ***
			strSQLinsAkt = "INSERT INTO aktiviteter "_
			&" (navn, beskrivelse, dato, editor, job, fakturerbar, "_
			&" projektgruppe1, projektgruppe2, projektgruppe3, "_
			&" projektgruppe4, projektgruppe5, projektgruppe6, "_
			&" projektgruppe7, projektgruppe8, projektgruppe9, "_
			&" projektgruppe10, aktstartdato, aktslutdato, "_
			&" budgettimer, faktor, aktstatus, fomr, incidentid) VALUES "_
			&"('"& strEmne &""& navn_ext &"', "_
			&"'"& strBesk &"', "_
			&"'"& strDatoSQL &"', "_ 
			&"'"& strEditor &"', "_
			&""& jobid &", "_ 
			&""& fakbar &", "_
			&""& projektgruppe1 &", "_ 
			&""& projektgruppe2 &", "_ 
			&""& projektgruppe3 &", "_ 
			&""& projektgruppe4 &", "_ 
			&""& projektgruppe5 &", "_
			&""& projektgruppe6 &", "_ 
			&""& projektgruppe7 &", "_ 
			&""& projektgruppe8 &", "_ 
			&""& projektgruppe9 &", "_ 
			&""& projektgruppe10 &", "_     
			&"'"& year(now) &"/"& month(now) &"/"& day(now) &"', "_ 
			&"'"& year(jobslutdato) &"/"& month(jobslutdato) &"/"& day(jobslutdato) &"', "_
			&""& dblEsttid & ", "& cdblFaktor &", 1, "& fomr &", "& id &")"
			
			
			'Response.write strSQLinsAkt
			'Response.end
			oConn.execute(strSQLinsAkt)
			
			'next
			
			end if
		
		    
		    '**** Til opdatering af stopur ****'
		    aktid = 0
		    strSQL = "SELECT id FROM aktiviteter WHERE id <> 0 ORDER BY id DESC "
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then
		    
		    aktid = oRec("id")
		    
		    end if
		    oRec.close
		    
		    strSQLstopur = "UPDATE stopur SET jobid = " & jobid &", aktid = " & aktid & " WHERE incident = " & id
		    oConn.execute(strSQLstopur)
		    '*********************************'
		
		Response.redirect "sdsk.asp?lastedit="&lastedit&"&bcc="&bcc
		
		
				
	
	end if
	
	
	
	
	
	
	
	'**************************************************
	'*** Hovemenu *************************************
	'**************************************************%>
	
	
	<%if kview <> "j" then '**** Kundelogin%>
	
	        
	        <%if print <> "j" then %>
	 
	        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	        <!--#include file="../inc/regular/topmenu_inc.asp"-->
	        <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	        <%call sdskmainmenu(1)%>
	        </div>
	        <div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	        <%
	        call sdsktopmenu()
	        %>
	        </div>
        	
	        <%
        	
	        leftVal = "20"
	        topVal = "150" 
	
	        else
	        %>
	        
	        <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	        <%
	        
	        leftVal = "20"
	        topVal = "20" 
	        
	        end if
	
	else
		
  
		
		'** Firma logo ***
		call visfirmalogo(990, 110, kontaktId)
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call kundelogin_mainmenu(5, lto, kontaktId)%>
	</div>
	<%
	
	leftVal = "20"
	topVal = "150" 
	%>
	
	<%end if
	
	 
	
	%>
	
	
	
	<div id="sideindhold" style="position:absolute; left:<%=leftVal%>; top:<%=topVal%>; visibility:visible;">
	<%select case func
	case "dbopr2"
	
	
	'**** Springer stetep 3 over ved incident oprettelse ***'
	if len(request("FM_tilknytjob")) <> 0 then
	tilknytjob = request("FM_tilknytjob")
	else
	tilknytjob = 0
	end if
	
	

	'**** Springer stetep 3 over ved incident oprettelse ***'
	if cint(tilknytjob) = 0 then
	Response.redirect "sdsk.asp?lastedit="&lastedit&"&FM_kontakt="&kontaktId&"&bcc="&bcc
	else		
	
	
	
	
	%>
	<h3>Opret Incident Step 3 - Jobtilknytning</h3>
		
	<%call tilfojtiljob()%>
		
		
	<br><br>
	<div style="position:relative; background-color:#ffffe1; visibility:visible; border:1px red dashed; border-right:2px red dashed; border-bottom:2px red dashed; padding:15px; width:550px;"><img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
	<b>Når en Incident oprettes som en aktivitet på et job, bliver den oprettet med følgende data:</b><br>
	- Navn sættes lig med navnet på Incidenten.<br>
	- Beskrivelse sættes lig med beskrivelsen på Incicenten.<br>
	- Datointerval: Startdato = dagsdato, slutdato = jobslutdato<br>
	- Faktor sættes lig med den faktor der er angivet under prioitet på Incidenten.<br>
	- Projektgrupper sættes lig med de grupper der er tilknyttet jobbet.<br>
	- Værdi sættes til 0.<br>
	- Status sættes altid til "Aktiv."<br>
	- Timepriser nedarves fra jobbet.<br>
	- Forkalk. timer sættes = Estimeret timeforbrug på Incident.
	</div>
	
	

	<%
	
	end if
	
	case "opr"
	kid = request("FM_kontakt")
	
	if len(trim(request("FM_sog"))) <> 0 then
	sogVal = replace(request("FM_sog"), "'", "")
	Response.Cookies("sdsk")("sogkun") = sogVal
	else
	    if trim(request.Cookies("sdsk")("sogkun")) <> "" then
	    sogVal = request.Cookies("sdsk")("sogkun")
	    else
	    sogVal = ""
	    end if
	end if
	%>
	<script type="text/javascript">
		$(document).ready(function() {
		$("#BtnCustDescUpdate").click(function() {
		$(this).hide();
		$("#LoadCustDescUpdate").show();
		$.post("?", { AjaxUpdateField: "true", id: $("#BtnCustDescUpdate").data("cust"), control: "FM_CustDesc", value: $("#custDesc").val() }, function() {
					TONotifie("Opdateringen i kundebeskrivelsen blev gemt", true);
					$("#BtnCustDescUpdate").show();
					$("#LoadCustDescUpdate").hide();
				});
			});
			$("#FM_kontakt").change(function() {
				$("#custDesc").val("henter text...");
				$("#BtnCustDescUpdate").data("cust", $(this).val());
				$.post("?", { control: "FN_getCustDesc", AjaxUpdateField: "true", cust: $(this).val() }, function(data) { $("#custDesc").val(data); });
			});
		});</script>
	<h3>Opret Incident Step 1 - Kontakt (kunde)</h3>
	<%call filterheader(0,0,600,pTxt) %>
	<table width=100% cellspacing=2 cellpadding=2 border=0>
	<form method="post" action="sdsk.asp?func=opr&FM_kontakt=0">
	<tr>
	<td align=right valign=top style="padding-top:6px;"><b>Søg:</b></td><td valign=top>
        <input id="FM_sog" name="FM_sog" type="text" value="<%=sogVal %>" size="40" /> <br />(% = wildcard)<br />
        søg på domæner, navn, id, tlf.</td>
        <td>
            <input id="Submit1" type="submit" value="Søg >>" /></td>
	</tr>
	</form>
	</table>
    <!-- filter div -->
	</td></tr></table>
	</div><br /><br />
	
	<h4>Vælg kontakt</h4>
	
	<table width=575 cellspacing=2 cellpadding=2 border=0>
	<form method=post action="sdsk.asp?func=opr1&id=<%=id%>">
	<tr><td valign=top><img src="../ill/ikon_kunder_16.png" />&nbsp;&nbsp;<font color=red size=2>*</font>&nbsp;<b>Kontakt:</b>
	&nbsp;&nbsp;
		(<a href="kunder.asp?func=opret&ketype=k&hidemenu=1&showdiv=onthefly&rdir=1" target="_blank" class=vmenu>Opret ny her..</a>)<br />
	<select name="FM_kontakt" id="FM_kontakt" style="width:478px;">
		
		<%
		if sogVal = "" then
		%>
		<option value="0">Vælg kontakt..</option>
		<%
		end if
		
		
		ketypeKri = " ketype <> 'e'"
		
		if sogVal <> "" then
		kundeKri = ""& ketypeKri &" AND sdskpriogrp <> 0 AND (kkundenr LIKE '"& sogVal &"' OR kkundenavn LIKE '"& sogVal &"%' OR url LIKE '"& sogVal &"' OR telefon LIKE '"& sogVal &"')"
		else
		kundeKri = ""& ketypeKri &" AND sdskpriogrp <> 0"
		end if
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& kundeKri &" ORDER BY Kkundenavn"
				'Response.write strSQL
				'Response.flush
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr")%>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		<option value="0">Ingen</option>
		</select>
		
		
		</td>
		<td align="left" valign=top style="padding-top:20px;"><input type="submit" value="Videre >> " ></td>
		</tr>
		<tr>
		<td><br />Beskrivelse:<br />
		<textarea cols="70" rows="8" style="width:478px;" id="custDesc"></textarea><br />
		<input type="button" id="BtnCustDescUpdate" value="^ Opdater Beskriv. ^"/><div id="LoadCustDescUpdate" style="display:none"><img src="../inc/jquery/images/ajax-loader.gif" alt="opdaterer beskrivelse" />pdaterer beskrivelse...</div>
		</td>
		</tr>
	</table>
	
	<div style="position:absolute; background-color:#ffffe1; top:35px; left:720px; visibility:visible; border:1px red dashed; padding:10px; width:350px;"><img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
	<b>Der skal angives en kontakt!</b><br>
	 - Hvis der ikke er nogen kontakter tilgængelige på listen,<br>
	er det fordi der ikke er angivet ServiceDesk prioitets gruppe på nogen kontakter.
	<br><br>Angiv ServiceDesk prioitets gruppe på de kontakter du ønsker at oprette incidents på.
	</div>
	<br /><br /><br /><br />
&nbsp;
	
	
	<%		
	case "red","opr1"
	
	
	
	
	
	
	if func = "red" then
		
		if id <> 0 then
		id = id
		else
		id = request.cookies("sdsk")("idredfromjob")
		end if
		
		response.cookies("sdsk")("idredfromjob") = id
		
		
		strSQL = "SELECT emne, besk, priotype, type, prioitet, kundeid, jobid, ansvarlig, "_
		&" esttid, dato, editor, tidspunkt, public, status, sogeord1, "_
		&" sogeord2, sogeord3, sogeord4, creator, kpers, kpers2, kpers2email, duedate, useduedate "_
		&" FROM sdsk WHERE id = "&id
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		
		kid = oRec("kundeid")
		strEmne = oRec("emne")
		strBesk = oRec("besk")
		intType = oRec("type")
		intPrio = oRec("prioitet")
		intPriotype = oRec("priotype")
		intJobid = oRec("jobid")
		intAns = oRec("ansvarlig")
		cdblEsttid = oRec("esttid")
		editor = oRec("editor")
		dtTidspunkt = oRec("tidspunkt")
		intPublic = oRec("public") 
		intStatus = oRec("status")
		
		strSogeord_1 = oRec("sogeord1")
		strSogeord_2 = oRec("sogeord2")
		strSogeord_3 = oRec("sogeord3")
		strSogeord_4 = oRec("sogeord4")
		
		ejer = oRec("creator")
		
		kpers = oRec("kpers")
		kpers2 = oRec("kpers2")
		kpers2email = oRec("kpers2email")
		
		duedate = formatdatetime(oRec("duedate"), 2)
		duetime = left(formatdatetime(oRec("duedate"), 3), 5)
		useduedate = oRec("useduedate")
		
		end if
		oRec.close
	
		lastedit = id
		nFunc = "dbred"
		
		
	else
	
		
		
		kid = request("FM_kontakt")
		if len(trim(request("FM_emne"))) <> 0 then
		strEmne = request("FM_emne")
		else
		strEmne = ""
		end if
		strBesk = ""
		intType = 0
		intPrio = 0
		intPriotype = 0
		intJobid = 0
		intAns = 0
		cdblEsttid = 0
		dtTidspunkt = ""
		nFunc = "dbopr" 
		intPublic = 0
		intStatus = 0
		strSogeord_1 = ""
		strSogeord_2 = ""
		strSogeord_3 = ""
		strSogeord_4 = ""
		
		duedate = day(now) & "-" & month(now) & "-" & year(now)
		duetime = left(formatdatetime(now, 3), 5)
		useduedate = 0
		
		ejer = session("mid")
		
		kpers = 0
		kpers2 = ""
		kpers2email = ""
	
	end if
	
	if useduedate = 1 then
	brugduedateStatus = "CHECKED"
	duedDis = ""
	else
	brugduedateStatus = ""
	duedDis = "DISABLED"
	end if 
	%>
	
	<%
	if func = "red" then%>
	<div style="position:absolute; left:550px; top:0px; width:300px; height:200px; overflow:auto;">
	<h3>Filer/Screenshots:</h3>
	<table>
	<tr>
		<td><a href="#" onClick="popitup('upload.asp?jobid=0&sdsk=1&type=pic&iid=<%=id%>')" class=vmenu>Vedhæft / Upload fil eller screenshot&nbsp;&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a></td>
	</tr>
	<%
	
	
	strSQL = "SELECT id, filnavn FROM filer WHERE incidentid = " & id  'AND incidentid <> 0 AND kundeid = " & kid & " AND kundeid <> 0"
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	%>
	<tr><td><img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;<a href="../inc/upload/<%=lto%>/<%=oRec("filnavn")%>" class='vmenulille' target="_blank"><%=oRec("filnavn")%></a>&nbsp;</td></tr>
	<%
	oRec.movenext
	wend
	
	oRec.close
	%>
	
	</table>
	</div>
	<div id="jobinfo" style="position:absolute; left:820px; top:0px; width:265; visibility:visible; padding:10px; border:1px red dashed; background-color:#ffffe1;">
			<table>
			<tr><td>
			<img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;&nbsp;<b>Upload - Sådan gør du:</b><br>
			
			1) Brug [Print Scr].<br>
			2) Åbn WordPad. (under Start - Tilbehør)<br>
			3) Paste [CRTL] + [V] billedet ind i WordPad.<br>
			4) Gem filen og upload den herefter til Timeout.
			</td></tr>
			</table>
		</div>
	<%end if%>
	
	
	
	<%if func = "red" then
	sVal = "Opdater"
	else
	sVal = "Opret"
	end if%>
	
	
	
	
	
	<%
	
	if kview = "j" then
	tTop = 100
	else
	tTop = 0
	end if
	
	tLeft = 0
	tWdth = 1200
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	<br />
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr><td style="padding: 20px 40px;" valign=top>
	<h3><%=sVal%> Incident</h3>
	<form method=post action="sdsk.asp?func=<%=nFunc%>&id=<%=id%>">
	<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=intJobid%>">
	<input type="hidden" name="usekview" id="Hidden1" value="<%=kview%>">
	<input type="hidden" name="FM_opr_status" id="FM_opr_status" value="<%=intStatus%>">
	<input type="submit" style="float:right; margin: 5px 0 5px 5px;" value="<%=sVal%> Incident >> ">
	<br class="clear" />
	<br />
	<div class="maintablesquare" style="overflow:visible;">
	<font color=red size=2>*</font>&nbsp;<b>Emne:</b>
		<br /><textarea style="width:500px; height:40px;" name="FM_emne" id="FM_emne"><%=strEmne%></textarea>
		<br><br /><b>Beskrivelse</b>
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
		                	<%if kview <> "j" then%>
		                	<table>
	<tr>
		<td style="padding-top:10px;" valign=top colspan=2><br><b>Søgeord:</b>(til knowledgebase) 
		
		<br>
		<input type="text" name="FM_sogeord_1" id="FM_sogeord_1" value="<%=strSogeord_1%>" style="width:175px;">&nbsp;&nbsp;
	    <input type="text" name="FM_sogeord_2" id="FM_sogeord_2" value="<%=strSogeord_2%>" style="width:175px;"><br />
	    <input type="text" name="FM_sogeord_3" id="FM_sogeord_3" value="<%=strSogeord_3%>" style="width:175px;">&nbsp;&nbsp;
		<input type="text" name="FM_sogeord_4" id="FM_sogeord_4" value="<%=strSogeord_4%>" style="width:175px;"></td>
	</tr>
	</table>
	<%end if%>
	</div>
	<div class="maintablesquare" style="overflow:visible;">
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	
	<tr>
	<td>
        <img src="../ill/ikon_kunder_16.png" />&nbsp;&nbsp;<font color=red size=2>*</font>&nbsp;<b>Kontakt:</b></td><td>
	<%
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE Kid = "& kid 
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then 
				Response.write oRec("Kkundenavn") &" ("&oRec("Kkundenr")&")"
			end if
			oRec.close 
	%>
		<input type="hidden" name="FM_kontakt" id="FM_kontakt" value="<%=kid%>">
		</td>
	</tr>
	<tr>
		<td valign=top><br /><b>Tilknyt kontaktperson:</b></td><td valign=top><br />
            <select name="FM_kpers" id="FM_kpers" style="width:272px; font-size:10px;">
            <option value="0">Ingen - eller vælg..</option>
            <%
			strSQL = "SELECT id, navn, email FROM kontaktpers WHERE kundeid = "& kid 
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			
			if cint(kpers) = oRec("id") then
			kpersSEL = "SELECTED"
			else
			kpersSEL = ""
			end if%>
			
		   <option value="<%=oRec("id") %>" <%=kpersSEL %>><%=oRec("navn") %>, <%=oRec("email") %></option>
			<%	
			oRec.movenext
			wend 
			oRec.close 
	%>
            </select>
            </td>
            </tr>
     <tr><td valign=top style="padding-top:10px;">       
            Angiv evt. en kontaktperson her:</td>
            <td valign=top style="padding-top:10px;">Navn: <input name="FM_kpers2" id="FM_kpers2" type="text" value="<%=kpers2 %>" style="width:240px; font-size:10px;" /><br />
            Email: <input name="FM_kpers2email" id="FM_kpers2email" type="text" value="<%=kpers2email %>" style="width:240px; font-size:10px;" />
	</td>
	</tr>
	<tr>
		<td valign=top><br><font color=red size=2>*</font>&nbsp;<b>Status:</b></td><td><br>
		<select name="FM_status" id="FM_status" style="width:278px;">
		<%
		if kview = "j" then
		lmt = " LIMIT 0, 1"
		else
		lmt = ""
		end if
		
		strSQL = "SELECT id, navn FROM sdsk_status ORDER BY navn" & lmt & ""
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		if cint(intStatus) = oRec("id") then
		stsel = "SELECTED"
		else
		stsel = ""
		end if 
		%>
		<option value="<%=oRec("id")%>" <%=stsel%>><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		<option value="0">Ingen</option>
		</select><br>&nbsp;
		
		</td>
	</tr>
	<%if kview <> "j" then 
	    
	    select case lto 
	    case "kits", "kringit"
	    advSel1 = "CHECKED"
	    advSel2 = "CHECKED"
	    advSel3 = ""
	    case else
	    advSel1 = ""
	    advSel2 = ""
	    advSel3 = ""
	    end select
	
	%>
	<tr>
	    <td colspan=2>
		<br>
		<br><b>Email adviser følgende personer ved oprettelse og status-skift.</b><br> (Incident- ansvarlig og -ejer bliver altid adviseret.) <br>
		<input type="checkbox" name="FM_sendtilprogrp" id="FM_sendtilprogrp" value="1" <%=advSel1%>> Adviser projektgruppe. (tilknyttet i status)<br>
		<input type="checkbox" name="FM_sendtilkogjans" id="FM_sendtilkogjans" value="1" <%=advSel2%>> Adviser 
		<%if func = "red" then%>
		jobansvarlige og 
		<%end if%>
		 primær kontaktansvarlig.<br>
		<input type="checkbox" name="FM_sendtilkontakter" id="FM_sendtilkontakter" value="1" <%=advSel3%>> Adviser tilknyttede kontaktpersoner. (Se øverst)<br>
		<br>&nbsp;</td>
	</tr>
	<%else %>
        <input name="FM_sendtilprogrp" id="FM_sendtilprogrp" value="1" type="hidden" />
         <input name="FM_sendtilkogjans" id="FM_sendtilkogjans" value="1" type="hidden" />
          <input name="FM_sendtilkontakter" id="FM_sendtilkontakter" value="1" type="hidden" />
	
	<%end if%>
	<tr>
		<td><font color=red size=2>*</font>&nbsp;<b>Kategori:</b></td><td>
		<select name="FM_type" id="FM_type" style="width:278px;">
		<%
		strSQL = "SELECT id, navn FROM sdsk_typer ORDER BY navn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		if cint(intType) = oRec("id") then
		tsel = "SELECTED"
		else
		tsel = ""
		end if 
		%>
		<option value="<%=oRec("id")%>" <%=tsel%>><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		<option value="0">Ingen</option>
		</select>
		</td>
	</tr>
	<tr><td><font color=red size=2>*</font>&nbsp;<b>Type:</b></td><td>
		<%
		if print <> "j" then
		strSQL = "SELECT p.id, p.navn, p.responstid, pt.navn AS typeNavn FROM kunder k "_
		&" LEFT JOIN sdsk_prio_grp g ON (g.id = k.sdskpriogrp)"_
		&" LEFT JOIN sdsk_prioitet p ON (priogrp = g.id) "_
		&" LEFT JOIN sdsk_prio_typ pt ON (pt.id = p.type) "_
		&" WHERE kid = "& kid &" ORDER BY p.navn"
		%>
		
		<select name="FM_prio" id="FM_prio" style="width:278px;">
		<%
		
		
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		if cint(intPrio) = oRec("id") then
		psel = "SELECTED"
		pPpSel = oRec("navn")
		else
		psel = ""
		end if 
		%>
		<option value="<%=oRec("id")%>" <%=psel%>><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		<option value="0">Ingen</option>
		</select>
		<% else
		if pPpSel = "" then
		pPpSel = Request.Form("FM_prio")
		end if
		Response.write pPpSel
		end if %>
		</td>
	</tr>
	<tr>
	<td>
	<b>Prioritet:</b>	</td>
	<td>
	<select name="FM_prio_typ" class="txtField">
	<option value="0">Ingen</option>
	<%
	strSQL = "SELECT id, navn FROM sdsk_prio_typ"
	oRec2.open strSQL, oConn, 3
	while not oRec2.EOF
		if cint(intPriotype) = oRec2("id") then
		tpSEL = "selected"
		else
		tpSel = "" 
		end if%>
		
	<option value="<%=oRec2("id")%>" <%=tpSel%>><%=oRec2("navn")%></option>
	<%
	oRec2.movenext
	wend
	
	oRec2.close
	%>
	</select>
	</td>
	</tr>
	<%if kview = "j" then%>
	<input type="hidden" name="FM_esttid" id="FM_esttidhidden" value="0">
	<%else
	
	StrTdato_alt = duedate
	%>
	<!--#include file="inc/dato2_alt.asp"-->
	<tr>
		<td style="padding-top:15px;" valign=top><b>Angiv dato for udførsel og<br />
		 estimeret tidsforbrug:</b>&nbsp;&nbsp;</td>
		 <td style="padding-top:15px;">
             <input name="FM_useduedate" id="FM_useduedate" value="1" <%=brugduedateStatus%> type="checkbox" onclick="visduedate()" /> ja 
             
             
             <select name="FM_start_dag_due" id="FM_useduedate_dag"  <%=duedDis %>>
				<option value="<%=strDag_alt%>"><%=strDag_alt%></option> 
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
				
				<select name="FM_start_mrd_due" id="FM_useduedate_md" <%=duedDis %>>
				<option value="<%=strMrd_alt%>"><%=strMrd_navn_alt%></option>
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
				
				
				<select name="FM_start_aar_due" id="FM_useduedate_aar" <%=duedDis %>>
				<option value="<%=strAar_alt%>">
				<%if id <> 0 OR nedarvdato = "j" then%>
				20<%=strAar_alt%>
				<%else%>
				<%=strAar_alt%>
				<%end if%></option>
				
			   	<%for x = -5 to 10
                useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
                  <option value="<%=right(useY, 2)%>"><%=useY%></option>
                <%next %>
				</select>
				&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=6')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
			    
			    &nbsp;kl. <input type="text" value="<%=duetime%>" name="FM_duetime" id="FM_duetime" style="width:50px;" <%=duedDis %>>  mm:ss </td>
	</tr>
	<tr>
	    <td>Estimeret tid:</td>
	    <td><input type="text"  value="<%=cdblEsttid %>" name="FM_esttid" id="FM_esttid" style="width:60px;"> timer
	    	<script type="text/javascript">
	    		$(document).ready(function() {
	    			$(":input[name=FM_esttid]").change(function() {
	    				$(":checkbox[name=FM_useduedate]").attr("checked", true);
	    			});
	    		});
	</script>
	    </td><!-- <=duedDis %> -->
	    
	</tr>
	<%end if%>
	
	<%if kview = "j" then%>
	<input type="hidden" name="FM_ans" id="FM_ans" value="0">
	<%else%>
	<tr>
		<td style="padding-top:15px;"><b>Incident Ansvarlig:</b> <br>
		Primær kontakansvarlig forvalgt.</td><td style="padding-top:15px;"><select name="FM_ans" id="FM_ans" style="width:278px; background-color:lightpink;">
		<option value="0">Ingen</option>
		<%
		strSQL = "SELECT mid, mnavn, mnr, init, k.kundeans1 FROM medarbejdere "_
		&" LEFT JOIN kunder k ON (k.kid = "& kid &") WHERE mansat <> 2 GROUP BY mid ORDER BY mnavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
		if cint(intAns) = cint(oRec("mid")) OR (func = "opr1" AND oRec("mid") = oRec("kundeans1")) then
		aSel = "SELECTED"
		else
		aSel = ""
		end if
		%>
		
		<option value="<%=oRec("mid")%>" <%=aSel%>><%=oRec("mnavn")%> (<%=oRec("mnr")%>)
		<%if len(oRec("init")) <> 0 then%>
		<%=" - "& oRec("init")%>
		<%end if%>
		</option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		</select></td>
	</tr>
	
	<tr>
		<td style="padding-top:5px;"><b>Incident Ejer:</b><br />
		(Creator)</td><td style="padding-top:5px;"><select name="FM_creator" id="Select2" style="width:278px; background-color:#ffffe1;">
		<option value="0">Ingen</option>
		<%
		strSQL = "SELECT mid, mnavn, mnr, init FROM medarbejdere "_
		&" WHERE mansat <> 2 GROUP BY mid ORDER BY mnavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
		if cint(ejer) = cint(oRec("mid")) then
		aSel = "SELECTED"
		else
		aSel = ""
		end if
		%>
		
		<option value="<%=oRec("mid")%>" <%=aSel%>><%=oRec("mnavn")%> (<%=oRec("mnr")%>)
		<%if len(oRec("init")) <> 0 then%>
		<%=" - "& oRec("init")%>
		<%end if%>
		</option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		</select></td>
	</tr>
	
	
	<%end if%>
	
	
	<%if func = "red" then%>
	<tr>
		<td colspan=2 style="padding-top:5px;"><input type="checkbox" name="FM_forny" id="FM_forny" value="1"><b>Opdater/forny</b><br>
		Opdater/forny responstidspunkt og oprettet af information.</td>
	</tr>
	<%end if%>
	
		<%
		if kview = "j" then
		%>
		<input type="hidden" name="FM_public" id="FM_public" value="1">
		<%
		else
		%>
		<tr><%
		if intPublic = 1 OR func = "opr1" then                      
		pubChecked = "CHECKED"
		else
		pubChecked = ""
		end if
		
		%>
		<td colspan=2 style="padding-top:5px;"><input type="checkbox" name="FM_public" id="FM_public" value="1" <%=pubChecked%>><b>Offentlig. </b><br>
		Incident tilgængelig for ekstern tilknyttet kontakt.</td>
	</tr>
	<%end if%>
	<!--<tr>
		<td colspan=2><input type="checkbox" name="FM_losning" id="FM_losning" value="1">Godkendt som løsning.</td>
	</tr>-->
	

	</tr>
	
	<%if func <> "red" AND kview <> "j" then%> 
	<tr>
	<%
	        if request.cookies("sdsk")("tilknytjob") = "1" then
	        tkjCHK = "CHECKED"
	        else
	        tkjCHK = ""
	        end if  %>
		<td colspan=2 style="padding-top:5px;">
            <input id="FM_tilknytjob" name="FM_tilknytjob" value="1" type="checkbox" <%=tkjCHK %> /> Tilknyt incident til job og opret aktivitet. (Step 3)</td>
	
	</tr>
	<%end if %>
	<tr>
		<td colspan=2 align=right style="padding-right:50px;"><br></td>
	
	</tr>
	
	</table>
	</div>
	<hr class="clear" style="margin: 5px 0;" />
		<div class="maintablesquare">
		<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
	<td colspan=2 valign=bottom>
	
	<h4>Incident log:</h4>
	<b><%=strEmne%></b> - Incident Id: <%=id%><br />
	</td>
	 <td>
        &nbsp;</td>
	</tr>
	<tr><td colspan=2 height=20>
        &nbsp;</td></tr>
	
	
	<%
	strSQL2 = "SELECT besk, sdskdato, sdsktidspunkt, id, editor, public, "_
	&" losning, editor2, dato2 FROM sdsk_rel WHERE sdsk_rel = "& id &" ORDER BY id DESC"
	
	'Response.Write strSQL2
	'Response.Flush
	
	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF
	
	%>
	<tr><td style="border-top:1px #C4C4C4 dashed; padding:10px 10px 10px 10px;">
	
    <%if oRec2("losning") = 1 then%>
	<img src="../ill/ac0060-16.gif" width="16" height="16" alt="Godkendt som løsning" border="0">
	<%end if%>
	
	<%if oRec2("public") = 1 then%>
	<img src="../ill/ac0058-16.gif" width="16" height="16" alt="Offentlig tilgængelig" border="0">
	<%end if%>
	
	<%if kview <> "j" then%>
	<b><%=left(weekdayname(weekday(oRec2("sdskdato"))), 2)%>. <%=oRec2("sdskdato")%>&nbsp;<%=formatdatetime(oRec2("sdsktidspunkt"), 3)%></b> - Logentry Id: <%=oRec2("id") %> <br>
	<a href="#" onclick="Javascript:window.open('sdsk_tilfoj.asp?sdskrelid=<%=id%>&id=<%=oRec2("id")%>&func=red&lastedit=<%=id%>&usekview=<%=kview%>&FM_kontakt=<%=kontaktId%>&rdir=redsdsk', '', 'width=600,height=650,resizable=yes,scrollbars=yes')" class=vmenu>
	<%=oRec2("besk")%></a>
	<br>
	<%else%>
	<b><%=left(weekdayname(weekday(oRec2("sdskdato"))), 2)%>.&nbsp;<%=oRec2("sdskdato")%>&nbsp;<%=formatdatetime(oRec2("sdsktidspunkt"), 3)%></b> - Logentry Id: <%=oRec2("id") %> <br>
	<i><%=oRec2("besk")%></i>
	<%end if%>
	
	
	
	<br /><font class=megetlillesilver>
		<i>Oprettet af: <%=oRec2("editor")%>
		<%if len(oRec2("editor2")) <> 0 then%>
		<br>Sidst redigeret af: <%=oRec2("editor2")%> d. <%=oRec2("dato2")%>
		<%end if%>
		</i>
	
	</td>
	
	<td valign=top style="border-top:1px #C4C4C4 dashed; padding:10px 10px 10px 10px;">
	<%if kview <> "j" AND print <> "j" then%>
    <a href="javascript:popUp('stopur_2008.asp?func=ins&incid=<%=id%>&logentry=<%=oRec2("id")%>&jobid=<%=jobid%>&FM_mid=<%=session("mid")%>','1000','650','20','20')" class=vmenu><img src="../ill/stopur_st_stop.png" alt="Tilføj Stopur's Entry" border=0 /></a>
    <%end if %>
	&nbsp;
	<br />
	</td></tr>
	
	<%
	
	'if s > 3 then
    'Response.End
    'end if
    
    
    
    s = s + 1
	
	oRec2.movenext
	wend
	oRec2.close
	%>
	</table>
	</div>
	<div class="maintablesquare afstemning">

	<%if func = "red" then
		
		if intJobid = 0 then%>
		<%
		kontaktId = kid
		call tilfojtiljob()%>
		<%else%>
		
		<h3><img src="../ill/ac0063-16.gif" width="16" height="16" alt="" border="0">&nbsp;Jobtilknytning.</h3>
		
		<table cellspacing=0 cellpadding=0 border=0>
		<tr><td><br><b>Job:</b>&nbsp;&nbsp;&nbsp;</td>
		<td><br>
				<%
				strSQL = "SELECT jobnavn, jobnr, id, jobans1, jobans2, jobknr FROM job WHERE id = "& intJobid 
				oRec.open strSQL, oConn, 3
				
				usejobnr = 0 
				if not oRec.EOF then 
				
				usejobnr = oRec("jobnr")
				
					'*** Tjekker rettigehder eller om man er jobanssvarlig ***
					editok = 0
					if level = 1 then
					editok = 1
					else
							if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
							editok = 1
							end if
					end if
					
					if editok = 1 then
					Response.write "<a href=""jobs.asp?menu=job&func=red&id="&intJobid&"&rdir=sdsk2&FM_kunde="&oRec("jobknr")&""" class=vmenu>"&oRec("jobnavn") &"&nbsp;("&oRec("jobnr")&")</a>"
					else
					Response.write oRec("jobnavn") &"&nbsp;("&oRec("jobnr")&")"
					end if
					
				end if
				oRec.close %>
				
		</td></tr>
		</table>
		<%end if%>
	<%end if%>
		   
		   
		   
		   <%if intJobid <> 0 then%>
				<a href="javascript:popUp('timereg_akt_2006.asp?jobid=<%=intJobid%>&usemrn=<%=session("mid")%>&showakt=1&fromsdsk=1&sdskkomm=<%="\n == Incident Id "& intJobid &" == Opr. "& datoogklokkeslet  &"\n"& strEmne%>','950','620','50','20');" target="_self"; class=vmenu>Til Timeregistrering.. </a><br />
				
				<h3>Timeregistreringer:</h3><br />
			
		    
		    <!-- Eksisterende time-registreringer -->
                <table cellpadding=1 cellspacing=0 border=0 width=95%>
              
               <%'*** Joblog denne uge ***' 
               
               strSQL = "SELECT tid, taktivitetnavn, timer, tfaktim, tjobnavn, tjobnr, tdato, "_
               &" tknr, tknavn, tmnavn, tmnr, kkundenr, godkendtstatus, timerkom FROM timer "_
               &" LEFT JOIN kunder K on (kid = tknr) WHERE tmnr <> 0 AND "_
               &" tdato BETWEEN '2002-01-01' AND '"& year(now) &"/"& month(now) &"/"& day(now ) &"' AND tjobnr = "& usejobnr &" ORDER BY tdato DESC "
               
               'Response.Write strSQL
               'Response.flush
               at = 0
               timertot = 0
               lwedaynm = ""
               timerDag = 0
               
               oRec.open strSQL, oConn, 3
               while not oRec.EOF 
                
                select case right(at, 1)
                case 0,2,4,6,8
                bgcol = "#ffffff"
                case else
                bgcol = "#EFF3ff"
                end select
                
                call akttyper2009prop(oRec("tfaktim"))
                
                
                if lwedaynm <> weekdayname(weekday(oRec("Tdato"))) then
                 if at <> 0 then%>
                <tr>
                    <td colspan=6 align=right><b><%=formatnumber(timerDag, 2) %></b></td>
                </tr>
                <%
                timerDag = 0
                
                end if %>
                <tr bgcolor="#ffffe1">
                    <td colspan=6><b><%=weekdayname(weekday(oRec("Tdato"))) %></b></td>
                </tr>
                <%
                end if
                
                %>
               <tr bgcolor="<%=bgcol%>">
               <td class=lille style="border-bottom:1px #cccccc solid;"><%=left(weekdayname(weekday(oRec("Tdato"))), 3) %>. <%=day(oRec("Tdato")) &" "&left(monthname(month(oRec("Tdato"))), 3) &". "& right(year(oRec("Tdato")), 2)%></td>
               <td class=lille style="border-bottom:1px #cccccc solid;"><%=oRec("tmnavn") %> (<%=oRec("tmnr") %>)</td>
               <td class=lille style="border-bottom:1px #cccccc solid;"><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</td>
               <td class=lille style="border-bottom:1px #cccccc solid;"><%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
               <td class=lille style="border-bottom:1px #cccccc solid;"><%=oRec("taktivitetnavn") %>
                <%
	            call akttyper(oRec("tfaktim"), 1)
	            Response.Write "<font class=megetlilleblaa> ("& akttypenavn & ")</font>"
	            %>
               
               </td>
               <td class=lille align=right style="border-bottom:1px #cccccc solid;">
               <%=formatnumber(oRec("timer"), 2) %>
               </td>
               </tr>
               
               <%if len(trim(oRec("timerkom"))) <> 0 then %>
               <tr><td class=lille colspan=6 valign=top style="padding:3px;">
               <%if len(oRec("timerkom")) > 255 then %>
               <%=left(trim(oRec("timerkom")), 255) %>..
               <%else %>
               <%=trim(oRec("timerkom")) %>
               <%end if %>
               
               </td></tr>
               <%end if %>
               
               <%
               lwedaynm = weekdayname(weekday(oRec("Tdato")))
               if cint(aty_real) = 1 then
               timertot = timertot + oRec("timer")
               timerDag = timerDag + oRec("timer")
               end if
               at = at + 1
               oRec.movenext
               wend
               oRec.close
               
               
               if at <> 0 then%>
               
               <tr>
                    <td colspan=6 align=right><b><%=formatnumber(timerDag, 2) %></b></td>
                </tr>
               <tr>
                    <td align=right colspan=6>Total: <b><%=formatnumber(timertot, 2) %></b></td>
               </tr>
               <tr>
               <td colspan=6 align=right><br />Kun registreringer på aktivitets typer der tæller med <br />
                i det daglige forventede timeforbrug er med i totaler.</td>
               </tr>
               <%else %>
                <tr>
                    <td colspan=6><br />Ingen registreringer på dette job.</td>
                </tr>
               <%end if %>
               </table>	
			
			
			
			
			<%end if%>
			
			
			
	</div>
	<br><br>
	<input type="submit" style="float:right; margin: 35px 0 5px 5px;" value="<%=sVal%> Incident >> ">
	</form>
	
	</td>
	
	</tr>
	</table>
	</div>

	<br /><br /><br />
	<br /><br /><br />
	<br /><br /><br />&nbsp;
	
	
	
	<%case else%>
	<script>

		function showbesk(besk_id) {
			var entryDiv = $("#b_" + besk_id);
	(entryDiv.css("display") == "none") ? entryDiv.stop().slideDown(200) : entryDiv.stop().slideUp(200);
	/*if (lastOpen != 0) {
	document.getElementById("b_"+lastOpen).style.display = "none";
	document.getElementById("b_"+lastOpen).style.visibility = "hidden";
	}
	
	document.getElementById("b_"+besk_id).style.display = "";
	document.getElementById("b_"+besk_id).style.visibility = "visible";
	
	document.getElementById("FM_div_id_txtbox_"+besk_id).focus()
	//document.getElementById("FM_div_id_txtbox_"+besk_id).scrollintoview()

	
	document.getElementById("divopen").value = besk_id*/
	return false;
	}
	
	function lukemldiv(){
	document.getElementById("emldiv").style.display = "none";
	document.getElementById("emldiv").style.visibility = "hidden";
	}
	
	</script>
	
	<%
	
	
	
	public statusKri
	function findstatusKri()
	
	
			
			statusKri = " AND ( s.status = 0 "
				
				
				strSQL = "SELECT id FROM sdsk_status WHERE vispahovedliste = 1 ORDER BY navn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				statusKri = statusKri & " OR s.status = "& oRec("id")
				
				
				oRec.movenext
				wend
				
				oRec.close
			
			statusKri = statusKri &")"
			
	
	end function
	
	
	
	
	'* Kunde
	if len(request("FM_kontakt")) <> 0 then
		if request("FM_kontakt") <> 0 then
		selKid = request("FM_kontakt")
		kontaktKri = " s.kundeid = "& selKid
		else
		selKid = 0
		kontaktKri = " s.kundeid <> 0"
		end if
	else
	
	
			if len(request.cookies("sdsk")("kontaktidthis")) <> 0 then
			selKid = request.cookies("sdsk")("kontaktidthis")
			else
			selKid = 0
			end if
			
			if cint(selKid) = 0 then 
			kontaktKri = " s.kundeid <> 0"
			else
			kontaktKri = " s.kundeid = "& selKid
			end if
	end if
	
	response.cookies("sdsk")("kontaktidthis") = selKid
	
	
	'* Type
	if len(request("FM_type")) <> 0 then
		
		if request("FM_type") <> 0 then
		typeSel = request("FM_type")
		typeKri = " AND s.type = "& typeSel
		else
		typeSel = 0
		typeKri = ""
		end if
	
		response.cookies("sdsk")("type") = typeSel
	
	else
				
			if request.cookies("sdsk")("type") <> "" then
				typeSel = request.cookies("sdsk")("type")
				
				if cint(typeSel) <> 0 then
				typeKri = " AND s.type = "& typeSel
				else
				typeKri = ""
				end if
				
			else
			typeSel = 0
			typeKri = ""
			end if
				
	end if
	
	
	
	'* status
	if len(request("FM_status")) <> 0 then
	
		if request("FM_status") <> 0 then
		    if cint(request("FM_status")) <> -1 then
		    statusSel = request("FM_status")
		    statusKri = " AND s.status = "& statusSel
		    else
		    statusSel = -1
			call findstatusKri()
		    end if
		else '*** Alle 
		statusSel = 0
		statusKri = ""
		end if
	else
		
		if len(request.cookies("sdsk")("status")) <> 0 then
		statusSel = request.cookies("sdsk")("status")
			if cint(statusSel) > 0 then
				statusKri = " AND s.status = "& statusSel
			else
			    if cint(statusSel) = 0 then
			    statusKri = ""
			    else
			    call findstatusKri()
			    end if
			end if
			
		else
			
			statusSel = -1
			call findstatusKri()
			
		end if
	end if
	response.cookies("sdsk")("status") = statusSel
	
	
	'Response.Write statusSel & "<br>"
	'Response.Write statusKri
	if len(request("FM_priotype")) <> "-1" then
				priotypeSel = request("FM_priotype")
				if priotypeSel <> "-1" AND priotypeSel <> "" then
				typeKri = " AND s.prioitet = "& priotypeSel
				else
				typeKri = " AND s.prioitet <> -1"
				end if		
			response.cookies("sdsk")("priotype") = priotypeSel
	else
				priotypeSel = request.Cookies("sdsk")("priotype")
				if len(priotypeSel) <> "-1" AND priotypeSel <> "" then
				typeKri = " AND s.prioitet = "& priotypeSel
				else
				typeKri = " AND s.prioitet <> -1"
				end if
	end if
	'* Prioitet
		
				

	
			if len(request("FM_prioitet")) <> 0 then
				prioSel = request("FM_prioitet")
				if prioSel <> "-1" and prioSel <> "" then
				prioitetKri = " AND s.priotype = "& prioSel
				else
				prioitetKri = " AND s.priotype <> -1"
				end if
			response.cookies("sdsk")("prioritet") = prioSel
			else
				prioSel = request.cookies("sdsk")("prioritet")
				if prioSel <> "-1" AND prioSel <> "" then
				prioitetKri = " AND s.priotype = "& prioSel
				else
				prioitetKri = " AND s.priotype <> -1"
				end if
				
			end if
		


		
	
	
	
	
	
	'* rspKri
	if len(request("FM_rsptid")) <> 0 then
		rspKri = request("FM_rsptid")
		response.cookies("sdsk")("rsptid") = rspKri
	else
		if request.cookies("sdsk")("rsptid") <> 0 then
		rspKri = request.cookies("sdsk")("rsptid")
		else
		rspKri = 0
		end if 
	end if
	
	'* Ansavrlig "" (tom) = alle, -1 = Ikke angive dvs 0.
	if len(request("FM_ansv")) <> 0 then
		
		if request("FM_ansv") <> 0 then
			
			if request("FM_ansv") <> "-1" then '** Ikke angivet
			ansvKri = " AND s.ansvarlig = "& request("FM_ansv")
			ansSel = request("FM_ansv")
			else
			ansvKri = " AND s.ansvarlig = 0"
			ansSel = "-1"
			end if
			
		else
			ansSel = 0
			ansvKri = ""
		end if
		
		response.cookies("sdsk")("ansv") = ansSel
		
	else
			if request.cookies("sdsk")("ansv") <> "" then
			ansSel = request.cookies("sdsk")("ansv")
				if ansSel <> 0 then
				ansvKri = " AND s.ansvarlig = "& ansSel
				else
				ansvKri = ""
				end if
			else
			ansSel = session("mid") '0
			ansvKri = " AND s.ansvarlig = "& ansSel
			end if
	end if
	
	'* sogeKri på ID
	'if len(request("FM_emne")) <> 0 then
	'sogeKri = " AND s.emne LIKE '"& request("FM_emne") &"%'"
	'showSogeKri = request("FM_emne") 
	'else
	'showSogeKri = ""
	'sogeKri = ""
	'end if
	
	if len(Request("sog")) <> 0 then
	sogVal = Request("sog")
	sogeKri = " s.id = "& sogVal 
	else
	    if len(request.cookies("sdsk")("sog")) <> 0 AND len(trim(request("FM_kontakt"))) = 0 then
	    sogVal = request.cookies("sdsk")("sog")
	    sogeKri = " s.id = "& sogVal 
	    else
	    sogVal = ""
	    sogeKri = ""
	    end if
	end if
	response.cookies("sdsk")("sog") = sogVal
	
	
	'* Dato interval **'
	if request("usedatokri") = "j" then
	
	datoKriJa = "CHECKED"
	datoKriNej = ""
	datoKri = " AND s.dato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"'"
	datoKrival = "j"
	
	else
	    
	    
	    
	    if request.cookies("sdsk")("bdato") <> "" then
	    
	    'Response.Write "her:"& request("usedatokri") & " cookies :" &  request.cookies("sdsk")("bdato")
	    
	        if request.cookies("sdsk")("bdato") = "j" AND request("usedatokri") <> "n" then
	        datoKriJa = "CHECKED"
	        datoKriNej = ""
	        datoKri = " AND s.dato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"'"
	        datoKrival = "j"
	        else
	        datoKri = ""
	        datoKriJa = ""
	        datoKriNej = "CHECKED"
	        datoKrival = "n"
	        end if
	     else
	        datoKri = ""
	        datoKriJa = ""
	        datoKriNej = "CHECKED"
	        datoKrival = "n"
	     end if
	end if
	
	response.cookies("sdsk")("bdato") = datoKrival
	
	
	%>
	
	<%
	if showbcc = 1 then
	%>
	<div id="emldiv" name="emldiv" style="position:relative; left:0px; width:400px; top:10px; border:2px #5582d2 dashed; padding:10px; background-color:#FFFFe1;">
	<b>Email advisering er udsendt til:</b><br />
	<font size=1><%=bcc%> </font>
	<br />
	Email udsendes kun 1 gang til hver af de ovenstående email adresser, selvom de optræder flere gange på listen.
	<br /><br /><a href="#" onclick="lukemldiv()" class="red">Luk vindue</a>
	</div>
	<%end if
	%>
	
	
	
	
	
	
	<%
	bgcStyle = ""
	%>
	
	
	<%
	
	if kview <> "j" then
	call filterheader(20,0,800,pTxt)
	else
	call filterheader(-28,210,750,pTxt)
	end if
	%>
	
	<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<form action="sdsk.asp" method="post">
	<input type="hidden" name="usekview" id="usekview" value="<%=kview%>">
	<tr><td valign=top style="padding-top:10px;">
	
	
	<table cellspacing=0 cellpadding=4 border=0>
	<!--<tr>
		<td align=right style="padding-right:5px;"><b>Søg på emne ell. søgeord:</b></td><td><input type="text" name="FM_emne" id="FM_emne" value="<%=showSogeKri%>" style="font-family: arial; font-size: 10px; width:240px; background-color:#999999;" maxlength=0></td>
	</tr>-->
	<tr>
		<td align=right style="padding-right:5px;"><b>Vælg kontakt:</b></td><td>
		
		<%if print <> "j" then %>
		<select name="FM_kontakt" style="font-family: arial; font-size: 10px; width:240px;" onchange=renssog();>
		<%end if %>
			
			<%
			ketypeKri = " ketype <> 'e'"
			
			if kview = "j" then
			wh = " kid = "& selKid
			else
			wh = ""& ketypeKri &" AND sdskpriogrp <> 0 "
			    
			    if print <> "j" then%>
			    <option value="0">Alle</option>
			    <%
			    pKundeSel = "Alle"
			    end if
			    
			end if
			
			
			
					strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& wh &" ORDER BY Kkundenavn"
					'Response.write strSQL
					'Response.flush
					oRec.open strSQL, oConn, 3
					while not oRec.EOF
					
					if cint(selKid) = cint(oRec("Kid")) then
					isSelected = "SELECTED"
					pKundeSel = oRec("Kkundenavn") &"(" &oRec("Kkundenr")&")"
					else
					isSelected = ""
					end if
					if print <> "j" then %>
					<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
					<%
					end if
					
					
					oRec.movenext
					wend
					oRec.close
					
					if print <> "j" then %>
			</select>
			<%else %>
			<%=pKundeSel %>
			<%end if %>
		</td>
	</tr>
	<tr>
		<td align=right style="padding-right:5px;"><b>Kategori:</b></td><td>
		            
		            <%if print <> "j" then %>
		            <select name="FM_type" style="font-family: arial; font-size: 10px; width:240px;" onchange=renssog();>
		            <option value="0">Alle</option>         
		            <%
		            pTyp = "Alle"
		            end if
							strSQL = "SELECT id, navn FROM sdsk_typer WHERE id <> 0 ORDER BY navn"
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							if cint(typeSel) = cint(oRec("id")) then
							isSelected = "SELECTED"
							pTyp = oRec("navn")
							else
							isSelected = ""
							end if
							
							if print <> "j" then%>
							<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
							<%
							end if
							oRec.movenext
							wend
							
							oRec.close
							
							if print <> "j" then
							%>
							</select>
							<%else %>
							<%=pTyp %>
							<%end if %>
	</td></tr>
	<tr><td align=right style="padding-right:5px;"><b>Antal timer til respons:</b></td><td>
	<%if print <> "j" then %>
	<select name="FM_rsptid" style="font-family: arial; font-size: 10px; width:90px;" onchange=renssog();>
	<%end if %>
						
						<%select case rspKri
						case "-1"
						sel_0 = ""
						sel_4 = ""
						sel_12 = ""
						sel_24 = ""
						sel_48 = ""
						sel__1 = "SELECTED"
						pRsp = "Overskreddet"
						case 0
						sel_0 = "SELECTED"
						sel_4 = ""
						sel_12 = ""
						sel_24 = ""
						sel_48 = ""
						sel__1 = ""
						pRsp = "Alle "
						case 4
						sel_0 = ""
						sel_4 = "SELECTED"
						sel_12 = ""
						sel_24 = ""
						sel_48 = ""
						sel__1 = ""
						pRsp = "< 4 timer"
						case 12
						sel_0 = ""
						sel_4 = ""
						sel_12 = "SELECTED"
						sel_24 = ""
						sel_48 = ""
						sel__1 = ""
						pRsp = "< 12 timer"
						case 24
						sel_0 = ""
						sel_4 = ""
						sel_12 = ""
						sel_24 = "SELECTED"
						sel_48 = ""
						sel__1 = ""
						pRsp = "< 24 timer"
						case 48
						sel_0 = ""
						sel_4 = ""
						sel_12 = ""
						sel_24 = ""
						sel_48 = "SELECTED"
						sel__1 = ""
						pRsp = "< 48 timer"
						case else
						sel_0 = "SELECTED"
						sel_4 = ""
						sel_12 = ""
						sel_24 = ""
						sel_48 = ""
						sel__1 = ""
						pRsp = "Alle"
						end select 
						
						
						    if print <> "j" then
						    %>
						
							<option value="0" <%=sel_0%>> Alle </option>
							<option value="4" <%=sel_4%>> < 4 timer</option>
							<option value="12" <%=sel_12%>> < 12 timer</option>
							<option value="24" <%=sel_24%>> < 24 timer</option>
							<option value="48" <%=sel_48%>> < 2 døgn</option>
							<option value="-1" <%=sel__1%>>Overskreddet</option>
							</select>
							<%else %>
							<%=pRsp %>
							<%end if %>
							</td>
							<td>
						</tr>
			
			<tr>
		<td align=right style="padding-right:5px;"><b>Status:</b></td><td>
		<%if print <> "j" then %>
		<select name="FM_status" style="font-family: arial; font-size: 10px; width:240px;" onchange=renssog();>
	    <option value="0">Alle</option>
		<%
		pForvalgt = "Alle"
		end if %>					
							
							<%if statusSel = "-1" then 
							forvalgteSEL = "SELECTED"
							pForvalgt = "Alle forvalgte"
							else
							forvalgteSEL = ""
							end if%>
							
							<%if print <> "j" then %>
							<option value="-1" <%=forvalgteSEL%>>Alle forvalgte</option>
							<%
							end if
							
							strSQL = "SELECT id, navn, vispahovedliste FROM sdsk_status WHERE id <> 0 ORDER BY navn"
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							if cint(statusSel) = cint(oRec("id")) then
							isSelected = "SELECTED"
							pForvalgt = oRec("navn")
							else
							isSelected = ""
							end if
							
							
							if oRec("vispahovedliste") = 1 then
							forvalgt = "(forvalgt)"
							else
							forvalgt = ""
							end if%>
							
							<%if print <> "j" then %>
							<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%>&nbsp;&nbsp<%=forvalgt %></option>
							<%end if
							
							
							oRec.movenext
							wend
							
							oRec.close
							
							if print <> "j" then %>
							</select>
							<%else %>
							<%=pForvalgt%>
							<%end if %>
	
	</td></tr>
	<tr>
	    <td align=right style="padding-right:5px;"><b>Søg på Incident Id:</b>&nbsp;<br />
	    (Ignorerer andre filter kriterier)</td>
	    <td>
	    <%if print <> "j" then %>
	    <input id="sog" name="sog" type="text" value="<%=sogVal %>" style="font-family: arial; font-size: 10px; width:100px;"/>
	    <%else %>
	    <%=sogVal %>
	    <%end if %>
	</td>
	</tr>
	</table>
	
	</td><td valign=top style="padding-top:10px;">
	
	<table cellspacing=0 cellpadding=4 border=0>
	
		<tr><td align=right style="padding-right:5px;"><b>Prioitet:</b></td><td>

		<input type="hidden" name="FM_prio_tp" id="FM_prio_tp" value="1">
		<%
		strSQL = "SELECT id, navn FROM sdsk_prio_typ"
		
		oRec.open strSQL, oConn, 3
							
		'response.write strSQL
		'response.flush
		
		                    if print <> "j" then%>
		                    <select name="FM_prioitet" style="font-family: arial; font-size: 10px; width:240px;" onchange=renssog();>
							<option value="-1">Alle</option>
							
							<%
							pPrio = "Alle"
							end if
							
							while not oRec.EOF
							
							if cint(prioSel) = oRec("id") then
							isSelected = "SELECTED"
							pPrio = oRec("navn")
							else
							isSelected = ""
							end if
							
							if print <> "j" then
%>
							    <option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
							<%
							end if
							
							oRec.movenext
							wend
							
							oRec.close
							
							if print <> "j" then%>
							<option value="0" >Ingen</option>
							</select>
							<%else %>
							<%=pPrio %>
							<%end if %>
	</td></tr>
	<tr><td align=right style="padding-right:5px;"><b>Type:</b></td>
	<td>
	<% if print <> "j" then %>
	<select name="FM_priotype" style="font-family: arial; font-size: 10px; width:240px;">
	<option value="-1">Alle</option>
	<% end if
	strSQL = "SELECT id, navn FROM sdsk_prioitet WHERE id  <> 0 ORDER BY navn"

		oRec.open strSQL, oConn, 3 
							
							while not oRec.EOF
							
							if cint(priotypeSel) = oRec("id") then
							isSelected = "SELECTED"
							ptPrio = oRec("navn")
							else
							isSelected = ""
							end if
							
							if print <> "j" then
							%>
							    <option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
    							
							    <%
							    
							
							end if
							
							oRec.movenext
							wend
							
							oRec.close
							
							if print <> "j" then%>
							</select>
							<%else 
							if ptPrio = "" then
							ptPrio = Request.Form("FM_priotype")
							end if %>
							<%=ptPrio %>
							<%end if %>
	</td>
	</tr>
	</td></tr>
	<tr><td align=right style="padding-right:5px;"><b>Ansvarlig (Intern):</b></td><td>
	            <%if print <> "j" then %>
	            <select name="FM_ansv" style="font-family: arial; font-size: 10px; width:240px;" onchange=renssog();>
							<option value="0">Alle</option>
							
							
							<%
							pAnsv = "Alle"
							end if
							
							
							strSQL = "SELECT mid, mnavn, mnr FROM medarbejdere WHERE mid <> 0 AND mansat <> 2 ORDER BY mnavn"
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							if cint(ansSel) = cint(oRec("mid")) then
							isSelected = "SELECTED"
							pAnsv = oRec("mnavn") & " ("& oRec("mnr") &")"
							else
							isSelected = ""
							end if
							
							if print <> "j" then%>
							<option value="<%=oRec("mid")%>" <%=isSelected%>><%=oRec("mnavn")%> (<%=oRec("mnr") %>)</option>
							<%end if
							
							oRec.movenext
							wend
							
							oRec.close
							
							if ansSel = "-1" then
							isSelected = "SELECTED"
							pAnsv = "Ikke angivet"
							else
							isSelected = ""
							end if
							
							if print <> "j" then%>
							<option value="-1" <%=isSelected%>>Ikke angivet</option>
							</select>
							<%else %>
							<%=pAnsv %>
							<%end if %>
	</td></tr>
	<tr>
		<td align=right style="padding-right:5px;"><b>Benyt datointerval:</b></td>
		<td>
		<%if print <> "j" then %>
		<input type="radio" name="usedatokri" value="j" <%=datoKriJa%>>&nbsp;Ja&nbsp;&nbsp;
		<input type="radio" name="usedatokri" value="n" <%=datoKriNej%>>&nbsp;Nej
		<%else %>
		    <%if datoKrival = "j" then %>
		    Ja
		    <%else %>
		    Nej
		    <%end if %>
		<%end if %>
		</td>
	
	</tr>
	<tr><%if print <> "j" then %>
	    <td colspan="2">
		    <table>
			    <tr><!--#include file="inc/weekselector_b.asp"--></tr>
		    </table>
		    <%else %>
		    <td>
		    <b style="float:right">Periode:</b></td><td> <%=formatdatetime(strDag &"/"& strMrd &"/"& strAar, 1) %> til <%=formatdatetime(strDag_slut &"/"& strMrd_slut &"/"& strAar_slut, 1) %>
		    <%end if %>
	    </td>
	</tr>
	</table>
	
	
	</td>
	</tr>
	
	
	<%
	if len(request("FM_sort")) <> 0 then
	sortBy = request("FM_sort")
	else
		if len(request.cookies("sdsk")("sortby")) <> 0 then
		sortBy = request.cookies("sdsk")("sortby")
		else
		sortBy = 1
		end if
	end if
	
	response.cookies("sdsk")("sortby") = sortBy
	response.cookies("sdsk").expires = date + 10
	
	
	if print <> "j" then%>
	
	<tr>
	
	<td colspan=2 align=right style="padding-right:50px;">
	Sorter efter:&nbsp;<select name="FM_sort" id="FM_sort" style="background-color:#ffff99; font-family: arial; font-size: 10px; width:200px;">
		
		<%
		sSel_1 = ""
		sSel_2 = ""
		sSel_3 = ""
		sSel_4 = ""
		sSel_5 = ""
		sSel_6 = ""
		sSel_7 = ""
		sSel_8 = ""
		sSel_10 = ""
		select case sortBy
		case 1
		sSel_1 = "SELECTED"
		case 2
		sSel_2 = "SELECTED"
		case 3
		sSel_3 = "SELECTED"
		case 4
		sSel_4 = "SELECTED"
		case 5
		sSel_5 = "SELECTED"
		case 6
		sSel_6 = "SELECTED"
		case 7
		sSel_7 = "SELECTED"
		case 8
		sSel_8 = "SELECTED"
		case 10
		sSel_10 = "SELECTED"
		case else '9 / 0
		sSel_9 = "SELECTED"
		end select %>
		
		
		<option value="1" <%=sSel_1%>>Dato (oprettelses dato)</option>
		<option value="7" <%=sSel_7%>>Udførelses Dato</option>
		<option value="6" <%=sSel_6%>>Jobnavn</option>
		<option value="2" <%=sSel_2%>>Id</option>
		<option value="3" <%=sSel_3%>>Status</option>
		<option value="4" <%=sSel_4%>>Kontakt</option>
		<option value="5" <%=sSel_5%>>Kategori</option>
		<option value="8" <%=sSel_8%>>Prioritet</option>
		<option value="10" <%=sSel_10%>>Type</option>
		<option value="9" <%=sSel_9%>>Sortering (Drag'n drop mode)</option>
		</select>&nbsp;&nbsp;
	<input type="submit" value="  Søg  "></td></tr>
	</form>
	<%end if %>
	
	</table>
	</td></tr></table>
	</div>
	
	 <br><br><br>
    
    
    <h3>Servicedesk - Incidentliste</h3>
    
    <script type="text/javascript">
    	$(document).ready(function() {
    		$("select[name*=ajax]").AjaxUpdateField({parent : "tr", subselector : "td:first > input[name=rowId]"});
<%if sortBy = "9" AND kview <> "j" AND print <> "j" then%>    		
$("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode : "td:first > :input[name=rowId]"});
<%end if %>
    	});</script>
	<%if kview = "j" then
	oprLinK = "sdsk.asp?func=opr1&usekview=j&FM_kontakt="&selKid
			
			'*** Gem rsp tider på kundeside?? **
			gemTider = 0
			strSQL = "SELECT k.sdskpriogrp, gp.gemtider FROM kunder k "_
			&"LEFT JOIN sdsk_prio_grp gp ON (gp.id = k.sdskpriogrp) WHERE kid = "& selKid
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			gemTider = oRec("gemtider")
			end if
			oRec.close
			
	else
	gemTider = 0
	oprLinK = "sdsk.asp?func=opr&FM_kontakt="&selKid
	end if
	
	estimerettidTOT = 0
	
	
	
	tTop = 0
	tLeft = 0
	tWdth = 1184
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	<table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#ffffff" id="incidentlist">
	<%if kview = "j" then
	oprBg = "#ffffff"
	else
	oprBg = "#ffffff"
	end if %>
	<%if print <> "j" then%>
	<tr>
	    <td bgcolor="<%=oprBg %>" style="padding:5px 10px 5px 0px;">
	        <%if kview <> "j" then %>
	        <a href="javascript:popUp('stopur_2008.asp?FM_mid=<%=session("mid")%>','1000','650','20','20')"><img src="../ill/stopur_32.png" alt="Se Stopurs Entries" border=0 /></a>
	        <%end if %>
            &nbsp;
	</td>
	    <td bgcolor="<%=oprBg %>" style="padding:5px 10px 5px 0px;">
	       <%if kview <> "j" then %>
	       <a href="javascript:popUp('stopur_2008.asp','1000','650','20','20')">Se Stopurs Entries</a>
	       <%end if %>
            &nbsp;
	</td>
	
	<%if kview <> "j" then 
	pdTop = 5
	else
	pdTop = 10
	end if
	%>
	
	
	
	    <td bgcolor="<%=oprBg %>" align=right colspan=10 style="padding:<%=pdTop%>px 10px <%=pdTop%>px 0px;"><a href="<%=oprLinK%>">Opret ny Incident</a></td>
	</td>
		<td bgcolor="<%=oprBg %>" style="padding:<%=pdTop%>px 10px <%=pdTop%>px 0px;"><a href="<%=oprLinK%>"><img src="../ill/add2.png" alt="" border="0"></a></td>
	</tr>
	<%end if %>
	<tr bgcolor="#8caae6">
	<td class=alt align=center height=40 valign=bottom><b>Id</b></td>
	<td class=alt valign=bottom><b>Opr. Dato & Kl.</b><br /> 
	
	Ønskes udført dato
	<%
	if len(gemTider) <> 0 then
	gemTider = gemTider
	else
	gemTider = 0
	end if
	
	if kview = "j" AND cint(gemTider) = 1 then%>
	<%else%>
	<br><b>Respons Dato & Kl.</b>
	<%end if%>
	</td>
	<td class=alt valign=bottom style="padding:0px 5px 0px 0px;"><b>Estim.</b></td>
	<td class=alt valign=bottom style="padding-left:10px;" ><b>Kontakt</b><br />kontaktpers.</td>
	<td class=alt valign=bottom><b>Emne</b><br />
	Beskrivelse<br />
	Ansvarlig<br />
	Opr. af</td>
	<td class=alt valign=bottom><b>Log</b></td>
	<td class=alt valign=bottom ><b>Status</b></td>
	<td class=alt valign=bottom ><b>Kategori</b></td>
	<td class=alt valign=bottom ><b>Type</b></td>
    <td class=alt valign=bottom ><b>Prioitet</b></td>
	<td class=alt valign=bottom ><b>Filer</b></td>
	<td class=alt valign=bottom >
	<b>Stopur</b>
	</td>
	<td class=alt valign=bottom><b>Til timreg.</b><br />
	(Job tilknytning)</td>
	<td class=alt valign=bottom>
        &nbsp;</td>
	</tr>
	
	<%
	
	expTxt = "Id;Opr. Dato & Kl.;Ansvarlig;Respons Dato & Kl.;Kontakt;Kontakt Id;Kontaktperson 1;Email 1;Kontaktperson 2;Email 2;Status;Kategori;Type;Prioitet;Emne;Beskrivelse;"
	expTxt = expTxt &"xx99123sy#z"
	
	'**** Firma normalåbningstider fre kontrolpanel ****
	
	strSQL = "SELECT normtid_st_man, normtid_sl_man, normtid_st_tir, normtid_sl_tir, "_
	&" normtid_st_ons, normtid_sl_ons, normtid_st_tor, normtid_sl_tor, normtid_st_fre, "_
	&" normtid_sl_fre, normtid_st_lor, normtid_sl_lor, normtid_st_son, normtid_sl_son, brugabningstid "_
	&" FROM licens WHERE id = 1"
	
	oRec.open strSQL, oConn, 3
    if not oRec.EOF then
        
        normtid_st_man  = oRec("normtid_st_man")
        normtid_sl_man = oRec("normtid_sl_man") 
        normtid_st_tir = oRec("normtid_st_tir")  
        normtid_sl_tir = oRec("normtid_sl_tir")
        normtid_st_ons = oRec("normtid_st_ons") 
        normtid_sl_ons = oRec("normtid_sl_ons") 
        normtid_st_tor = oRec("normtid_st_tor") 
        normtid_sl_tor = oRec("normtid_sl_tor") 
        normtid_st_fre = oRec("normtid_st_fre") 
        normtid_sl_fre = oRec("normtid_sl_fre")
        normtid_st_lor = oRec("normtid_st_lor")
        normtid_sl_lor = oRec("normtid_sl_lor") 
        normtid_st_son = oRec("normtid_st_son") 
        normtid_sl_son = oRec("normtid_sl_son") 
        
        brugabningstid = oRec("brugabningstid")
        
    end if
    oRec.close
	
	
	'Response.Write dato
	'Response.Flush
	
	nu = dato &" "& tid
	
	'*** D ****
	manTimer = dateDiff("h", normtid_st_man, normtid_sl_man)
	tirTimer = dateDiff("h", normtid_st_tir, normtid_sl_tir) 
	onsTimer = dateDiff("h", normtid_st_ons, normtid_sl_ons) 
	torTimer = dateDiff("h", normtid_st_tor, normtid_sl_tor) 
	freTimer = dateDiff("h", normtid_st_fre, normtid_sl_fre) 
	lorTimer = dateDiff("h", normtid_st_lor, normtid_sl_lor) 
	sonTimer = dateDiff("h", normtid_st_son, normtid_sl_son) 
	
	'Response.Write "<br>her"&  manTimer
	'Response.Write "<br>her"&  tirTimer
	'Response.Write "<br>her"&  onsTimer
	'Response.Write "<br>her"&  torTimer
	'Response.Write "<br>her"&  freTimer
	'Response.Write "<br>her"&  lorTimer
	'Response.Write "<br>her"&  sonTimer
	
	Dim useDayHours, useDayStTime, useDayEndTime
	Redim useDayHuors(7), useDayStTime(7), useDayEndTime(7)
	
	for i = 1 to 7
	
	    select case i
	    
	    case 2
	    useDayHuors(2) = manTimer
	    useDayStTime(2) = normtid_st_man
	    useDayEndTime(2) = normtid_sl_man
	    
	    case 3
	    useDayHuors(3) = tirTimer
	    useDayStTime(3) = normtid_st_tir
	    useDayEndTime(3) = normtid_sl_tir
	    
	    case 4
	    useDayHuors(4) = onsTimer
	    useDayStTime(4) = normtid_st_ons
	    useDayEndTime(4) = normtid_sl_ons
	    
	    case 5
	    useDayHuors(5) = torTimer
	    useDayStTime(5) = normtid_st_tor
	    useDayEndTime(5) = normtid_sl_tor
	   
	    case 6
	    useDayHuors(6) = freTimer
	    useDayStTime(6) = normtid_st_fre
	    useDayEndTime(6) = normtid_sl_fre
	    
	    case 7
	    useDayHuors(7) = lorTimer
	    useDayStTime(7) = normtid_st_lor
	    useDayEndTime(7) = normtid_sl_lor
	    
	    case 1
	    useDayHuors(1) = sonTimer
	    useDayStTime(1) = normtid_st_son
	    useDayEndTime(1) = normtid_sl_son
	    end select
	    
	next
	  
	
	select case sortBy
	case 1
	grpAndOrderBy  = " GROUP BY s.id ORDER by s.dato DESC, s.tidspunkt DESC"
	case 2
	grpAndOrderBy  = " GROUP BY s.id ORDER by s.id DESC"
	case 3
	grpAndOrderBy  = " GROUP BY s.id ORDER by s.status, s.id DESC"
	case 4
	grpAndOrderBy  = " GROUP BY s.id ORDER by k.kkundenavn, s.id DESC"
	case 5
	grpAndOrderBy  = " GROUP BY s.id ORDER by s.type, s.id DESC"
	case 6
	grpAndOrderBy  = " AND s.jobid <> 0 GROUP BY s.id ORDER by j.jobnavn, s.id DESC"
	case 7
	grpAndOrderBy  = " AND s.duedate > '2001-1-1' AND useduedate = 1 GROUP BY s.id ORDER by s.duedate, s.dato, s.tidspunkt"
	case 8
	grpAndOrderBy = " GROUP BY s.id ORDER by s.priotype ASC"
	case 9
	grpAndOrderBy = " GROUP BY s.id ORDER by sortorder ASC"
	case 10
	grpAndOrderBy = " GROUP BY s.id ORDER by s.prioitet ASC"
	end select
	
	
	strSQL = "SELECT s.editor, s.priotype, s.public, s.id, s.emne, s.besk, s.dato, s.tidspunkt,"_
	&" p.navn AS pnavn, p.responstid, p.type, p.kunweekend, s.type, s.ansvarlig, "_
	&" t.navn AS type, k.kkundenavn, k.kkundenr, k.kid, m.mnavn, m.mnr, m.init, "_
	&" s.kundeid, s.prioitet, s.status, st.navn AS statusnavn, st.farve, "_
	&" count(rel.id) AS antallog, s.editor2, s.dato2, "_
	&" s.jobid, j.jobnavn, j.jobnr, s.kpers, kp.navn AS kpers1, kp.email AS kpers1email, "_
	&" s.kpers2, s.kpers2email, s.duedate, s.useduedate, s.esttid, s.sortorder FROM sdsk s "_
	&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet) "_
	&" LEFT JOIN sdsk_prio_typ pt ON (pt.id = p.type) "_
	&" LEFT JOIN sdsk_typer t ON (t.id = s.type) "_
	&" LEFT JOIN kunder k ON (k.kid = s.kundeid) "_
	&" LEFT JOIN kontaktpers kp ON (kp.id = s.kpers) "_
	&" LEFT JOIN job j ON (j.id = s.jobid) "_
	&" LEFT JOIN sdsk_status st ON (st.id = s.status) "_
	&" LEFT JOIN medarbejdere m ON (m.mid = s.ansvarlig)"_
	&" LEFT JOIN sdsk_rel rel ON (rel.sdsk_rel = s.id)"
	
	if len(sogeKri) <> 0 then
	strSQL = strSQL & " WHERE " & sogeKri & grpAndOrderBy 
	else
	strSQL = strSQL &" WHERE "& kontaktKri  & typeKri & prioitetKri & ansvKri & datoKri & statusKri & typeKri & ""_
    & grpAndOrderBy 
	end if
	
	
	'if lto = "accendo" then
	'response.write strSQL &"<br><br>"
	'response.flush
	'end if
	
	d = 0
	i = 0
	s = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
		
		
		rspDagogTidUse = ""
		
		
		dag_kl_OprNew = oRec("dato") &" "& formatdatetime(oRec("tidspunkt"), 3)
		dagOpr = oRec("dato")
		
		aftaleRspTimer = oRec("responstid")
		thisWeekday = oRec("dato")
		 tilfojTimerIalt = 0
		
		if len(oRec("responstid")) <> 0 then
		addTimerTrsp = oRec("responstid")
		else
		addTimerTrsp = 0
		end if  
		
		rspDagogTid = dateadd("h", addTimerTrsp, dag_kl_OprNew)
		
		if cint(oRec("kunweekend")) = 1 AND cint(brugabningstid) = 1 then
		call response_dagogTid(rspDagogTid)
		rspDagogTid = rspDagogTidUse
		'ab = "rspDagogTid" & rspDagogTid & "a"
		else
		len_rspDagogTid = len(rspDagogTid)
		left_rspDagogTid = left(rspDagogTid, len_rspDagogTid - 3)
		rspDagogTid = left_rspDagogTid
		'ab = "b"
		end if     
		
		nutid = date() & " " & time()
		timerTilrespons = datediff("h", nutid, rspDagogTid, 2, 3)           
		               
		
        
		
		
	if cint(rspKri) = 0 OR (cint(rspKri) > 0 AND cint(rspKri) > cint(timerTilrespons) AND cint(timerTilrespons) >= 0) OR (cint(rspKri) = -1 AND cint(timerTilrespons) < 0) then%>
	
	<%
	            if cint(lastedit) = cint(oRec("id")) then 'OR lastedit = 0 AND i = 0
	            bgThis = "#ffff99"
            	
	            vzb = "hidden" 
	            dsp = "none"
            	
	            lastopenDiv = lastedit
	            else
	                select case right(i, 1)
	                case 2,4,6,8,0
	                bgThis = "#ffffff"
	                case else
            	    bgThis = "#EFF3FF"
            	    end select
	            vzb = "hidden" 
	            dsp = "none"
            	
	            lastopenDiv = 0
	            end if
	
	
	
	
	%>
    
    <tr bgcolor="<%=bgThis%>">
		<td valign=top align=center bgcolor="<%=bgThis%>" style="padding:5px 2px 2px 2px; border-top:1px #C4C4C4 solid;"><input type="hidden" name="SortOrder" value="<%=oRec("sortorder")%>" />
		<input type="hidden" name="rowId" value="<%=oRec("id")%>" /><%=oRec("id")%><br />
            <img src="../ill/ticket_green.gif" />
	    </td>
		
		  <td valign=top style="padding:5px 0px 2px 2px; white-space:nowrap; width:160px; border-top:1px #C4C4C4 solid;">
		<%
		datoogklokkeslet = left(weekdayname(weekday(oRec("dato"))), 3) &". "&formatdatetime(oRec("dato"), 2) &" "& left(formatdatetime(oRec("tidspunkt"), 3), 5)
		%>
		
	
		 
		
		<table cellspacing=0 cellpadding=0 border=0 width=100%><tr>
		         <td valign=top style="padding-top:1px;"><img src="../ill/bullet_square_green.png" alt="Modtaget d."/> </td>
		         <td valign=top style="white-space:nowrap;"><font class=lgreen><%=datoogklokkeslet%></font></td>
		         </tr>
		
		
		
		
		<%if kview <> "j" then %>
		    <%if oRec("useduedate") = 1 then %>
    		
		    <%
		    duedate = left(weekdayname(weekday(oRec("duedate"))), 3) &". "&formatdatetime(oRec("duedate"), 2) &" "& left(formatdatetime(oRec("duedate"), 3), 5)
		    %>
		    <tr>
		             <td valign=top style="padding-top:1px;"><img src="../ill/bullet_square_grey.png" alt="Ønskes udført d." /></td>
		             <td valign=top style="white-space:nowrap;"><font class=lillesort><%=duedate%></font></td>
		             </tr>
    	     
	        <%end if %>
	    <%end if %>
		
		
		
		
	    <%
		'*** resp tid ***
		addTimerTrspShow= ""
		if kview = "j" AND cint(gemTider) = 1 then%>
		
		<%else%>
		
	          
		
		        <%if cint(addTimerTrsp) <> 0 then %>
		       
		        <% 
		        addTimerTrspShow = left(weekdayname(weekday(rspDagogTid)), 3)&". "& rspDagogTid
		        %>
		         <tr>
		         <td valign=top style="padding-top:1px;"><img src="../ill/bullet_square_red.png" alt="Incident skal være påbegyndt inden d." /></td>
		         <td valign=top style="white-space:nowrap;"><font class=lroed><%=addTimerTrspShow%></font></td>
		         </tr>
		        <%end if %>
		
		
		<%end if%>
		
		
		</table>
	
		</td>
		
		<td style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid; white-space:nowrap;" align=right valign=top class=lille>
		<%if oRec("esttid") <> 0 then %>
		<%=formatnumber(oRec("esttid"), 1) %>
		<%
		estimerettidTOT = estimerettidTOT + oRec("esttid")
		else %>
		&nbsp;
		<%end if %></td>
		<td valign=top class=lille style="padding:5px 5px 2px 10px; border-top:1px #C4C4C4 solid;">
	    <b><%=oRec("kkundenavn")%></b> (<%=oRec("kkundenr")%>)
	    
	    
	   
	    <%if oRec("kpers") <> 0 then %>
	    <br /><%=oRec("kpers1") %> <a href="mailto:<%=oRec("kpers1email") %>" class=rmenu><%=oRec("kpers1email") %></a>
	    <%end if %>
	    
	    <%if len(oRec("kpers2")) <> 0 then %>
	    <br /><%=oRec("kpers2") %> <a href="mailto:<%=oRec("kpers2email") %>" class=rmenu><%=oRec("kpers2email") %></a>
	    <%end if %>
	 
	    </td>
		
		  <td valign=top style="padding:5px 10px 2px 2px; width:400px; border-top:1px #C4C4C4 solid;"> 
		 <%if oRec("public") = 1 then%>
		&nbsp;<img src="../ill/ac0058-16.gif" width="16" height="16" alt="Offentlig tilgængelig" border="0">
		<%end if%>
		
		<%if oRec("status") = 1 OR oRec("status") = 4 OR oRec("status") = 18 then  %>
		<s>
		<%end if %>
		
		<%if kview = "j" OR print = "j" then%>
		<b><%=oRec("emne")%></b>
		<%else%>
		<a href="sdsk.asp?func=red&id=<%=oRec("id")%>" class="vmenu"><i><%=oRec("emne")%></i></a>
		<%end if%>
		
		
		<i>
		
		<%
		strBesk = oRec("besk")
	    %>
		
		<%if len(strBesk ) > 150 then %>
		<%=left(strBesk , 150) %>..
		<%else %>
		<%=strBesk  %>
		<%end if %></i>
		
		<%if oRec("status") = 1 OR oRec("status") = 4 OR oRec("status") = 18 then  %>
		</s>
		<%end if %>
		<br />
		
		 <%if len(oRec("mnavn")) <> 0 then
		ansvrh = oRec("mnavn") & " ("&oRec("mnr")&") "
		else
		ansvrh = " (ansv. ikke angivet) "
		end if%>
		 
		<%if len(oRec("init")) <> 0 then
		ansvrh = ansvrh & " - "& oRec("init")
		end if%>
		
		<font class=megetlillesilver><i> Ansv. <%=ansvrh %>
		&nbsp;Opr. af: <%=oRec("editor") %> 
		
		<% 
	if len(oRec("editor2")) <> 0 then %>
    &nbsp;Sidst red.: <%=oRec("editor2")%> d. <%=oRec("dato2")%>
    <%
	end if
	%></font>
		
		
		</td>
		<td style="padding:5px 2px 2px 2px; border-top:1px #C4C4C4 solid; white-space:nowrap; width:120px;" valign=top><a href="#" onClick="showbesk('<%=oRec("id")%>');return false;" class=vmenu>
            Log entries (<%=oRec("antallog")%>) </a>
            
        <%if print <> "j" AND kview <> "j" then %><br />
		<a href="javascript:popUp('sdsk_tilfoj.asp?sdskrelid=<%=oRec("id")%>&id=0&func=opr&lastedit=<%=oRec("id")%>&usekview=<%=kview%>&FM_kontakt=<%=kontaktId%>','600','650','250','30')" class=lgron>Tilføj entry +</a>
		<%end if %>
		</td>
	  
		
		<% if kview <> "j" and print <> "j" then 
		'updateabilities for incidents todo: add multiple update
		%>
		<td valign=top style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid;" class=lille>
		<select name="FM_ajaxstatus" class="txtField" style="width:100px;">
		<%
		if kview = "j" then
		lmt = " LIMIT 0, 1"
		else
		lmt = ""
		end if
		
		strSQL = "SELECT id, navn FROM sdsk_status ORDER BY navn" & lmt & ""
		oRec2.open strSQL, oConn, 3
		while not oRec2.EOF
		if oRec("status") = oRec2("id") then
		stsel = "selected"
		else
		stsel = ""
		end if 
		%>
		<option value="<%=oRec2("id")%>" <%=stsel%>><%=oRec2("navn")%></option>
		<%
		oRec2.movenext
		wend
		oRec2.close
		%>
		<option value="0">Ingen</option>
		</select></td>    
	    <td valign=top style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid;" class=lille>	
	    <select name="FM_ajaxtype" class="txtField" style="width:100px;">
		<%
		strSQL = "SELECT id, navn FROM sdsk_typer ORDER BY navn"
		oRec2.open strSQL, oConn, 3
		while not oRec2.EOF
		if oRec("type") = oRec2("navn") then
		tsel = "selected"
		else
		tsel = ""
		end if
		
		%>
		<option value="<%=oRec2("id")%>" <%=tsel%>><%=oRec2("navn")%></option>
		<%
		oRec2.movenext
		wend
		oRec2.close
		%>
		<option value="0">Ingen</option>
		</select></td>
			<td valign=top style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid;" class=lille><%
		strSQL = "SELECT p.id, p.navn, p.responstid, pt.navn AS typeNavn FROM kunder k "_
		&" LEFT JOIN sdsk_prio_grp g ON (g.id = k.sdskpriogrp)"_
		&" LEFT JOIN sdsk_prioitet p ON (priogrp = g.id) "_
		&" LEFT JOIN sdsk_prio_typ pt ON (pt.id = p.type) "_
		&" WHERE kid = "& oRec("kundeid") &" ORDER BY p.navn"
		%>
		
		<select name="FM_ajaxprio" style="width:100px;" class="txtField">
		<%
		
		
		oRec2.open strSQL, oConn, 3
		while not oRec2.EOF
		if oRec("prioitet") = oRec2("id") then
		psel = "selected"
		else
		psel = ""
		end if 
		%>
		<option value="<%=oRec2("id")%>" <%=psel%>><%=oRec2("navn")%></option>
		<%
		oRec2.movenext
		wend
		oRec2.close
		%>
		<option value="0">Ingen</option>
		</select></td>
		<td valign=top style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid;" class=lille>
		<select name="FM_ajaxprio_typ" class="txtField">
		<option value="0">Ingen</option>
	<%
	strSQL = "SELECT id, navn FROM sdsk_prio_typ"
	oRec2.open strSQL, oConn, 3
	while not oRec2.EOF
		if cint(oRec("priotype")) = cint(oRec2("id")) then
		tpSEL = "SELECTED"
		else
		tpSel = ""
		end if%>
		
	<option value="<%=oRec2("id")%>" <%=tpSel%>><%=oRec2("navn")%></option>
	<%
	oRec2.movenext
	wend
	
	oRec2.close
	%>
	</select>
            &nbsp;</td>
		<%else %>
		<td valign=top style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid;" class=lille><%=oRec("statusnavn")%></td>    
		<td valign=top style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid;" class=lille><b><%=oRec("type")%></b></td>
	    <td valign=top style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid;" class=lille><%=oRec("pnavn")%></td>
	<%
	if oRec("priotype") <> 0 then
	strSQL = "SELECT id, navn FROM sdsk_prio_typ WHERE id = " & oRec("priotype")
	oRec2.open strSQL, oConn, 3
	while not oRec2.EOF
	strPrioTypeNavn = oRec2("navn")
	oRec2.Movenext
	wend
	oRec2.close
	response.Flush
	else
	strPrioTypeNavn ="Ikke sat"
	end if%>
		<td valign=top style="padding:5px 5px 2px 2px; border-top:1px #C4C4C4 solid;" class=lille><%=strPrioTypeNavn%></td>
	
	    <%
	    end if
	    call htmlparseCSV(strBesk)
		strBesk = htmlparseCSVtxt
		
		
	    
	    expTxt = expTxt & oRec("id") &";"& oRec("dato") &" "& formatdatetime(oRec("tidspunkt"), 3) &";"_
	    & ansvrh &";"_
	    & rspDagogTid &";"&oRec("kkundenavn")&";"&oRec("kkundenr")&";"&oRec("kpers1") &";"& oRec("kpers1email") &";"&oRec("kpers2") &";"& oRec("kpers2email") &";"_
	    & oRec("statusnavn") &";"_
	    & oRec("type") &";" & oRec("pnavn") &";"& oRec("priotype") &";"& replace(oRec("emne"), Chr(34), "''") &";"& replace(strBesk, Chr(34), "''") &";"
		expTxt = expTxt &"xx99123sy#z"
	   %>
	  
	   <td valign=top style="padding:0px 5px 2px 2px; border-top:1px #C4C4C4 solid;">
	 
	   
	   <%if print <> "j" then %>
	   <a href="#" onClick="popitup('upload.asp?jobid=0&sdsk=1&type=pic&iid=<%=oRec("id")%>')" class=lgron>Tilføj fil +</a>
	   <br>
	   <%end if %>
	   <br>
		<%
		strSQLf = "SELECT fi.filnavn FROM filer fi WHERE fi.incidentid = "& oRec("id") & " AND fi.incidentid <> 0" ' AND fi.kundeid <> 0
		'Response.Write strSQLf
		oRec3.open strSQLf, oConn, 3
		while not oRec3.EOF
		%>
		<a href="../inc/upload/<%=lto%>/<%=oRec3("filnavn")%>" class='vmenulille' target="_blank"><%=left(oRec3("filnavn"), 15)%>..</a><br>
		<%
		oRec3.movenext
		wend
		
		oRec3.close
		%></td>
	    
	   
	   <td style="padding:5px 2px 2px 2px; border-top:1px #C4C4C4 solid;" valign=top>
		<%if print <> "j" AND kview <> "j" then%>
		    &nbsp;<a href="javascript:popUp('stopur_2008.asp?func=ins&incid=<%=oRec("id")%>&logentry=0&jobid=<%=oRec("jobid") %>&FM_mid=<%=session("mid")%>','1000','650','20','20')" class=vmenu><img src="../ill/stopur_st_stop.png" alt="Tilføj Stopur's Entry" border="0" /></a>
		    
		<%end if%>
           &nbsp;
		</td> 
	    
	   <td style="padding:5px 0px 5px 2px; border-top:1px #C4C4C4 solid;" valign=top>
	        <%if oRec("jobid") <> 0 then%>
				<%if kview <> "j" AND print <> "j" then%>
				<a href="javascript:popUp('timereg_akt_2006.asp?jobid=<%=oRec("jobid")%>&usemrn=<%=session("mid")%>&showakt=1&fromsdsk=1&sdskkomm=<%="<br><br> == Incident Id "& oRec("id") &" == <br> Opr. "& datoogklokkeslet  &"<br>"& oRec("emne")%>','950','620','50','20');" target="_self"; class=rmenu><%=oRec("jobnavn") & "&nbsp;("& oRec("jobnr") &")"%></a>
				<%else %>
				<%=oRec("jobnavn") & "&nbsp;("& oRec("jobnr") &")"%>
				<%end if%>
			<%end if%>
           &nbsp;</td>
	   
	   
	   <%if level = 1 AND print <> "j" AND kview <> "j" then%>
	   
		<td style="padding:5px 0px 5px 2px; border-top:1px #C4C4C4 dashed;" valign=top><a href="sdsk.asp?func=slet&id=<%=oRec("id")%>">
            <img src="../ill/slet_16.gif" border="0" /></a></td>
		<%end if%>
	 
	
	</tr>
	<tr>

	<td colspan=13 style="padding:10px;">
      
	
	<div id="b_<%=oRec("id")%>" name="b_<%=oRec("id")%>" style="position:relative; display:<%=dsp%>; border:0px #c4c4c4 solid; padding:20px; width:100%;">
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr><td align=right colspan=2><a onClick="showbesk(<%=oRec("id")%>); return false;" class=red>Luk denne log [x] </a></td></tr>
	<tr>
	<td colspan=2 valign=bottom>
	
	<h4>Incident log:</h4>
	<b><%=oRec("emne")%></b> - Incident Id: <%=oRec("id")%><br />
	<%=strBesk %>
	
	</td>
	 <td>
        &nbsp;</td>
	</tr>
	<tr><td colspan=2 height=20>
        &nbsp;</td></tr>
	
	
	<%
	strSQL2 = "SELECT besk, sdskdato, sdsktidspunkt, id, editor, public, "_
	&" losning, editor2, dato2 FROM sdsk_rel WHERE sdsk_rel = "& oRec("id") &" ORDER BY id DESC"
	
	'Response.Write strSQL2
	'Response.Flush
	
	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF
	
	%>
	<tr><td style="border-top:1px #C4C4C4 dashed; padding:10px 10px 10px 10px;">
	
    <%if oRec2("losning") = 1 then%>
	<img src="../ill/ac0060-16.gif" width="16" height="16" alt="Godkendt som løsning" border="0">
	<%end if%>
	
	<%if oRec2("public") = 1 then%>
	<img src="../ill/ac0058-16.gif" width="16" height="16" alt="Offentlig tilgængelig" border="0">
	<%end if%>
	
	<%if kview <> "j" then%>
	<b><%=left(weekdayname(weekday(oRec2("sdskdato"))), 2)%>. <%=oRec2("sdskdato")%>&nbsp;<%=formatdatetime(oRec2("sdsktidspunkt"), 3)%></b> - Logentry Id: <%=oRec2("id") %> <br>
	<a href="javascript:popUp('sdsk_tilfoj.asp?sdskrelid=<%=oRec("id")%>&id=<%=oRec2("id")%>&func=red&lastedit=<%=oRec("id")%>&usekview=<%=kview%>&FM_kontakt=<%=kontaktId%>','600','650','250','30')" class=vmenu><i><%=oRec2("besk")%></i></a>
	<br>
	<%else%>
	<b><%=left(weekdayname(weekday(oRec2("sdskdato"))), 2)%>.&nbsp;<%=oRec2("sdskdato")%>&nbsp;<%=formatdatetime(oRec2("sdsktidspunkt"), 3)%></b> - Logentry Id: <%=oRec2("id") %> <br>
	<i><%=oRec2("besk")%></i>
	<%end if%>
	
	
	
	<br /><font class=megetlillesilver>
		<i>Oprettet af: <%=oRec2("editor")%>
		<%if len(oRec2("editor2")) <> 0 then%>
		<br>Sidst redigeret af: <%=oRec2("editor2")%> d. <%=oRec2("dato2")%>
		<%end if%>
		</i>
	
	</td>
	
	<td valign=top style="border-top:1px #C4C4C4 dashed; padding:10px 10px 10px 10px;">
	<%if kview <> "j" AND print <> "j" then%>
	<a href="javascript:popUp('stopur_2008.asp?func=ins&incid=<%=oRec("id")%>&logentry=<%=oRec2("id")%>&jobid=<%=oRec("jobid") %>&FM_mid=<%=session("mid")%>','1000','650','20','20')" class=vmenu><img src="../ill/stopur_st_stop.png" alt="Tilføj Stopur's Entry" border=0 /></a>
	<%end if %>
	&nbsp;
	<br />
	</td></tr>
	
	<%
	
	'if s > 3 then
    'Response.End
    'end if
    
    
    
    s = s + 1
	
	oRec2.movenext
	wend
	oRec2.close
	%>
	</table>
	
	
	
	
	<br><br><br><br><br><br><br><br><br>&nbsp;
	<form>
	<input type="text" name="FM_div_id_txtbox_<%=oRec("id")%>" id="FM_div_id_txtbox_<%=oRec("id")%>" value="" style="border:0px; background-color:#FFFFFF;"></form>
	</div> 
	</td></tr>
	
	
	<%
	end if
	
	
	Response.flush
	
	i = i + 1
	oRec.movenext
	wend
	
	oRec.close
	
	%>
	
	</table>
	
	<!-- slut table div-->
	</div> 
	
	<br>
	Der er ialt fundet: <b><%=i%></b> Incidents.
	<br>
	<%if estimerettidTOT <> 0 then %>
	Tid estimeret på viste incidents: <b><%=formatnumber(estimerettidTOT, 1) %></b> timer.
	<%end if %>
	<br />
	<br>
    **) Åbningstider ignoreret.
	
	<form><input type="hidden" name="divopen" id="divopen" value="<%=lastopenDiv%>"></form>
	
	
	
	<%
	if print <> "j" then
            
            if kview <> "j" then
            ptop = 20
            pleft = 810
            pwdt = 160
            else
            ptop = 102
            pleft = 0
            pwdt = 200
            end if
           
            call eksportogprint(ptop, pleft, pwdt)
            %>
               
			
            <tr>
                <td align=center>
                <a href="sdsk.asp?func=eksport"><input type="image" src="../ill/export1.png"></a>
                </td>
                </td><td>.csv fil eksport</td>
                </tr>
                </form>
                
                <tr>
                <td align=center>
                
			    <a href="sdsk.asp?print=j" target="_blank"><img src="../ill/printer3.png" border=0></a>
    			
		       </td><td>Print version</td>
               </tr>
               </form>
	            
               </table>
            </div>
            
            <%
            else
            Response.Write("<script language=""JavaScript"">window.print();</script>")
            end if%>
	
	<%end select%>
	</div>
	
	
<%
if request("func") = "eksport" then
	
	ekspTxt = expTxt
	ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	'kspTxt = replace(ekspTxt, ";", chr(34))
	datointerval = request("datointerval")
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\sdsk_eksport.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\sdsk_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\sdsk_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\sdsk_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\sdsk_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "sdsk_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				objF.WriteLine(ekspTxt)
				
	%>
	<script type="text/javascript">
	document.location.href = <%= "'../inc/log/data/"& file &"'"%>
	</script>
	<%
	end if	 

end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
