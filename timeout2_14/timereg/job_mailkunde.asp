<%Response.buffer = true%>
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
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	
	
	if len(request("Kid")) <> 0 then
	kid = request("Kid")
	else
	kid = 0
	end if
	
	'Response.Write kid & "ff " & func
	
	'if len(request("jnr")) <> 0 then
	'strjnr = request("jnr")
	'else
	'strjnr = 0
	'end if
	
	if len(request("jnavn")) <> 0 then
	strjnavn = request("jnavn")
	else
	strjnavn = ""
	end if
	
	select case func
	case "sendmail"
	
	
	kids = split(kid, ",")
	for x = 0 to UBOUND(kids)
	
	if kids(x) <> 0 then
	
	strKundeId = kids(x)
	strModtagere = ""
	
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	' Sætter Charsettet til ISO-8859-1 
	Mailer.CharSet = 2
	
	
	'*** Afsender ***
	strSQLe = "SELECT email, mnavn FROM medarbejdere WHERE mid = " & session("mid")
	oRec2.open strSQLe, oConn, 3 
	if not oRec2.EOF then
		
		afsNavn = oRec2("mnavn")
		afsEmail = oRec2("email")
		
	end if
	oRec2.close
	
	' Afsenderens navn 
	Mailer.FromName = afsNavn
	' Afsenderens e-mail 
	Mailer.FromAddress = afsEmail
	
	'*** Suppl afsender oplysninger ***
	strSQLk = "SELECT kkundenavn, adresse, postnr, city, telefon, land FROM kunder WHERE useasfak = 1"
		oRec2.open strSQLk, oConn, 3 
		if not oRec2.EOF then
			afsKundenavn = oRec2("kkundenavn")
			afsKundeadr = oRec2("adresse")
			afsKundepostnr = oRec2("postnr")
			afsKundecity = oRec2("city")
			afsKundeland = oRec2("land")
			afsKundeTlf = oRec2("telefon")
	end if
	oRec2.close 
	
	Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
	'*** Modtagere ***
	'*** Kundehovedmail ***
	
	'Response.Write "her" & "  " & strKundeId & " - "& kids(x) & "<br>" 
	if request("FM_sendtil_kunde") = "1" then
	
		strSQLm = "SELECT email, kkundenavn FROM kunder WHERE kid = " & strKundeId
		
		
		oRec3.open strSQLm, oConn, 3 
		if not oRec3.EOF then 
            modNavn = oRec3("kkundenavn")
			modEmail = oRec3("email")
			
			strModtagere = oRec3("kkundenavn") & ", " & oRec3("email") &"<br>"
			
			Mailer.AddRecipient modNavn, modEmail						
											
		end if 
		oRec3.close 							
	end if
	
	
	intKpersHvilke = Split(request("FM_sendtil_kpers"), ", ")
	For i = 0 to Ubound(intKpersHvilke)
	
	strSQLm = "SELECT navn, email AS kpersemail FROM kontaktpers WHERE id ="& intKpersHvilke(i)
	oRec.open strSQLm, oConn, 3 
	if not oRec.EOF then
 			modkpersNavn = oRec("navn")
			modKpersEmail = oRec("kpersemail")
			
			strModtagere = strModtagere & oRec("navn") & ", " & oRec("kpersemail") & "<br>"
			Mailer.AddRecipient modkpersNavn, modKpersEmail
			
	end if
	oRec.close 	
	
	next									
											
								
											
											
										
											
				'Mailer.AddBCC "Support", "support@outzource.dk" 
				' Mailens emne
				Mailer.Subject = strjnavn &" er oprettet i TimeOut."
				
				' Selve teksten
				 Mailer.BodyText = "Jobbet: " & strjnavn &" er blevet oprettet i vores online time/sag system TimeOut."& vbCrLf & vbCrLf _ 
				& "Som en service for vore kunder kan I følge med i timeforbruget på jobbet her: https://outzource.dk/timeout_xp/wwwroot/ver2_1/login_kunder.asp?key="& strLicenskey &"&lto="& lto &"" & vbCrLf & vbCrLf _ 
				& "Kontakt os venligst hvis du/i ikke er bekendt med login og password."& vbCrLf & vbCrLf _
				& "Med venlig hilsen"& vbCrLf & vbCrLf & afsNavn & vbCrLf _  
				& afsKundenavn & vbCrLf _
				& afsKundeadr & vbCrLf _
				& afsKundepostnr & " "& afsKundecity & vbCrLf _
				& afsKundeland & vbCrLf _
				& "Tlf: " & afsKundeTlf
				
				on error resume next
				Mailer.SendMail 
											
	
	end if
	next 'kids											
												
	%>
	<div name="indhold" id="indhold" style="position:absolute; display:; visibility:visible; left:20; top:20; background-color:#eff3ff; width:220; z-index:10000; border:1px #5582d2 solid; padding:5px;">
	
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="300">
	<tr><td>Mail er afsendt til de valgte modtagere.<br><br>
	<%=strModtagere%><br><br>
	<a href="Javascript:window.close()">Luk vindue</a></td></tr>
	</table>
	</div>
	<%										
	'Response.Write("<script language=""JavaScript"">window.close();</script>")
   
   
	
	
	case else
	
	
	
	%>
	<div name="indhold" id="indhold" style="position:absolute; display:; visibility:visible; left:20; top:20; background-color:#eff3ff; width:220; z-index:10000; border:1px #5582d2 solid; padding:5px;">
	
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="300">
	<tr><td colspan=3><b>Send mail til følgende kontaktpersoner vedr. oprettelse af job:</b></td></tr>
	<form action="job_mailkunde.asp?func=sendmail" method="post" name="sendmail" id="sendmail">
	<input type="hidden" name="kid" id="kid" value="<%=kid%>">
	<!--<input type="hidden" name="jnr" id="jnr" value="<=strjnr%>">-->
	<input type="hidden" name="jnavn" id="jnavn" value="<%=strjnavn%>">
	<%
	
	t = 0
	
	kids = split(kid, ",")
	for x = 0 to UBOUND(kids)
    if kids(x) <> 0 then
	
	strSQLk = "SELECT email, kkundenavn FROM kunder WHERE kid = " & kids(x)
	oRec3.open strSQLk, oConn, 3 
	
	if not oRec3.EOF then
	 	
		if instr(oRec3("email"), "@") <> 0 then
		t = t + 1%>
		<tr><td><input type="checkbox" name="FM_sendtil_kunde" id="FM_sendtil_kunde" value="1"></td><td><%=oRec3("kkundenavn") %></td><td><%=oRec3("email")%></td></tr> 
		<%end if
			
			strSQLm = "SELECT id, navn, email AS kpersemail FROM kontaktpers WHERE kundeid ="& kids(x)
			oRec.open strSQLm, oConn, 3 
			while not oRec.EOF 
		 	if instr(oRec("kpersemail"), "@") <> 0 then
			t = t + 1%>
			<tr><td><input type="checkbox" name="FM_sendtil_kpers" id="FM_sendtil_kpers" value="<%=oRec("id")%>"></td><td><%=oRec("navn") %></td><td><%=oRec("kpersemail")%></td></tr> 
			<%		
			end if				
			oRec.movenext	
			wend 
			oRec.close 
	end if
	oRec3.close 
	
	end if
	Next
	
	if t <> 0 then%>
	<tr><td colspan=2 align=center><br><input type="image" src="../ill/sendmailpil.gif"></td></tr>
	<%else%>
	<tr><td colspan=2><br><font color="red">Der findes ingen gyldige email adresser på den valgte kunde eller på dennes kontakt personer.</td></tr>
	<%end if%>
	</form>
	</table>
	</div>	
	<%
	end select
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
