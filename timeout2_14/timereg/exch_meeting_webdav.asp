
<%
function oprExchange(rcDomKonto, exch_appid, dato, sttid, sltid)
		
		
		' Exchange server name.
		strSQL3 = "SELECT owa, kontonavn, kontopw, dom FROM licens WHERE id = 1"
		oRec3.open strSQL3, oConn, 3 
		if not oRec3.EOF then
			ExchangeServerURL = oRec3("owa")
			ExchangeServerDOM = oRec3("dom")
			ExchangeServerBruger = replace(oRec3("kontonavn"), "#", "\")
			ExchangeServerPW = oRec3("kontopw")
		end if
		oRec3.close 
		
		
		
		
		
		if len(ExchangeServerURL) <> 0 AND len(ExchangeServerDOM) <> 0 AND len(ExchangeServerBruger) <> 0 AND len(ExchangeServerBruger) <> 0 then
		
		
		'**** Sætter de 2 FBA cookies ****
		Set req = Server.CreateObject("Msxml2.SERVERXMLHTTP") '("Microsoft.XMLhttp")
		servername = ExchangeServerURL '"mail.kringitsolutions.dk"
		mailbox = ExchangeServerBruger '"timeout"
		domain = ExchangeServerDOM' "KRING"
		strpassword = ExchangeServerPW '"bgt567ujmko0"
		strusername =  domain & "\" & mailbox '"KRING\timeout"
		szXml = "destination=https://" & servername & "/exchange&flags=0&username=" & strusername
		szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
		req.setOption 2, 13056
		req.Open "post", "https://" & servername & "/exchweb/bin/auth/owaauth.dll", False
		
		Response.write szXml
		Response.flush
		
		req.send szXml
		reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
		for i = lbound(reqhedrarry) to ubound(reqhedrarry)
			if instr(lcase(reqhedrarry(i)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(i),len(reqhedrarry(i))-12)
			if instr(lcase(reqhedrarry(i)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(i),len(reqhedrarry(i))-12)
		next
		'*************************************
		
		
		
		tidszoneTilpasningStart = dateadd("h", -1, ""& dato &" "& sttid &"") 
		tidszoneTilpasningSlut = dateadd("h", -1, ""& dato &" "& sltid &"") 
		
		'tidszoneTilpasningStart = dateadd("h", -2, "17-11-2005 14:00:00") 
		'tidszoneTilpasningSlut = dateadd("h", -2, "17-11-2005 15:00:00") 
		
		tzSt = tidszoneTilpasningStart
		tzSl = tidszoneTilpasningSlut
		
		
		'*** Dage funktioner ***
		if day(tzSt) < 10 then
		tzStDag = "0"&day(tzSt)
		else
		tzStDag = day(tzSt)
		end if
		
		if day(tzSl) < 10 then
		tzSlDag = "0"&day(tzSl)
		else
		tzSlDag = day(tzSl)
		end if
		
		'*** Måned funktioner ***
		if month(tzSt) < 10 then
		tzStMonth = "0"&month(tzSt)
		else
		tzStMonth = month(tzSt)
		end if
		
		if month(tzSl) < 10 then
		tzSlMonth = "0"&month(tzSl)
		else
		tzSlMonth = month(tzSl)
		end if
		
		strTilPasStart = year(tzSt)&"-"&tzStMonth&"-"&tzStDag&"T"&right(tidszoneTilpasningStart, 8)&".000Z"
		strTilPasSlut = year(tzSl)&"-"&tzSlMonth&"-"&tzSlDag&"T"&right(tidszoneTilpasningSlut, 8)&".000Z"
		
		'Response.write strTilPasStart
		
		t = 100
		if t = 100 then
		
		
		'Option Explicit
		'Private Sub CreateAppointment()
		
		
		 ' Variables
		 'Dim strApptURL As String
		 'Dim strExchSvrName  As String
		 'Dim strMailbox As String
		 'Dim strApptItem As String
		 'Dim strUserName As String
		 'Dim strPassWord As String
		 'Dim strApptRequest As String
		 'Dim strCalInfo As String
		 'Dim strHeaderInfo As String
		 'Dim strMailInfo As String
		 'Dim strXMLNSInfo As String
		
		 ' Use MSXML 4.0
		 'Dim xmlReq As MSXML2.XMLHTTP40
		
		
		 
		strExchSvrName = ExchangeServerURL  '"outzource-747" '"localhost/exchange" '"www.danexchange.dk" '"Exchangeservername"
		'Response.write strExchSvrName & "<br>"
		
		 ' Mailbox folder name.
		 'strMailbox = ExchangeServerBruger
		'strMailbox = "timeout@kringitsolutions.dk" '"administrator" '"administrator@arusha.outzource.dk" '"user"
		strMailbox = rcDomKonto '"Timeout"
		
		'Response.write "Email modtager:"&  strMailbox & "<br>"
		'Response.flush
		 
		 ' Appointment item.
		 '* Skal ændres hver gang ***
		 strApptItem = exch_appid &".eml" '"timeout_"& now &".eml" 
		
		 ' Username and password of appointment creator.
		 strUserName = ExchangeServerDOM &"\"& ExchangeServerBruger '"arusha\administrator"
		 strPassWord = ExchangeServerPW 
		 
		 'strUserName = "arusha\administrator"
		 'strPassWord = "Sok!2637"
		 
		 'Response.write strUserName & ", " & strPassWord &"<br>"
		 'Response.flush
		 
		
		 ' URL of the appointment item. ' "http://" & '& "/exchange/"
		 strApptURL = "https://"& strExchSvrName  &"/Exchange/"& strMailbox & "/Calendar/" & strApptItem
		 'strApptURL = "http://"& strExchSvrName &"/ews/Exchange.asmx"
		 'strApptURL = strExchSvrName  &"/"& strMailbox & "/"
		 Response.write "URL:<br>"& strApptURL & "<br><br>"
		 'Response.flush
		
		 ' XML namespace info of the WebDAV request.
		 strXMLNSInfo = "xmlns:g=""DAV:"" " & _
		    "xmlns:e=""http://schemas.microsoft.com/exchange/"" " & _
		    "xmlns:mapi=""http://schemas.microsoft.com/mapi/"" " & _
		    "xmlns:mapit=""http://schemas.microsoft.com/mapi/proptag/"" " & _
		    "xmlns:x=""xml:"" xmlns:cal=""urn:schemas:calendar:"" " & _
		    "xmlns:dt=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"" " & _
			"xmlns:header=""urn:schemas:mailheader:"" " & _
		    "xmlns:mail=""urn:schemas:httpmail:"">"
		
		 ' Set the appointment item properties.  The reminder time is set in seconds.
		 ' To create an all-day meeting, set the dtstart/dtend range for 24 hours
		 ' or more and set the alldayevent property to 1.  See the documentation
		 ' on the properties in the urn:schemas:calendar: namespace for more information.
		 strCalInfo = "<cal:location>"& kundenavnogadr &"</cal:location>" & _
		    "<cal:dtstart dt:dt=""dateTime.tz"">"& strTilPasStart &"</cal:dtstart>" & _
		    "<cal:dtend dt:dt=""dateTime.tz"">"& strTilPasSlut &"</cal:dtend>" & _
		    "<cal:instancetype dt:dt=""int"">0</cal:instancetype>" & _
		    "<cal:busystatus>BUSY</cal:busystatus>" & _
		    "<cal:meetingstatus>CONFIRMED</cal:meetingstatus>" & _
		    "<cal:alldayevent dt:dt=""boolean"">0</cal:alldayevent>" & _
			"<cal:responserequested dt:dt=""boolean"">1</cal:responserequested>" & _
		    "<cal:reminderoffset dt:dt=""int"">900</cal:reminderoffset>"& _
			"<cal:timezoneid dt:dt=""int"">15</cal:timezoneid>"
			'  & _
			'"<mapi:0x00008223 dt:dt=""boolean"">1</mapi:0x00008223>"   'Recorrence
		
		 ' Set the required attendee of the appointment.
		  strHeaderInfo = "<header:to>" & strMailbox & "</header:to>"
		
		 ' Set the subject of the appointment.
		 strMailInfo = "<mail:subject>TimeOut Booking: kl. " & sttid & " - "& sltid &"</mail:subject>" & _
		    "<mail:htmldescription>"_
			& "Kontaktperson: " & kpers &""& vbCrLf _  
			& "Email: " & kpers_email &""& vbCrLf _
			& "Dir. Tlf: "& kpers_dirtlf &""& vbCrLf _
			& "Mobil Tlf: "& kpers_mobiltlf &  vbCrLf & vbCrLf _ 
			& "Job: "& jobnavnognr &"" & vbCrLf _ 
			& "Beskrivelse: "& jobbesk &"" & vbCrLf & vbCrLf _ 
			& "</mail:htmldescription>"
		
		 ' Build the XML body of the PROPPATCH request.
		 strApptRequest = "<?xml version=""1.0""?>" & _
		  "<g:propertyupdate " & strXMLNSInfo & _
		  "<g:set><g:prop>" & _
		    "<g:contentclass>urn:content-classes:appointment</g:contentclass>" & _
		    "<e:outlookmessageclass>IPM.Appointment</e:outlookmessageclass>" & _
		    strMailInfo & _
		    strCalInfo & _
		    strHeaderInfo & _
		    "<mapi:finvited dt:dt=""boolean"">1</mapi:finvited>" & _
		   "</g:prop></g:set>" & _
		 "</g:propertyupdate>"
		
		 ' Create the DAV PROPPATCH request.
		 Set xmlReq = server.CreateObject("Msxml2.SERVERXMLHTTP") '("Microsoft.XMLHTTP")
		 xmlReq.open "PROPPATCH", strApptURL, False, strUserName, strPassWord
		
		 '*** De 2 FBA cookies ****'
		 xmlReq.SetRequestHeader "cookie", reqsessionID
		 xmlReq.SetRequestHeader "cookie", reqCadata   
		 
		 xmlReq.setRequestHeader "Content-Type", "text/xml"
		 xmlReq.send strApptRequest
		
					 '' The PROPPATCH request was successful.
					 If (xmlReq.Status >= 200 And xmlReq.Status < 300) Then
					 Response.write "Appointment created successfully."
					
					 '' Display error info.
					 Else
					
					 Response.write "PROPPATCH status: " & xmlReq.Status & vbCrLf & _
					 "Status text: " & xmlReq.statusText
					 End If
		
		' Clean up.
		Set xmlReq = Nothing
		Set req = Nothing
		'Exit Sub
		
		'' Implement custom error handling here.
		'ErrHandler:
		   ' ' Display error info.
		   'Response.write "Error creating appointment/meeting: " & Err.Number & ": " & Err.Description
	
	    ' Clean up.
	    'Set xmlReq = Nothing
	
	'End Sub
	
	end if '** t = 100
	end if '*** Exchangeoplysninger i Licens


end function
%>




