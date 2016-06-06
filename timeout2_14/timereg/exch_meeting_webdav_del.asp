
<%
function delExchange(rcDomKonto, exch_appid)

' Exchange server name.
strSQL2 = "SELECT owa, kontonavn, kontopw, dom FROM licens WHERE id = 1"
oRec2.open strSQL2, oConn, 3 
if not oRec2.EOF then
	ExchangeServerURL = oRec2("owa")
	ExchangeServerDOM = oRec2("dom")
	ExchangeServerBruger = replace(oRec2("kontonavn"), "#", "\")
	ExchangeServerPW = oRec2("kontopw")
end if
oRec2.close 


if len(exch_appid) <> 0 AND len(ExchangeServerURL) <> 0 AND len(ExchangeServerDOM) <> 0 AND len(ExchangeServerBruger) <> 0 AND len(ExchangeServerBruger) <> 0 then


'**** Sætter de 2 FBA cookies ****
Set req = Server.CreateObject("Microsoft.XMLhttp")
servername = ExchangeServerURL '"mail.kringitsolutions.dk"
mailbox = ExchangeServerBruger '"timeout"
domain = ExchangeServerDOM' "KRING"
strpassword = ExchangeServerPW '"bgt567ujmko0"
strusername =  domain & "\" & mailbox
szXml = "destination=https://" & servername & "/exchange&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
req.Open "post", "https://" & servername & "/exchweb/bin/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for i = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(i)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(i),len(reqhedrarry(i))-12)
	if instr(lcase(reqhedrarry(i)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(i),len(reqhedrarry(i))-12)
next
'*************************************



 ' Username and password of appointment creator.
 strUserName = ExchangeServerDOM &"\"& ExchangeServerBruger  '"arusha\administrator"
 strPassWord = ExchangeServerPW 
 
 strExchSvrName = ExchangeServerURL '"outzource-747" '"localhost/exchange" '"www.danexchange.dk" '"Exchangeservername"

 ' Mailbox folder name.
 strMailbox =  rcDomKonto '"administrator" '"administrator@arusha.outzource.dk" '"user"

 ' Appointment item.
 if len(exch_appid) <> 0 then '*** Hvis tom slettes hele kalenderen
 exch_appid = exch_appid
 else
 exch_appid = "xxxx"
 end if
 strApptItem = exch_appid &".eml" '"testappointment10.eml" '* Skal ændres hver gang ***

 
' URL of the appointment item.
 strApptURL = "https://" & strExchSvrName & "/exchange/" & _
 strMailbox & "/Calendar/" & strApptItem

'// Initialize the XMLHTTP request object.
 Set xmlReq = server.CreateObject("Microsoft.XMLHTTP")
 xmlReq.open "DELETE", strApptURL, false, strUserName, strPassWord
 
 '*** De 2 FBA cookies ****
 xmlReq.SetRequestHeader "cookie", reqsessionID
 xmlReq.SetRequestHeader "cookie", reqCadata  
 
 xmlReq.setRequestHeader "Content-Type", "text/xml"
 xmlReq.send()



end if '*** Exchangeoplysninger i Licens

end function
%>


