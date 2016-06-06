
<%


   


 Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
    	   ' Sætter Charsettet til ISO-8859-1 
	Mailer.CharSet = 2
    Mailer.FromName = "TimeOut"
	Mailer.FromAddress = "timeout_no_reply@outzource.dk"
				
                    'if instr(request.servervariables("HTTP_HOST"), "timeout2.outzource.dk") <> 0 then
                    Mailer.RemoteHost = "formrelay.rackhosting.com" '"mailrelay3.rackhosting.com" 
                    'else
					'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" RACK: mailrelay3.rackhosting.com
                    'end if
	

    Mailer.AddRecipient "Support", "support@outzource.dk" 
	Mailer.AddBCC "Support", "support@outzource.dk" 
	Mailer.Subject = "TimeOut - Medarbejder profil"

   	Mailer.BodyText = ""& time &"<br> Bla. Bla."

 If Mailer.SendMail Then
        	
    Response.Write "Ok...<br>" & Mailer.Response

    Else
   Response.Write "Fejl...<br>" & Mailer.Response
 End if

     Set Mailer = nothing

     Response.write "<br><br>URL: " & request.servervariables("HTTP_HOST") & "<br><br>"
%>