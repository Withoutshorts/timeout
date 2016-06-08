
<%
'Response.write session.sessionId

Response.buffer = True 


 
'** Bruges også i 
'** timereg.asp
'** ressourcer.asp 
'** til response.flush

if thisfile <> "printversion.asp" then

if len(request("key")) <> "0" then
session("lto") = request("key")
strLicenskey = request("key")

else
strLicenskey = session("lto") 'Request.Cookies("licenskey")
end if

else

strLicenskey = session("lto") 'Request.Cookies("licenskey")
 
end if


'* finder load tid ***
timeA = now
	

'ODBC
'Dim objXMLHTTP, objXMLDOM, i
'Opretter en instans af Microsoft.XMLHTTP, så det er muligt at få fat på dokumentet
'Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
'Opretter en instans af Microsofts XML-parser, XMLDOM
'Set objXMLDOM = Server.CreateObject("Microsoft.XMLDOM")





'Response.Write " session(lto) :" & session("lto")
'Response.Write " strLicenskey" & strLicenskey



if len(strLicenskey) <> 0 then
	select case strLicenskey
	case "2.052-xxxx-B000" 'demo
		strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_demo;"
		'strConnThis = "driver={MySQL ODBC 5.00 Driver};server=192.168.23.45; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
		addnewclient = 1
		lto = "demo"
	case "2.152-0416-B001", "2.152-1604-B001" 'outzource
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
		lto = "outz"
        strConnThis = "timeout_intranet64"
        'response.redirect "http://timeout2.outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.152-0416-B001&lto=outz"


        
    
    'ODBC 3.51 Driver
	
	case else 'ktv udvikling (hvis der ikke er en cookie)
		
         'strConnThis = "timeout_intranet_351_30_64"
        'strConnThis = "timeout_intranet32" 
        lto = "intranet - local"

        'strConnThis = "mySQL_timeOut_intranet"
        'strConnThis = "timeout_wwf"
		'lto = "intranet - local"
        'lto = "tec"
        'strConnThis = "timeout_intranet64"
        strConnThis = "timeout_nt"
        'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet;"
	    'response.write strConnThis
        'response.flush
	
       'lto = "outz"
	   'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=192.168.1.35; Port=3306; uid=root;pwd=;database=timeout_intranet;"
	
    
    	
        'lto = "kejd_pb"
        'Ny IP Pr 5.4.2014
        '195.189.130.210
		'NY IP: 62.182.173.226
		''lto = "epi"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"
	    
	    'lto = "kejd_pb"
		'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kejd_pb;"
	
		
		
	end select 



   


else
	if instr(request.servervariables("HTTP_HOST"), "localhost") <> 0 then
	
	'Zepto
	'strConnThis = "mySQL_timeOut_intranet"
    'strConnThis = "timeout_intranet_351_30_64"
    'strConnThis = "timeout_intranet32"
    lto = "intranet - local"
    'lto = "tec"
	'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet;"
	' strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet; OPTION=32"
	   'response.write strConnThis
       ' response.flush

    'strConnThis = "timeout_intranet64"
    strConnThis = "timeout_nt"


     
    ' lto = "dencker"
    ' strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"
	

    'lto = "kejd_pb"

    'NY IP: 62.182.173.226
    'lto = "cst"
    'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=81.19.249.35; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cst;"
	 
    'lto = "wwf"
    'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wwf_BAK;"

     'lto = "dencker"
     'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"

     'lto = "tec"
	 'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_tec;"
	   

	else
	
    call showError(5)
	response.end
	
	'Response.redirect "../login_fejlet.asp"
	
	'Response.write "Sessionen er udløbet"
	'Response.end
	end if
end if

'objXMLHTTP.Send

'Lægger indholdet af dokumentet over i vores XML-parser objekt
'Set objXMLDOM = objXMLHTTP.ResponseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("modul")

'Henter indholdet af alle tags med navnet 'url'
'Set objDBConn = objXMLDOM.getElementsByTagName("dbconn")

strConn = strConnThis	'objDBConn(0).text  'Connection


strA1 = "-" 	'objModuler(0).text   'Timeregistrering
strA2 = "-"		'objModuler(1).text   'Medarbejder log



'*** ObjModuler(2).text   'Skift medarbejder 
select case lto
case "kringit", "outz", "worldiq"', "intranet"
strA3 = "on"
case else
strA3 = "off"
end select 		


strA4 = "on"		'objModuler(3).text   'Se timereg for andre medarb.
strB1 = "-"   	'Joboversigt og Aktiviter. Mulighed for at lukke akt. 
strB2 = "-"		'objModuler(5).text   'Projektgrupper
strB3 = "-"		'objModuler(6).text   'Stamaktiviteter
strB4 = "-"		'objModuler(7).text   'Jobplanner
strB5 = "-"		'objModuler(8).text   'Km regnskab
strB6 = "-"		'objModuler(9).text   'Udgifter paa job
strB7 = "-"		'objModuler(10).text   'Ikke planlagt
strB8 = "-"		'objModuler(11).text   'Upload af dokumenter
strB9 = "-"		'objModuler(12).text   'Incident tool
strB10 = "-"	'objModuler(13).text	'Aktivitetsskabeloner

strC1 = "-" 	'objModuler(14).text  'Kunder + filter og sog paa kunde
strC2 = "on" 	'objModuler(15).text  'Kunde login
strC3 = "-"		'objModuler(16).text   'Upload logoer paa kunder

strD1 = "-" 	'objModuler(17).text  'Medarbejdere
strD2 = "-" 	'objModuler(18).text  'Brugergrupper
strD3 = "-"		'objModuler(19).text  'Medarbejdertyper

strE1 = "-"		'objModuler(20).text  'Statistik og Fakturering + filter
strE2 = "-"		'objModuler(21).text  'Joblog_z_b, joblog, top 5 job
strE3 = "-"		'objModuler(22).text  'Aars omsaetning
strE4 = "-"		'objModuler(23).text  'Export
strE5 = "on"		'objModuler(24).text	 'Faktura oversigt

strH1 = "-"		'objModuler(25).text  'Kalender
strH2 = "-"		'objModuler(26).text  'Firmaer
strH3 = "-"		'objModuler(27).text	 'Aktions historik + sogefunktion
strH4 = "-"		'objModuler(28).text  'Status
strH5 = "-"		'objModuler(29).text  'Emner
strH6 = "-"		'objModuler(30).text  'Kontaktform
strH7 = "-"		'objModuler(31).text  'Outlook integration

'Set objXMLHTTP = Nothing
'Set objXMLDOM = Nothing

'if lto = "outz" then
'Response.Write "lto" & lto & "<br>"
'Response.Write "conn:" & strConn
'Response.flush
'end if

if len(strConn) <> 0 then
	strConnect_DBConn = strConn
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRec = Server.CreateObject ("ADODB.Recordset")
	Set oRec2 = Server.CreateObject ("ADODB.Recordset")
	Set oRec3 = Server.CreateObject ("ADODB.Recordset")
	Set oRec4 = Server.CreateObject ("ADODB.Recordset")
	Set oRec5 = Server.CreateObject ("ADODB.Recordset")
	Set oRec6 = Server.CreateObject ("ADODB.Recordset")
	Set oRec7 = Server.CreateObject ("ADODB.Recordset")
    Set oRec8 = Server.CreateObject ("ADODB.Recordset")
    Set oCmd = Server.CreateObject ("ADODB.Command")
	
	oConn.open strConnect_DBConn
	
	
	sub closeDB
	'now close and clean up
			oRec.Close
			oConn.close
			
			Set oRec = Nothing
			set oConn = Nothing
	
	end sub




                if thisfile <> "printversion.asp" then

                 '** Hvis der er skiftet version ****
                '** Finder bruger i ny version ****'
                if request("eksterntlnk") = "aaQWEIOC345345DFNEfjsdf7890sdfv" then
   

                        if len(trim(request("email"))) <> 0 then
                        emailThis = request("email")
                        else
                        emailThis = "99999999"
                        end if


                        'session.abandon
                        'response.flush

                        session("user") = ""
			            session("login") = ""
			            session("mid") = ""
			            session("rettigheder") = ""


                        newUSerId = 0
                        strSQLnewuser = "SELECT mid, mnavn, m.brugergruppe, b.rettigheder FROM medarbejdere AS m "_
                        &" LEFT JOIN brugergrupper AS b ON (b.id = m.brugergruppe) WHERE email = '"& emailThis & "'"
    
                        'Response.write strSQLnewuser
                        'Response.flush

                        oRec4.open strSQLnewuser, oConn, 3
                        if not oRec4.EOF then


        


                        newUSerId = oRec4("mid")

                        session("user") = trim(oRec4("mnavn"))
			            session("login") = newUSerId
			            session("mid") = newUSerId
			            session("rettigheder") = oRec4("rettigheder")
            

                        end if 
                        oRec4.close

                usemrn = newUSerId
   

                if newUSerId <> 0 then


                Response.write "<div style=""top:40px; left:40px,"">Du er godkendt til at skifte version, <a href=""timereg_akt_2006.asp"">klik her</a> for at komme videre...</div>"

                Response.end


                else

                'Response.write "her"
                'Response.end
                'Response.redirect "../login_fejlet.asp"
    
                call showError(5)
                response.end


                end if
          	
    

                end if
                end if 'printversion
   

end if

%>
