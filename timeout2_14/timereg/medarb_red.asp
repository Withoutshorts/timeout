<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/medarb_func.asp"-->
<!--#include file="inc/smiley_inc.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
					sub medarb_opr_ok
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					


	            
	
					<!-------------------------------Sideindhold------------------------------------->
					<div id="sindhold" style="position:absolute; left:90px; top:102px;">
					

                    <%
	
	oimg = "ikon_medarb_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Medarbejdere"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	%>

					<table cellspacing="0" cellpadding="0" border="0" width="400" bgcolor="#FFFFFF">
					<tr bgcolor="#5582D2">
					<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
					<td valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
					</tr>
                    <form method="post" action="medarb.asp?menu=medarb&lastmedid=<%=lastmedid %>">
					<tr bgcolor="#5582D2">
						<td class=alt><b>Medarbejder info:</b></td>
					</tr>
					<tr>
						<td style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
						<td><br>
                        <h4>Tillykke <img src="../ill/check2.png" /></h4>
						Du har netop oprettet <b><%=strNavn%></b> i TimeOut.<br><br>
						Der er afsendt en email til medarbejderen med login og password oplysninger.<br>
						Emailen bør være fremme indenfor et par minutter.<br>
						<br><br>
						<img src="../ill/ac0001-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Timepriser på stam-aktiviteter:</b><br>
						Systemet har oprettet timepriser for <b><%=strNavn%></b> på alle de eksisterende stamaktiviteter.<br>
						<br>Timepriserne er baseret på de allerede gældende timerpriser, på hver enkelt stamaktivitet, for andre medarbejdere af samme <b>medarbejdertype</b>.<br><br />
						
                        <!--<br>Hvis du ønsker at gennemse alle stam-aktiviteter for at se om medarbejderen er oprettet med korrekt timepris, 
						<a href="akt_gruppe.asp?menu=job&func=favorit">gennemse stam-aktiviteterne her.</a><br><br>
                        -->
                        <br />
						<input type="submit" value="Fortsæt >>" /> 
						<br><br>&nbsp;</td>
						<td style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					</tr>
                    </form>
					<tr>
						<td height=20 style="border-right:1px #003399 solid; border-left:1px #003399 solid; border-bottom:1px #003399 solid;" colspan=3><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					</tr>
					</table>
					<%
					end sub
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	
	<%
	
	slttxt = "Du er ved at <b>slette</b> en medarbejder. <br>"_
	&"Du vil samtidig slette alle timeregistreringer, projektgruppe relationer mm. "_
	&"for denne medarbejder.<br>"_
    &"Data vil <b>ikke kunne genskabes</b>.<br>"_
	&"<br>Er du sikker på at du vil slette denne medarbejder?<br><br>"
	
	slturl = "medarb_red.asp?menu=medarb&func=sletok&id="&id
		
	slturlalt = ""
	slttxtalt = ""
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	
	case "sletok"
	'*** Her slettes en medarbejder ***
	
	strSQL = "SELECT mid, mnavn, mnr FROM medarbejdere WHERE mid = "& id &"" 
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
		
		'*** Indsætter i delete historik ****'
	    call insertDelhist("med", oRec("mid"), oRec("mnr"), oRec("mnavn"), session("mid"), session("user"))
		
	oRec.movenext
	wend
	oRec.close
	
	
	oConn.execute("DELETE FROM budget_medarb_rel WHERE medid = "& id &"")
	oConn.execute("DELETE FROM medarbejdere WHERE Mid = "& id &"")
	oConn.execute("DELETE FROM progrupperelationer WHERE MedarbejderId = "& id  &"") 'projektgruppeId = 10 AND
	oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& id  &"") 
    oConn.execute("DELETE FROM timer WHERE tmnr = "& id &"")
	oConn.execute("DELETE FROM ressourcer WHERE mid = "& id &"")
	oConn.execute("DELETE FROM ressourcer_md WHERE medid = "& id &"")
	oConn.execute("DELETE FROM timepriser WHERE medarbid = "& id &"")
	
   

  

    if level <> 1 then	
	Response.redirect "medarb.asp?menu=medarb"
    else
    Response.redirect "timereg_akt_2006.asp"
    end if
	
	case "dbred", "dbopr" 
	'*** Her tjekkes om alle required felter er udfyldt. ***
	if len(Request("FM_login")) = 0 OR len(Request("FM_pw")) = 0 OR len(Request("FM_mnr")) = 0 then
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%
	errortype = 9
	call showError(errortype)
	
	else
	 
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(Request("FM_mnr"))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		
			<%
			errortype = 22
			call showError(errortype)
			
			isInt = 0
			else
	        
	        
	        if len(request("ansatdato")) <> 0 AND isDate(request("ansatdato")) then
			ansatdato = request("ansatdato")
			ansatdato = year(ansatdato)&"/"&month(ansatdato)&"/"&day(ansatdato)
			else
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
			<%
			errortype = 114
			call showError(errortype)
			Response.end
			end if	
			
			if len(request("opsagtdato")) <> 0 AND isDate(request("opsagtdato")) then
			opsagtdato = request("opsagtdato")
			opsagtdato = year(opsagtdato)&"/"&month(opsagtdato)&"/"&day(opsagtdato)
			else
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
			<% 
			errortype = 115
			call showError(errortype)
			Response.end
			end if	
			
			
		
		
		
	    
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
			strNavn = SQLBless(Request("FM_navn"))
			strAnsat = Request("FM_ansat")
			
			strInit = SQLBless(request("FM_init"))
			
			if len(request("FM_timereg")) <> 0 then
			intTimereg = request("FM_timereg") 
			else
			intTimereg = 0
			end if
			
			
			
			strLogin = Request("FM_login")
			strPw = Request("FM_pw")
			strMnr = Request("FM_mnr")
			strEditor = session("user")
			strDato = session("dato")
			strMedinfo = SQLBless(Request("FM_medinfo"))
			strBrugergruppe = Request("FM_bgruppe")
			strMedarbejdertype = Request("FM_medtype")
			strEmail = Request("FM_email")
			
			madr = replace(request("FM_adr"), "'", "''")
			mpostnr = replace(request("FM_postnr"), "'", "''")
			mcity = replace(request("FM_city"), "'", "''")
			mland = request("FM_land")
			mtlf = replace(request("FM_tel"), "'", "''")
			mcpr = replace(request("FM_cpr"), "'", "''")
			mkoregnr = replace(request("FM_regnr"), "'", "''")
			
			'*** Kun admin brugere skal sættes til at modtage nyhedsbrev ved opret. ***'
			if func = "dbred" then
			nyhedsbrev = request("nyhedsbrev")
			else
			    if cint(strBrugergruppe) = 3 then
			    nyhedsbrev = 1
			    else
			    nyhedsbrev = 0
			    end if
			end if
			
			dagsdato = year(now) &"/"& month(now) &"/"& day(now)
			
			if len(request("FM_tsacrm")) <> 0 then 
			inttsacrm = request("FM_tsacrm")
			else
			inttsacrm = 0
			end if
			
			if len(request("FM_smiley")) <> 0 then
			intSmil = request("FM_smiley")
			else
			intSmil = 0
			end if
			
			strExch = replace(request("FM_exch"), "\", "#")
			
			sprog = request("FM_sprog")

            if len(trim(request("FM_visskiftversion"))) <> 0 then
            visskiftversion = request("FM_visskiftversion")
            else
            visskiftversion = 0
            end if
			
            '** Vis timer som start / stop tid
            if len(trim(request("FM_timer_ststop"))) <> 0 then
            timer_ststop = 1
            else
            timer_ststop = 0
            end if

                
			
			'*** Her tjekkes for dubletter i db ***
			strSQL = "SELECT Mnr FROM medarbejdere WHERE Mnr ="& strMnr &" AND Mid <> "& id
			oRec.open strSQL, oConn, 3	
			if oRec.EOF then
			strMnrOK = "y"
			end if
			oRec.close
			
			strSQL = "SELECT login FROM medarbejdere WHERE login ='"& strLogin &"' AND Mid <> "& id
			oRec.open strSQL, oConn, 3	
			if oRec.EOF then
			strLoginOK = "y"
			end if
			oRec.close
			
			'Checker for dubletter i DB før record opdateres eller indsættes
			if strMnrOK = "y" AND strLoginOK = "y" then
			        

                '** Finder medarbejdertypegrp (hvis findes) **'
                    '** Gemmes på medarbejder pga perfomance på statistikker
                    intMedarbejdertype_grp = 0
                    strSQlmtyp = "SELECT mgruppe FROM medarbejdertyper WHERE id = "& strMedarbejdertype
                    oRec2.open strSQlmtyp, oConn, 3
                    if not oRec2.EOF then
                    intMedarbejdertype_grp = oRec2("mgruppe")
                    end if
                    oRec2.close
			        '****'
			        
			        
			        
			        if func = "dbred" then
			        
			        
			        '*** Medarbejdertype historik ***'
			        nyMedarbType = 0
			        strSQLmth = "SELECT medarbejdertype FROM medarbejdere WHERE mid = "& id 
			        oRec.open strSQLmth, oConn, 3
			        if not oRec.EOF then
			        
			            if cint(oRec("medarbejdertype")) <> cint(strMedarbejdertype) then
			            
			                    
			                    '*** Er det første ændring? ***'
			                    
			                    typeHist = 0
			                    strSQL1 = " SELECT mid FROM medarbejdertyper_historik WHERE mid = "& id
			                    oRec2.open strSQL1, oConn, 3
			                    if not oRec2.EOF then
			                        
			                        typeHist = 1
			                    
			                    end if
			                    oRec2.close
			                    
			                    
			                    '*** Opretter den forladte type ved første sktift,  **'
			                    '*** så man ved hvilken type medarbejder            **'
			                    '*** har været fra ansættelses dato                 **'
			                    '*** til første skift.                              **'
			                    
			                    if typeHist = 0 then
			                    
			                    strSQLmthupd1 = "INSERT INTO medarbejdertyper_historik (mid, mtype, mtypedato) "_
			                    &" VALUES ("& id &", "& oRec("medarbejdertype") &", '"& ansatdato &"') "
			                    oConn.execute(strSQLmthupd1)
			                    
			                    
			                    end if
			                    
			                    
			            
			            strSQLmthupd3 = "INSERT INTO medarbejdertyper_historik (mid, mtype, mtypedato) "_
			            &" VALUES ("& id &", "& strMedarbejdertype &", '"& dagsdato &"') "
			            
			            oConn.execute(strSQLmthupd3)
			            
			            end if

                    nyMedarbType = oRec("medarbejdertype")
			        
			        end if
			        oRec.close
			        
			        
			        
			                   
			        
			        

                    '** Opdaterer datoer på eksisterende ****'
                    mtyphist_ids = split(request("FM_mtyphist_id"), ", ")
                    mtyphist_datoer = split(request("FM_mtyphist_dato"), ", ")

                        
                    useOpdStDatoTp = year(now) & "-" & month(now) & "-" & day(now)
                    useOpdMtypeGrp = 0
                    aktiveI = 0

                    for i = 0 to UBOUND(mtyphist_ids)
                                    

                                    if isDate(mtyphist_datoer(i)) = true then

                                    mtyphist_datoer(i) = year(mtyphist_datoer(i)) & "-"& month(mtyphist_datoer(i)) & "-" & day(mtyphist_datoer(i))
                                    strSQLmthupd2 = "UPDATE medarbejdertyper_historik SET mtypedato = '"& mtyphist_datoer(i) &"' WHERE id = "& mtyphist_ids(i)
                                    'Response.write strSQLmthupd2
                                    'Response.end
			                        oConn.execute(strSQLmthupd2)

                                    end if
                        
                                aktiveI = i
                                
                             

                    next
                            
                            
                                  
                                
                                       
                
                                '*** Kost og timepriser opdatering ***********'
                                '*** SKal eksisterede kostprsier opdateres ***'
                                if len(trim(request("FM_opd_mty_kp"))) <> 0 then

                                useOpdStDatoTp = mtyphist_datoer(aktiveI)
                                useOpdMtypeGrp = nyMedarbType
                        
                                nyKost = 0
                                nyKurs = 100
                                nyValuta = 1

                                strSQLsmtyKp = "SELECT m.kostpris FROM medarbejdertyper AS m WHERE m.id = "& useOpdMtypeGrp
                            
                                'response.write "strSQLsmtyKp: "& strSQLsmtyKp
                                'response.flush 

                                oRec6.open strSQLsmtyKp, oConn, 3 
                                if not oRec6.EOF then
                            
                                nyKost = replace(oRec6("kostpris"), ",", ".") 
                                        
                                end if
                                oRec6.close

                                strSQLupdKp = "UPDATE timer SET kostpris = "& nyKost &" WHERE tmnr = "& id &" AND tDato >= '"& useOpdStDatoTp &"'"

                                'response.write "<br>strSQLupdKp: "& strSQLupdKp
                                'response.end

                                oConn.execute(strSQLupdKp)        

                                end if 
                
                
                                '*** SKal eksisterede timepriser opdateres ***'
                                if len(trim(request("FM_opd_mty_tp"))) <> 0 then
                                
                                useOpdStDatoTp = mtyphist_datoer(aktiveI)
                                useOpdMtypeGrp = nyMedarbType


                                nyTp = 0
                                nyKurs = 100
                                nyValuta = 1

                                strSQLsmtyKp = "SELECT m.timepris, m.tp0_valuta, v.kurs FROM medarbejdertyper AS m "_
                                &"LEFT JOIN valutaer AS v ON (v.id = tp0_valuta) WHERE m.id = "& useOpdMtypeGrp
                            
                                'response.write "strSQLsmtyKp: "& strSQLsmtyKp
                                'response.flush 

                                oRec6.open strSQLsmtyKp, oConn, 3 
                                if not oRec6.EOF then
                            
                                nyTp = replace(oRec6("timepris"), ",", ".") 
                                nyValuta = oRec6("tp0_valuta") 
                                nyKurs = replace(oRec6("kurs"), ",", ".")           

                                end if
                                oRec6.close

                                strSQLupdTp = "UPDATE timer SET timepris = "& nyTp &", valuta = "& nyValuta &", kurs = "& nyKurs &" WHERE tmnr = "& id &" AND tDato >= '"& useOpdStDatoTp &"'"
                            
                                'if session("mid") = 1 then
                                'response.write "<br>strSQLupdTp: "& strSQLupdTp
                                'response.end
                                'end if

                                oConn.execute(strSQLupdTp)      

                                
                
                                end if            
                                '**********************************************************************************
                                



                       
                      'if session("mid") = 1 then
                      'response.end
                      'end if









                    '*** dato skal altid følge ansatdato på første type **'
			        strSQLmthupd1 = "SELECT id FROM medarbejdertyper_historik WHERE mid =  "& id &" ORDER BY id"
			                    
			        oRec2.open strSQLmthupd1, oConn, 3
			        if not oRec2.EOF then
			                        
			                strSQLmthupd2 = "UPDATE medarbejdertyper_historik SET mtypedato = '"& ansatdato &"' WHERE id = "& oRec2("id")
			            oConn.execute(strSQLmthupd2)
                                    
			                    
			        end if
			        oRec2.close

                                

                               
                            

                    
			        
			        
                    '*** Hvis medarbejder DE-aktiveres Sætter opsigelsesdato til dd. PGHA normtids beregning ***'
                    'Response.write "strAnsat" & strAnsat
                    if strAnsat = "2" then
                    
                    'Response.write "her"
                    useToBeAnsat = "0"
                    strSQLtjk = "SELECT mansat FROM medarbejdere WHERE mid = "& id
                    oRec5.open strSQLtjk, oConn, 3
                    if not oRec5.EOF then

                    useToBeAnsat = oRec5("mansat")
                    'Response.write "<br>useToBeAnsat"& useToBeAnsat 

                    end if
                    oRec5.close
                    
                    if useToBeAnsat <> "2" then
                    opsagtdato = year(now)&"/"&month(now) &"/"& day(now)
                    end if
                    
                    
                    
                    end if
                    '*** end 

                   
                    'Response.end
                    


					
					strSQL = "UPDATE medarbejdere SET"_
					&" Mnavn = '"& strNavn &"',"_
					&" Mnr = "& strMnr &","_
					&" Mansat = '"& strAnsat &"',"_
					&" login = '"& strLogin &"',"
					
					if strPw <> "KEEPTHISPW99" then
					strSQL = strSQL &" pw = MD5('"& strPw &"'),"
					end if
					
					strSQL = strSQL &" brugergruppe = "& strBrugergruppe &","_
					&" Medarbejdertype = "& strMedarbejdertype &","_
					&" editor = '"& strEditor &"',"_
					&" dato = '"& strDato &"',"_
					&" Medarbejderinfo = '"& strMedinfo &"', "_
					&" Email = '"& strEmail &"',"_
					&" tsacrm = "& inttsacrm &", smilord = "& intSmil &", exchkonto = '"& strExch &"', "_
					&" init = '"& strInit &"', timereg = "& intTimereg &","_
					&" ansatdato = '"& ansatdato &"', "_
					&" opsagtdato = '"& opsagtdato &"', sprog = "& sprog &", nyhedsbrev = "& nyhedsbrev &", "_
					&" madr = '"& madr & "', mpostnr = '"& mpostnr &"', mcity = '"& mcity &"', mland = '"& mland &"', "_
					&" mtlf = '"& mtlf &"', mcpr = '"& mcpr &"', mkoregnr = '"& mkoregnr &"', "_
                    &" visskiftversion = "& visskiftversion &", medarbejdertype_grp = "& intMedarbejdertype_grp &", timer_ststop = "& timer_ststop &""_
			        &" WHERE Mid = "& id &""
					
					oConn.execute(strSQL)
					
					
					'*** projektgruppe relationer ***'
					call prgrel(id, func)

                    'Response.end
					
					
					'**** timereg_usejob ****'
                    del = 1 '1 slet fra timereg_usejob / 0: tilføj kun
					call tilfojtilTU(id, del) 
					
					
					
					'*** Når der ændres PW ****'
					'*** Sender mail til medarb. at profil er blevet opdateret ***'
					
                    if cint(strAnsat) = 1 then 'kun mail til aktive brugere

					if strPw <> "KEEPTHISPW99" then
					
					if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\medarb_red.asp" then
					
					Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					
					' Sætter Charsettet til ISO-8859-1 
					Mailer.CharSet = 2
					' Afsenderens navn 
					Mailer.FromName = "TimeOut"
					' Afsenderens e-mail 
					    'select case lto
                        'case "kejd_pb"
                        'Mailer.FromAddress = "joedan@okf.kk.dk"
                        'case else
					    Mailer.FromAddress = "timeout_no_reply@outzource.dk"
				        'end select

					Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
					' Modtagerens navn og e-mail
					Mailer.AddRecipient strNavn, strEmail
					'Mailer.AddBCC "Support", "support@outzource.dk" 
					' Mailens emne
					Mailer.Subject = "TimeOut - Medarbejder profil opdateret"

                    if lto = "mi" then
                    Mailer.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20120106.pdf" 
                    else
                    Mailer.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20130102.pdf" 
                    end if
					
					' Selve teksten
					Mailer.BodyText = "" & "Hej "& strNavn & vbCrLf _ 
					& "Din medarbejderprofil er blevet opdateret." & vbCrLf _ 
					& "Dit brugernavn er: " & strLogin & " og dit password er: " & strPw & vbCrLf & vbCrLf _ 
					& "Gem disse oplysninger, til du skal logge ind i TimeOut."  & vbCrLf _ 
					& "Du kan altid selv ændre dem når du er logget på systemet." & vbCrLf & vbCrLf _ 
					& "Adressen til TimeOut er: https://outzource.dk/"&lto&""& vbCrLf & vbCrLf _ 
					& "Med venlig hilsen"& vbCrLf & vbCrLf & strEditor & vbCrLf 
					
					    If Mailer.SendMail Then
        				
    					Else
    				    Response.Write "Fejl...<br>" & Mailer.Response
  					    End if
					
					end if
					
					end if '** Ændre PW **'

                    end if '** bruger aktiv
					
					  level = session("rettigheder")
  
                     if level = 1 then	
	                 Response.Redirect "medarb.asp?menu=medarb&lastmedid="&id
                     else
                     Response.redirect "timereg_akt_2006.asp"
                     end if
					
					end if
					
					
					
					
					
		        '*** Opret medarbejer i DB **'
			    if func = "dbopr" then
    					
    					
                        if cint(strAnsat) = 1 then 'kun mail til aktive brugere

					    if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\medarb_red.asp" then
    					
					    Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
    					
					    ' Sætter Charsettet til ISO-8859-1 
					    Mailer.CharSet = 2
					    ' Afsenderens navn 
					    Mailer.FromName = "TimeOut"
					    ' Afsenderens e-mail 
                        select case lto
                        case "kejd_pb"
                        Mailer.FromAddress = "joedan@okf.kk.dk"
                        case else
					    Mailer.FromAddress = "timeout_no_reply@outzource.dk"
				        end select

                	    Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
					    ' Modtagerens navn og e-mail
					    Mailer.AddRecipient strNavn, strEmail
					    Mailer.AddBCC "Support", "support@outzource.dk" 
					    ' Mailens emne
					    Mailer.Subject = "TimeOut - Medarbejder profil"

                        
                        
                            
                            if lto = "mi" then
                            Mailer.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20120106.pdf"
                            else
                            Mailer.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20130102.pdf" 
                            end if

                      
					    ' Selve teksten
					    Mailer.BodyText = "" & "Hej "& strNavn & vbCrLf _ 
					    & "Velkommen til TimeOut." & vbCrLf & vbCrLf _
					    & "Du er blevet oprettet som bruger i TimeOut" & vbCrLf _ 
					    & "Dit brugernavn er: " & strLogin & " og dit password er: " & strPw & vbCrLf & vbCrLf _ 
					    & "Gem disse oplysninger, til du skal logge ind i TimeOut."  & vbCrLf _ 
					    & "Du kan altid selv ændre dem når du er logget på systemet." & vbCrLf & vbCrLf _ 
					    & "Adressen til TimeOut er: https://outzource.dk/"&lto&""& vbCrLf & vbCrLf _ 
					    & "Med venlig hilsen"& vbCrLf & vbCrLf & strEditor & vbCrLf 
    					
					    end if

                        end if
    					
    					
						    strSQLminsert = "INSERT INTO medarbejdere"_
						    &" (Mnavn, Mnr, Mansat, login, pw, Brugergruppe, Medarbejdertype, "_
						    &" editor, dato, Medarbejderinfo, Email, tsacrm, smilord, exchkonto, init, "_
						    &" timereg, ansatdato, opsagtdato, sprog, nyhedsbrev, "_
						    &" madr, mpostnr, mcity, mland, mtlf, mcpr, mkoregnr, visskiftversion, medarbejdertype_grp, timer_ststop) VALUES ("_
						    &" '"& strNavn &"',"_
						    &" "& strMnr &","_
						    &" '"& strAnsat &"',"_
						    &" '"& strLogin &"',"_
						    &" MD5('"& strPw &"'),"_
						    &" "& strBrugergruppe &","_
						    &" "& strMedarbejdertype &","_
						    &" '"& strEditor &"',"_
						    &" '"& strDato &"',"_
						    &" '"& strMedinfo &"',"_
						    &" '"& strEmail &"',"_
						    &" "& inttsacrm &", "& intSmil &", "_
						    &" '"& strExch &"', '"& strInit &"',"_
						    &" "& intTimereg &","_
						    &" '"& ansatdato &"', '"& opsagtdato &"', "& sprog &", "_
						    &" "& nyhedsbrev &", "_
						    &" '"& madr & "', '"& mpostnr &"', '"& mcity &"', '"& mland &"', "_
					        &" '"& mtlf &"', '"& mcpr &"', '"& mkoregnr &"', "& visskiftversion &", "& intMedarbejdertype_grp &", "& timer_ststop &")"
    						
                            'Response.write strSQLminsert
                            'Response.flush
                            oConn.execute(strSQLminsert)
						    strSQL = "SELECT Mid FROM medarbejdere WHERE Mnr = " & strMnr &""
						    oRec.open strSQL, oConn, 3
						    if not oRec.EOF then
							    intMid = oRec("Mid")
						    end if
						    oRec.close
    						
						    '*** Medarbejder  type historik ****
						    strSQLmthupd = "INSERT INTO medarbejdertyper_historik (mid, mtype, mtypedato) "_
			                &" VALUES ("& intMid &", "& strMedarbejdertype &", '"& ansatdato &"') "
			                oConn.execute(strSQLmthupd)
    						
						    
    						'*** projektgruppe relationer ***'
					        call prgrel(intMid, func)
					        
					        '**** Sætter guiden aktive job aktive ****'
					        'guideasyids = 0
                            'forvalgt = 0 'off / 1 = on

                            'Response.Write "guidjobids:"& guidjobids

                            'Response.end
					        'call setGuidenUsejob(intMid, guidjobids, 1, guideasyids, forvalgt, 1)

                            '**** timereg_usejob ****'
                            call positiv_aktivering_akt_fn() 'wwf
                            if cint(positiv_aktivering_akt_val) <> 1 then
					            call tilfojtilTU(intMid, 0) 
					        end if
    						
						    '*** Opdaterer timepriser på stamaktiviteter ***
    						
						    medtypeid = 0
						    strSQL = "SELECT mid FROM medarbejdere WHERE Medarbejdertype = "& strMedarbejdertype & " AND mid <> " & intMid
						    oRec.open strSQL, oConn, 3 
						    while not oRec.EOF 
						    medtypeid = oRec("mid")
						    oRec.movenext
						    wend
						    oRec.close 
    						
    						
                            ct = 0
                            strAktIDsWrt = ""
						    strSQL = "SELECT aktid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = 0 AND medarbid = "& medtypeid &" GROUP BY aktid, timeprisalt" 
						    'Response.write strSQL & "<br>"
    						
						    oRec.open strSQL, oConn, 3 
						    while not oRec.EOF 

                                tprisUse = replace(oRec("6timepris"), ".", "")
                                tprisUse = replace(tprisUse, ",", ".")
    						
							    strSQLinsTp = "INSERT INTO timepriser ( "_
							    &" jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
							    &" VALUES (0, "& oRec("aktid") &", "& intMid &", "& oRec("timeprisalt") &", "& tprisUse &", "& oRec("6valuta") &")"

                                'Response.Write strSQLinsTp & "<br>"
                                'Response.flush

                                oConn.execute(strSQLinsTp)
    						    


                                strAktIDsWrt = strAktIDsWrt & " AND id <> " & oRec("aktid")
                                ct = ct + 1
						    oRec.movenext
						    wend
						    oRec.close 
    						
    					'Response.write strAktIDsWrt & "<br><br> "

                        '*** Indlæser default timepris på de aktiviteter der ikke bliver fundet *****'
                        '*** finder tp på type ***
                        tprisUse = 0
                        valutaUse = 1
                        strSQLmtyp = "SELECT timepris, tp0_valuta FROM medarbejdertyper WHERE id = " & strMedarbejdertype

                        oRec.open strSQLmtyp, oConn, 3 
						if not oRec.EOF then

                          tprisUse = replace(oRec("timepris"), ".", "")
                          tprisUse = replace(tprisUse, ",", ".")

                          valutaUse = oRec("tp0_valuta")

                        end if
                        oRec.close

                        strSQLstamAktiF = "SELECT id FROM aktiviteter WHERE job = 0 AND aktfavorit <> 0 " & strAktIDsWrt
                        'Response.Write strSQLstamAktiF
                        'Response.flush

                        oRec.open strSQLstamAktiF, oConn, 3 
						while not oRec.EOF 

                        strSQLinsTpAktiF = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
                        &" VALUES (0, "& oRec("id") &", "& intMid &", 6, "& tprisUse &", "&valutaUse&")" 

                        'Response.Write strSQLinsTpAktiF
                        'Response.flush

                        oConn.execute(strSQLinsTpAktiF)

                        oRec.movenext
						wend
						oRec.close 

                        'Response.Write "<br>ct:"& ct
                        'Response.end

    					lastmedid = intMid	
    					
    					
					    if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\medarb_red.asp" then
    					
    					        
                                if len(trim(strEmail)) <> 0 then

					            If Mailer.SendMail Then
        				
						            call medarb_opr_ok
    						
					            Else
    				            Response.Write "Fejl...<br>" & Mailer.Response
  					            End if

                                else

                                    call medarb_opr_ok

                                end if
    					
					    else
    						
						    call medarb_opr_ok
    						
					    end if
    						
			    end if
			else
			'*** Hvis dublet fandtes ***
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			if strMnrOK <> "y" then
			errortype = 10
			call showError(errortype)
			else
			
			if strLoginOK <> "y" then
			errortype = 11
			call showError(errortype)
						end if
					end if
				end if
			end if
		end if
	case else
	'**************************** Medarbejder Data ************************************
	%>
	
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	
	function showdisable() {
	//if ( == ) {
	//document.getElementBy("FM_ansat2").status = "checked"
	alert("Når en medarbejder bliver deaktiveret kan denne ikke længere logge sig ind i TimeOut. Medarbejderen vil ikke længere indgå som bruger når antallet af aktive brugere bliver opgjort. (TimeOut Service Aftale). Alle medarbejderens data vil forblive som de er, og vil stadigvæk være tilgængelige.")
	//}
	}
	
	function rensfelt(){
	document.getElementById("pw").value = ""
	}



	$(document).ready(function () {

	$("#FM_medtype").change(function () {

        //alert("her")

	    $("#FM_opd_mty_kp").attr("disabled", true);
	    $("#FM_opd_mty_tp").attr("disabled", true);


	});

	});



	</script>
	
	
	

                        <%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px;">
	<%
	
	'oimg = "ikon_medarb_48.png"
	'oleft = 0
	'otop = 0
	'owdt = 300
	'oskrift = "Medarbejdere"
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	%>
	
	<%
	
	 tTop = 0
	 tLeft = 0
	 tWdth = 700
	       
        	
        	
	        call tableDiv(tTop,tLeft,tWdth)

    if level = 1 OR session("mid") = id then

        if func = "red" then
        dbfuncTxt = "rediger"
        else
        dbfuncTxt = "opret"
        end if

	
	%>
	
	<table cellspacing="0" cellpadding="5" border="0" width="100%" bgcolor="#eff3ff">
    <tr bgcolor="#ffffff"><td colspan="2"><h4>Medarbejdere<span style="font-size:10px; font-weight:normal;"> - <%=dbfuncTxt %></span></h4></td></tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 class=alt style="height:30px;"><b>Stamdata:</b></td>
	</tr>
	<%
	if func = "red" then
		strSQL = "SELECT mid, mnavn, mnr, mansat, login, pw, lastlogin, brugergruppe, "_
		&" medarbejdertype, type, navn, medarbejdere.editor, medarbejdere.dato, "_
		&" Medarbejderinfo, email, tsacrm, smilord, exchkonto, init, timereg, ansatdato, "_
		&" opsagtdato, sprog, nyhedsbrev,  "_
		&" madr, mpostnr, mcity, mland, mtlf, mcpr, mkoregnr, visskiftversion, timer_ststop "_
		&" FROM medarbejdere, brugergrupper, medarbejdertyper "_
		&" WHERE mid = "& id &" AND brugergrupper.id = brugergruppe AND medarbejdertyper.id = medarbejdertype"
		
		oRec.open strSQL, oConn, 3
		
		if not oRec.EOF then
			
			strNavn = oRec("Mnavn")
			strAnsat = oRec("Mansat")
			strLogin = oRec("login")
			strPw = oRec("pw")
			strMnr = oRec("Mnr")
			strEditor = oRec("editor")
			strDato = oRec("dato")
			strMedinfo = oRec("Medarbejderinfo")
			dbfunc = "dbred"
			varsubval = "Opdater"
			strEmail = oRec("email")
			intCRM = oRec("tsacrm")
			intSmil = oRec("smilord")
            
            if len(trim(oRec("exchkonto"))) <> 0 then
			strExch = replace(oRec("exchkonto"), "#", "\")
            else
            strExch = ""
            end if
			
            strInit = oRec("init")
			intTimereg = oRec("timereg")
			ansatdato = oRec("ansatdato")
			opsagtdato = oRec("opsagtdato")
			sprog = oRec("sprog")
			nyhedsbrev = oRec("nyhedsbrev")
			
			madr = oRec("madr")
			mpostnr = oRec("mpostnr")
			mcity = oRec("mcity")
			mland = oRec("mland")
			mtlf = oRec("mtlf")
			mcpr = oRec("mcpr")
			mkoregnr = oRec("mkoregnr")

            visskiftversion = oRec("visskiftversion")
            
            medarbejdertype = oRec("medarbejdertype")

            timer_ststop = oRec("timer_ststop")
			
		end if
		oRec.close
		
		'if strAnsat = "Yes" then
		'strAnsat = "checked"
		'else
		'strAnsat = ""
		'end if
         
       
		
		
	%>
	<tr>
		
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
		
	</tr>
	<%
	else
		strSQL = "SELECT Mnr FROM medarbejdere ORDER BY Mnr"
		oRec.open strSQL, oConn, 3
		
		oRec.movelast
		
		strNavn = ""
		strAnsat = ""
		strLogin = ""
		strPw = ""
		strMnr = (oRec("Mnr") + 1)
		strEditor = ""
		strDato = ""
		strMedinfo = ""
		dbfunc = "dbopr"
		varSubVal = "Opret" 
		strEmail = ""
		strCRMcheckedTSA = "checked"
		intSmil = 1
		strExch = ""
		strInit = ""
		intTimereg = "1"
		ansatDato = date()
		opsagtdato = "01-01-2044"
		nyhedsbrev = 1
        visskiftversion = 0

        medarbejdertype = 0
        timer_ststop = 0

       
		
		oRec.close


        select case lto 
        case "esn", "tec", "intranet - local"
        intCRM = 3
        case else
        intCRM = 0
        end select

	end if
	



          strCRMcheckedCRM = ""
		strCRMcheckedTSA = ""
		strCRMcheckedRes = ""
         strCRMcheckedTSA_3 = ""
        strCRMcheckedTSA_4 = ""
        strCRMcheckedTSA_5 = ""
        strCRMcheckedTSA_6 = ""
		
		select case intCRM 
		case 1
		strCRMcheckedCRM = "checked"
		case 2
		strCRMcheckedRes = "checked"
		case 3
        strCRMcheckedTSA_3 = "CHECKED"
        case 4
        strCRMcheckedTSA_4 = "CHECKED"
        case 5
        strCRMcheckedTSA_5 = "CHECKED"
        case 6
        strCRMcheckedTSA_6 = "CHECKED"
		case else
		strCRMcheckedTSA = "checked"
		end select


	
            if cint(visskiftversion) = 1 then
            visskiftversionCHK = "CHECKED"
            else
            visskiftversionCHK = ""
            end if

            if cint(timer_ststop) = 1 then
            timer_ststopCHK = "CHECKED"
            else
            timer_ststopCHK = ""
            end if

            

		if intSmil = 1 then
			chkSmil = "CHECKED"
		else
			chkSmil = ""
		end if
		
		
		if intTimereg = "1" then
		strTregChk0 = ""
		strTregChk1 = "CHECKED"
		else
		strTregChk0 = "CHECKED"
		strTregChk1 = ""
		end if
	%>
	<form action="medarb_red.asp?menu=medarb&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<tr bgcolor="#ffffff">
	    <td colspan=2>&nbsp;</td>
	</tr>
	
	<tr bgcolor="#ffffff">
	    <td style="padding-top:5px; width:200px;"><font color=red size=2>*</font> <b>Navn:</b></td>
		<td style="padding-top:5px; width:500px;"><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" >&nbsp;&nbsp;
		Initialer: <input type="text" name="FM_init" value="<%=strInit%>" size="5" ></td>
		
	</tr>
	<%if level <= 2 OR level = 6 then%>
	<tr bgcolor="#ffffff">
		
		<td><font color=red size=2>*</font> <b>Medarbejder nr.</b></td>
		<td><input type="text" name="FM_Mnr" value="<%=strMnr%>" size="10" ></td>
		
	</tr>
	<%
	else%>
	<input type="hidden" name="FM_Mnr" value="<%=strMnr%>">
	<%end if
	
	
	if level = 1 then%>
	<tr bgcolor="#ffffff">
		
		<td valign=top><br /><b>Status:</b><br />
        <span style="color:#999999; font-size:10px;">
		<b>Deaktiveret</b> - Kan ikke logge ind og der kan ikke registreres timer. Tæller ikke med i licensregnskab.<br /><br />
		<b>Passiv</b> - Kan ikke logge ind, men der kan godt registreres timer på medarbejderen.</span></td>
		<%
		select case strAnsat
		case "2"
		chk1 = ""
		chk2 = "CHECKED"
		chk3 = ""
		case "3"
		chk3 = "CHECKED"
		chk2 = ""
		chk1 = ""
		case else
		chk1 = "CHECKED"
		chk2 = ""
		chk3 = ""
		end select 
		
	    %>
		<td valign=top><br /><input type="radio" name="FM_ansat" id=FM_ansat1 value="1" <%=chk1%>>Aktiv <%if func <> "red" then %>
            (der udsendes mail til medarbejder med logind oplysninger)
            <%end if %><br />
		<input type="radio" name="FM_ansat" id="FM_ansat2" value="2" <%=chk2%>>Deaktiveret (fratrædelsesdato bliver sat = d.d. når en medarb. deaktiveres)<br />
		<input type="radio" name="FM_ansat" id="FM_ansat3" value="3" <%=chk3%>>Passiv
		<br />
        &nbsp;
		</td>
		
	</tr>
	<%
	else
		%>
		<input type="hidden" name="FM_ansat" value="<%=strAnsat%>">
		<%
	end if%>
	<tr bgcolor="#ffffff">
		
		<td><font color=red size=2>*</font><b> Login:</b></td>
		<td><input type="text" name="FM_login" value="<%=strLogin%>" size="30" ></td>
		
	</tr>
	<tr bgcolor="#ffffff">
		
		<td valign=top><font color=red size=2>*</font> <b>Password:</b></td>
		<%if func = "red" then 
		strPw = "KEEPTHISPW99"
		else
		strPw = ""
		end if %>
		<td><input type="password" id="pw" name="FM_pw" value="<%=strPw %>" size="30"  onfocus="rensfelt()">
		<br />
		<%if func = "red" then %>
		(Hvis der ændres password, vil der blive udsendt en email til medarbejderen herom. Gælder kun hvis medarbejderen er aktiv.)<br />
		<%end if %>
            &nbsp;
		</td>
		
	</tr>
	
	
		
	<tr bgcolor="#ffffff">
	    <td>Adresse</td>
	    <td>
            <input id="Text3" name="FM_adr" type="text" value="<%=madr%>" style="width:200px;" /></td>
	</tr>
	<tr bgcolor="#ffffff">
	    <td>Postnr.</td>
	    <td>
            <input id="Text4" name="FM_postnr" type="text" value="<%=mpostnr%>" style="width:200px;" /> 
            By:  <input id="Text6" name="FM_city" value="<%=mcity%>" type="text" style="width:200px;" /></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td>Land:</td>
		<td><select name="FM_land" style="width:200px;">
		<%if func = "red" then%>
		<option SELECTED><%=mland%></option>
		<%else%>
		<option SELECTED>Danmark</option>
		<%end if%>
		<!--#include file="inc/inc_option_land.asp"-->
		</select>
		<!--<input type="text" name="FM_land" value="<=strLand%>" size="20" style="border:1px #86B5E4 solid;">--></td>
	</tr>
	<tr bgcolor="#ffffff">
	    <td>Tel.</td>
	    <td>
            <input id="Text5" name="FM_tel" value="<%=mtlf%>" type="text" style="width:200px;" /></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td>Email:</td>
		<td><input type="text" name="FM_email" value="<%=strEmail%>" style="width:200px;"> (Medarbejder info sendes til denne email adr.)</td>
		
	</tr>
	<tr bgcolor="#ffffff">
	    <td>Cpr. nr</td>
	    <td>
            <input id="Text7" name="FM_cpr" value="<%=mcpr%>" type="text" style="width:200px;" /></td>
	</tr>
	<tr bgcolor="#ffffff">
	    <td>Registrerings nr. på køretøj</td>
	    <td>
            <input id="Text8" name="FM_regnr" type="text" value="<%=mkoregnr%>" style="width:200px;" /></td>
	</tr>
	
	
	<tr bgcolor="#ffffff">
		
		<td valign="top">Note / beskrivelse</td>
		<td><textarea cols="55" rows="5" name="FM_medinfo"><%=strMedinfo%></textarea></td>
		
	</tr>
	
	<%if level = 1 OR func <> "red" then %>
	<tr bgcolor="#ffffff">
		
		<td valign=top style="padding-top:5px;"><b>Ansat dato:</b><br />
		<span style="color:#999999; font-size:10px;">TimeOut start dato. Normtimer på medarb. regnes fra denne dato, dog ikke længere tilbage end licens st. dato.</span>
	    </td>
		<td><input type="text" name="ansatdato" id="ansatdato" value="<%=ansatdato%>" size="9" >
		dd-mm-åååå <br />
		Den første type i ens medarb.type historik (dvs. normtimer og timepriser) følger altid ansatdato,<br /> og vil blive opdateret hvis ansatdato ændres.</td>
		
	</tr>
	<tr bgcolor="#ffffff">
		
		<td valign=top style="padding-top:5px;"><b>Fratrådt dato:</b></td>
		<td><input type="text" name="opsagtdato" id="opsagtdato" value="<%=opsagtdato%>" size="9" >
		dd-mm-åååå (01-01-2044 = Stadigvæk ansat) <br /></td>
		
	</tr>
	<%else %>
	<tr bgcolor="#ffffff">
	<td valign=top>Ansat dato:<br />
	(TimeOut start dato)
	<br />
	Fratrådt dato:</td>
	<td valign=top> <%=ansatdato%><br />
	<%=opsagtdato%>
	</td>
	</tr>
	<input type="hidden" name="ansatdato" id="Text2" value="<%=ansatdato%>">
	<input type="hidden" name="opsagtdato" id="Text1" value="<%=opsagtdato%>">
	<%end if %>
	
	
	<tr bgcolor="#ffffff">
		
		<td>Sprog:</td>
		<td><select name="FM_sprog" style="width:200px;">
		<%
		
			strSQL = "SELECT sproglabel, id FROM sprog WHERE id < 3"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF 
			 if sprog = oRec("id") then
			 sprogSEL = "SELECTED"
			 else
			 sprogSEL = ""
			 end if
			 %>
			<option value="<%=oRec("id")%>" <%=sprogSEL %>><%=oRec("sproglabel")%></option>
			<%
			oRec.movenext 
			wend
			oRec.close
	    %>
	    </select></td>
		
	</tr>
	
	<%if func = "red" then %>
	
	<tr bgcolor="#ffffff">
	<% 
	if cint(nyhedsbrev) = 0 then
	chkny0 = "CHECKED"
	chkny1 = ""
	else
	chkny0 = ""
	chkny1 = "CHECKED"
	end if 
	%>	
		<td valign=top style="padding-top:5px;">
		<b>Nyhedsbrev</b><br />
        Jeg ønsker at modtage OutZourCE TimeOut nyhedsbrev<br /><br />&nbsp;</td>
		<td valign=top style="padding-top:5px;">Ja <input id="nyhedsbrev" name="nyhedsbrev" type="radio" value="1" <%=chkny1 %> />&nbsp;&nbsp;Nej 
            <input id="nyhedsbrev" name="nyhedsbrev" type="radio" value="0" <%=chkny0 %> /></td>
		
	</tr>
	<%end if %>
	<tr>
		
		<td valign=top><br><b>Exchange/Domænekonto navn:</b><br>
		<span style="color:#999999; font-size:10px;">Bruges ved integretion til firmaets<br> egen Exchange server.</span></td>
		<td valign=top><br><input type="text" name="FM_exch" value="<%=strExch%>" size="30" >
		</td>
		
	</tr>
    <%if level = 1 then %>
 

    <tr>
		
		<td valign=top><b>Tillad medarbejder skifter TimeOut version.</b> (hvis findes)</td>
		<td valign=top><input type="checkbox" name="FM_visskiftversion" value="1" <%=visskiftversionCHK %>/> Ja, må gerne skifte internt (fra timereg. siden), uden at skulle logge ind igen</td>
		
	</tr>
    <%else %>
    <input type="hidden" value="<%=visskiftversion %>" name="FM_visskiftversion" />
    <%end if %>

	<%if cint(level) = 1 then
	        
	        if func = "red" then
	        antalhist = 0
	        strSQLantalth = "SELECT COUNT(id) AS antalhist FROM medarbejdertyper_historik WHERE mid = "& id &" GROUP BY mid"
	        oRec.open strSQLantalth, oConn, 3
	        if not oRec.EOF then
	            
	            antalhist = oRec("antalhist") 
	            
	        end if
	        oRec.close
	        
	        if antalhist >= 30 then
	        dis = "DISABLED"
	        else
	        dis = ""
	        end if
	        
	        else
	        dis = ""
	        end if
	        
	        
	%>
	<tr bgcolor="#ffffff">
		
		<td valign="top"><br /><b>Medarbejdertype:</b> (timepriser & normtid)<br />
		<span style="color:#999999; font-size:10px;">Bruges til at bestemme hvor mange timer man er ansat og til hvilke timepriser medarbejderen faktures.</span><br />

		<a href="medarbtyper.asp" target="_blank" class=vmenu>Se medarbejdertyper..</a></td>
		<td valign=top><br />
            <%	strSQL = "SELECT medarbejdertype, mt.type, mt.id, mgruppe, "_
            &" mtg.navn AS mtgnavn "_ 
            &" FROM medarbejdere AS m, medarbejdertyper AS mt"_
            &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mgruppe) WHERE mid = "& id &" AND mt.id = medarbejdertype AND mid = 99999999"
			%>
            
            <select <%=dis %> name="FM_medtype" id="FM_medtype" style="width:400px;">
		<%
		if func = "red" then
		
		
            
            oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			        
            if oRec("mgruppe") <> 0 then
            mtgNavn = " [hovedgruppe:  "& oRec("mtgnavn") &"]"
            else
            mtgNavn = ""
            end if
            %>
			<option value="<%=oRec("id")%>" SELECTED><%=oRec("type")%> <span style="font-size:8px;"><%=mtgNavn%></span></option>
			<%
			end if
			oRec.close
		end if
		
		'strSQL = "SELECT id, type FROM medarbejdertyper ORDER BY type"
		    strSQL = "SELECT mt.type, mt.id, mgruppe, "_
            &" mtg.navn AS mtgnavn "_ 
            &" FROM medarbejdertyper AS mt"_
            &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mgruppe) WHERE mt.id <> 0 ORDER BY mtg.navn, type"
        oRec.open strSQL, oConn, 3
		While not oRec.EOF
                
            if oRec("mgruppe") <> 0 then
            mtgNavn = " [ hovedgruppe:  "& oRec("mtgnavn") &"]"
            else
            mtgNavn = ""
            end if

            if cint(medarbejdertype) = oRec("id") then
            mtypSEL = "SELECTED"
            else
            mtypSEL = ""
            end if

            if lastMtypHgrp <> oRec("mtgNavn") then
            %>  
                <option value="0" disabled></option>
                <option value="0" disabled>Hovedgruppe: <%=oRec("mtgNavn")%></option>
            <%
            
            end if

                %>
		<option value="<%=oRec("id")%>" <%=mtypSEL%>><%=oRec("type")%></option>
		<%

        lastMtypHgrp = oRec("mtgNavn")
		oRec.movenext
		Wend
		oRec.close
		%>
		</select>
        <br /><br />
        <% 
        uWdt = 200
        uTxt = "Husk at tilpasse ferie- / ferie fridage -afholdt, hvis der ændres medarbejdertype / normtid med tilbagevirkende kraft."
        call infoUnisport(uWdt, uTxt)
        %>
        </td>
		
	</tr>
	<%
	else
		strSQL = "SELECT medarbejdertype, type, id FROM medarbejdere, medarbejdertyper WHERE Mid = "& id &" AND medarbejdertyper.id = medarbejdertype"

		oRec.open strSQL, oConn, 3
		if not oRec.EOF then%>
		<input type="hidden" name="FM_medtype" value="<%=oRec("id")%>">
		<%
		end if
		oRec.close
	end if%>
	
	
	<%if func = "red" then %>
	<tr bgcolor="#ffffff">
		
		<td style="height:40px;">
            &nbsp;</td>
		<td style="padding:4px 0px 4px 0px;" valign=top class=lille><b>Medarbejdertype historik:</b> (maks 30)<br />
        Første type skal altid følge ansat dato. Derfor kan den ikke opdateres.<br />
        Kun normtid følger historikken. Realiserede timer følger altid den gruppe der er valgt for medarbejdederen. Der kan ændres timepriser realiserede timer ved at redigere medarbejdertypen. <br /><br />
		<%
		
			strSQL = "SELECT mtypedato, mt.type AS mttype, mth.id AS mtypeid FROM medarbejdertyper_historik AS mth, "_
			&" medarbejdertyper mt WHERE mth.mid = "& id &" AND mt.id = mth.mtype ORDER BY mth.id"

            'Response.Write strSQL
            'Response.flush
            mth = 1
			oRec.open strSQL, oConn, 3
			while not oRec.EOF 
			
			%>
			<%=oRec("mttype")%> - fra d. 

            <%if level = 1 then
               
            %>
            <input id="Hidden4" type="hidden" name="FM_mtyphist_id" value="<%=oRec("mtypeid") %>" />
            
            <%if cint(mth) = 1 then %>
            <input id="Hidden6" type="hidden" name="FM_mtyphist_dato" value="<%=oRec("mtypedato") %>" /><%=oRec("mtypedato") %><br />
            <%else %>
            <input id="Text9" type="text" name="FM_mtyphist_dato" style="font-size:9px; width:80px;" value="<%=oRec("mtypedato") %>" /><br />
            <%end if %>


            <%else %>
            <input id="Hidden5" type="hidden" name="FM_mtyphist_id" value="0" />
            <input id="Text10" type="hidden" name="FM_mtyphist_dato" value="<%=oRec("mtypedato") %>" /><%=oRec("mtypedato") %><br />
            <%end if %>
			<% 

            mth = mth + 1
			oRec.movenext
			wend 
			oRec.close
			
			if mth = 1 then
			%>
			(Ingen)
			<%
			end if
		
		

        if level = 1 AND func = "red" then
        %>
        <br />
        <b>Opdater eksisterende timepriser?</b> (på den aktuelle type)<br />
       Opdaterer priser på alle timeregistreringer, på alle job uanset status, for denne medarbejder til at følge den nuværende medarbejdertype indstilling, fra den dato medarbejderen er oprettet som denne type.<br />
            <input type="checkbox" name="FM_opd_mty_kp" id="FM_opd_mty_kp" value="1"/>Opdater eksisterende <b>kostpriser</b> 
        <br />
             <input type="checkbox" name="FM_opd_mty_tp" id="FM_opd_mty_tp" value="1"/>Opdater eksisterende <b>timepriser</b> (overskriver ikke indstillinger på job, men opdaterer timeprisen på eksisterende registreringer)
           <br />&nbsp;
        
            <%
        end if
                %>

	</td>
		
	</tr>
    <%end if %>
	
	
	
	<tr>
		
		<td><br /><b>Brugergruppe:</b> (rettigheder)<br />
        <span style="color:#999999; font-size:10px;">
		Bruges til at bestemme hvilke rettigheder og hvilke områder af TimeOut medarbejderen 
		har adgang til.</span>
		
		</td>
		<%if level = 1 then %>
		<td valign=top><br /><select name="FM_bgruppe" style="width:200px;">
		<%
		'** Finder den allerede valgte brugergruppe **
		if func = "red" then
			strSQL = "SELECT brugergruppe, navn, id FROM medarbejdere, brugergrupper WHERE Mid = "& id &" AND brugergrupper.id = brugergruppe"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			%>
			<option value="<%=oRec("id")%>" SELECTED><%=oRec("navn")%></option>
			<%
			end if
			oRec.close
		end if
		
		'** Hvis der oprettes ny medarbejder sættes brugergruppe 2 () til default ***
		'if licensType = "CRM" then
		strWHEREKlaus = " "
		'else
		'strWHEREKlaus = " WHERE rettigheder <= 2 OR rettigheder = 6 "
		'end if
		
		strSQL = "SELECT id, navn FROM brugergrupper "& strWHEREKlaus &" Order by id"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF
		if oRec("id") = 7 AND func <> "red" then%>
		<option value="<%=oRec("id")%>" SELECTED><%=oRec("navn")%></option>
		<%else%>
		<option value="<%=oRec("id")%>"><%=oRec("navn")%></option>
		<%end if
		oRec.movenext
		Wend
		oRec.close
		%>
		</select>
	<%
	'*** Hvis bruger ikke har admin rettigheder ****
	else
	        rettighedsnavn = ""
			strSQL = "SELECT m.brugergruppe, b.navn AS bnavn, b.id FROM medarbejdere m LEFT JOIN brugergrupper b ON (b.id = m.brugergruppe) WHERE Mid = "& id 
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			rettighedsnavn = oRec("bnavn") %>
			<input type="hidden" name="FM_bgruppe" value="<%=oRec("id")%>">
			<%
			end if
			oRec.close
	
	%>
	<td><b><%=rettighedsnavn %></b>
	<%	
	end if%>
	</td>
		
	</tr>
	
	<%if level = 1 then
	projGrpDis = ""
	else 
	projGrpDis = "DISABLED"
	end if %>
	<tr bgcolor="#ffffff">
	
	<td valign=top><br /><b>Projektgrupper</b> (til job)<br />
	<span style="color:#999999; font-size:10px;">Medarbejderen skal være medlem af en eller flere projektgrupper for at kunne registrere timer. 
	En medarbejder bliver automatisk tilmeldt "Alle" projektgruppen ved oprettelse.</span> <br />
    <%if level = 1 then %>
	<a href="projektgrupper.asp?menu=job" class=vmenu target="_blank">Se projektgrupper..</a>
    <%end if %>
    </td>
    <td valign=top><br /><b>Projektgruppe medlemsskaber:</b>
    (ud over "Alle" gruppen)<br />
    Du kan ikke angive nye teamleder egenskaber på projektgrupper fra denne liste.
    
    
        <input id="Hidden3" name="FM_progrp" type="hidden" value="10" /><!-- Alle gruppen -->
        <select id="FM_progrp" name="FM_progrp" multiple style="width:200px; height:200px;" <%=projGrpDis %>>
            
    <%
    
    selprogrp = ""
    
    if func = "opret" then
    thisMid = 0
    else
    thisMid = id
    end if
    
    strNOTpgids = ""
    
    '*** Henter projektgrupper grupper medarbejder er med i ***'
    strSQLpg = "SELECT Mid, MedarbejderId, p.navn AS pgnavn, p.id AS pid, pr.teamleder FROM progrupperelationer pr "_
    &" LEFT JOIN medarbejdere AS m ON (Mid = MedarbejderId) "_
    &" LEFT JOIN projektgrupper AS p ON (p.id = pr.projektgruppeId) WHERE Mid = "& thisMid &" ORDER BY p.navn"
	oRec2.open strSQLpg, oConn, 3 
	pgr = 1
	
	while not oRec2.EOF 
	
	if len(trim(oRec2("pid"))) <> 0 then
	strNOTpgids = strNOTpgids & " AND pg.id <> " & oRec2("pid")
	end if
	
	if oRec2("pid") <> 10 then
        
        if oRec2("teamleder") <> 0 then
        teaml = "_1"
        teamlTxt = " (teamleder)"
        else
        teaml = ""
        teamlTxt = ""
        end if


	%>
    <option value=<%=oRec2("pid") & teaml %> SELECTED><%=oRec2("pgnavn") %> <%=teamlTxt %></option>
    <%
    
    selprogrp = selprogrp & ","& oRec2("pid")
    
	end if
	
	pgr = pgr + 1
	oRec2.movenext
	wend
	oRec2.close
	
	
	
	'*** Henter resterende grupper ***'
	strSQLrpg = "SELECT pg.navn, pg.id AS pid FROM projektgrupper pg WHERE pg.id <> 10 " & strNOTpgids & " ORDER BY pg.navn"
	'Response.Write strSQLrpg
	'Response.flush
	oRec2.open strSQLrpg, oConn, 3 
	while not oRec2.EOF 
	%>
    <option value=<%=oRec2("pid")%>><%=oRec2("navn") %></option>
    <%
	oRec2.movenext
	wend
	oRec2.close
	%>
    </select>
	
    <%if func = "red" AND projGrpDis = "DISABLED" then %>
        <input id="FM_progrp" name="FM_progrp" type="hidden" value ="<%=selprogrp%>" />
    <%end if %>   
	
	
    <br />Er medlem af: <b><%=pgr-1 %></b> projektgrupper incl. "Alle" gruppen.
    
   
    
    </td>
	
	</tr>
	
	
	
	
        <%if level = 1 then %>
	<tr>
		
		<td colspan="2"><br><b>Startside</b>:<br>
		 <input type="radio" name="FM_tsacrm" value="0" <%=strCRMcheckedTSA%>> Timeregistrering (default)  <input type="checkbox" name="FM_timer_ststop" value="1" <%=timer_ststopCHK%> /> Indtast timer som start og stop tid <br>
          
         <input type="radio" name="FM_tsacrm" value="6" <%=strCRMcheckedTSA_6%>> Ugeseddel<br>


            <%call erStempelurOn()
                
            if cint(stempelurOn) = 1 then %>
          <input type="radio" name="FM_tsacrm" value="3" <%=strCRMcheckedTSA_3%>> Stempelur (Komme / Gå)  <br>
            <%end if %>
		
		 <input type="radio" name="FM_tsacrm" value="2" <%=strCRMcheckedRes%>> Ressourceplanner<br>
        <input type="radio" name="FM_tsacrm" value="4" <%=strCRMcheckedTSA_4%>> Igangværende job<br>
         <input type="radio" name="FM_tsacrm" value="5" <%=strCRMcheckedTSA_5%>> Joblisten<br>
         <%if licensType = "CRM" then%>
         <input type="radio" name="FM_tsacrm" value="1" <%=strCRMcheckedCRM%>> CRM Kalender<br>
        <%end if%>

		</td>

		
	</tr>
        <%end if %>
	
	
	
	<%
	'*** Smiley ***
	'call ersmileyaktiv()
	
	'if cint(smilaktiv) = 1 then
    %>
    <!--
	<tr>
		
		<td colspan="2"><br><b>Smiley ordning</b>:<br>
		<%if cint(intSmil) = 1 then%>
		<b><i><font color=green>V</font></i></b>&nbsp;Du er tilmeldt Smileyordningen. Smileyordningen kan kun slås til og fra på firma-niveau.
		<%else%>
		<b><i><font color=red>-</font></i></b>&nbsp;Du er IKKE tilmeldt Smileyordningen, men du vil automatisk blive tilmeldt<br />
		når du opdaterer din profil, da I som firma tilvalgt Smileyordningen.
		<%end if %>
		<br><font class=megetlillesort>Smiley ordningen omfatter daglige reminders på timeregistreringssiden 
		vedr. magnlende registreringer, samt giver mulighed for at færdigmelde uger.</font></td>
	    <input type="hidden" name="FM_smiley" id="FM_smiley" value="1">
	    
	</tr>
	<end if%>
	
	-->
	
	<!--<tr>
		
		<td colspan="2"><br><b>Timereg. side</b>:<br>
		Jeg ønsker at benytte:<br>
		 <input type="radio" name="FM_timereg" value="0" <%=strTregChk0%>> Ver. 02 (gammel version) <br>
		 <input type="radio" name="FM_timereg" value="1" <%=strTregChk1%>> Ver. 06 (ny version)</td>
		
	</tr>
	-->
	
	<input type="hidden" name="FM_smiley" id="Hidden1" value="1">
	<input type="hidden" name="FM_timereg" id="Hidden2" value="1">
	
	<tr>
		
		<td colspan="2" align=center><br><br><input type="submit" value="<%=varsubval%> >> "><br><br>&nbsp;</td>
		
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</form>
	</table>
    </div>
    <br><br>
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br>
	
    <%else 'level%>

     <table><tr><td>Du har ikke adgang til at se denne side.</td></tr></table></div>

     <%end if %>

	<!-- table div -->
	</div>
	
	

</div>
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
