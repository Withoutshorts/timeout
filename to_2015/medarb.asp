


<%response.buffer = true 
Session.LCID = 1030
%>
			        

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->



<%'** ER SESSION UDLØBET  ****
    if len(session("user")) = 0 then
	
	errortype = 5
	call showError(errortype)
    response.end
    end if %>

<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%
   
    func = request("func")

	if func = "dbopr" then
	id = 0
	else
        if len(trim(request("id"))) <> 0 then
	    id = request("id")
        else
        id = 0
        end if
	end if    
    
    media = request("media")
   

    if media <> "eksport" then


    if media <> "print" then    
        call menu_2014
    end if


     %>
    <div id="wrapper">
        <div class="content">

    <%

    end if

    select case func
    case "slet"
    %>
            <!--slet sidens indhold-->
        <div class="container" style="width:500px;">
            <div class="porlet">
            
            <h3 class="portlet-title">
               <u>Medarbejder</u>
            </h3>
            
                <div class="portlet-body">
                    <div style="text-align:center;">Du er ved at <b>SLETTE</b> en medarbejder. Er dette korrekt?
                    </div><br />
                   <div style="text-align:center;"><a class="btn btn-primary btn-sm" role="button" href="medarb.asp?menu=tok&func=sletok&id=<%=id%>">&nbsp Ja &nbsp</a>&nbsp&nbsp&nbsp&nbsp<a class="btn btn-default btn-sm" role="button" href="medarb.asp?">Nej</a>
                    </div>
                    <br /><br />
                 </div>

            </div>
        </div>
    
      <%
         case "sletok"
        '*** Her slettes en medarbejder ***

	oConn.execute("DELETE FROM budget_medarb_rel WHERE medid = "& id &"")
	oConn.execute("DELETE FROM medarbejdere WHERE Mid = "& id &"")
	oConn.execute("DELETE FROM progrupperelationer WHERE MedarbejderId = "& id  &"") 'projektgruppeId = 10 AND
	oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& id  &"") 
    oConn.execute("DELETE FROM timer WHERE tmnr = "& id &"")
	oConn.execute("DELETE FROM ressourcer WHERE mid = "& id &"")
	oConn.execute("DELETE FROM ressourcer_md WHERE medid = "& id &"")
	oConn.execute("DELETE FROM timepriser WHERE medarbid = "& id &"")
	
    	
    Response.Redirect "medarb.asp?menu=medarbejder&lastmedid="&id
                     
					
					
    

    
    case "dbred", "dbopr" 


	'*** Her tjekkes om alle required felter er udfyldt. ***
      


	if len(trim(Request("FM_login"))) = 0 OR len(trim(Request("FM_pw"))) = 0 OR len(trim(Request("FM_mnr"))) = 0 OR len(trim(request("FM_navn"))) = 0 then 
	
	
	errortype = 9
	call showError(errortype)

       
    response.end
	
	else
	        %>
			<!--#include file="../timereg/inc/isint_func.asp"-->
			<%
		    isInt = 0
	        call erDetInt(trim(request("FM_mnr")))

                'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br>SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS #"& request("FM_mnr") & "# isInt: "& isInt &"<br>"
                'Response.write 
                'Response.end


			if isInt > 0 then
			
			errortype = 22
			call showError(errortype)

           
            response.end
			
			isInt = 0
			else
	        
	        
	        if len(request("ansatdato")) <> 0 AND isDate(request("ansatdato")) then
			ansatdato = request("ansatdato")
			ansatdato = year(ansatdato)&"/"&month(ansatdato)&"/"&day(ansatdato)
			else
		

			errortype = 114
			call showError(errortype)
			Response.end
			end if	
			
			if len(request("opsagtdato")) <> 0 AND isDate(request("opsagtdato")) then
			opsagtdato = request("opsagtdato")
			opsagtdato = year(opsagtdato)&"/"&month(opsagtdato)&"/"&day(opsagtdato)
			else
			

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
			strMnr = trim(Request("FM_mnr"))
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
                    

                    if len(trim(request("FM_deakdd_cancel"))) <> 0 then
                    deakdd_cancel = 1
                    else
                    deakdd_cancel = 0
                    end if

                    if useToBeAnsat <> "2" AND deakdd_cancel = 0 then
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

                        Set myMail=CreateObject("CDO.Message")
                        myMail.Subject="TimeOut - Medarbejder profil opdateret"
                        myMail.From="timeout_no_reply@outzource.dk"

                        if len(trim(strEmail)) <> 0 then
                        myMail.To= ""& strNavn &"<"& strEmail &">"
                        end if

                        if lto = "mi" then
                        myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20120106.pdf" 
                        else
                        myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20130102.pdf" 
                        end if
                    
                        myMail.TextBody= "" & "Hej "& strNavn & vbCrLf _ 
					    & "Din medarbejderprofil er blevet opdateret." & vbCrLf _ 
					    & "Dit brugernavn er: " & strLogin & " og dit password er: " & strPw & vbCrLf & vbCrLf _ 
					    & "Gem disse oplysninger, til du skal logge ind i TimeOut."  & vbCrLf _ 
					    & "Du kan altid selv ændre dem når du er logget på systemet." & vbCrLf & vbCrLf _ 
					    & "Adressen til TimeOut er: https://outzource.dk/"&lto&""& vbCrLf & vbCrLf _ 
					    & "Med venlig hilsen"& vbCrLf & vbCrLf & strEditor & vbCrLf 

                        
                        myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                        'Name or IP of remote SMTP server
                                    
                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                       smtpServer = "webout.smtp.nu"
                                    else
                                       smtpServer = "formrelay.rackhosting.com" 
                                    end if
                    
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                        'Server port
                        myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                        myMail.Configuration.Fields.Update

                        if len(trim(strEmail)) <> 0 then
                        myMail.Send
                        end if
                        set myMail=nothing
                    
					
					
					
					end if
					
					end if '** Ændre PW **'

                    end if '** bruger aktiv

					
                  
  
	                
					
                    
					
					end if
					
					
					
					
					
		        '*** Opret medarbejer i DB **'
			    if func = "dbopr" then
    					
    					
                        
    					
    					
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





                         if cint(strAnsat) = 1 then 'kun mail til aktive brugere

                                
                              if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\medarb.asp" then


                            
                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject="Timeout - Medarbejder profil opdateret"
                                
                                            select case lto
                                            case "kejd_pb"
                                            myMail.From = "joedan@okf.kk.dk"
                                            case else
					                        myMail.From = "timeout_no_reply@outzource.dk"
				                            end select

                                    'myMail.To=strEmail
                                    if len(trim(strEmail)) <> 0 then
                                    myMail.To= ""& strNavn &"<"& strEmail &">"
                                    end if

                                    if lto = "mi" then
                                    myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20120106.pdf" 
                                    else
                                    myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20130102.pdf" 
                                    end if
                    
                              
                                    mailtextbody= "" & "Hej "& strNavn & vbCrLf _ 
					                & "Velkommen til TimeOut." & vbCrLf & vbCrLf _
					                & "Du er blevet oprettet som bruger i TimeOut" & vbCrLf _ 
					                & "Dit brugernavn er: " & strLogin & " og dit password er: " & strPw & vbCrLf & vbCrLf _ 
					                & "Gem disse oplysninger, til du skal logge ind i TimeOut."  & vbCrLf _ 
					                & "Du kan altid selv ændre dem når du er logget på systemet." & vbCrLf & vbCrLf 

                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
					                mailtextbody = mailtextbody & "Adressen til TimeOut er: https://outzource.dk/"&lto&""& vbCrLf & vbCrLf 
                                    else
                                    mailtextbody = mailtextbody & "Adressen til TimeOut er: http://timeout.cloud/"&lto&""& vbCrLf & vbCrLf 
                                    end if

					                mailtextbody = mailtextbody & "Med venlig hilsen"& vbCrLf & vbCrLf & strEditor & vbCrLf 

                        
                                    myMail.TextBody= mailtextbody

                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                   
                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                       smtpServer = "webout.smtp.nu"
                                    else
                                       smtpServer = "formrelay.rackhosting.com" 
                                    end if
                    
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update
                    
                                    if len(trim(strEmail)) <> 0 then
                                    myMail.Send
                                    end if
                                    set myMail=nothing


    					
					   
    					
					            end if
                        

                        end if
                    


    					

                     
    					
					                   ' if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\medarb_red.asp" then
    					
    					        
                                                'if len(trim(strEmail)) <> 0 then

					                            'If Mailer.SendMail Then
        				
						                '            call medarb_opr_ok
    						
					                            'Else
    				                            'Response.Write "Fejl...<br>" & Mailer.Response
  					                            'End if

                                                'else

                                                    '***call medarb_opr_ok***

                                                'end if
    					
					                    'else
    						
						                    '***call medarb_opr_ok***
    						
					                    'end if
    						
			                    end if
			                else
			                    '*** Hvis dublet fandtes ***
			

			                    if strMnrOK <> "y" then
			                    errortype = 10
			                    call showError(errortype)
                                response.end
			                    else
			
			                    if strLoginOK <> "y" then
			                    errortype = 11
			                    call showError(errortype)
                                response.end

						                    end if
					                    end if
				                    end if
			                    end if
		                end if

                %>
                <div class="container" style="width:500px;">
            <div class="porlet">
            
            <h3 class="portlet-title">
               <u>Medarbejder oprettet</u>
            </h3>
            
                <div class="portlet-body">
                    <div>Du har oprettet / opdateret en medarbejder. 
                    </div><br />
                    <div>
                       
                        <%if request("medarbtypligmedarb") = "1" then 'Medarbejdertype:medarbejder 1:1%> 
                        <b><a href="medarbtyper.asp?lastmedid=<%=id %>&mtypenavnforvlgt=Mtyp: <%=strNavn %>&func=opret">Videre til medarbejdertype (timepriser og normtid) >></a>    
                        <%else%>
                        <b><a href="medarb.asp?menu=medarbejder&lastmedid=<%=id %>">Videre >></a>
                        <%end if %>
                       
                    </div>
                    <br /><br />
                 </div>

            </div>
        </div>
                <%
               

        'Response.write "Medarbejder oprettet. <a href='medarb.asp?menu=medarbejder&lastmedid="&id&"'>Videre..</a>"
        'Response.Redirect "medarb.asp?menu=medarbejder&lastmedid="&id

         Response.end
                
                
        
    %>


<%
    case "red", "opret"

    %>
    <script src="js/medarb_jav.js" type="text/javascript"></script>        
    <%
     
	if func = "red" then

        HeaderTxt = "Redigér"


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
			
            strId =oRec("mid")
            strInit = oRec("init")
			intTimereg = oRec("timereg")
			ansatdato = formatdatetime(oRec("ansatdato"), 2) 'day(oRec("ansatdato")) &"-"& month(oRec("ansatdato")) &"-"& year(oRec("ansatdato"))
			opsagtdato = formatdatetime(oRec("opsagtdato"),2) 'day(oRec("opsagtdato")) &"-"& month(oRec("opsagtdato")) &"-"& year(oRec("opsagtdato"))
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



        else

        HeaderTxt = "Opret"

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
		strCRMcheckedTSA = "CHECKED"
		intSmil = 1
		strExch = ""
		strInit = ""
		intTimereg = "1"
		ansatDato = date()
		opsagtdato = "01-01-2044"
		nyhedsbrev = 0
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
        strCRMcheckedTSA_7 = ""
		
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
        case 7
        strCRMcheckedTSA_7 = "CHECKED"
		case else
		strCRMcheckedTSA = "CHECKED"
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

    

     <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=HeaderTxt & " Medarbejder"%></u>
        </h3>
        

    <form action="medarb.asp?func=<%=dbfunc%>" method="post">
    <input type="hidden" name="FM_smiley" id="Hidden1" value="1">
	<input type="hidden" name="FM_timereg" id="Hidden2" value="1">
	<input type="hidden" name="id" value="<%=id%>">
    
            <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2 pad-b10">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
            </div>
            </div>

        <div class="portlet-body">
         <section>
                <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well">Stamdata</h4>
                        </div>
                    </div>

                    <div class="row">
                         <div class="col-lg-1">&nbsp;</div>

                        <div class="col-lg-2 pad-t5">Navn:&nbsp<span style="color:red;">*</span>
                           

                        </div>
                        <div class="col-lg-3">   
                            <input name="FM_navn" type="text" class="form-control input-small" value="<%=strNavn %>" />
                        </div>
                        <!-- Sortering -->

                        <div class="col-lg-1 pad-t5">Intialer:</div>
                        <div class="col-lg-2">   
                            <input name="FM_init" type="text" class="form-control input-small" value="<%=strInit%>" />

                        </div>
                                
                           
                      
                        <!-- Id -->
                        <%if level <= 2 OR level = 6 then%>
                        <div class="col-lg-1 pad-t5">Medarb.nr:&nbsp<span style="color:red;">*</span></div>
                        <div class="col-lg-1">   
                            <input name="FM_Mnr" type="text" class="form-control input-small" value="<%=strMnr%>" /> 
                           
                        </div>
                        <%else %>
                        <input type="hidden" name="FM_Mnr" value="<%=strMnr%>">
                        <%end if %>
                    </div>
                    
                
            <!--Log ind-->
           
                     <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                </div>
                                <div class="col-lg-2 pad-t5">Login:&nbsp<span style="color:red;">*</span></div>
                                <div class="col-lg-3">   
                                    <input name="FM_login" type="text" class="form-control input-small" value="<%=strLogin %>" />
                                </div>
                                <div class="col-lg-1 pad-t5">Password:&nbsp<span style="color:red;">*</span></div>
                                <div class="col-lg-4">   
                                    <%if func = "red" then 
		                                strPw = "KEEPTHISPW99"
		                                else
		                                strPw = ""
		                                end if %>

                                    <input name="FM_pw" id="pw" type="password" class="form-control input-small" value="<%=strPw%>" />
                                </div>     
                     </div>

                    <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Status:</div>
                        <div class="col-lg-3">
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
		                        <input type="radio" name="FM_ansat" id="FM_ansat1" value="1" <%=chk1%>> Aktiv <%if func <> "red" then %>
                                <span style="font-size:10px;">(Der udsendes mail med logind info)</span>
                                <%end if %>
		                    <br /><input type="radio" name="FM_ansat" id="FM_ansat2" value="2" <%=chk2%>> Deaktiveret 
		                    <br /><input type="radio" name="FM_ansat" id="FM_ansat3" value="3" <%=chk3%>> Passiv
		                    <br />

                        </div>


                        <div class="col-lg-6">

                        <div id="deaktivnote" style="visibility:hidden; display:none; width:450px; padding:20px; background-color:#cccccc;">
                                Vær opmærksom på at når en medarbejder skifter til de-aktiveret, bliver opsagtdato sat til d.d.<br /><br />
                                <input type="checkbox" value="1" name="FM_deakdd_cancel" /> Klik af her, hvis opsagtdato ikke <b>skal sættes til d.d</b>

                            </div>
                         </div>

                    </div>
            
                </div>
             </section>

            <section>
                <!-- Accordion -->
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <!-- PersonData -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">
                            Persondata
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">
                        <div class="panel-body">
                                <div class="row">
                                 <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Adresse:</div>
                        <div class="col-lg-3">   
                                    <input name="FM_adr" type="text" class="form-control input-small" value="<%=madr%>"  />
                            </div>
                                    <div class="col-lg-6">&nbsp;</div> 
                        </div>
                <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                    </div>
                    <div class="col-lg-2">PostNr:</div>
                    <div class="col-lg-1">   
                                    <input name="FM_postnr" type="text" class="form-control input-small" value="<%=mpostnr%>" />           
                    </div>
                    <div class="col-lg-1">By:</div>
                    <div class="col-lg-3">   
                                    <input name="FM_city" type="text" class="form-control input-small" value="<%=mcity%>" />
                    </div>
                    <div class="col-lg-4">&nbsp;</div>   
                </div>

                    <div class="row">
                                    <div class="col-lg-1">   
                                            &nbsp;
                                            </div>
                        <div class="col-lg-2">Land:</div>
                        <div class="col-sm-3">
                        <select class="form-control input-small" name="FM_land" style="width:200px;">
						            <!--#include file="../timereg/inc/inc_option_land.asp"-->
                                <%if func = "red" then%>
                        		<option SELECTED><%=mland%></option>
		                        <%else%>
		                        <option SELECTED>Danmark</option>
		                        <%end if%>

		                </select>

                         </div>
                    </div>
                    <div class="row">
                                    <div class="col-lg-1">   
                                            &nbsp;
                                            </div>
                        <div class="col-lg-2">Telefon:</div>
                        <div class="col-lg-2"> 
                        <input name="FM_tel" type="text" class="form-control input-small" value="<%=mtlf%>"/>
                                  </div>
                        <div class="col-lg-7">&nbsp;</div>    
                    </div>
                         <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                    </div>
                    <div class="col-lg-2">Email:</div>
                    <div class="col-lg-3">   
                                    <input name="FM_email" type="text" class="form-control input-small" value="<%=strEmail%>" />
                                    <span style="color:#999999; font-size:10px; line-height:10px;">Hvis email er indtastet, og medarbejder er aktiv, bliver der sendt mail til modtager med logind information.</span>               
                    </div>
                                    <div class="col-lg-5">&nbsp;</div>     
                </div>
                    <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                    </div>
                    <div class="col-lg-2">Cpr.nr:</div>
                    <div class="col-lg-3">   
                                    <input name="FM_cpr" type="text" class="form-control input-small" value="<%=mcpr%>" />           
                    </div>
                                    <div class="col-lg-6">&nbsp;</div>
                </div>
                    <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                    </div>
                    <div class="col-lg-2">Registrerings nr. på køretøj:</div>
                    <div class="col-lg-3">   
                                    <input name="FM_regnr" type="text" class="form-control input-small" value="<%=mkoregnr%>" />  
                                          <div class="col-lg-8">   
                                          &nbsp;
                                         </div>         
                   
                    </div>
                </div>
                    
                          </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                      </div> <!-- /.panel -->

                    <!-- Medarbejder -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo">
                            Medarbejderindstillinger
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">
                          <div class="panel-body">
                            
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2">Ansat:</div>
                                  <div class="col-lg-2">
                                        <div class='input-group date'>
                                              <input class="form-control input-small" type="text" name="ansatdato" id="ansatdato" value="<%=ansatdato%>"/>
                                              <!--<span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>-->
                                        </div>

                                  </div>
                                  <div class="col-lg-1">Fratrådt:</div>
                                  <div class="col-lg-2">
                                      <div class='input-group date'>
                                      <input class="form-control input-small" type="text" name="opsagtdato" id="opsagtdato" value="<%=opsagtdato%>" />
                                           <!--<span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>-->
                                        </div>
                                  </div>
                                        <div class="col-lg-4">&nbsp</div>
                              </div>
                              
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2">Sprog:</div>
                                  <div class="col-lg-2"><select class="form-control input-small" name="FM_sprog" style="width:160px;">
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
	                            %></select>
                                  </div>

                              </div>
                              <% 
	                            if cint(nyhedsbrev) = 0 then
	                            chkny0 = "CHECKED"
	                            chkny1 = ""
	                            else
	                            chkny0 = ""
	                            chkny1 = "CHECKED"
	                            end if 
	                            %>	
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2 pad-t20 pad-b20">Nyhedsbrev:</div>
                                  <div class="col-lg-2 pad-t20 pad-b20">Ja tak <input id="nyhedsbrev" name="nyhedsbrev" type="radio" value="1" <%=chkny1 %> />&nbsp;&nbsp;Nej 
                                    <input id="nyhedsbrev" name="nyhedsbrev" type="radio" value="0" <%=chkny0 %> /> </div>
                              </div>

                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2">Exchange/Domæne Navn:</div>
                                  <div class="col-lg-6">
                                      <input class="form-control input-small" type="text" name="FM_exch" value="<%=strExch%>" placeholder="DomæneNavn"/>
                                  </div>
                                  
                              </div>
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2 pad-b20">Tillad medarbejder skifter TimeOut version. (hvis findes):</div>
                                  <div class="col-lg-3"><input type="checkbox" name="FM_visskiftversion" value="1" <%=visskiftversionCHK %>/> Ja, må gerne skifte internt (fra timereg. siden), uden at skulle logge ind igen  </div> 
                               <div class="col-lg-6">&nbsp</div>  
                              </div>
                              

                              <%if cint(level) = 1 then %>
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2">
                                      <%if level = 1 then %>
                                      <a href="medarbtyper.asp" target="_blank">Medarb.type:</a>
                                      <%else %>
                                      Medarb.type:
                                      <%end if %>
                                      &nbsp&nbsp<a data-toggle="modal" href="#styledModalSstGrp20"><span class="fa fa-info-circle"></span></a>
                                      
                                  </div>
                                    
                                    <div id="styledModalSstGrp20" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                        <div class="modal-dialog">
                                            <div class="modal-content" style="border:none !important;padding:0;">
                                              <div class="modal-header">
                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                  <h5 class="modal-title">Medarbejdertype</h5>
                                                </div>
                                                <div class="modal-body">
                                                Bruges til at bestemme hvor mange timer man er ansat og til hvilke timepriser medarbejderen faktures.

                                                </div>
                                            </div>
                                        </div>
                                 </div>

                                  <div class="col-lg-8"> 

                                      <% if func = "red" then
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
	                                    end if %>

                                  <select class="form-control input-small" <%=dis %> name="FM_medtype" id="FM_medtype" style="width:540px;">
                                     
                                    <%	strSQL = "SELECT medarbejdertype, mt.type, mt.id, mgruppe, "_
                                    &" mtg.navn AS mtgnavn "_ 
                                    &" FROM medarbejdere AS m, medarbejdertyper AS mt"_
                                    &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mgruppe) WHERE mid = "& id &" AND mt.id = medarbejdertype" 'AND mid = 99999999
			                        %>
            
            
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
                                      
                                      
                                      <%if func = "red" AND level = 1 then %>
                                      <a href="medarbtyper.asp?func=red&id=<%=medarbejdertype %>" target="_blank">Rediger medarbejdertype (timepris, normtid mm.)</a>
                                      <%else
                                          
                                      call medarbtypligmedarb_fn()%>
                                      <input type="hidden" name="medarbtypligmedarb" value="<%=medarbtypligmedarb %>" />   
                                      <%end if %>
                                  
                                
                              <%
                              else
		                        strSQL = "SELECT medarbejdertype, type, id FROM medarbejdere, medarbejdertyper WHERE Mid = "& id &" AND medarbejdertyper.id = medarbejdertype"

		                        oRec.open strSQL, oConn, 3
		                        if not oRec.EOF then%>
		                        <input type="hidden" name="FM_medtype" value="<%=oRec("id")%>">
		                        <%
		                        end if
		                        oRec.close
	                        end if
                            %>




                                  
                                <br /><br /><b>Medarbejdertype historik:</b> (maks 30)<br />
                                <span style="color:#999999;">Første type skal altid følge ansat dato. Derfor kan den ikke opdateres.<br />
                                Kun normtid følger historikken. Realiserede timer følger altid den gruppe der er valgt for medarbejdederen. Der kan ændres timepriser realiserede timer ved at redigere medarbejdertypen.</span> <br /><br />
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
                            <br /><br /><br />
                            Opdater eksisterende timepriser? (på den aktuelle medarbejdertype)<br />
                                 <span style="color:#999999;">Opdaterer priser på alle timeregistreringer, på alle job uanset status, for denne medarbejder til at følge den nuværende medarbejdertype indstilling, fra den dato medarbejderen er oprettet som denne type.</span><br />
                                <input type="checkbox" name="FM_opd_mty_kp" id="FM_opd_mty_kp" value="1"/> Opdater eksisterende <b>kostpriser</b> 
                            <br />
                                 <input type="checkbox" name="FM_opd_mty_tp" id="FM_opd_mty_tp" value="1"/> Opdater eksisterende <b>timepriser</b> (overskriver ikke indstillinger på job, men opdaterer timeprisen på eksisterende registreringer)
                               <br />&nbsp;
        
                                <%
                            end if
                                    %>
                             </div>
                             </div><!-- ROW -->


                               <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2"><br /><br />Systemrettigheder:&nbsp&nbsp<a data-toggle="modal" href="#styledModalSstGrp3"><span class="fa fa-info-circle"></span></a></div>
                                    <div id="styledModalSstGrp3" class="modal modal-styled fade" style="top:60px;">
                                                    <div class="modal-dialog">
                                                        <div class="modal-content" style="border:none !important;padding:0;">
                                                            <div class="modal-header">
                                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                                <h5 class="modal-title">Systemrettigheder</h5>
                                                            </div>
                                                            <div class="modal-body">
                                                                Bruges til at bestemme hvilke rettigheder og hvilke områder af TimeOut medarbejderen har adgang til. 
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>

                                  <div class="col-lg-4"><br /><br />
                                     <select class="form-control input-small" name="FM_bgruppe">
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
		                                if level <> 1 then
		                                strWHEREKlaus = " WHERE rettigheder = 7"
		                                else
		                                strWHEREKlaus = " WHERE rettigheder <> 0" '= 2 OR rettigheder = 6 "
		                                end if
		
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
                                  </div>
                                   <div class="col-lg-5">&nbsp</div>
                              </div>
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                   
                                  <div class="col-lg-2 pad-t10">
                                      <%if level = 1 then %>
                                      <a href="projektgrupper.asp?" target="_blank">Projektgrupper:</a>
                                      <%else %>
                                      Projektgrupper:
                                      <%end if %>&nbsp&nbsp<a data-toggle="modal" href="#styledModalSstGrp40"><span class="fa fa-info-circle"></span></a>

                                        <div id="styledModalSstGrp40" class="modal modal-styled fade" style="top:60px;">
                                                    <div class="modal-dialog">
                                                        <div class="modal-content" style="border:none !important;padding:0;">
                                                            <div class="modal-header">
                                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                                <h5 class="modal-title">Projektgrupper</h5>
                                                            </div>
                                                            <div class="modal-body">
                                                                 Projektgrupper bruges til at bestemme hvilke job og 
                                                                aktiviteter medarbejderen har adgang til på timeregisteringssiden, samt hvis Du er Teamleder hvilke medarbejdere Du har adgang til at se timer for.<br /><br />
                                                                Alle medarbejdere er <u>ALTID</u> med i "Alle-gruppen")  
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>

                                      <span style="font-size:10px;"></span>
                                  </div>
                                  
                                   <div class="col-lg-4 pad-t10">
                                       <input id="Hidden3" name="FM_progrp" type="hidden" value="10" /><!-- Alle gruppen -->
                    
                                       
                        <select id="FM_progrp" name="FM_progrp" multiple style="height:200px;" class="form-control input-small" <%=projGrpDis %>>
            
                        <%
    
                        selprogrp = ""
    
                        if func = "opret" then
                        thisMid = 0
                        else
                        thisMid = id
                        end if
    
                        strNOTpgids = ""
    
                        '*** Henter projektgrupper grupper medarbejder er med i ***'
                        strSQLpg = "SELECT Mid, MedarbejderId, p.navn AS pgnavn, p.id AS pid, pr.teamleder, pr.notificer FROM progrupperelationer pr "_
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

                            if oRec2("notificer") <> 0 then
                            notf = "-1"
                            notfTxt = " (notificer)"
                            else
                            notf = ""
                            notfTxt = ""
                            end if


	                    %>
                        <option value=<%=oRec2("pid") & teaml & notf %> SELECTED><%=oRec2("pgnavn") %> <%=teamlTxt %> <%=notfTxt %></option>
                        <%
    
                        selprogrp = selprogrp & ","& oRec2("pid")
    
	                    end if
	
	                    pgr = pgr + 1
	                    oRec2.movenext
	                    wend
	                    oRec2.close
	
	


	
	                        '*** Henter resterende grupper ***'
	                        strSQLrpg = "SELECT pg.navn, pg.id AS pid, opengp FROM projektgrupper pg WHERE pg.id <> 10 " & strNOTpgids & " ORDER BY pg.navn"
	                        'Response.Write strSQLrpg
	                        'Response.flush
	                        oRec2.open strSQLrpg, oConn, 3 
	                        while not oRec2.EOF 

                            if cint(level) = 1 OR oRec2("opengp") = 1 then
                            %>
                            <option value=<%=oRec2("pid")%>><%=oRec2("navn") %></option>

                            <%else
                                
                                if lto <> "wwf" then%>
                                <option value=<%=oRec2("pid")%> DISABLED><%=oRec2("navn") %></option>

                                <%end if
                             end if

	                        oRec2.movenext
	                        wend
	                        oRec2.close
	                        %>
                            </select>


                                   </div>
                                <div class="col-lg-7">&nbsp</div>      
                              </div>
                               <%if level = 1 then %>
                              <div class="row">
                                   <div class="col-lg-12"><br />&nbsp</div>
                                </div>
                              <div class="row">
                                  
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2 pad-t5">Medarbejders startside:</div>
                                  <div class="col-lg-5">
                                       <input type="radio" name="FM_tsacrm" value="0" <%=strCRMcheckedTSA%>> Timeregistrering (default) &nbsp  <input type="checkbox" name="FM_timer_ststop" value="1" <%=timer_ststopCHK%> /> Indtast timer som start og stop tid <br>
          
                                     <input type="radio" name="FM_tsacrm" value="6" <%=strCRMcheckedTSA_6%>> Ugeseddel<br>


                                        <%call erStempelurOn()
                
                                        if cint(stempelurOn) = 1 then %>
                                      <input type="radio" name="FM_tsacrm" value="3" <%=strCRMcheckedTSA_3%>> Stempelur (Komme / Gå)  <br>
                                        <%end if %>
		
		                             <input type="radio" name="FM_tsacrm" value="2" <%=strCRMcheckedRes%>> Ressourceplanner<br>
                                    <input type="radio" name="FM_tsacrm" value="4" <%=strCRMcheckedTSA_4%>> Igangværende job<br>
                                     <input type="radio" name="FM_tsacrm" value="5" <%=strCRMcheckedTSA_5%>> Joblisten<br>
                                      <input type="radio" name="FM_tsacrm" value="7" <%=strCRMcheckedTSA_7%>> Dashboard<br>

                                     <%if licensType = "CRM" then%>
                                     <input type="radio" name="FM_tsacrm" value="1" <%=strCRMcheckedCRM%>> CRM Kalender<br>
                                    <%end if%>
                                  </div>
                              <%end if %>
                              </div>



                                
                <%if func = "red" then %>
                              <br />
                <div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>

                            <%if func = "red" then %>
                            <span style="font-size:9px; color:#999999;"> 
                            TimeOut Id: <%=id %>
                            <%end if %></span><br /><br />&nbsp;

                <%else %>
                <br /><br />&nbsp;
                <%end if %>
                         


                          </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                      </div> <!-- /.panel -->


                   


   
             

                     </div> <!-- /.accordion -->
                
               </section>

                  <div style="margin-top:15px; margin-bottom:15px;">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                     
         
                    <div class="clearfix"></div>
                 </div>


        </div> <!-- /.portlet-body -->
        </form>
      </div> <!-- /.portlet -->
    </div> <!-- /.container -->

            <%

    case else 'list

                %>
                <script src="js/medarb_list_jav.js" type="text/javascript"></script>
                <%


                if len(trim(request("lastmedid"))) <> 0 then
                lastmedid = request("lastmedid")
                else
                lastmedid = 0
                end if

                if len(trim(request("visikkemedarbejdere"))) <> 0 then
                visikkemedarbejdere = 1
                else
                visikkemedarbejdere = 0
                end if


                if len(trim(request("sogsubmitted"))) <> 0 then
                sogsubmitted = 1
                else
                sogsubmitted = 0
                end if

   
	            if cint(sogsubmitted) = 1 then
                ''**AND len(request("ktype")) <> 0 
	            '** request("ktype") tjekker at form e submitted **'
	            thiskri = request("FM_soeg")
	            useKri = 1

                response.cookies("medarb_2015")("medarbsogekri") = thiskri
	           
	            else
	        
	                    if request.cookies("medarb_2015")("medarbsogekri") <> "" then
	                    thiskri = request.cookies("medarb_2015")("medarbsogekri")
	                    useKri = 1
	                    visikkemedarbejdere = 0
	                    else
	                    thiskri = ""
	                    useKri = 0
	                    end if
	        
	            end if

                if cint(sogsubmitted) = 1 then
                    if len(trim(request("FM_vispasogluk"))) <> 0 then
                    vispasogluk = 1
                    else
                    vispasogluk = 0
                    end if
                else
                    if request.cookies("medarb_2015")("vispasogluk") <> "" then
                    vispasogluk = request.cookies("medarb_2015")("vispasogluk")
                    else
                    vispasogluk = 0
                    end if
                end if

                response.cookies("medarb_2015")("vispasogluk") = vispasogluk
                response.cookies("medarb_2015").expires = date + 1

                if cint(vispasogluk) = 1 then
                vispasoglukCHK = "CHECKED"
                else
                vispasoglukCHK = ""
                end if

   
                
         if media <> "eksport" then %>



 <!--Medarbejder liste-->
           
<script src="js/Medarbejder.js" type="text/javascript"></script>

    <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Medarbejdere</u>
        </h3>
        
          <%if level = 1 then 
              
              
              strSQL = "SELECT licens.key, klienter FROM licens WHERE id = 1"
		    oRec.Open strSQL, oConn, 0, 1, 1
		    if not oRec.EOF then
		    'ikl = left(oRec("key"), 4)
		    'akl = right(ikl, 2)
		
		    akl = oRec("klienter")
		
		        call antalmedarb_fn(2)

		    end if
		    oRec.close
              
              if cint(kli) < cint(akl) then '* kun admin brugere kan oprette medarb.
              
              %>
          <form action="medarb.asp?menu=medarber&func=opret" method="post">
                <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                            <button class="btn btn-sm btn-success pull-right"><b>Opret ny +</b></button><br />&nbsp;
                            </div>
                        </div>
                </section>
         </form>
          <%end if %>

          <%end if %>

         
          <section>

                 <form action="medarb.asp?vismedarbejdere=1&sogsubmitted=1" method="POST">
                <div class="well">
                         <h4 class="panel-title-well">Søgefilter</h4>
                         <div class="row">
                             <div class="col-lg-3">
                                 Søg på Navn: <br />
			                    <input type="text" name="FM_soeg" id="FM_soeg" class="form-control input-small" placeholder="% Wildcard" value="<%=thiskri%>">
                               </div>
                            <div class="col-lg-9"><br />
                            <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis medarbejdere >></b></button>
                            </div>
                             </div>
                    <div class="row">
                        <div class="col-lg-3"><input id="jq_vispasogluk" name="FM_vispasogluk" type="checkbox" <%=vispasoglukCHK %> value="1" /> Vis alle (passive og de-aktiverede)</div>
                    </div>
                            
                    </div>
              </form>
 

        <div class="portlet-body">

           <table id="medarb_list" class="table dataTable table-striped table-bordered table-hover ui-datatable">
       
           
            <thead>
              <tr>
                <th style="width: 22%">Navn</th>
                <th style="width: 5%">Status</th>
                <th style="width: 15%">Type</th>
                <th style="width: 18%">Rettigheder</th>
                <th style="width: 15%">Email</th>
                <th style="width: 10%">Ansat</th>
                <th style="width: 10%">Login</th>
                <th style="width: 5%">Joblog</th>
                <th style="width: 5%">Slet</th>

              </tr>
            </thead>
              <tbody>

                <%

                end if 'media


                    'sort = Request("sort")
	                ekspTxt = ""
	
		
		            '** Søge kriterier **
		            if cint(visikkemedarbejdere) <> 1 then
		
			            if usekri = 1 then '**SøgeKri
			            sqlsearchKri = " (mnavn LIKE '"& thiskri &"%' OR init LIKE '"& thiskri &"%' OR (mnr LIKE '"& thiskri &"%' AND mnr <> '0'))" 
			            else
			            sqlsearchKri = " (mid <> 0)"
			            end if
			
		            else
		            sqlsearchKri = " (mid = 0)"
		            end if

                    if cint(vispasogluk) = 1 then
                    vispasoglukKri = " AND mansat <> -1"
                    else
                    vispasoglukKri = " AND mansat = 1"
                    end if

                     if level <> 1 then 
                     strSQLAdminRights = " AND mid = " & session("mid") 
                     end if
                     %>
                



                   <%strSQL = "SELECT mnavn, mid, email, lastlogin, init, mnr, mansat, brugergruppe, medarbejdertype, mt.type AS mtypenavn, b.navn AS brugergruppenavn, ansatdato FROM medarbejdere AS m "_
                    &"LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) "_
                    &"LEFT JOIN brugergrupper AS b ON (b.id = m.brugergruppe) WHERE "& sqlsearchKri &" "& vispasoglukKri &" "& strSQLAdminRights &" ORDER BY mnavn LIMIT 4000" 
        
                    'response.write strSQL
                    'response.flush

                    oRec.open strSQL, oConn, 3
        
                    while not oRec.EOF
                    
                       if len(trim(request("jq_visalle"))) <> 0 then
                      jq_vispasluk = request("jq_visalle")
                      else
                      jq_vispasluk = 0
                      end if

                      if jq_vispasluk <> "1" then
                      visAlleSQLval = " AND mansat = 1 "
                      else
                      visAlleSQLval = " AND mansat <> -1 "
                      end if

                        select case oRec("mansat")
                        case 3
                           mStatus = "Passiv"
                        case 1
                           mStatus = "Aktiv"
                        case 2
                           mStatus = "De-aktiv"
                        end select

                       mtypenavn = oRec("mtypenavn")
                       mBrugergruppe = oRec("brugergruppenavn")
                   

                          
                     if media <> "eksport" then
                       
                       
                       if cint(lastmedid) = cint(oRec("mid")) then
                       trBgCol = "#FFFFE1" 
                       else
                       trBgCol = ""
                       end if%>

                        <tr style="background-color:<%=trBgCol%>;">
                            
                            <td> <a href="medarb.asp?func=red&id=<%=oRec("mid") %>"><%=left(oRec("mnavn"), 30) %>
                                
                                   <%if len(trim(oRec("init"))) <> 0 then%>
                                   &nbsp;[<%=oRec("init") %>]
                                   <%end if %>
                                
                                </a></td>
                            <td><%=mStatus %></td>
                            <td><%=left(mtypenavn, 15) %></td>
                            <td><%=left(mBrugergruppe, 20) %></td>
                            <td><a href="mailto:<%=oRec("email") %>"><%=oRec("email") %></a></td>
                            <td><%=oRec("ansatdato") %></td>
                            <td><%=oRec("lastlogin") %></td>
                            <td style="text-align:center;"> <a href="../timereg/joblog.asp?menu=timereg&FM_medarb=<%=oRec("Mid")%>&FM_job=0&selmedarb=<%=oRec("Mid")%>" target="_blank"><span class="fa fa-external-link"></span></a></td>
                            <td style="text-align:center;"><a href="medarb.asp?menu=tok&func=slet&id=<%=oRec("mid")%>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
                        </tr>
               
                  <% end if 'media
                      
                      ekspTxt = ekspTxt & oRec("mnavn") & ";" & oRec("mnr") & ";" & oRec("init") & ";" & mStatus & ";" & oRec("mtypenavn") & ";" & oRec("brugergruppenavn") & ";" & oRec("email") & ";"& oRec("ansatdato") &"; xx99123sy#z"
                      oRec.movenext
                      wend
                      
                      oRec.close 
                      
                      
                      
             if media <> "eksport" then          
            %>


            </tbody>
               <tfoot>
                   <tr>
                    <th>Navn</th>
                    <th>Status</th>
                    <th>Type</th>
                    <th>Rettigheder</th>
                    <th>Email</th>
                    <th>Ansat</th>
                    <th>Sidste login</th>
                    <th>Joblog</th>
                    <th>Slet</th>
                   </tr>
               </tfoot>
          </table>

        </div> <!-- /.portlet-body -->
        </section>
        <%end if %>



<%
'************************************************************************************************************************************************
'******************* Eksport **************************' 

                if media = "eksport" then

    
   
    
                    call TimeOutVersion()
    
    
                    ekspTxt = ekspTxt 'request("FM_ekspTxt")
	                ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	
	
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\medarb.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                end if
				
				
				
				                file = "medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				                '**** Eksport fil, kolonne overskrifter ***
				
			
				                strOskrifter = "Medarbejder; Nr.; Init; Status; Type; Brugergruppe; Email; Ansatdato"
				
				
				
				
				                objF.WriteLine(strOskrifter & chr(013))
				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
				                <div style="position:absolute; top:100px; left:200px; width:300px;">
	                            <table border=0 cellspacing=1 cellpadding=0 width="100%">
	                            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                            <img src="../ill/outzource_logo_200.gif" />
	                            </td>
	                            </tr>
	                            <tr>
	                            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
                
                             
	                            <a href="../inc/log/data/<%=file%>" target="_blank" >Din CSV. fil er klar >></a>
	                            </td></tr>
	                            </table>
                                </div>
	            
	          
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	
				



                end if  
 '************************************************************************************************************************************************
              
          %>




          <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b>Funktioner</b>
                        </div>
                    </div>
                    <form action="medarb.asp?media=eksport" method="Post" target="_blank">
                  
                    <div class="row">
                     <div class="col-lg-12 pad-r30">
                         
                    <input id="Submit5" type="submit" value="Eksport til csv." class="btn btn-sm" /><br />
                    <!--Eksporter viste kunder og kontaktpersoner som .csv fil-->
                         
                         </div>


                </div>
                </form>
                
            </section>

      </div><!-- /.portlet -->
    </div> <!-- /.container -->








<%end select %>

  </div> <!-- .content -->
</div> <!-- /#wrapper -->


<!--#include file="../inc/regular/footer_inc.asp"-->