<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/kontakter_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->




<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	
	if len(request("id")) <> 0 then
	id = request("id")
	else
	id = 0
	end if
	
	if len(request("kid")) <> 0 then
	kid = request("kid")
	else
	kid = 0
	end if
	
	intKid = kid
	
	rdir = request("rdir")
	

    public skptypeTxt
    function kptyp_sb(kptype)

        skptypeTxt = ""
        if cint(kptype) <> 0 then
            
            if cint(kptype) = 1 then
            kptypeTxt = "Faktura adr."
            else
            kptypeTxt = "Lev. adr."
            end if

        end if 
        
        if len(trim(kptypeTxt)) <> 0 then
        Response.write "<br>"& kptypeTxt
        end if

    end function

	
	select case func 
	case "slet"
	
	strSQL = "DELETE FROM kontaktpers WHERE id = "& id
	oConn.execute(strSQL)
	
	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
				
	
	case "dbred", "dbopr"
	
	
	function SQLBless2(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, "'", "''")
	SQLBless2 = tmp
	end function
			
			
			if len(trim(request("FM_navn_kpers"))) = 0 then
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<%
			useleftdiv = "j2"
			errortype = 138
			call showError(errortype)
			
			Response.end
			end if
			
			
			'******************** Opretter kontaktpersoner ***********************
				
				
					useUpdateDato = year(now)&"/"&month(now)&"/"&day(now)
	
					thiskpersId = id
					
					if thiskpersId <> 0 then
					'** finder eksiterede pw, så der kan tjekkes om det er ændret og om der skal afsendes email ***'
					            
					             ekspw = ""
					             
					             strSQL = "SELECT password FROM kontaktpers WHERE id = " & thiskpersId
                                 oRec.open strSQL, oConn, 3
                                 if not oRec.EOF then
                                 ekspw = oRec("password")
                                 end if
                                            
                                 oRec.close
					
					end if
					
					
					kundeId = kid
					
					strKpers = SQLBless2(request("FM_navn_kpers"))
					strKperstit = SQLBless2(request("FM_titel_kpers"))
					strKpersdTlf = SQLBless2(request("FM_tlf_kpers"))
					strKpersMobil = SQLBless2(request("FM_mtlf_kpers"))
					strKpersPrivat = SQLBless2(request("FM_ptlf_kpers"))
					strKpersEmail = SQLBless2(request("FM_email_kpers"))
					strKperspw = SQLBless2(request("FM_pw_kpers"))
					strAdr_kpers = SQLBless2(request("FM_adr_kpers"))
					strPostnr_kpers = SQLBless2(request("FM_postnr_kpers"))
					strCity_kpers = SQLBless2(request("FM_city_kpers"))
					strLand_kpers = SQLBless2(request("FM_land_kpers"))
					strAf_kpers = SQLBless2(request("FM_afdeling_kpers"))
					strBeskkpers = SQLBless2(Request("FM_besk_kpers"))
					
                    kpean = SQLBless2(request("FM_pkean"))
                    kpcvr = SQLBless2(request("FM_pkcvr"))
					kptype = request("FM_kptype")

                    '**** Nulstiller forvalgte faktura adr / Leverings adr *****'
                    if cint(kptype) = 1 then
                    strSQLfvFak = "UPDATE kontaktpers SET kptype = 0 WHERE kptype = 1 AND kundeId = "& kundeId
                    oConn.execute(strSQLfvFak)
                    end if

                    if cint(kptype) = 2 then
                    strSQLfvFak = "UPDATE kontaktpers SET kptype = 0 WHERE kptype = 2 AND kundeId = "& kundeId
                    oConn.execute(strSQLfvFak)
                    end if

                    '************************************************************


					if func = "dbopr" then
					
					strSQLKpers = "INSERT INTO kontaktpers ("_
					&" kundeid, editor, dato, navn, titel, adresse, "_
					&" postnr, town, land, afdeling, email, password, "_
					&" dirtlf, mobiltlf, privattlf, beskrivelse, kpean, kptype, kpcvr) VALUES "_
					&" ("& kundeId &", '"& session("user") &"', '" & useUpdateDato &"', '" & strKpers &"', '"& strKperstit &"', "_
					&" '"& strAdr_kpers &"', '"& strPostnr_kpers &"', '"& strCity_kpers &"', "_
					&" '"& strLand_kpers &"', '"& strAf_kpers &"', "_ 
					&" '"& strKpersEmail &"', '"& strKperspw &"', '"& strKpersdTlf &"', "_
					&" '"& strKpersMobil &"', '"& strKpersPrivat &"', '" & strBeskkpers &"', '"& kpean &"', "& kptype &", '"& kpcvr &"')"
					
					else
					
					strSQLKpers = "UPDATE kontaktpers SET "_
					&" kundeid = "& kundeId &", editor ='"& session("user") &"', dato = '" & useUpdateDato &"', "_
					&" navn = '" & strKpers &"', titel = '"& strKperstit &"', adresse = '"& strAdr_kpers &"', "_
					&" postnr = '"& strPostnr_kpers &"', town = '"& strCity_kpers &"', land = '"& strLand_kpers &"', "_
					&" afdeling = '"& strAf_kpers &"' , email = '"& strKpersEmail &"', password = '"& strKperspw &"', "_
					&" dirtlf = '"& strKpersdTlf &"', mobiltlf = '"& strKpersMobil &"', privattlf = '"& strKpersPrivat &"', "_
					&" beskrivelse = '" & strBeskkpers &"', kpean = '"& kpean &"', kptype = "& kptype &", kpcvr = '"& kpcvr &"'"_
					&" WHERE id = " & thiskpersId 
					
					end if
					
					'response.write strSQLKpers
					
                    'Response.flush
                    oConn.execute(strSQLKpers)
					
					
					 'if lto = "outz" then        
                               
					if len(trim(strKpersEmail)) <> 0 AND len(trim(strKperspw)) <> 0 AND ekspw <> strKperspw then
					 '*** Sender mail, hvis der er givet adgang til kunde område for kpers..**
                               '***** Oprettter Mail object ***
                               if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\kontaktpers.asp" then
                                
                              
           
                                 Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
                                 ' Sætter Charsettet til ISO-8859-1 
                                 Mailer.CharSet = 2
                                 ' Afsenderens navn 
                                 Mailer.FromName = "TimeOut ServiceDesk"
                                 ' Afsenderens e-mail 
                                 Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                 Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
                                 Mailer.ContentType = "text/html"
                                 
                                 
                                  '** Henter key ***
                                 lkey = ""
                                 strSQL = "SELECT l.key as lkey FROM licens l WHERE id = 1"
                                 oRec.open strSQL, oConn, 3
                                 if not oRec.EOF then
                                 lkey = oRec("lkey")
                                 end if
                                            
                                 oRec.close
                                
                          
                          
                                '** Henter Afsender ***
                                 strSQL = "SELECT mnavn, mnr, init, email FROM medarbejdere WHERE mid = " & session("mid")
                                 oRec.open strSQL, oConn, 3
                                 if not oRec.EOF then
                                            
                                 afsNavn = oRec("mnavn") & "("& oRec("mnr") &") - " & oRec("init")
                                 afsEmail = oRec("email")
                                            
                                 
                                 end if
                                            
                                 oRec.close
                                
                                
                                'response.Write "ok - afsender navn:" & afsNavn
                                'response.end 
                                
                            
                                 
                                 
                                'Mailens emne
                                 Mailer.Subject = "Du er blevet tildelt adgang til vores ServiceDesk"
                                 ' Modtagerens navn og e-mail
                                 Mailer.AddRecipient strKnavn, strKpersEmail
                                 
                          
                                 
                                 ' Selve teksten
                                                      Mailer.BodyText = "Hej  "& strKpers &",<br><BR>" _ 
                                                      &" Du modtager denne mail fordi du er blevet tildelt adgang / har fået opdateret adgang til vores ServiceDesk.<br><br>"_
                                                      & "I vores ServiceDesk kan du oprette Incidents og følge med i status på igangværende Incidents.<BR>" _
                                                      & "Du kan også følge med i timeforbrug på job og aftaler hvortil du er blevet tildelt adgang.<br><br>" _
                                                      & "Du finder vores ServiceDesk her:<br>"_
                                                      & "<a href=""https://outzource.dk/timeout_xp/wwwroot/ver2_1/login_kunder.asp?key="& lkey &"&lto="& lto &""">"_
                                                      & "https://outzource.dk/timeout_xp/wwwroot/ver2_1/login_kunder.asp?key="& lkey &"&lto="& lto &"</a><br><br>"_
                                                      & "Dit logind og password er:<br>"_
                                                      & ""& strKpersEmail &" / " & strKperspw & "<br>"_
                                                      & "<br><br><br>"_
                                                      & "Med venlig hilsen<br>" & afsNavn & ", "& afsEmail & "<br>"_ 
                                                      '& "Incident ansvarlig: " & ansvNavn & ", " & ansvEmail & vbCrLf & vbCrLf _ bcc
                                                      '& "Adressen til TimeOut er: https://outzource.dk/"&lto&"" & vbCrLf & vbCrLf _ 
                                            
                                                      If Mailer.SendMail Then
                                            
                                                      Else
                                                      Response.Write "Fejl...<br>" & Mailer.Response
                                                      End if
                                                      
                             
                            
                             end if     '** C drev: mailer ** 		
							 
							 %>
							 <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
                             <script language="javascript">
                                 function lukogvidere() {
                                     window.opener.location.reload();
                                     window.close();
                                 }
            </script>


							 <%
							 
							 call sideinfo(50,40,400)
							 %>
                             
                             Der er afsendt en email til <b><%=strKpers%></b> med adgangs <br />
                             oplysninger til TimeOut ServiceDesk kundeområde.
                             
                             <a href="#" onclick="lukogvidere()" class=vmenu>Ok, videre >></a>
                             </td></tr>
                             </table>
                             </div>
                             
                             <%
							
				
				    else	
					
					if rdir <> "treg" AND rdir <> "fak" then
					Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
				    end if
				    
				    Response.Write("<script language=""JavaScript"">window.close();</script>")
						
				    
				    end if 'pw + email
				   'end if
				
				
				
	case "list"
	
	        %>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		   
			
			 <%
		
                 
               call menu_2014()  
                 
                 	 
	'oimg = "ikon_kontaktpers_48.png"
	oleft = 40
	otop = 70
	owdt = 300
	oskrift = "Kontaktpersoner / filialer"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
			 
    'url = "#"
	'text = "Opret ny kontaktperson / filial"
	'otoppx = 40
    'oleftpx = 610
    'owdtpx = 210
    'java = "popUp('kontaktpers.asp?id=0&kid=0&func=opr','550','450','150','120');"
	
	'call opretNyJava(url, text, otoppx, oleftpx, owdtpx, java)
    
        %>
<div style="position:absolute; left:740px; top:90px;">

<%    

                nWdt = 180
                nTxt = "Opret ny kontaktperson / filial"
                nLnk = "kontaktpers.asp?id=0&kid=0&func=opr"
                nTgt = "_blank"
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) 

    %>
    </div>
    
    <%


    tTop = 82
	tLeft = 40
	tWdth = 800
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	
	<table width=100% cellspacing=0 cellpadding=6 border=0>
	<tr bgcolor="#8caae6">
	<td class=alt>Kontakt</td>
	<td class=alt>Kontaktperson</td>
	<td class=alt>Email</td>
	<td class=alt>Tel. dir.</td>
	<td class=alt>Tel. mobil</td>
	<td class=alt>Adgang til ServiceDesk</td>
	</tr>
	<%
	
		strSQLKpers = "SELECT id, kundeid, navn, kp.adresse, kp.postnr, kp.town, "_
		&" kp.land, kp.afdeling, kp.email, kp.password, kp.dirtlf, kp.mobiltlf, kp.privattlf, "_
		&" kp.titel, kkundenavn, kid, kptype FROM kontaktpers kp "_
		&" LEFT JOIN kunder k ON (kid = kundeid) WHERE id <> 0 ORDER BY kkundenavn, navn"
		
		'Response.Write strSQLKpers
		'Response.flush
		x = 0
		oRec2.open strSQLKpers, oConn, 3	
		while not oRec2.EOF 				
		
		    select case right(x,1)
		    case 0,2,4,6,8
		    bgcol = "#ffffff"
		    case else
		    bgcol = "#Eff3ff"
		    end select
		    
		strKpers = oRec2("navn") 'strKpers1
		strKperstit = oRec2("titel") 'str_titel_kpers1
		strKpersTlf = oRec2("privattlf")
		strKpersmTlf = oRec2("mobiltlf")
		strKpersdTlf = oRec2("dirtlf")
		strKpersEmail = oRec2("email")
		strKperspw = oRec2("password")
		strAdr_kpers = oRec2("adresse")
		strPostnr_kpers = oRec2("postnr")
		strCity_kpers = oRec2("town")
		strLand_kpers = oRec2("land")
		strAf_kpers = oRec2("afdeling")
		thisKid = oRec2("kundeid")
        kptype = oRec2("kptype")
		

		%>
		<tr bgcolor="<%=bgcol %>">
		<td style="border-bottom:1px #cccccc solid; padding:4px;"><%=oRec2("kkundenavn") %>&nbsp;
        <b><%
       call kptyp_sb(kptype)
        %></b></td>
		<td style="border-bottom:1px #cccccc solid;">
		<a href="javascript:popUp('kontaktpers.asp?id=<%=oRec2("id")%>&kid=<%=oRec2("kid")%>&func=red','650','750','150','120');" target="_self" class=vmenu><%=strKpers %></a>&nbsp;</td>
		<td style="border-bottom:1px #cccccc solid;"><%=strKpersEmail %>&nbsp;</td>
		<td style="border-bottom:1px #cccccc solid;"><%=strKperdTlf %>&nbsp;</td>
		<td style="border-bottom:1px #cccccc solid;"><%=strKpersmTlf %>&nbsp;</td>
		<td style="border-bottom:1px #cccccc solid;">
		<% if len(trim(oRec2("password"))) <> 0 AND len(trim(oRec2("email"))) <> 0 then%>
		Ja
		<%else %>
		<%end if %>
		&nbsp;</td>
		</tr>
		<%
		
		x = x + 1
		oRec2.movenext
		wend
		oRec2.close		
	    
	    %>
	    </table>
	    </div>
	    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	    &nbsp;
	    
	    <%
	    
	case "opr", "red"
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%

	'***************** Kontaktpersoner *************************************************
	%>
	<div id=kpers name=kpers style="position:absolute; left:20; top:5; visibility:visible; display:;">
	<%
	
	
		strSQLKpers = "SELECT id, kundeid, navn, adresse, postnr, town, land, afdeling, email, password, dirtlf, mobiltlf, "_
        &" privattlf, beskrivelse, titel, kpean, kptype, kpcvr FROM kontaktpers WHERE id ="& id
		oRec2.open strSQLKpers, oConn, 3	
		if not oRec2.EOF then				
		
		strKpers = oRec2("navn") 'strKpers1
		strKperstit = oRec2("titel") 'str_titel_kpers1
		strKpersTlf = oRec2("privattlf")
		strKpersmTlf = oRec2("mobiltlf")
		strKpersdTlf = oRec2("dirtlf")
		strKpersEmail = oRec2("email")
		strKperspw = oRec2("password")
		strAdr_kpers = oRec2("adresse")
		strPostnr_kpers = oRec2("postnr")
		strCity_kpers = oRec2("town")
		strLand_kpers = oRec2("land")
		strAf_kpers = oRec2("afdeling")
		strBeskkpers = oRec2("beskrivelse")
		thisKid = oRec2("kundeid")
		kpean = oRec2("kpean") 
        kptype = oRec2("kptype")
        kpcvr = oRec2("kpcvr")

		end if
		oRec2.close
		
		if id <> 0 then
		dbfunc = "dbred"
		thisKid = thisKid
		else
		dbfunc = "dbopr"
		thisKid = kid
		end if


        kptypSEL0 = ""
        kptypSEL1 = ""
        kptypSEL2 = ""
        select case kptype 
        case 0
        kptypSEL0 = "SELECTED"
        case 1
        kptypSEL1 = "SELECTED"
        case 2
        kptypSEL2 = "SELECTED"
        end select
		
		
	
	%>
	<br>
	<h3>Kontaktperson / Filial:</h3>
	

    <%
    
     tTop = 0
	tLeft = 0
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
    
    %>

    <form method=post action="kontaktpers.asp?func=<%=dbfunc%>&id=<%=id%>&rdir=<%=rdir %>">
	<table cellspacing="0" cellpadding="1" border="0" width=100% bgcolor="#EFF3FF">
	
	<tr>
		<td valign="top" rowspan=18 style=""><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td style="padding-left:8px; "><font color="red">*</font> Navn:&nbsp;<br />
        <span style="color:#999999;">Navn eller firmanavn,<br /> &lt;br&gt; for linjeskift</span></td>
		<td style="padding-left:8px; padding-top:5px; "><textarea name="FM_navn_kpers" rows="2" style="width:300px;"><%=strKpers%></textarea></td>
		<td valign="top" rowspan=18 align="right" style="">&nbsp;
        <%if id <> 0 then%>
	<a href="kontaktpers.asp?func=slet&id=<%=id%>" class=red>[X]</a>
	<%end if%>
    &nbsp;
        </td>
	</tr>
    	<tr>
		<td style="padding-left:8px;">Titel:&nbsp;<br />
         <span style="color:#999999;">Ell. beskrivelse / kundenr.</span></td>
		<td style="padding-left:8px;"><input type="text" name="FM_titel_kpers" value="<%=strKperstit%>" size="15" style="width:300px;"></td>
	</tr>
	<tr>
		<td style="padding-left:8px;">Kunde tilhørsforhold:&nbsp;</td>
		<td style="padding-left:8px;">
		<select id="kid" name="kid" style="width:300px;">
			<%
			strSQL = "SELECT kkundenavn, kkundenr, kid FROM kunder WHERE kid <> 0 ORDER BY kkundenavn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
				if cint(thisKid) = cint(oRec("kid")) then
				selThis = "SELECTED"
				else
				selThis = ""
				end if%>
			<option value="<%=oRec("kid")%>" <%=selThis %>><%=oRec("kkundenavn")%>  (<%=oRec("kkundenr")%>)</option>
			<%
			oRec.movenext
			wend
			oRec.close 
			%>
		</select>
		&nbsp;<a href="kunder.asp?func=opret&ketype=k&hidemenu=1&showdiv=onthefly&rdir=1" target="_blank" class=vmenu>Opret ny her..</a>
		</td>
	</tr>
	<tr>
		<td style="padding-left:8px; padding-top:5px;" valign=top>Adresse:&nbsp;</td>
		<td style="padding-left:8px;"><textarea name="FM_adr_kpers" style="height:50px; width:300px;"><%=strAdr_kpers%></textarea></td>
	</tr>
	<tr>
		<td style="padding-left:8px;">Postnr:&nbsp;</td>
		<td style="padding-left:8px;"><input type="text" name="FM_postnr_kpers" value="<%=strPostnr_kpers%>" size="5" style="width:200px;"></td>
	</tr>
    <tr>
		<td style="padding-left:8px;">By:&nbsp;</td>
		<td style="padding-left:8px;"><input type="text" name="FM_city_kpers" value="<%=strCity_kpers%>" size="32" style="width:300px;"></td>
	</tr>
	<tr>
		<td style="padding-left:8px;">Land:&nbsp;</td>
		<td style="padding-left:8px;"><select name="FM_land_kpers" style="width:300px;">
		<%if func = "red" then%>
		<option checked><%=strLand_kpers%></option>
		<%else%>
		<option checked>Danmark</option>
		<%end if%>
		<!--#include file="inc/inc_option_land.asp"-->
		</select></td>
	</tr>
	<tr>
		<td style="padding-left:8px;">Afdeling&nbsp;</td>
		<td style="padding-left:8px;"><input type="text" name="FM_afdeling_kpers" value="<%=strAf_kpers%>" size="44" style="width:300px;"></td>
	</tr>


    
    	<tr>
		<td style="padding-left:8px;">EAN:&nbsp;</td><td style="padding-left:8px;"><input type="text" name="FM_pkean" value="<%=kpean%>" size="18" style="width:300px;"></td>
	</tr>
    	<tr>
		<td style="padding-left:8px;">CVR:&nbsp;</td><td style="padding-left:8px;"><input type="text" name="FM_pkcvr" value="<%=kpcvr%>" size="18" style="width:300px;"></td>
	</tr>
    	<tr>
		<td style="padding-left:8px;">Type&nbsp;</td><td style="padding-left:8px;">
        <select name="FM_kptype" style="width:300px;">
            <option value="0" <%=kptypSEL0 %>>Kontaktperson</option>
            <option value="1" <%=kptypSEL1 %>>Faktura adr. (forvalgt)</option>
            <option value="2" <%=kptypSEL2 %>>Leverings adr. (forvalgt)</option>
        </select></td>
	</tr>


	<tr>
		<td style="padding-left:8px;">Tlf: (direkte)&nbsp;</td><td style="padding-left:8px;"><input type="text" name="FM_tlf_kpers" value="<%=strKpersdTlf%>" size="18" style="width:300px;"></td>
	</tr>
	<tr>		
		<td style="padding-left:8px;">Tlf: (privat)&nbsp;</td><td style="padding-left:8px;"><input type="text" name="FM_ptlf_kpers" value="<%=strKpersTlf%>" size="18" style="width:300px;"></td>
	</tr>
	<tr>
		<td style="padding-left:8px;">Mobil:&nbsp;</td><td style="padding-left:8px;"><input type="text" name="FM_mtlf_kpers" value="<%=strKpersmTlf%>" size="18" style="width:300px;"></td>
	</tr>
	<tr>
		<td style="padding-left:8px;">Email: (login)&nbsp;</td><td style="padding-left:8px;"><input type="text" name="FM_email_kpers" value="<%=strKpersEmail%>" size="30" style="width:300px;"></td>
	</tr>
	<tr>
		<td style="padding-left:8px;">Password:&nbsp;</td><td style="padding-left:8px;"><input type="text" name="FM_pw_kpers" value="<%=strKperspw%>" size="18" style="width:300px;"></td>
	</tr>
	<tr>
		<td colspan=2 style="padding-left:8px;">(Password medfører adgang til TimeOuts kunde-del)</td>
	</tr>
	<tr>
		<td colspan=2 valign=top style="padding-left:8px;"><br>Beskrivelse:<br>
		<textarea cols="50" rows="4" name="FM_besk_kpers"><%=strBeskkpers%></textarea><br />&nbsp;</td>
	</tr>
    </table>

	<table width=100%>
    <tr>
		<td colspan=2 align=right style="padding:20px 30px 2px 2px;">
		<input type=submit value=" Gem >> " /></td>
	</tr>
	</table>
	
	</form>
	</div>
	</div>	
<%
case else
	
	%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
 

	<%
    call menu_2014()

	call kundemenu("kpers")
	
	'***************** Kontaktpersoner *************************************************
    
    
    if kid <> 0 then

    'url = "#"
	'text = "Opret ny kontaktperson / filial"
	'otoppx = 102
    'oleftpx = 830
    'owdtpx = 210
    'java = "popUp('kontaktpers.asp?id=0&kid="&kid&"&func=opr','650','750','150','120');"
	
	'call opretNyJava(url, text, otoppx, oleftpx, owdtpx, java)

        %>
        <div style="position:absolute; left:920px; top:124px;">
            <%
                nWdt = 180
                nTxt = "Opret ny kontaktperson / filial"
                nLnk = "kontaktpers.asp?id=0&kid="&kid&"&func=opr"
                nTgt = "_blank"
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) 
	%></div><%
	end if
	
	
	%>
	<div id=kpers name=kpers style="position:absolute; left:90px; top:126px; visibility:visible; display:; width:800px;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan=2 valign=top style=" "><img src="../ill/blank.gif" width="8" height="1" alt="" border="0"></td>
		<td valign="top" colspan=2 style=""><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan=2 valign=top style=" "><img src="../ill/blank.gif" width="8" height="1" alt="" border="0"></td>
	</tr>
	
	</table>
	
	<%
	'**Søg
	if len(request("FM_kpers_sog")) <> 0 then
	sogKri = request("FM_kpers_sog")
	strSQLWclaus = " (p.navn LIKE '"& sogKri &"%' OR p.email LIKE '"& sogKri &"%' OR p.afdeling LIKE '"& sogKri &"%') "
	else
	sogKri = ""
	strSQLWclaus = " p.id <> 0 "
	end if
	
	if len(request("FM_kundkunde")) <> 0 then
	strSQLWclaus = strSQLWclaus
	vallkunderChk = "CHECKED"
	else
	vallkunderChk = ""
	strSQLWclaus = strSQLWclaus & " AND p.kundeid = "& kid &""
	end if
	
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
	<form action="kontaktpers.asp?menu=kund&kid=<%=kid%>&FM_soeg=<%=thiskri%>" method="post">
	<tr bgcolor="5582D2">
		<td class='alt' style="  padding:20px;"><b>Søg:</b> (Navn, email, afd.) <input type="text" value="<%=sogKri%>" name="FM_kpers_sog" id="FM_kpers_sog" style="width:200px;">&nbsp;<input type="submit" value=" Søg >> ">
		<br><input type="checkbox" name="FM_kundkunde" id="FM_kundkunde" value="1" <%=vallkunderChk%>> Vis alle kontaktpersoner uanset kunde-tilhørsforhold. </td>
	</tr>
	</form>
	</table>
	
	<table cellspacing="1" cellpadding="0" border="0" width=100% bgcolor="#cccccc">
	
	<%
		
		
		strSQLKpers = "SELECT p.id, p.kundeid, p.navn, p.adresse, p.postnr, p.town, "_
		&" p.land, p.afdeling, p.email, p.password, p.dirtlf, p.mobiltlf, p.privattlf, "_
		&" p.beskrivelse, p.titel, k.kkundenavn, kpean, kptype FROM kontaktpers p "_
		&" LEFT JOIN kunder k ON (k.kid = p.kundeid) WHERE "& strSQLWclaus &" AND navn <> '0' ORDER BY p.navn"  '"& useThisKpers
		oRec2.open strSQLKpers, oConn, 3	
		t = 0
		while not oRec2.EOF 			
		
		intKpersId = oRec2("id")
		strKpers = oRec2("navn") 'strKpers1
		strKperstit = oRec2("titel") 'str_titel_kpers1
		strKpersTlf = oRec2("privattlf")
		strKpersmTlf = oRec2("mobiltlf")
		strKpersdTlf = oRec2("dirtlf")
		strKpersEmail = oRec2("email")
		strKperspw = oRec2("password")
		strAdr_kpers = oRec2("adresse")
		strPostnr_kpers = oRec2("postnr")
		strCity_kpers = oRec2("town")
		strLand_kpers = oRec2("land")
		strAf_kpers = oRec2("afdeling")
		strBeskkpers = oRec2("beskrivelse")
		strKundenavn = oRec2("kkundenavn")
		
        kpean = oRec2("kpean")
        kptype = oRec2("kptype")

		select case right(t, 1)
		case 0
		%><tr><%
		case 4
		%></tr><tr><%
		t = 0
		end select%>
		
		<td bgcolor="#ffffff" width=150 height=140 valign=top style="padding:5px;">
		<a href="javascript:popUp('kontaktpers.asp?id=<%=intKpersId%>&kid=<%=kid%>&func=red','650','750','150','120');" target="_self" class=vmenu><%=strKpers%></a><br>
		<b><%=strKundenavn%></b>
        <span style="background-color:#FFFFe1;"><b><%call kptyp_sb(kptype) %></b></span>
		<br>
		<%if len(strKperstit) <> 0 then%>
		<u>Titel:</u> <%=strKperstit%><br>
		<%end if%>
		<%if len(strAf_kpers) <> 0 then%>
		<u>Afdeling:</u> <%=strAf_kpers%><br>
		<%end if%>
		
		<%if len(strAdr_kpers) <> 0 then%>
		<u>Adresse:</u><br>
		<%=strAdr_kpers%><br>
		<%=strPostnr_kpers%>, <%=strCity_kpers%><br>
		<%=strLand_kpers%><br>
		<%end if%>

        <%if len(trim(kpean)) <> 0 then %>
		EAN: <%=kpean %><br />
        <%end if %>
		
        

		<u>Kontakt info:</u><br>
		<%if len(strKpersdTlf) <> 0 then%>
		Tlf (dir.): <%=strKpersdTlf%><br>
		<%end if%>
		
		<%if len(strKpersmTlf) <> 0 then%>
		Mobil: <%=strKpersmTlf%><br>
		<%end if%>
		
		<%if len(strKpersEmail) <> 0 then%>
			<a href="mailto:<%=strKpersEmail%>" class=kal_g>
			<%if len(strKpersEmail) > 23 then
			Response.write left(strKpersEmail, 23) & ".."
			else
			Response.write strKpersEmail
			end if%></a><br>
		<%end if%>
		
		<%
		
		t = t + 1
		oRec2.movenext
		wend 
		oRec2.close
	
	
	if t = 0 then%>
	<tr><td height=100 bgcolor="#ffffff" style="padding:10px;">Der er ikke oprettet nogen kontaktpersoner på denne kunde.<br>
	Opret ny kontaktperson ved at klikke på linket oppe til højre.
		
		<%if kid = 0 then%>
		<font color=red><br><br><b>Der kan først tilknyttes kontaktpersoner<br> når kunden er oprettet.</b></font>
		<%end if%>
		
	</td>
	<%end if%>
	</tr>
	</table>
	</div>

       
		
	<%
end select
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->