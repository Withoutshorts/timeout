


    <%
        
        public matUnderMinLagerTXT
        function matUnderMinLager(lto)


            '*** kun udesende 1 gang om dagen ***'
            dd = year(now) & "-" & month(now) & "-"& day(now)                                                        
            notificerEmailMuMinLager = 0
            matlageruminemailnotDate = "2002-01-01"

            strSQLdd = "SELECT matlageruminemailnot FROM licens WHERE id = 1"
            
            oRec.open strSQLdd, oConn, 3
            if not oRec.EOF then
            
            notificerEmailMuMinLager = 1  
            matlageruminemailnotDate = oRec("matlageruminemailnot")

            end if
            oRec.close


            if cDate(dd) > cDate(matlageruminemailnotDate) then
            

            strSQLmedarb = "SELECT mid, email, mnavn FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"
            oRec3.open strSQLmedarb, oConn, 3
            
            while not oRec3.EOF


                antalM = 0
                matUnderMinLagerTXT = ""
                strSQL = "SELECT m.navn, m.antal, m.minlager, m.leva, la.navn AS leverandora, lb.navn AS leverandorb, mg.navn AS matgrpnavn FROM materialer AS m "_
                &" LEFT JOIN leverand AS la ON (la.id = leva)"_
                &" LEFT JOIN leverand AS lb ON (lb.id = levb)"_
                &" LEFT JOIN materiale_grp AS mg ON (mg.id = m.matgrp)"_
                &" WHERE m.id > 0 AND m.antal < m.minlager AND m.minlager <> -1 AND medarbansv = "& oRec3("mid") &" ORDER BY mg.navn, m.navn LIMIT 1000"

                'response.write strSQL
                'response.flush

                oRec.open strSQL, oConn, 3
                while not oRec.EOF 
            
                matUnderMinLagerTXT = matUnderMinLagerTXT & "<tr><td>"& oRec("matgrpnavn") &"</td><td>"& oRec("navn") & "</td><td>"& oRec("antal") & "</td><td>" & oRec("minlager") & "</td><td>"& oRec("leverandora") &"</td><td>"& oRec("leverandorb") &"</td></tr>" 

                antalM = antalM + 1 
                oRec.movenext
                wend
                oRec.close


                          if antalM <> 0 AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
					  	            
                                       'Sender notifikations mail
		                                Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                                ' Sætter Charsettet til ISO-8859-1
		                                Mailer.CharSet = 2
		                                Mailer.FromName = "TimeOut Email Service - Mimimumslager status"
            
                                        Mailer.ContentType = "text/html"
                                        Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                        Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            

                                        'if lto = "dencker" then
                                        'Mailer.AddRecipient "Anders Dencker", "ad@dencker.net"
                                        'Mailer.AddRecipient "Per Kristensen", "pk@dencker.net"
                                        'Mailer.AddRecipient "Jes Bligård", "jb@dencker.net"
                                        'Mailer.AddRecipient "SK", "support@outzource.dk"
                                        'end if

                                        Mailer.AddRecipient ""& oRec3("mnavn") &"", ""& oRec3("email") &""
                                   

            					        Mailer.Subject = "TimeOut - Mimimumslager status"


		                                strBody = "Hej "& oRec3("mnavn") &" (materialegruppe ansvarlig)<br><br>"
                                        strBody = strBody &"Følgende materialer er under deres angivne minimumslager: <br>"
                                        strBody = strBody &"<table cellpadding=2 cellspacing=0 border=1><tr><td>Gruppe</td><td>Navn</td><td>Antal på lager</td><td>Minimumslager</td><td>Leverandør A</td><td>Leverandør B</td></tr>"
                                    
                                        strBody = strBody &""& matUnderMinLagerTXT &"</table><br><br>&nbsp;"

		                        
            		
		                                Mailer.BodyText = strBody
            		
		                                Mailer.sendmail()
		                                Set Mailer = Nothing



                                '** Opdater status for seneste udsendning  ***'
                                strSQLUpd = "UPDATE licens SET matlageruminemailnot = '"& dd &"' WHERE id = 1"
                                oConn.execute(strSQLUpd)
                                   
				            
                                end if' x


                oRec3.movenext
                wend
                oRec3.close
            


               end if 'dd


        end function
   
        
        
        
        function XsenstematReg(usemrn, useSog, sogBilagOrJob, showonlypers, vasallemed, sogliste)
        %>

        <table width="100%" cellspacing="0" cellpadding="2" border="0">
		
		<tr bgcolor="#5582d2">
		       
		    <td valign=bottom class=alt><b><%=tsa_txt_230 %></b><br />
		    <%=tsa_txt_231%><br />
		    <b><%=tsa_txt_236%></b><br />
		    
		    ~ <%=tsa_txt_218 %>
		    <%if level <= 2 OR level = 6 then %>
			 <%=tsa_txt_213 %>
			<%end if %>
		    </td>
		    
			<td valign=bottom width=80 class=alt><b><%=tsa_txt_183 %></b><br />
			<%=tsa_txt_077 %></td>
			
			<td valign=bottom align=right class=alt><b><%=tsa_txt_202 %></b></td>
			<td valign=bottom style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_209 %></b><br />Lokation</td>
			<td valign=bottom align=center style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_217 %></b></td>
			
			
			
			
			<td valign=bottom align=right style="padding:2px 5px 2px 5px;" class=alt>
			<b><%=tsa_txt_219 %></b><br />
			
			<%if level <= 2 OR level = 6 then %>
			<%=tsa_txt_220 %>
			<%end if %>
			
			</td>
			
			<td valign=bottom align=right style="padding:2px 5px 2px 5px;" class=alt>
			<b><%=tsa_txt_248 %></b>
			</td>
			
			
			<td valign=bottom class=alt>
			<%=left(tsa_txt_251, 3) %>.</td>
			<td valign=bottom class=alt><%=tsa_txt_221 %></td>
			
			
		</tr>
	<%
	'if cint(useSog) = 1 then
        'if cint(sogBilagOrJob) = 0 then
	    sqlWh = "mf.bilagsnr LIKE '"& sogliste &"' OR j.jobnr LIKE '"& sogliste &"' OR j.jobnavn LIKE '"& sogliste &"%' OR matnavn LIKE '"& sogliste &"%'"
        'else
        'sqlWh = "j.jobnr LIKE '"& sogliste &"'"
        'end if
	'else
	'sqlWh = "j.jobnr <> 0"
	'end if
	
	if cint(showonlypers) = 1 then
	strSQLper = " AND personlig = 1"
	else
	strSQLper = ""
	end if

    if cint(vasallemed) = 1 then
    sqlWh = sqlWh &  " AND (usrid <> 0) "
    else
    sqlWh = sqlWh & " AND (usrid = "& usemrn & ")" 
    end if
	
	strSQLmat = "SELECT m.mnavn AS medarbejdernavn, mnr, init, mf.id AS mfid, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
	&" mf.matnavn AS navn, mf.matantal AS antal, mf.dato, mf.editor, "_
	& "mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, "_
	&" mf.usrid, mf.forbrugsdato, j.id, j.jobnr, j.jobnavn, "_
	&" mg.av, f.fakdato, k.kkundenavn, "_
	&" k.kkundenr, mf.valuta, mf.intkode, mf.bilagsnr, v.valutakode, mf.personlig, j.serviceaft, "_
    &" s.navn AS aftnavn, s.aftalenr, mf.kurs, ma.lokation, j.risiko, a.navn AS aktnavn, a.id AS aktid "_
	&" FROM materiale_forbrug mf"_
	&" LEFT JOIN materialer ma ON (ma.id = matid) "_
    &" LEFT JOIN materiale_grp mg ON (mg.id = mf.matgrp) "_
	&" LEFT JOIN medarbejdere m ON (mid = usrid) "_
	&" LEFT JOIN job j ON (j.id = mf.jobid) "_
    &" LEFT JOIN aktiviteter a ON (a.id = mf.aktid) "_
	&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft) "_
    &" LEFT JOIN fakturaer f ON (f.jobid = mf.jobid AND f.faktype = 0) "_
	&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
	&" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
	&" WHERE "& sqlWh &" "& strSQLper &" GROUP BY mf.id ORDER BY mf.id DESC, mf.forbrugsdato DESC, f.fakdato DESC LIMIT 100" 	
	
	'mf.jobid = "& id &"
	'response.write strSQLmat
	'Response.flush
	matkobsprisialt = 0
	s = 0
	oRec.open strSQLmat, oConn, 3
	while not oRec.EOF
	 
	 
	 if len(oRec("fakdato")) <> 0 then
	 fakdato = oRec("fakdato")
	 else
	 fakdato = "01/01/2002"
	 end if
	 
	 if oRec("mfid") = cdbl(lastId) then
	 bgthis = "#FFFF99"
	 else
	     select case right(s, 1) 
	     case 0,2,4,6,8
	     bgthis = "#EFf3FF"
	     case else
	     bgthis = "#FFFFff"
	     end select
	 end if
	 %>
	 
	<tr bgcolor="<%=bgthis %>" class=lille>
	    <%
	    
	    useBr = ""
	    %>
	    <td valign=top style="padding:5px 2px 2px 2px; width:200px; border-bottom:1px d6dff5 solid;" class=lille>
        <%if len(oRec("bilagsnr")) <> 0 then%>
		<%=oRec("bilagsnr") %> 
		<%
		useBr = "<br />"
		end if %>
		
		
        
        <%=useBr %><b> <%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>)</b><br />
        <%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>) 

            <%if oRec("aktid") <> 0 then %>
            <br /><span style="color:#5582d2;"><%=oRec("aktnavn") %></span>
            <%end if %>
        
        <%if oRec("serviceaft") <> "0" then %>
        <br /><span style="color:#999999;">Aft: <%=oRec("aftnavn") %> (<%=oRec("aftalenr") %>)</span>
        <%end if %>

         <font class="megetlillesort">
		 <%if len(oRec("gnavn")) <> 0 then %>
                    <br /> ~ <%=oRec("gnavn")%>
                    <%if level <= 2 OR level = 6 then %>
                    (<%=oRec("av") %>%)
                    <%end if %>
            <%end if %>
		</font>&nbsp;
        
            </td>
	    
		<td valign=top class=lille style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid; white-space:nowrap;">
		<%if len(oRec("forbrugsdato")) <> 0 then%>
		<b><%=formatdatetime(oRec("forbrugsdato"), 2)%></b><br />
		<%end if%>
		
	    <span style="color:#5C75AA; font-size:9px;"><%=left(oRec("medarbejdernavn"), 15) & " ("& oRec("mnr") &")"%></span>

            <span style="color:#999999; font-size:9px;">
		
		<%select case oRec("intkode")
		case 0
	    intKode = "-"
		case 1
		intKode = tsa_txt_239 'intern
		case 2
		intKode = tsa_txt_240 'ekstern
		'case 3
		'intKode = tsa_txt_241
		end select %>
		<br />

		<%if intKode <> "-" then %>
		<%=intKode%> - 
		<%
		
		end if %>
		
		<%if oRec("personlig") <> 0 then %>
		<%=tsa_txt_234 %>
		<%
		
		end if %>
        
        </span>


		</td>
	
		<td valign=top align=right style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid; width:75px;" class=lille><b><%=oRec("antal")%></b>&nbsp;
		
		<%if len(oRec("enhed")) <> 0 then
		enh = oRec("enhed")
		else
		enh = tsa_txt_222 '"Stk."
		end if %>
		
		<%=enh%>
		</td>
	    
	    <td valign=top style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;" class=lille><b><%=oRec("navn")%></b>&nbsp;
        <%if len(trim(oRec("lokation"))) <> 0 then%>
        <span style="font-size:9px; color:#999999;"><%=oRec("lokation")%></span>
        <%end if %>
        </td>
		<td valign=top align=right style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;" class=lille><%=oRec("varenr")%></td>
			
		
		<td valign=top align=right style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;" class=lille>
	    <b><%=formatnumber(oRec("matkobspris"), 2) %></b>
        
        <%
        matkobsprisIalt = matkobsprisialt + (oRec("matkobspris")/1) * (oRec("kurs")/100)
        
        
        if level <= 2 OR level = 6 then %>
        <br />
		<%=formatnumber(oRec("matsalgspris"), 2) %>
        <%end if %>
		</td>
		
		<td valign=top align=right style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid; width:120px;" class=lille>
		<%kobsprisialt = formatnumber(oRec("antal") * oRec("matkobspris"), 2) %>
        <b><%=kobsprisialt %> &nbsp;<%=oRec("valutakode") %></b>
        
        <%if level <= 2 OR level = 6 then %>
        <br />
		<%salgsprisialt = formatnumber(oRec("antal") * oRec("matsalgspris"), 2) %>
        <%=salgsprisialt %> &nbsp;<%=oRec("valutakode") %>
        
		<%end if %>
		
		</td>
		
		
	    <td valign=top style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;">
	    
	    <%
	       '*** Er uge alfsuttet af medarb, er smiley og autogk slået til
                erugeafsluttet = instr(afslUgerMedab(usemrn), "#"&datepart("ww", oRec("forbrugsdato"),2,2)&"_"& datepart("yyyy", oRec("forbrugsdato")) &"#")
                
                'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                'Response.flush
              

                  strMrd_sm = datepart("m", oRec("forbrugsdato"), 2, 2)
                strAar_sm = datepart("yyyy", oRec("forbrugsdato"), 2, 2)
                strWeek = datepart("ww", oRec("forbrugsdato"), 2, 2)
                strAar = datepart("yyyy", oRec("forbrugsdato"), 2, 2)

                if cint(SmiWeekOrMonth) = 0 then
                usePeriod = strWeek
                useYear = strAar
                else
                usePeriod = strMrd_sm
                useYear = strAar_sm
                end if

                
                call erugeAfslutte(useYear, usePeriod, usemrn)
		        
		        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		        'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		        'Response.Write "tjkDag: "& tjkDag & "<br>"
		        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		         call lonKorsel_lukketPer(oRec("forbrugsdato"), oRec("risiko"))
		         
                'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                 if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) = year(now) AND DatePart("m", oRec("forbrugsdato")) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("forbrugsdato")) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
				
				
				
				if (ugeerAfsl_og_autogk_smil = 0 _
				OR ugeerAfsl_og_autogk_smil = 1 AND level = 1) _
				AND cdate(fakdato) < cdate(oRec("forbrugsdato")) OR (oRec("intkode") <> 2)  then 'intkode <> 2 ekstern
				
		
		'*** Kun materialer der ikke er oprettet på laver skal kunne redigeres ***'
		'*** Ændre denne så man vælger ved flueben fra matreg. om det skal oprettes på lager **'
		
		if oRec("varenr") = "0" then	%>
		<a href="materialer_indtast.asp?id=<%=oRec("id")%>&func=red&matregid=<%=oRec("mfid")%>&lastid=<%=oRec("mfid")%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aftid%>&vis=otf&mid=<%=usemrn%>"><img src="../ill/ac0059-16.gif" alt="<%=tsa_txt_251 %>" border=0 /></a>&nbsp;
		<%else %>
		&nbsp;
		<%end if %>
		
		</td>
		<td valign=top style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;">
	    &nbsp;<a href="materialer_indtast.asp?id=<%=oRec("id")%>&func=slet&matregid=<%=oRec("mfid")%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aftid%>&vis=<%=vis%>&mid=<%=usemrn%>"><img src="../ill/slet_16.gif" alt="<%=tsa_txt_221 %>" border=0 /></a>
	    <%else%>
	    &nbsp;</td>
		<td>
	    &nbsp;
	    <%end if %>
	    
	    </td>
		
	</tr>
	
	<%
	s = s + 1
	oRec.movenext
	wend
	
	oRec.close
	%>
    <tr>
        <td colspan=9 style="padding:5px 2px 2px 2px;">Ialt: <%=s %> registreringer, Købspris ialt: <%=formatnumber(matkobsprisialt, 2) &" "& basisValISO %> </td>
    </tr>
	</table>
        
        
        <%
     end function %>    

   



    <%
public strMedarbOptionsHTML
 sub selmedarbOptions

        'visAlleMedarb = 1

        'call medarb_teamlederfor

        'strSQLmids

        if len(trim(jqHTML)) <> 0 then
        jqHTML = 1
        else
        jqHTML = 0
        end if

        if cint(jqHTML) = 1 then
        strMedarbOptionsHTML = "<select name='mid' id='matreg_medid' style='width:260px;'>"
        else
        %>

        
        <select name="mid" id="matreg_medid" style="width:260px;">
        <%end if


    '*** medarb ****
    if level <= 2 OR level = 6 then
        
        if vis_passive = 1 then
        strSQLmansat = " (mansat = 1 OR mansat = 3)"
        else
        strSQLmansat = " mansat = 1"
        end if
	    
        strSQLM = "SELECT mid, mnavn, mnr, init, mansat FROM medarbejdere WHERE "& strSQLmansat &" ORDER BY mnavn"
    
    else
    strSQLM = "SELECT mid, mnavn, mnr, init, mansat FROM medarbejdere WHERE mid = "& usemrn
    end if

   
    	oRec.open strSQLM, oConn, 3 

        while not oRec.EOF 
        
        if cint(usemrn) = cint(oRec("mid")) then
        mSEL = "SELECTED"
        else
        mSEl = ""
        end if
        
        if oRec("mansat") = 3 then
        msat = " - Passiv"
        else
        msat = ""
        end if

        if cint(jqHTML) = 1 then
        strMedarbOptionsHTML = strMedarbOptionsHTML & "<option value="& oRec("mid") &" "&mSel&">"& oRec("mnavn") &" (" & oRec("mnr") &")" & msat &"</option>" 
        else
        %>
        <option value="<%=oRec("mid") %>" <%=mSel %>><%=oRec("mnavn") %> (<%=oRec("mnr") %>) <%=msat %></option>
        <%
        end if

    

        oRec.movenext
        wend
        oRec.close

        
        if cint(jqHTML) = 1 then
        strMedarbOptionsHTML = strMedarbOptionsHTML & "<option value='0'>Ingen</option></select>"
        else%>
        <option value="0">Ingen</option>
        </select>
        <%end if %>
      
       

    <%
    end sub





	'***************** Jobliste ***************************************
	function jobliste(vis, vispasogluk, ignprojekgrp, jobsogVal, jobid)
	
	
	
	if len(trim(jobsogVal)) <> 0 then
	jobsogvalKri = " AND (jobnavn LIKE '%"& jobsogval &"%' OR jobnr = '"& jobsogval &"' OR kkundenavn LIKE '"& jobsogval &"%')" 
	else
	jobsogvalKri = " AND j.id = 0"
	end if


     if jobid <> 0 then
        jobid = jobid
    else 
        jobid = 0
    end if

    jobsogvalKri = jobsogvalKri & " OR (id = "& jobid &")"

 
	
    if cint(vispasogluk) = 1 then
	vispasoglukSQL = "OR jobstatus = 0 OR jobstatus = 2 OR jobstatus = 3 OR jobstatus = 4 "
	else
	vispasoglukSQL = ""
	end if
	
    if cint(ignprojekgrp) = 1 then
    strPgrpSQLkri = ""
    else
    strPgrpSQLkri = strPgrpSQLkri
    end if

	'*** Det valgte job ***
	strSQL = "SELECT id, jobnavn, jobnr, kkundenavn, kkundenr, kid, jobstatus, serviceaft, risiko FROM job j "_
	&" LEFT JOIN kunder k ON (kid = jobknr) WHERE (jobstatus = 1 "& vispasoglukSQL &") "& strPgrpSQLkri &" "& jobsogvalKri &" ORDER BY kkundenavn, jobnavn LIMIT 20"
	
	'jobstatus = 2 = passive
	'Response.Write strSQL
	'response.end
    strJoblisteOptions = ""
	oRec.open strSQL, oConn, 3 
	

	
    lastknr = 0
	fj = 0
	while not oRec.EOF
	    
	    if oRec("id") = cint(jobid) OR (jobid = 0 AND oRec("risiko") = -2) then
	    'strValgtJob = oRec("jobnavn") & "&nbsp;("& oRec("jobnr") &")</b>"
	    jCHK = "SELECTED"
	    else
	    jCHK = ""
	    end if
	     
	     if lastknr <> oRec("kid") then
	          if cint(fj) <> 0 then 
	          strJoblisteOptions = strJoblisteOptions &"<option value="& oRec("id") &" disabled></option>"
	          end if
	          
        
               strJoblisteOptions = strJoblisteOptions &"<option value="& oRec("id") &" disabled>"& oRec("kkundenavn") &" ("& oRec("kkundenr") &")</option>"
	         
	
	     end if
	        
	        jstTxt = ""
	        if oRec("jobstatus") <> 1 then
	            
	            if oRec("jobstatus") = 2 then
	            jstTxt = " - " & tsa_txt_355 
	            end if
	            
	             if oRec("jobstatus") = 0 then
	            jstTxt = " - " & tsa_txt_020
	            end if 

                 if oRec("jobstatus") = 3 then
	            jstTxt = " - Tilbud"
	            end if 

                if oRec("jobstatus") = 4 then
	            jstTxt = " - Gennemsyn"
	            end if
	        
	        end if
	        
	        strJoblisteOptions = strJoblisteOptions & "<option "& jCHK &" value="& oRec("id") &">"& oRec("jobnavn") &" ("&oRec("jobnr") &")" & jstTxt &" .................. "& oRec("kkundenavn") &" ("&oRec("kkundenr") &") </option>"	
	
	lastknr = oRec("kid")
    '** Opdater aftId hvis der findes aftid på job
    aftid = oRec("serviceaft")
	fj = fj + 1
	oRec.movenext
	wend
	oRec.close 
	
	if fj = 0 then
	
	strJoblisteOptions = strJoblisteOptions & "<option value=0>Ingen</option>"
	
	end if
	
         '*** ÆØÅ **'
        call jq_format(strJoblisteOptions)
        strJoblisteOptions = jq_formatTxt
        response.write strJoblisteOptions


    end function




    function matregLagerheader(vis)

        %>
         <tr>
                     
                     <%if cint(vis) = 1 then %>
                     <td>Dato</td>
                     <%end if %>
                   
                     <td>Antal</td>
                     <td><%=tsa_txt_227 %> (mat. nr.)<br />Leverand&oslash;r<br />Lokation</td>

                       
                     <td>Pris</td>
                     <td><%=tsa_txt_225 &"<br>"& tsa_txt_226 &"<br>"& tsa_txt_213 %></td>
                     <td>Salgspris</td>

                     <td>Valuta</td>
                     <td>Intern kode</td>
                     <!--<td>Bilagsnr.</td>
                     <td>&nbsp;</td>-->
                 </tr>
        <%
    end function



    '***************************************************************************************************************
    '**************************** Indlæs mat forbrug func **********************************************************
    '***************************************************************************************************************

    function indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matid, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, func, matreg_opdaterpris, matava)
	
	
		
        'response.write "otf: "& otf &" Dato: "& regdato &" Navn: "& navn & " aktid:" & aktid &"<br>"
        'Response.Write "jobid" & jobid & " antal: "& intAntal & "<br>"
        'response.end

        if len(trim(aftid)) <> 0 then
        aftid = aftid
        else
        aftid = 0
        end if
          

        '** Valutakode '***' 
        if valuta <> 0 AND len(trim(valuta)) <> 0 then
        valuta = valuta
        else
        valuta = 0
        end if

        valutaKode = ""
        strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer WHERE id = "& valuta
    		
    	    oRec3.open strSQL3, oConn, 3 
		    if not oRec3.EOF then
        
            valutaKode = oRec3("valutakode")

            oRec3.close
            end if 

        


       



        '***** Loop ********************
        'for e = 0 to endPoint


        'if cint(otf) = 1 then

        'matId = matId
		

		
			    
			    if isDate(regdato) = false then
			    
			    errortype = 142
			    call showError(errortype)
			    
			    Response.end
			    end if 

		        
		       
		                
		        regDatoSQL = year(regdato) &"/"& month(regdato) &"/"& day(regdato)

      
	


		        '**** Hvis materiale forbrug  / Udlæg skal videre faktureres på job **'
                '**** Tjekker om peridoe er lukket 
		        if cint(intkode) = 2 OR cint(intkode) = 0  then  '** Ekstern / Ikke angivet **'
		        

                strMrd_sm = datepart("m", regdato, 2, 2)
                strAar_sm = datepart("yyyy", regdato, 2, 2)
                strWeek = datepart("ww", regdato, 2, 2)
                strAar = datepart("yyyy", regdato, 2, 2)

                if cint(SmiWeekOrMonth) = 0 then
                usePeriod = strWeek
                useYear = strAar
                else
                usePeriod = strMrd_sm
                useYear = strAar_sm
                end if

        
                
                call erugeAfslutte(useYear, usePeriod, medid)
		        
		        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		        'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		        'Response.Write "tjkDag: "& tjkDag & "<br>"
		        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
                  

		         call lonKorsel_lukketPer(regdato, 1)
		         
                'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                 if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                 (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) = year(now) AND DatePart("m", regdato) < month(now)) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) = 12)) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) <> 12) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", regdato) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
                  
                    ugeerAfsl_og_autogk_smil = 1
                    else
                    ugeerAfsl_og_autogk_smil = 0
                    end if 
                
		        
		                
		            '*** Skal ikke længere tjekke om er foreligger en faktura, da TimeOut nu holder øje med 
		            '*** Om materiale er faktureret ***'    
		            '*** Tjekker sidste fakdato ***'
		            lastFakdato = "01/01/2002"
		         
    		        
		            if ugeerAfsl_og_autogk_smil = 1 then ' OR cdate(lastFakdato) >= cdate(regdato) then
    		        
		            
    			    
			        'if lastFakdato = "01/01/2002" then
			        'lastFakdato = "(ingen)"
			        'end if
    			    
			        useleftdiv = "c"
			        errortype = 108
			        call showError(errortype)
    		        
		            Response.end
		            end if
    		        
		        
		        
                
                end if '** Ekstern ***'       


			            matGrp = gruppe
			            matVarenr = varenr
                      
			            
			            
			           '** Skal mat / udlæg oprettes på lager??
		               if len(trim(opretlager)) <> 0 AND opretlager <> 0 then
		               opretlager = 1
		               else
		               opretlager = 0
		               end if 
		               
		               
	                         if cint(opretlager) = 1 AND matVarenr = "0" then
			                 
				                errortype = 143
				                call showError(errortype)
                			
    			            
			                Response.end
			                end if
			          
			            '*** Findes en vare / et materiale allerede med det valgte varenr'    
			            '*** I den valgte gruppe'
			            findesallerede = 0
			        

			            if matVarenr <> "0" AND cint(otf) = 1 AND cint(opretlager) = 1 then
			            
			            strSQLfindes = "SELECT id, varenr, navn FROM materialer WHERE varenr = '"& matVarenr &"'"_
			            &" AND matgrp = " & matGrp
			            
                       

			            oRec.open strSQLfindes, oConn, 3
	                    if not oRec.EOF then
	                    findesallerede = 1
	                    end if
	                    oRec.close
	                    
	                         if cint(findesallerede) = 1 then
			                
                				
				                errortype = 74
				                call showError(errortype)
                			
    			            
			                Response.end
			                end if
			            
			            end if
	                    
	                    
	                    '*******************************************'
	                    '*** Købs og Salgspris + Ava beregnning ****'
	                    '*******************************************'
	                    
	                                    
		                matNavn = replace(navn, "'", "''") 
		                
                        dblKobsPris = replace(pris, ".","")
		                dblKobsPris = replace(dblKobsPris, ",",".")
            	
                        dblSalgsPris = replace(salgspris,".","")
		                dblSalgsPris = replace(dblSalgsPris,",",".")
			           
		                betegnelse = replace(betegnelse, "'", "''") 
		                
		                '*** End købs og salgspris ***'
		                
		              
              

		                
		                      '**** Opretter on the fly på lager **'
		                      if  matVarenr <> "0" AND cint(opretlager) = 1 AND cint(otf) = 1 then
		                            
		                            intprvAntal = 0
		                            intMatgrp = matGrp
		                            
		                            strSQLins = "INSERT INTO materialer (editor, dato, navn, "_
		                            &" varenr, matgrp, antal, indkobspris, salgspris, "_
						            &" enhed, betegnelse) VALUES "_
						            &"('"& strEditor &"', '"& strDato &"', "_
						            &" '"& matNavn &"', '"& matVarenr &"',"_
						            &""& matGrp &", 0, "_
						            &""& dblKobsPris &", "& dblSalgsPris  &", 'stk.', '"& betegnelse &"')"
            						
						            'Response.write strSQLins & "<br>"
						            'Response.end
            						
						            oConn.execute(strSQLins)
            						
						            matId = 0
						            strSQLmid = "SELECT id FROM materialer WHERE id <> 0 ORDER BY id DESC"
						            oRec.open strSQLmid, oConn, 3
						            if not oRec.EOF then
						            matId = oRec("id")
						            end if
						            oRec.close
            						
            			    end if
            			    '*** End opret på lager **'
						
						
						
						
			            'end if 
						'**** End On the fly ***'
						
		
			                         
                   
			
			
			'*************************************************************'
			'*** Henter Materiale data og indlæser materiale forbrug  ****'
			'*************************************************************'
			

        

			'if cint(otf) = 1 then
			    
			    strVarenr = matVarenr
			    strNavn = matNavn
			    avaGrp = matGrp
			    avaProcent = av
			    strEnhed = betegnelse
			    
			   
    			
			'else
			
			'**************************************************'
			'*** indtasting fra lager ***'
			'**************************************************'
			
			        
			        if matreg_opdaterpris <> 0 then
			        opdaterPris = 1
			        else
			        opdaterPris = 0
			        end if
			        
			        


			        
			        '** Opdaterer priser på mat **'
			        if cint(opdaterPris) = 1 then
			        strSQLpris = "UPDATE materialer SET salgspris = "& dblSalgsPris &", indkobspris = "& dblKobsPris &" WHERE id = "& matId
				    oCOnn.execute(strSQLpris)
				    end if
			        
			
			
			'end if
			'*** END henter mat. stamdata ***'
			

		
			  
			
			
			'Response.Write "dblKobsPris # dblSalgsPris " & dblKobsPris &" # " & dblSalgsPris
			'Response.end 
			

                '*************************************************************'
			    '************* aftid *****************************************'
			    '*************************************************************'
			   
                if jobid <> 0 then
                strSQLserviceaft = "SELECT serviceaft FROM job WHERE id = "& jobid
                oRec6.open strSQLserviceaft, oConn, 3
                if not oRec6.EOF then

                    aftid = oRec6("serviceaft")

                oRec6.movenext
                end if
                oRec6.close

                end if

                  
			   
			    '*************************************************************'
			    '**** Indlæser materiale forbrug                           ***'
			    '*************************************************************'
			
               
      
			    '*** Valuta kurs ***'
			    intValuta = valuta
			    call valutaKurs(intValuta)
			   

              

                                  if (dblSalgsPris) <> 0 then 

         
              

                                       if matava <> 0 then 'hentes fra job NT
                                       ava = matava / 100
                                       else
                
                                       dblKobsPrisBeregn = replace(dblKobsPris, ".", ",")
                                       dblSalgsPrisBeregn = replace(dblSalgsPris, ".", ",")

                                       ava = 1 - formatnumber(((dblKobsPrisBeregn/1) / (dblSalgsPrisBeregn/1)), 4)
                                       end if
                                       'response.write "dblKobsPris: " & dblKobsPris & "<br>"
                                       'response.write "dblSalgsPris: " & dblSalgsPris & "<br>" 
                                       'response.write "ava: "& ava & "<br>"
                                       'response.end

                                   ava = replace(ava, ",", ".")
                                   else
                                   ava = 0
                                   end if



			    if func = "dbopr" then
			   
			    if len(trim(personlig)) <> 0 then
			    personlig = personlig
			    else 
			    personlig = 0
			    end if
			   

                                  
            
                        if cdbl(intAntal) > 0 then
        
                            'response.write "55 otf: "& otf &" func: "& func &" Dato: "& regdato &" Navn: "& navn & "<br>"
                            'response.end
           	   
                            'aftid = 0
                            'strEnhed = "stk"

             
                                'response.write "OK"
            	
                               'if session("mid") = 1 then
                               'Response.Write "HER OK 4: "& matId &","& intAntal &","& strNavn &","& strVarenr&",K:"& dblKobsPris &","& dblSalgsPris &","& strEnhed &","& jobid &","& strEditor &","& strDato &","& medid &","& avaGrp &","& regDatoSQL &",a:"& aftid &","& intValuta &","& intkode &","& bilagsnr &","& dblKurs &","& personlig & "ava:" & ava
			                   'Response.end 
                               'end if

                                 call insertMat_fn(matId, intAntal, strNavn, strVarenr, dblKobsPris, dblSalgsPris, strEnhed, jobid, aktid, strEditor, strDato, medid, avaGrp, regDatoSQL, aftid, intValuta, intkode, bilagsnr, dblKurs, personlig, ava)

                           
			  
				
		                        '*** lastid *** '
                                lastId = 0
			                    strSQLlid = "SELECT id FROM materiale_forbrug WHERE id <> 0 ORDER BY id DESC"
			                    oRec.open strSQLlid, oConn, 3
			                    if not oRec.EOF then
			                     lastId = oRec("id")
			                    end if
			                    oRec.close



						
                     end if

			
			else			
				'*** Opdaterer, KUN materialer der ikke er indlæst på lager pga lagerstyrring **'
				'*** er allerede foretaget ***'
				'*** ved indtastning fra lager kan man kun slette indtastning **'
				
         

                  if cdbl(intAntal) > 0 then

				strSQL = "UPDATE materiale_forbrug SET "_
				&" matid = "& matId &", matantal = "& intAntal &", matnavn = '"& strNavn &"', "_
				&" matvarenr = '"& strVarenr &"', matkobspris = "& dblKobsPris &", matsalgspris = "& dblSalgsPris &", "_
				&" matenhed = '"& strEnhed &"', jobid = "& jobid &", "_
				&" editor = '"& strEditor &"', dato = '"& strDato &"', usrid = "& medid &", "_
				&" matgrp = "& avaGrp &", forbrugsdato = '"& regDatoSQL &"', serviceaft = "& aftid &", "_
				&" valuta = "& intValuta &", intkode = "& intkode &", bilagsnr = '"& bilagsnr &"', "_
				&" kurs = "& dblKurs &", personlig = "& personlig &", aktid = "& aktid &", ava = "& ava &" WHERE id = " & matregid
				
				
				
				'Response.Write strSQL
				'Response.end
				oConn.execute(strSQL)
		    
		    
		    
		    lastId = matregid
		    
		        end if
		    
		    end if
						
				    '**********************'
				    '** Opdaterer antal / lager status ***'
				    '**********************'		
					
				  if strVarenr <> "0" then	
						
				          '**** Flytter lager hvis der er valgt et andet lager end default ***'
			            
			              strSQL2 = "UPDATE materialer SET antal = (antal - "&intAntal&") WHERE id = "& matId
                          'intprvAntal-
				          oCOnn.execute(strSQL2)
			             
			      
			      end if
				
		
        
        
        
        
        
        '**error tjk
	    'end if 'dato
		'end if 'antal


        'next 'antal mat. loop
        
        
        				
		
		strIndlaest = "" & intAntal & " " & strEnhed &" "& navn & " "& formatnumber(pris, 2) &" "& valutaKode &" <span style=""color:yellowgreen;""><b><i>V</i></b></span><br>" 
		
        '*** ÆØÅ **'
        call jq_format(strIndlaest)
        strIndlaest = jq_formatTxt

		'Response.Write "request(FM_sog)" & request("FM_sog")
		'Response.end
		
		sogKri = replace(request("FM_sog"), "%", "wildcardprocent")
		
        if thisfile <> "erp_fak.asp" then
        response.write strIndlaest 
        end if
        'response.end


		'Response.redirect "materialer_indtast.asp?vis="&vis&"&mid="&usemrn&"&lastid="&lastId&"&id="&jobid&"&aftid="&aftid&"&FM_indlaest="&strIndlaest&"&FM_sog="&sogKri
	    
	    

		
	
	
	
	end function 'indlaes_mat




 sub mat_lager_sogeKri
       
            select case lto
            case "dencker", "jttek"
            sz = 10
            case else
            sz = 8
            end select
             %>

        <tr>
	    <td><b><%=tsa_txt_225 &" - "& tsa_txt_226 &" "&tsa_txt_213 %>:</b> <br />
	    <select id="sogmatgrp" size="<%=sz %>" name="sogmatgrp" style="width:220px;">  <!-- onchange="submit();"-->
                <option value="0" selected>Alle</option>
            <% 
	     
                   strSQL3 = "SELECT mg.id, mg.navn, mg.av, kkundenavn, kkundenr, mgkundeid, mg.nummer AS grpnrtxt FROM materiale_grp AS mg "_
                   &" LEFT JOIN kunder AS k ON (k.kid = mgkundeid) "_
                   &" WHERE mg.id <> 0 ORDER BY kkundenavn, mg.navn"
           
                   oRec3.open strSQL3, oConn, 3 
                    mg = 0
                    lastmgkundeid = 0
                   while not oRec3.EOF 

                            if cint(oRec3("mgkundeid")) <> 0 AND cint(lastmgkundeid) <> cint(oRec3("mgkundeid")) then %>
                        
                            <%if cdbl(mg) > 0 then %>
                            <option value="0" disabled></option>
                            <%end if %>
                            <option value="0" disabled> <%=oRec3("kkundenavn") & " ("& oRec3("kkundenr") &")" %></option>
            
                            <%end if 

                        

                        if cint(sogKrimgrp) = oRec3("id") then
                        smgSEL = "SELECTED"
                        else
                        smgSEL = ""
                        end if
                        
                        if oRec3("av") <> 0 then
                        avproc = "("& oRec3("av") & "%)"
                        else
                        avproc = ""
                        end if

                        if len(trim(oRec3("grpnrtxt"))) <> 0 then
                        grpnrtxt = " - " & left(oRec3("grpnrtxt"), 20)
                        else
                        grpnrtxt = ""
                        end if

                    %>
                   <option value="<%=oRec3("id") %>" <%=smgSEL %>><%=oRec3("navn") &" "& grpnrtxt &" "& avproc  %></option>
                   <%

                    lastmgkundeid = oRec3("mgkundeid") 
                       mg = mg + 1
                   oRec3.movenext
                   wend
                   oRec3.close
            %> 
            </select>       
            </td>
	
	    <td><b><%=tsa_txt_264 %>:</b> <br />
            
            
            
            <select size="<%=sz %>" id="soglev" name="soglev" style="width:200px;" > <!-- onchange="submit();"-->
                <option value="0" SELECTED>Alle</option>
            <% 
	     
                   strSQL3 = "SELECT id, navn FROM leverand WHERE id <> 0 ORDER BY navn"
           
                   oRec3.open strSQL3, oConn, 3 
                   while not oRec3.EOF 
                   
                        if cint(sogKrilev) = oRec3("id") then
                        slevSEL = "SELECTED"
                        else
                        slevSEL = ""
                        end if
                   
                   %>
                   <option value="<%=oRec3("id") %>" <%=slevSEL %>><%=oRec3("navn") %></option>
                   <%
                   oRec3.movenext
                   wend
                   oRec3.close
            %> 
            </select>       
            </td>
	
		<td valign="top">
		<b>Materiale:</b><br />
		<input type="hidden" name="FM_indlaest" id="FM_indlaest" value="<%=strIndlaest%>">
		<input type="text" name="FM_sog" id="FM_sog" value="<%=sogeKri%>" style="width:220px; border:2px yellowgreen solid;" placeholder="Søg materiale navn eller nummer" onFocus="rydSog();"> &nbsp;
        </td>
		</tr>
       <!-- <tr><td colspan="3" style="color:#999999;">    <input id="Submit2" type="submit" value="<%=tsa_txt_078 %> >> " /><br />
		<%=tsa_txt_206 %><br />
		Maks. 100 resultater pr. søgning.</td></tr>

-->

		<%


        end sub 



  '***************************************************************************************************************'
    '******************************** Materiale felter INDTAST FRA LAGER *************************************************
    '***************************************************************************************************************
    public matLmt
    function matfelter_lager(vis)

    '** Henter fra lager '***

            select case lto
            case "dencker", "jttek", "micmatic"
            matLmt = 50
            case else
            matLmt = 10
            end select
	

	if useSog = 1 then 
            if len(trim(sogeKri)) <> 0 then
	        sogKri = "("& sogMatgrpSQLkri &") "& sogLevSQLkri &" AND (m.navn LIKE '%"& sogeKri &"%' OR m.varenr LIKE '"& sogeKri &"')"
            else
            sogKri = "("& sogMatgrpSQLkri &") "& sogLevSQLkri &""
            end if
	else
	sogKri = " m.matgrp = -1 " 'sogMatgrpSQLkri &" "& sogLevSQLkri 
	end if
	
	strSQL = "SELECT m.id, m.navn, m.varenr, m.antal, mg.navn AS gnavn, m.matgrp, mg.nummer, "_
	&" m.enhed, m.pic, m.minlager, mg.av, m.indkobspris, m.salgspris, m.lokation, la.navn AS levanavn, "_
    &" lb.navn AS levbnavn, leva, levb, kkundenavn, kkundenr, mgkundeid "_
	&" FROM materialer m "_
    &" LEFT JOIN materiale_grp mg ON (mg.id = m.matgrp) "_
    &" LEFT JOIN kunder AS k ON (k.kid = mgkundeid) "_
    &" LEFT JOIN leverand AS la ON (la.id = m.leva)"_
    &" LEFT JOIN leverand AS lb ON (lb.id = m.levb)"_
	&" WHERE "& sogKri &" AND varenr <> '0' ORDER BY m.matgrp, m.navn LIMIT "& matLmt
	

    'if session("mid") = 21 then
    'Response.write strSQL
    'response.flush
    'end if

    if cint(vis) = 1 then
        cspan = 8
        else
        cspan = 7
    end if

	x = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
    select case right(x, 1)
    case 0,2,4,6,8
    bgthis = "#EFf3FF"
    case else
    bgthis = "#ffffff"
    end select


	
	%>

     <%    '*** ÆØÅ **'
        call jq_format(oRec("navn"))
        matNavn = jq_formatTxt %>

    <!-- faste felter der skal være udfyldt -->
    <input type="hidden" name="" id="matreg_navn_<%=oRec("id")%>" value="<%=matNavn%>" />
    <input type="hidden" name="" id="matreg_matid_<%=oRec("id")%>" value="<%=oRec("id")%>" />
    <input type="hidden" name="" id="matreg_varenr_<%=oRec("id")%>" value="<%=oRec("varenr")%>" />
    <input type="hidden" name="" id="matreg_bilagsnr_<%=oRec("id")%>" value="" />

    
    <!-- Betegnelse -->
    <input type="hidden" name="" id="matreg_betegn_<%=oRec("id")%>" value="<%=oRec("enhed")%>" />	


    <!--
	<input type="hidden" name="FM_matid" id="FM_matid_<%=x %>" value="<%=oRec("id")%>">
    -->


    <%if oRec("matgrp") <> lastmatgrp then %>

         <%    '*** ÆØÅ **'
        call jq_format(oRec("gnavn"))
        gruppeNavn = jq_formatTxt 
         

         if len(trim(oRec("nummer"))) <> 0 then
         call jq_format(oRec("nummer"))
         gruppeNR = " - "& jq_formatTxt   
         else
         gruppeNR = ""
         end if
            
         gruppeNavn = gruppeNavn &" "& gruppeNR  

             
        if oRec("mgkundeid") <> 0 then
            call jq_format(oRec("kkundenavn"))
            gruppeNavn = left(jq_formatTxt, 20)  &" ("& oRec("kkundenr") &"): "& gruppeNavn
        end if%>

        <tr bgcolor="#ffffff">
            <td colspan="<%=cspan %>"><br /><b style="font-size:16px; line-height:18px;"><%=gruppeNavn %></b></td>
        </tr>

    <%
    lastmatgrp = oRec("matgrp")     
    end if %>

      

	<tr bgcolor="<%=bgthis%>">
	    
        <%if cint(vis) = 1 then %>
		<td valign=top style="padding:4px; border-bottom:0px silver solid;"><input type="text" name="matreg_regdato_<%=oRec("id")%>" id="matreg_regdato_<%=oRec("id")%>" value="<%=formatdatetime(now,2) %>" style="width:75px;">
		<%else %>
            <input type="hidden" name="matreg_regdato_<%=oRec("id")%>" id="matreg_regdato_<%=oRec("id")%>" value="<%=formatdatetime(now,2) %>">
        <%end if %>
		
		<td valign=top style="padding:4px; border-bottom:0px silver solid; width:60px; word-wrap:break-word;"><input type="text" name="FM_antal_<%=oRec("id")%>" id="matreg_antal_<%=oRec("id")%>" style="width:50px;">
		<br /><span style="font-size:9px; color:#999999;"><%=oRec("antal")%> <!--/<%=oRec("minlager") %>&nbsp;--><%=oRec("enhed")%></span>
		</td>

       

		<td valign=top style="padding:4px; border-bottom:0px silver solid; width:120px; overflow-wrap:break-word;"><b><a href="preview.asp?pic=<%=oRec("pic")%>&matid=<%=oRec("id")%>" class='vmenu' target="_blank"><%=matNavn%></a></b>
		    <%    '*** ÆØÅ **'
        call jq_format(oRec("varenr"))
        matVarenr = jq_formatTxt %>

            (<%=matVarenr%>)

            <%if oRec("leva") <> 0 then 
                
                 call jq_format(oRec("levanavn"))
                 levNavnA = jq_formatTxt 
            %>
            <br />  <span style="font-size:10px; line-height:11px; color:#999999;"><%=left(levNavnA, 20) %></span>
            <%end if %>

                 <%if oRec("levb") <> 0 then 
                 call jq_format(oRec("levbnavn"))
                 levNavnB = jq_formatTxt 
                 %>
            <br />  <span style="font-size:10px; line-height:11px; color:#999999;"><%=left(levNavnB, 20) %></span>
            <%end if %>

        <%if len(trim(oRec("lokation"))) <> 0 then 
            
             call jq_format(oRec("lokation"))
             lokTxt = jq_formatTxt 
            
            %>
        <br /><span style="font-size:10px; line-height:11px; color:#999999;">Lok.: <%=left(lokTxt, 20)%></span>
        <%end if %></td>
        
        <%




            if level <= 2 OR level = 6 then%>
            <td valign=top style="padding:4px; border-bottom:0px silver solid;">
            <input class="beregnsalgspris_txt" id="pris_<%=oRec("id")%>" name="FM_kobspris_<%=oRec("id")%>" type="text" value="<%=oRec("indkobspris") %>" style="width:50px;" /><!-- onkeyup="beregnsalgsprisOTF(<%=oRec("id")%>)"  -- onkeyup="beregnsalgsprisOTF(<%=oRec("id")%>)" <!-- FM_kobspris_ // beregnsalgspris('<%=oRec("id")%>') -->
            
                    <br /><input name="FM_opdaterpris_<%=oRec("id")%>" id="FM_opdaterpris_<%=oRec("id")%>" type="checkbox" /><span style="font-size:11px; color:#999999;"><%=left(tsa_txt_332,4) %></span>
            
            </td>

         
		    
		     <td valign=top style="padding:4px; width:90px; border-bottom:0px silver solid;">
		    <select name="gruppe_<%=oRec("id")%>" id="gruppe_<%=oRec("id")%>" class="s_matgrp" style="width:100px;"><!-- onchange="beregnsalgsprisOTF(<%=oRec("id")%>)" --><!-- FM_avagruppe_ // beregnsalgspris() -->
		    <option value="0"><%=tsa_txt_200 %></option>
		    
		    
		    <%
		    
		    matgrpVal = "<input id=""avagrpval_0"" name=""avagrpval_0"" type=""hidden"" value=""0"" />"
    		
		    strSQL3 = "SELECT id, navn, av, nummer FROM materiale_grp ORDER BY navn" 'id = "& oRec("matgrp") &" tilføejt 20141222 Det giver ikke mening at skifte gruppe
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(oRec3("id")) = cint(oRec("matgrp")) then
		    matGrpCHK = "SELECTED"
		    else
		    matGrpCHK = ""
		    end if
    		
    		if x = 0 then
    		matgrpVal = matgrpVal &  "<input id=""avagrpval_"& oRec3("id") &""" name=""avagrpval_"& oRec3("id") &""" type=""hidden"" value="& oRec3("av") &" />"
    		end if


                if len(trim(oRec3("nummer"))) <> 0 then
                grpNummerTxt = " - "& left(oRec3("nummer"), 10)
                else
                grpNummerTxt = ""
                end if
    		
		    %>

                 <%    '*** ÆØÅ **'
                call jq_format(oRec3("navn"))
                avagrpNavn = jq_formatTxt %>

		    <option value="<%=oRec3("id")%>" <%=matGrpCHK %>><%=left(avagrpNavn, 10) & grpNummerTxt%>&nbsp;(<%=oRec3("av") %>%)</option>
		    
                
		    
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
		    
		     
        
           
		    
            </td>
            <%if x = 0 then %>
            <%=matgrpVal %>
            <%end if %> 

                <td valign=top style="padding:4px; border-bottom:0px silver solid;">
		     <input class="salgspris_txt" id="FM_salgspris_<%=oRec("id")%>" name="FM_salgspris_<%=oRec("id")%>" type="text" value="<%=oRec("salgspris") %>" style="width:50px;" />
		    </td>
            
           
            
              <%else 
           
          
           
                   avaGrp = 0
                   strSQL3 = "SELECT id, navn, av FROM materiale_grp WHERE id = "& oRec("matgrp") &" ORDER BY navn"
           
                   oRec3.open strSQL3, oConn, 3 
                   if not oRec3.EOF then
                   avaGrp = oRec3("id")
                   avaGrpnavn = oRec3("navn")
                   end if
                   oRec3.close
                   %> 
            
                        <td valign=top colspan="3" align=right style="padding:3px; padding-right:6px; border-bottom:0px silver solid;">
                        <b><%=formatnumber(oRec("indkobspris"), 2) %></b><br />
                        <img src="ill/blank.gif" border=0 width=1 height=3 /><br />
                        <%=avaGrpnavn %></td>
                        <input id="gruppe_<%=oRec("id")%>" name="gruppe_<%=oRec("id")%>" value="<%=avaGrp%>" type="hidden" /><!-- FM_avagruppe_ -->
                        <input id="FM_kobspris_<%=oRec("id")%>" name="FM_kobspris_<%=oRec("id")%>" type="hidden" value="<%=oRec("indkobspris") %>"/>
                        <input id="FM_salgspris_<%=oRec("id")%>" name="FM_salgspris_<%=oRec("id")%>" type="hidden" value="<%=oRec("salgspris") %>" />
           
            
            <%end if %>
            


            
            <td valign=top style="padding:3px; border-bottom:0px silver solid;">
            <%if level <= 2 OR level = 6 then %>
            <select name="FM_valuta_<%=oRec("id")%>" id="matreg_valuta_<%=oRec("id")%>" class="s_valuta" style="width:50px;">
		    <option value="0"><%=tsa_txt_229 %></option>
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(oRec3("grundvaluta")) = 1 then
		    valGrpCHK = "SELECTED"
		    else
		    valGrpCHK = ""
		    end if
    		
    		
		    %>
		    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode")%></option>
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
		    <%else %>
		    <input name="FM_valuta_<%=oRec("id")%>" id="matreg_valuta_<%=oRec("id")%>" value="1" type="hidden" /><%=basisValISO %><br />
		    <%end if %>

            <%
                ik_sel0 = ""
                ik_sel1 = ""
                ik_sel2 = ""

                select case lto
                case "dencker", "hestia"
                ik_sel2 = "SELECTED"
                case else
                ik_sel0 = "SELECTED"
                end select
                
             %>



		    </td>
        <td valign=top style="padding:4px; border-bottom:0px silver solid;"><select name="FM_intkode_<%=oRec("id")%>" id="intkode_<%=oRec("id")%>" class="s_internkode" style="width:50px;">
		    <%if lto <> "execon" AND lto <> "immenso" AND lto <> "unik" then %>
		    <option value="0" <%=ik_sel0 %>><%=tsa_txt_235 %></option><!-- ingen -->
		    <%end if%>
		    
		    <option value="1" <%=ik_sel1 %>><%=tsa_txt_232 %></option><!-- intern -->
		    <option value="2" <%=ik_sel2 %>><%=tsa_txt_233 %></option><!-- ekstern -->
		    
		    <!--
		    <option value="3"><%=tsa_txt_234 %></option>
		    -->
		    
		    
    		</select>

            <!--
    		
               <br /> <input id="FM_personlig_<%=oRec("id")%>" name="FM_personlig_<%=oRec("id")%>" value="1" type="checkbox" /> <%=left(tsa_txt_234, 4) %>.
            -->
		    
		   
            </td>
		   
            
         
           <input type="hidden" class="s_regdato" name="regdato_<%=oRec("id")%>" value="<%=formatdatetime(date, 2) %>">

		<!--<td valign=top style="width:100px; padding:4px; border-bottom:0px silver solid;"><input id="matreg_bilagsnr_<%=oRec("id")%>" name="FM_bilagsnr_<%=oRec("id")%>" type="text" value="<%=bilagsnr %>" style="width:60px;" /></td> --><!-- FM_bilagsnr_ -->
	    <td valign=top style="width:40px; padding:4px; border-bottom:0px silver solid;"><input type="button" id="<%=oRec("id")%>" class="matreglg_sb" value=">>" /></td>
	</tr>
	
	<%
	x = x + 1
	oRec.movenext
	wend
    oRec.close

        %>
        <input type="hidden" value="<%=(x-1)%>" id="antal_x" />
        <%
	
	if x = 0 then%>
	
	<tr bgcolor="#ffffff">
		
		<td height=50 colspan=<%=cspan%> style="color:#999999;"><br /><br />S&oslash;g p&aring; materialer i boksene ovenfor.</td>
		
	</tr>

    <%else %>
    
	<tr bgcolor="#ffffff">
		
		<td height=50 colspan=<%=cspan%> style="color:#999999;"><br /><br />Der vises maks <%=matLmt %> resultater i s&oslash;gningen.</td>
		
	</tr>

    <%end if
        
        
        end function 'mat fra lager felter










    '***************************************************************************************************************'
    '******************************** Materiale felter INDTAST OTF *************************************************
    '***************************************************************************************************************
    sub matStFelter
    %>
        <tr>
		<td align=right style="padding-top:5px;"><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_202 %>:</td>
		<td style="padding-left:5px; padding-top:5px;"><input type="text" name="FM_antal" id="matreg_antal_0" value="<%=matantal %>" size="10"> <%=tsa_txt_203 %>
             <input type="hidden" name="" id="matreg_matid_0" value="0" />
		</td>
	</tr>
	
	<tr>
		<td align=right valign=top style="padding-top:5px;"><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_194 %>:</td>
		<td style="padding-left:5px;"><textarea  name="navn" id="matreg_navn_0" cols="30" rows="2"><%=matnavn %></textarea></td>
	</tr>
	<tr>
		<td align=right><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_201 %>:</td>
		<td style="padding-left:5px;"><input class="beregnsalgspris_txt" type="text" id="pris_0" name="pris" value="<%=matkobspris %>" size="10">&nbsp;<!--onkeyup="beregnsalgsprisOTF(0)" -->
		<select name="FM_valuta" id="matreg_valuta_0" style="width:55px;">
		    <!--<option value="0"><=tsa_txt_229 %></option>-->
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(valuta) = oRec3("id") then
		    valGrpCHK = "SELECTED"
		    else
		    valGrpCHK = ""
		    end if
		    
		   
		    %>
		    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode")%></option>
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
		</td>
	</tr>
	<%if lto = "execon" OR lto = "intranet - local" OR lto = "epi" OR lto = "epi_no" OR lto = "epi_ab" OR lto = "epi_sta" OR lto = "unik" OR lto = "adra" then 'lto = "intranet - local" OR%>
    <input id="matreg_betegn_0" type="hidden" name="FM_betegn" value="<%=matenhed %>" />
    <input type="hidden" name="varenr" id="matreg_varenr_0" value="0" />
	<%else %>
	<tr>
		<td align=right>&nbsp;<%=tsa_txt_245 %>:</td>
		<td style="padding-left:5px;"><input type="text" id="matreg_betegn_0" name="FM_betegn" value="<%=matenhed %>" size="30" maxlenght="20"><span style="color:#999999; font-size:9px;"> (<%=LCASE(tsa_txt_195) %>)</span></td>
	</tr>
	<tr>
		<td valign=top align=right style="padding-top:5px;"><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_196 %>:</td>
		<td valign=top style="padding-left:5px;"><input type="text" name="varenr" id="matreg_varenr_0" value="<%=matvarenr %>" size=20">
		<input type="checkbox" id="matreg_opretlager_0" name="opretlager" value="1" /> Opret på lager
		<br /><span style="color:#999999; font-size:10px;"><%=tsa_txt_197 %></span></td>
	</tr>
	<%end if %>






    <%
        
    '** Avancegruppe **'    
    if lto <> "epi" AND lto <> "epi_no" AND lto <> "epi_ab" AND lto <> "epi_sta" AND lto <> "xintranet - local" AND lto <> "unik" then  %>
	    <tr>
		    <td align=right>
           
		    <%if level <= 2 OR level = 6 then %>
		        <%if lto <> "execon" AND lto <> "intranet - local" AND lto <> "unik" then  %>
		        <%=tsa_txt_198 & " "& tsa_txt_199 %>
		        <%else %>
		        <%=tsa_txt_198 %>
		        <%end if %>
		    <%else %>
		    <%=tsa_txt_198 %>
		    <%end if %>
		    :</td>
		    <td style="padding-left:5px;">
		    <!--onChange="visgrp()"-->
		
		    <% 
		    matgrpVal = "<input id=""avagrpval_0"" name=""avagrpval_0"" type=""hidden"" value=""0"" />"
            %>
		    
    		
    	    <select class="s_matgrp" name="gruppe" id="gruppe_0" style="width:200px;"><!-- onchange="beregnsalgsprisOTF(0)" -->
		    <option value="0"><%=tsa_txt_200 %></option>
		    <%
		    strSQL = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		    oRec.open strSQL, oConn, 3 
		
		            while not oRec.EOF 
		
		            if cint(matgrp) = oRec("id") then
		            matgrpSel = "SELECTED"
		            else
		            matgrpSel = ""
		            end if
		
		
    		            matgrpVal = matgrpVal &  "<input id=""avagrpval_"&oRec("id")&""" name=""avagrpval_"&oRec("id")&""" type=""hidden"" value="& oRec("av") &" />"
    		
		
		            %>
		            <option value="<%=oRec("id")%>" <%=matgrpSel %>><%=oRec("navn")%>
		            <%if level <= 2 OR level = 6 then %>
		            &nbsp;(<%=oRec("av") %>%)
		            <%end if %></option>

		    <%
		    oRec.movenext
		    wend
		    oRec.close %>
		    </select>
		
		    <%=matgrpVal %>
		
		    </td>
        </tr>
    <%else %>
    <input id="avagrpval_0" name="avagrpval_0" type="hidden" value="0" />
    <input id="gruppe_0" name="gruppe" type="hidden" value="0" />
	<%end if  '*** Avancegruppe **' %>



    <%
    '*************** INERN KODE ******************    
    if lto = "adra" OR lto = "intranet - local" then
        %>
    <input type="hidden" name="FM_intkode" id="intkode_0" value="0">
    <input id="matreg_personlig_0" name="FM_personlig" value="0" type="hidden"/> 
        <%
    else %>


    <tr><td valign=top style="padding-top:4px;" align=right><%=tsa_txt_231 %>:</td>
		    <td valign=top style="padding-left:5px;">
		   <%
			    
                ikSEL0 = ""
			    ikSEL1 = ""
			    ikSEL2 = ""
			    ikSEL3 = ""

			    select case intKode
			    case 1
	            ikSEL1 = "SELECTED"
			    case 2
			    ikSEL2 = "SELECTED"
			    case 3
			    ikSEL3 = "SELECTED"
			    case else
			    
			    'if lto <> "execon" AND lto <> "immenso" AND lto <> "unik" then
			    'ikSEL0 = "SELECTED"
			    'ikSEL1 = ""
			    'else
			    'ikSEL0 = ""
			    'ikSEL1 = "SELECTED"
			    'end if
			    
			    'ikSEL2 = ""
			    'ikSEL3 = ""
			   
			    ikSEL2 = "SELECTED"
			 

			    end select
			    
			    if cint(personlig) = 1 then
			    persCHK = "CHECKED"
			    else
			    persCHK = ""
			    end if
			
			call licKid()

           
			%>
			
			<select name="FM_intkode" id="intkode_0" style="width:180px;">
			 <%if lto <> "execon" AND lto <> "immenso" AND lto <> "unik" then %>
		    <option value="0" <%=ikSEL0 %>><%=tsa_txt_235 %></option>
		    <%end if %>
		    <option value="1" <%=ikSEL1 %>><%=tsa_txt_232 %> </option> <!-- (<=licensindehaverKnavn%>) -->
		    <option value="2" <%=ikSEL2 %>><%=tsa_txt_233 %> </option> <!-- (<=lcase(tsa_txt_243) %>)-->
		    
		    <!--
		    <option value="3" <%=ikSEL3 %>><=tsa_txt_234 %></option>
		    -->
		    
    		</select>
    		&nbsp;
                <input id="matreg_personlig_0" name="FM_personlig" value="1" type="checkbox" <%=persCHK %> /> <%=tsa_txt_234 %>
		    </td>
           </tr> 
           <%end if'lto %>
	
	
	    
	                <%
                    '*** Salgspris ******'    
                    if (level <= 2 OR level = 6) then 
	    
	                if lto = "execon" OR lto = "intranet - local" OR (lto = "epi" OR lto = "epi_no" OR lto = "epi_as" OR lto = "epi_uk") OR lto = "unik" OR lto = "adra" then
	                slgsprVZB = "hidden"
	                else
	                slgsprVZB = "visible"
	                end if 
	    

                            if slgsprVZB = "visible" then
	                        %>
	                        <tr id="tr_slgs" style="visibility:<%=slgsprVZB%>;">
	                            <td align=right><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_261%>:</td>
	                            <td style="padding-left:5px;"><input class="salgspris_txt" id="FM_salgspris_0" name="FM_salgspris" type="text" value="<%=matsalgspris %>" style="width:80px;" /> <span style="color:#999999; font-size:9px;">(bruges ved fakturering)</span></td>

	                        </tr>
                            <%else %>
                               
                             <input id="FM_salgspris_0" name="FM_salgspris" type="hidden" value="0" />
                                
                            <%end if %>


	                <%else %>
	                <input id="FM_salgspris_0" name="FM_salgspris" type="hidden" value="0" />
	                <%end if %>
	    
	    
		
                   <%if lto <> "execon" AND lto <> "intranet - local" AND lto <> "epi" AND lto <> "epi_no" AND lto <> "epi_sta" AND lto <> "epi_ab" AND lto <> "unik" AND lto <> "adra" then %>    
                   <tr>
                        <td align=right><%=tsa_txt_230 %>: </td>
                        <td style="padding-left:5px;"><input id="matreg_bilagsnr_0" name="FM_bilagsnr" type="text" value="<%=bilagsnr %>" style="width:80px;" /> (<%=tsa_txt_256%>)</td>
                   </tr> 
                    <%else %>
                    <input id="matreg_bilagsnr_0" name="FM_bilagsnr" type="hidden" value="<%=bilagsnr %>"/>
                  <%end if %>   


    <%
    end sub








    public strMatNavn, strMatVarenr, avaProcent, strMatEnhed, intMatgrp, intprvAntal, dblKobsPris, dblSalgsPris
    function matidStamdata(matId, avaGrp, prs)


                                         '*** Henter StamData på materiale ***'
			                            strSQL2 = "SELECT id, navn, editor, dato, varenr, antal, "_
			                            & "salgspris, indkobspris, enhed, matgrp FROM materialer WHERE id=" & matId
                            			
			                            'Response.Write strSQL2
			                            'Response.flush
                            			
			                            oRec2.open strSQL2, oConn, 3
			                            if not oRec2.EOF then
                            			        
                            			      
			                            strMatNavn = oRec2("navn")
			                            strMatVarenr = oRec2("varenr")
                            			
			                            strMatEnhed = oRec2("enhed")
			                            intMatgrp = oRec2("matgrp")
			                            intprvAntal = oRec2("antal")

                                        if cint(prs) = 1 then
                                        dblKobsPris = replace(oRec2("indkobspris"),",",".")
			                            dblSalgsPris = replace(oRec2("salgspris"),",",".")
                            			end if
                            			
			                            end if
			                            oRec2.close
                            			
			        
			        
			                                dblKobsPris = replace(dblKobsPris,",",".")
			                                dblSalgsPris = replace(dblSalgsPris,",",".")



                                            avaProcent = 0
                                			if cint(avaGrp) <> 0 then
			                                 strSQL3 = "SELECT av FROM materiale_grp WHERE id = "& avaGrp
			                                 oRec3.open strSQL3, oConn, 3
			                                 if not oRec3.EOF then
			                                 avaProcent = oRec3("av")
			                                 end if
			                                 oRec3.close
                                            end if


  end function
    
  
    
  function  insertMat_fn(matId, intAntal, strNavn, strVarenr, dblKobsPris, dblSalgsPris, strEnhed, jobid, aktid, strEditor, strDato, usemrn, avaGrp, regDatoSQL, aftid, intValuta, intkode, bilagsnr, dblKurs, personlig, ava)

                         strSQL = "INSERT INTO materiale_forbrug "_
				            &" (matid, matantal, matnavn, matvarenr, matkobspris, matsalgspris, matenhed, jobid, "_
				            &" editor, dato, usrid, matgrp, forbrugsdato, serviceaft, valuta, intkode, bilagsnr, kurs, personlig, aktid, ava) VALUES "_
				            &" ("& matId &", "& intAntal &", '"& strNavn &"', '"& strVarenr &"', "_
				            &" "& dblKobsPris &", "& dblSalgsPris &", '"& strEnhed &"', "& jobid &", "_
				            &" '"& strEditor &"', '"& strDato &"', "& usemrn &", "_
				            &" "& avaGrp &", '"& regDatoSQL &"', "& aftid &", "& intValuta &", "_
				            &" "& intkode &", '"& bilagsnr &"', "& dblKurs &", "& personlig &", "& aktid &", "& ava &")"
				
                            'if session("mid") = 1 then
				            'Response.Write strSQL
				            'Response.flush
                            'end if
				            
                            oConn.execute(strSQL)

    end function
      
%>