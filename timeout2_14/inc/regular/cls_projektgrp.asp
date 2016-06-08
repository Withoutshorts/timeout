


<%
 sub progrpHeader
    %>
<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
<tr bgcolor="#5582D2">
		<td class="alt" valign="bottom" style="width:400px"><b>Navn</b></td>
		<td class="alt" valign="bottom" style="padding-right:20px;"><b>Tilføj/Fjern?</b></td>
		<td class="alt" valign="bottom" style="padding-right:20px;"><b>Teamleder?</b><br />
		<span style="font-color:#999999;">Kan trække statistik på <br />andre medarbejdere i gruppen<br />(kun admin. kan oprette teamledere)</span></td>
        <td class="alt" valign="bottom" style="padding-right:20px;"><b>Notificer</b> <br />
        <span style="font-color:#999999;">Modtag mail ved sygdom og ferie registreringer</span></td>
	</tr>

<%
end sub


 
     function erTeamLederForvlgtMedarb(level, usemrn, medidloggetpaa)
       
      erTeamlederForVilkarligGruppe = 0
      
        if level <= 2 OR level = 6 then

              call medariprogrpFn(usemrn)

              progrpArr = split(medariprogrpTxt, ",#")
              for pgArr = 0 to UBOUND(progrpArr)

                  progrpArr(pgArr) = replace(progrpArr(pgArr), "#", "")
                  progrpArr(pgArr) = replace(progrpArr(pgArr), " ", "")

                  call fTeamleder(session("mid"), progrpArr(pgArr))

              next


        end if

    end function


 public medariprogrpTxt
	function medariprogrpFn(medarbid)
		
		medariprogrpTxt = "#10#"
		strSQL = "SELECT ProjektgruppeId, MedarbejderId FROM projektgrupper AS p "_
        &" LEFT JOIN progrupperelationer ON (ProjektgruppeId = p.id AND MedarbejderId = "& medarbid &") "_
        &" WHERE MedarbejderId ="& medarbid &" GROUP BY ProjektgruppeId" 
		
        'response.Write "strSQL" & strSQL 
        'response.flush
    
    	oRec8.Open strSQL, oConn, 0, 1 
			
			While Not oRec8.EOF
			
             medariprogrpTxt = medariprogrpTxt & ",#"& oRec8("ProjektgruppeId") &"#"
				
		    oRec8.MoveNext
			Wend
		
		oRec8.Close
    end function


'**** BRUGES PÅ TIEMREG OMRÅDET VED FLUEBEN til at LAVE listen af medaerb. man er TEAMLEDER FOR
public strSQLmids, teamLederEditOK, strSQLmidsTjkRettighed  ', medarbgrpIdSQLkri
sub medarb_teamlederfor

     teamLederEditOK = 1 'Ændret til 1 da man kun ha adgang til medarbejderlisten, hvis man i forvejen er teamleder, så man behøver ikke tjekke igen om man er teamleder for den nekelte medarbejder.
   
     if cint(visAlleMedarb) = 1 then   
        '(cint(teamleder_flad) = 0 AND thisfile = "ugeseddel_2011.asp") ==== Må man godkende egne timer? kun som admin eller teamleder
       

        call projgrp(-1,level,session("mid"),1)
	    
	     medarbgrpIdSQLkri = "AND (mid = "& session("mid")
    
	    
	    for p = 0 to prgAntal
	     
        'response.write "prjGTeamleder(: "& prjGTeamleder(p) & "<br>"

	     if prjGoptionsId(p) <> 0 then
	        call medarbiprojgrp(prjGoptionsId(p), session("mid"), 0, -1)


            if cint(teamleder_flad) = 0 then '0: Bruger Teamleder niveau , 1: Flat - alle kan se alle uanset om de er teamledere 
                    'call fTeamleder(session("mid"), prjGoptionsId(p))
                    if cint(prjGTeamleder(p)) = 1 then 'Er han teamleder for gruppen?
                    midsTjkRettighed = midsTjkRettighed & medarbgrpIdSQLkri
                    end if
           else        

                    midsTjkRettighed = midsTjkRettighed & medarbgrpIdSQLkri

            end if
	     end if
	    
	    next 
	    
        'response.write "HER: teamleder_flad : "& teamleder_flad &" midsTjkRettighed: "& midsTjkRettighed 
        '& " medarbgrpIdSQLkri &" prgAntal: "& prgAntal
        'response.end

	     medarbgrpIdSQLkri = medarbgrpIdSQLkri & ")"
	    
	strSQLmids = medarbgrpIdSQLkri '" AND mid = "& usemrn
    
    else

    strSQLmids = "AND mid = "& usemrn

    end if


      strSQLmidsTjkRettighed = replace(midsTjkRettighed, "OR", "#")
    strSQLmidsTjkRettighed = replace(strSQLmidsTjkRettighed, ")", "#")
    strSQLmidsTjkRettighed = replace(strSQLmidsTjkRettighed, "(", "")
    strSQLmidsTjkRettighed = replace(strSQLmidsTjkRettighed, "AND", "")
    strSQLmidsTjkRettighed = replace(strSQLmidsTjkRettighed, " ", "")
    'strSQLmidsTjkRettighed = replace(strSQLmidsTjkRettighed, "mid="&session("mid")&"#", "") 

    'response.write "strSQLmidsTjkRettighed: "& strSQLmidsTjkRettighed

    'if instr(strSQLmidsTjkRettighed, "mid="&usemrn&"#") <> 0 then
    'teamLederEditOK = 1
    'else
    'teamLederEditOK = 0
    'end if
  

end sub



'***** Tilføj Medarbrejder til timereg.usejob for alle de job ahn/hun ertilmedt via sine projektgrupper ***

function tilfojtilTU(mid, del)



'mid <> 0 enjelt medarb / alle medab

if mid <> 0 then
midSQLKri = "mid = "& mid
else
midSQLKri = "mid <> 0"
end if

strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE "& midSQLKri &" AND mansat <> 2 ORDER BY mnavn"

'Response.Write strSQL & "<hr>"
'Response.flush

j = 0
oRec5.open strSQL, oConn, 3
while not oRec5.EOF 
        
 
 strSQLDontDelJob = " AND (jobid <> 0 "      


'Response.write "<u>"& oRec5("mnavn") & "</u><br>"
        

     call hentbgrppamedarb(oRec5("mid"))
	
    '*** Søger efter job først, da timereg_usejob kun indeholder de job der er aktive lige nu. 
    '*** Du kan godt være tilmeldt et job via dine projektgrupper uden det er med i timereg_usejob

	lastKid = 0
	strSQLj = "SELECT j.id AS jid, j.jobstatus, j.jobnavn, j.jobnr, j.jobknr, Kid, Kkundenavn, Kkundenr, "_
	&" useasfak, g.jobid AS gjid, g.easyreg, forvalgt FROM job j "_
    &" LEFT JOIN kunder ON (kunder.Kid = j.jobknr) "_
	&" LEFT JOIN timereg_usejob AS g ON (g.medarb = "& oRec5("mid") &" AND g.jobid = j.id) "_
	&" WHERE (jobstatus = 1 OR jobstatus = 3) AND kunder.Kid = j.jobknr "& strPgrpSQLkri &""_
	&" GROUP BY j.id, kid ORDER BY Kkundenavn, jobnavn"
	
    'Response.write strSQLj & "<br><br>"	
	'Response.flush
	
    	
    oRec4.open strSQLj, oConn, 3
    while not oRec4.EOF 

    jobFandtes = 0
    
    
    'Response.write oRec4("jobnavn") & "<br>"
            
            strSQLTU = "SELECT jobid FROM timereg_usejob WHERE jobid = "& oRec4("jid") & " AND medarb = "& oRec5("mid")
            oRec3.open strSQLTU, oConn, 3
            if not oRec3.EOF then

            jobFandtes = 1
            
            strSQLDontDelJob = strSQLDontDelJob & " AND jobid <> "&  oRec4("jid")

            end if
            oRec3.close
    
            '**** Opretter job i TU **
            if jobFandtes = 0 then

              strSQLins = "INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& oRec5("mid") &", "& oRec4("jid") &")"

              'Response.write "<b>"& strSQLins &"</b><br>"
              oConn.execute(strSQLins)

              strSQLDontDelJob = strSQLDontDelJob & " AND jobid <> "&  oRec4("jid")
            end if


   j = j + 1
    oRec4.movenext						
	wend 							
    oRec4.close
                         


    
        '** Renser ud i timereg.use job for den valgte medarbejder ****
        strSQLDontDelJob = strSQLDontDelJob & ")"
        if cint(del) = 1 then
        strSQLdel = "DELETE FROM timereg_usejob WHERE medarb = "& oRec5("mid") & strSQLDontDelJob
        'Response.write strSQLdel
        oConn.execute(strSQLdel)
        'Response.end
        end if    

oRec5.movenext
wend
oRec5.close



end function



'*** Hvis alle nødvendige er udfyldt ***'
		public guidjobids, guideasyids
	    function prgrel(id, func)
	    
	    'Response.Write request("FM_progrp") & "<br>"
        'Response.write "f: "& func
	    'Response.end
	    
	    guidjobids = ""
	    guideasyids = ""
	    
	                strSQLpgDel = "DELETE FROM progrupperelationer WHERE medarbejderid = " & id
					oConn.execute(strSQLpgDel)
					pgr = split(request("FM_progrp"), ",")
					
					for p = 0 to UBOUND(pgr,1)
					
					if len(trim(pgr(p))) <> 0 then

                        
                        '** NOTIFICER ***'
                        if instr(pgr(p), "-1") <> 0 then
                        notificer = 1
                        else
                        notificer = 0
                        end if 
                    

                        '*** TEAMLEDER ***
                        if instr(pgr(p), "_1") <> 0 then
                        teamleder = 1
                        pgr_len = len(pgr(p))

                        if (notificer) = 1 then
                        pgr_left = left(pgr(p), pgr_len - 4)
                        else
                        pgr_left = left(pgr(p), pgr_len - 2)
                        end if
                        pgr(p) = pgr_left 
                        else
                        teamleder = 0
                        pgr(p) = pgr(p)
                        end if 

                        
					
                    	'Response.Write strSQLpgIns & "<br>"
					    
					
                    if pgr(p) <> 1 then 'Ingen gruppen
					strSQLpgIns = "INSERT INTO progrupperelationer (projektgruppeId, medarbejderId, teamleder, notificer) VALUES ("& pgr(p) &", "& id &", "& teamleder &", "& notificer &")"
					oConn.execute(strSQLpgins)
                    end if
				
					    if func = "dbopr" then
					    
					
					        call projktgrpPaaJobids(pgr(p))
					       
					       
					     end if '** dbopr 
					
					end if

                    'Response.write pgr(p) & "<br>"
					
					next
					
					
					
					'Response.end
					
		end function



         function projktgrpPaaJobids(pid)
        
                            '*** Sætter guiden aktive job ready ***'
					        strSQLjob = "SELECT id AS jid FROM job WHERE "_
					        &" projektgruppe1 = "& pid &" OR " _
					        &" projektgruppe2 = "& pid &" OR " _
					        &" projektgruppe3 = "& pid &" OR " _
					        &" projektgruppe4 = "& pid &" OR " _
					        &" projektgruppe5 = "& pid &" OR " _
					        &" projektgruppe6 = "& pid &" OR " _
					        &" projektgruppe7 = "& pid &" OR " _
					        &" projektgruppe8 = "& pid &" OR " _
					        &" projektgruppe9 = "& pid &" OR " _
					        &" projektgruppe10 = "& pid 
					        
					        oRec4.open strSQLjob, oConn, 3
					        while not oRec4.EOF
					        
					        guidjobids = guidjobids & oRec4("jid") &", #, "
					        
					        oRec4.movenext
					        wend
					        oRec4.close
					        
					        
					        '*** Sætter guiden aktive job Easyreg ready ***'
					        strSQLeasy = "SELECT j.id AS jid, COUNT(a.id) AS antal_aeasy FROM job AS j "_
					        &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.easyreg = 1) WHERE "_
					        &" j.projektgruppe1 = "& pid &" OR " _
					        &" j.projektgruppe2 = "& pid &" OR " _
					        &" j.projektgruppe3 = "& pid &" OR " _
					        &" j.projektgruppe4 = "& pid &" OR " _
					        &" j.projektgruppe5 = "& pid &" OR " _
					        &" j.projektgruppe6 = "& pid &" OR " _
					        &" j.projektgruppe7 = "& pid &" OR " _
					        &" j.projektgruppe8 = "& pid &" OR " _
					        &" j.projektgruppe9 = "& pid &" OR " _
					        &" j.projektgruppe10 = "& pid & " GROUP BY j.id" 
					        
					        'Response.Write strSQLeasy
					        'Response.flush
					        
					        oRec4.open strSQLeasy, oConn, 3
					        while not oRec4.EOF
					        
					        if oRec4("antal_aeasy") <> "0" then
					        guideasyids = guideasyids & oRec4("jid") &", #, "
					        else
					        guideasyids = guideasyids &"0, #, "
					        end if
					        
					        oRec4.movenext
					        wend
					        oRec4.close
        
        
        end function





        function setGuidenUsejob(medid, useJob, useAkt, del, useEasy, useForvalgt, dothis)
		
        dtNow = year(now) & "-"& month(now) & "-"& day(now)
		'Response.Write useJob & "<hr>"
		'Response.Write useEasy & "<hr>"
        'Response.Write useForvalgt & "<hr>"

        'Response.end

        '**Forvalgt 
        '1  : Forvalgt, vises på aktivlisten. Altid = 1 ved Positiv tildeling
        '0  : Neutral 
        '-1 : Skjult
        '2  : Låst, vises altid , selvom der er gået mere end 14  dage / 2 md.

        '*** Fra Guiden er forvalg altid = 0 (job), eller 1 (aktiviteter)
         
		
		if len(trim(useJob)) <> 0 then
		len_useJob = len(useJob)
		left_useJob = left(useJob, (len_useJob-3))
		useJob = left_useJob
		end if
		

        if len(trim(useAkt)) <> 0 then
		len_useAkt = len(useAkt)
		left_useAkt = left(useAkt, (len_useAkt-3))
		useAkt = left_useAkt
		end if

        if len(trim(useEasy)) <> 0 then
		len_useEasy = len(useEasy)
		left_useEasy = left(useEasy, (len_useEasy-3))
		useEasy = left_useEasy
		end if

		'Response.Write useForvalgt & "<br>"
		'Response.end 
		
	

        intuseJob = Split(useJob, "#, ")
        intuseAkt = Split(useAkt, "#, ")
        intUseEasy = split(useEasy, "#, ")

       

        select case dothis 
        case 0

        if len(trim(useForvalgt)) <> 0 then
		len_useForvalgt = len(useForvalgt)
		left_useForvalgt = left(useForvalgt, (len_useForvalgt-3))
		useForvalgt = left_useForvalgt
		end if

        intUseForvalgt = split(useForvalgt, "#, ")

        case 1,2 'Ved ikke om denne bruges mere

         '** Hvis forvalgt ikke er udfyldt sættes den = med job Array. 
        intUseForvalgt = Split(useJob, "#, ")
        for j = 0 TO UBOUND(intUseForvalgt)
        intUseForvalgt(j) = "0, "
        next
           

        end select

        '** Hvis forvalgt ikke er udfyldt sættes den = med job Array. 
        '** Nedenfor sættes useForvalgt(j) = 0 Hvis den er forskellig fra 1
        'if len(trim(useForvalgt)) <> 0 then
        'useForvalgt = useForvalgt
        'else
        'useForvalgt = useJob
        'end if
		
		'Response.flush
        'DEL = 0 
        '*** Renser ud i ovenskydende timereg.use job aktiviter på lukkede job / Dubletter mv.
        '*** Sletter eksisteredne tilhørsforhold !! PAS PÅ DENNE '***
        '*** Sletter IKKE Easyreg aktiviteter. De har deres egen side og  ***'
		if cint(del) = 0 then
		oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& medid &"")
		end if
		
		
		
		j = 0
		
    
	   	For j = 0 to Ubound(intuseJob)
	   	    
	   	    'Response.Write len(trim(intuseJob(j))) & "Job: " & intuseJob(j) &" Easy: "& intUseEasy(j) &"<br>"
	   	    'Response.flush
	   	    
	   	    if len(trim(intuseJob(j))) > 1 then
	   	    intuseJob(j) = trim(left(intuseJob(j), len(intuseJob(j)) - 2))
	   	    else
	   	    intuseJob(j) = -1
	   	    end if
	   	    
	   	  
            if len(trim(intUseAkt(j))) > 1 then
	   	    intUseAkt(j) = trim(left(intUseAkt(j), len(intUseAkt(j)) - 2))
	   	    else
	   	    intUseAkt(j) = -1
	   	    end if

            if len(trim(intUseEasy(j))) > 1 then
	   	    intUseEasy(j) = trim(left(intUseEasy(j), len(intUseEasy(j)) - 2))
	   	    else
	   	    intUseEasy(j) = -1
	   	    end if

           
            'Response.Write "her:"& intuseJob(j) &" - "& intUseForvalgt(j) & "<br>"
            'Response.flush

            if len(trim(intUseForvalgt(j))) > 1 then
	   	    intUseForvalgt(j) = trim(left(intUseForvalgt(j), len(intUseForvalgt(j)) - 2))
	   	    else
	   	    intUseForvalgt(j) = 0
	   	    end if
	   	


	   	'Response.Write " intusejob DB: " & intuseJob(j) & "<br>"
	   	
	   	if intuseJob(j) <> -1 AND intuseAkt(j) <> -1 then

            if intUseForvalgt(j) <> 0 then
            intUseForvalgt(j) = 1
            else
            intUseForvalgt(j) = 0
            end if


        '** Findes record i forvejen ****'
        job_akt_medid_findes = 0
         strSQLfi = "SELECT aktid, jobid, medarb FROM timereg_usejob WHERE aktid = "& intuseAkt(j) &" AND jobid = "& intuseJob(j) &" AND medarb = "& medid 
         oRec4.open strSQLfi, oConn, 3
         if not oRec4.EOF then

         job_akt_medid_findes = 1
              
        
        end if
        oRec4.close


                if cint(job_akt_medid_findes) = 0 then
                    
                    if intUseForvalgt(j) = 1 then
                    strSQL = "INSERT INTO timereg_usejob (medarb, jobid, easyreg, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt, aktid) VALUES "_
                    &" ("& medid &", "& intuseJob(j) &", "& intuseEasy(j) &", "& intUseForvalgt(j) &", 0, "& session("mid") &", '"& dtNow &"', "& intuseAkt(j) &")"
                    else
		            strSQL = "INSERT INTO timereg_usejob (medarb, jobid, easyreg, forvalgt, aktid) VALUES ("& medid &", "& intuseJob(j) &", "& intuseEasy(j) &", "& intUseForvalgt(j) &", "& intuseAkt(j) &")"
		            end if

                'Response.write strSQL & "<br>"
		        'Response.flush
		        oConn.execute(strSQL)

                end if

       
		end if
		next
	    
	    'Response.end
	    
		end function



public erTeamleder, erTeamlederForVilkarligGruppe
function fTeamleder(medid, prgid)
    

    erTeamleder = 0
    strSQLt = "SELECT teamleder FROM progrupperelationer WHERE medarbejderID = "& medid &" AND teamleder = 1 AND ProjektgruppeId = "& prgid

    'Response.write strSQLt
    'Response.flush

    oRec3.open strSQLt, oConn, 3
    if not oRec3.EOF then
    erTeamleder = 1
    erTeamlederForVilkarligGruppe = 1
    end if
    oRec3.close
    


end function



public erNotificer
function fnotificer(prgid)
    

    erNotificer = "0"
    strSQLt = "SELECT notificer, medarbejderID FROM progrupperelationer WHERE notificer = 1 AND ProjektgruppeId = "& prgid

    'Response.write strSQLt
    'Response.flush

    oRec3.open strSQLt, oConn, 3
    while not oRec3.EOF 
    erNotificer = erNotificer &",#"& oRec3("medarbejderID") & "#"
    oRec3.movenext
    wend
    oRec3.close
    


end function



public prgNavnTxt
function prgNavn(pgid, cnt)
        
    
    prgNavnTxt = ""
    strSQL = "SELECT navn FROM projektgrupper "_
    &" WHERE id = "& pgid
    
    'Response.Write strSQL & "<br>"
    'Response.flush
    
    oRec3.Open strSQL, oConn, 0, 1
    if Not oRec3.EOF then
    
            
    prgNavnTxt = oRec3("navn")
   
    
    end if
    oRec3.close



    if cnt <> 200 then

    if prgNavnTxt <> "Ingen" then
     if cnt > 1 then 
      Response.write ", "& prgNavnTxt
     else
      Response.write prgNavnTxt
     end if
    end if

    else

    end if



end function



public prgNavnTxtShort
function prgNavnShort(pgid, cnt)
        
    
    prgNavnTxt = ""
    strSQL = "SELECT navn FROM projektgrupper "_
    &" WHERE id = "& pgid
    
    'Response.Write strSQL & "<br>"
    'Response.flush
    
    oRec3.Open strSQL, oConn, 0, 1
    if Not oRec3.EOF then
    
            
    prgNavnTxt = left(oRec3("navn"), 10)
   
    
    end if
    oRec3.close



    if cnt <> 200 then

    if prgNavnTxt <> "Ingen" then
     if cnt > 1 then 
      Response.write ", "& prgNavnTxt
     else
      Response.write prgNavnTxt
     end if
    end if

    else

    end if



end function


public prjGoptionsId, prjGoptionsTxt, prjGoptionsAntal, prgAntal, prjGTeamleder 
function projgrp(progrp,level,medarbid,visning)
    
    redim prjGoptionsId(350), prjGoptionsTxt(350), prjGoptionsAntal(350), prjGTeamleder(350)
    call teamleder_flad_fn()
    
    'level = session("rettigheder")
    '** Admin, eller indtil projgrp er sat op
    if cint(level) = 1 OR cint(teamleder_flad) = 1 then  '( (lto = "epi" OR lto = "epi_no" OR lto = "epi_ab" OR lto = "epi_sta" OR lto = "epi_uk") AND (level <=2 OR level = 6)) OR lto = "mi" then
        teamlederKri = "" 
        grpid = "pgrel1.ProjektgruppeId"
        medarbIdKri = "" 
      
            

               strSQL = "SELECT id AS ProjektgruppeId, navn AS pgnavn FROM projektgrupper WHERE id <> 0 ORDER BY navn"


    else
    
            
             
              teamlederKri = " AND (pgrel1.teamleder = 1 "  
              'AND pgrel1.MedarbejderId = "& medarbid
    
             'if thisfile = "feriekalender" then '*** Alle må gerne se andres ferie
             'teamlederKri = teamlederKri &" OR ProjektgruppeId = 10)"
             'else
             teamlederKri = teamlederKri & ")"
             'end if

             medarbIdKri = " AND pgrel1.medarbejderId = " & medarbid
    
             grpid = "pgrel1.ProjektgruppeId"


             progrpKri = " pgrel1.ProjektgruppeId <> 0"
             grpby = "pgrel1.projektgruppeid"
    
   
    
            strSQL = "SELECT "& grpid &", "_
            &" pgrel1.teamleder, pgrel1.medarbejderid, "_
            &" pg.id AS pid, pg.navn AS pgnavn FROM "_
            &" progrupperelationer AS pgrel1 "_
            &" LEFT JOIN projektgrupper AS pg ON (pg.id = pgrel1.projektgruppeid) "_
            &" WHERE "& progrpKri & " AND pg.id = pgrel1.projektgruppeid "& medarbIdKri &" "& teamlederKri &" GROUP BY "& grpby &" ORDER BY pg.navn"
    
        

    end if
    
       
       'if session("mid") = 1 then
       'Response.Write strSQL & "<br><br>"
       'end if

       'Response.flush
    
    oRec3.Open strSQL, oConn, 0, 1
    
    
    p = 0		
    While Not oRec3.EOF
    
            
    
    prjGoptionsAntal(p) = 0
    prjGTeamleder(p) = 0
            
            'call antalMediPgrp(oRec3("projektgruppeid"))
            
            
            prjGoptionsAntal(p) = 1 'antalMediPgrpX
           
            if prjGoptionsAntal(p) <> 0 then
            prjGoptionsId(p) = oRec3("projektgruppeid") 'pid   
            prjGoptionsTxt(p) = oRec3("pgnavn")
        
            if cint(teamleder_flad) = 1 OR cint(level) = 1 then
            prjGTeamleder(p) = 1
            else
            prjGTeamleder(p) = oRec3("teamleder")
            end if
            
            prgAntal = p
            p = p + 1    
            else
            prjGoptionsId(p) = 0    
            prjGoptionsTxt(p) = ""
            prjGTeamleder(p) = 0
            end if
    
    oRec3.MoveNext
    wend
    oRec3.close
    
     
End Function

public medarbgrpIdSQLkri, instrMedidProgrpThisGrp
public medarbgrpId, medarbgrpTxt, antalMedgrp, instrMedidProgrp, strOptionsJq
function medarbiprojgrp(progrp, medid, mtypesorter, seloptions)

'Response.Write "progrp" & progrp & "<br>"
'Response.end
    
    instrMedidProgrpThisGrp = ""

    redim medarbgrpId(2150), medarbgrpTxt(2150)
        
    'if progrp = "0, 0, 0" then
    'progrp = "0, 10, 0" 
    'else
    'progrp = progrp
    'end if
      

      if len(trim(medid)) <> 0 then
      medid = medid
      else
      medid = session("mid")
      end if

       prgKri = "(ProjektgruppeId = 0 "

       progrp = "0, " & progrp & ", 0"
       arr_progrp = split(progrp, ",")
       for u = 0 TO UBOUND(arr_progrp)
       
           if arr_progrp(u) <> "0" AND len(trim(arr_progrp(u))) <> 0 then
           prgKri = prgKri &  " OR ProjektgruppeId = "& arr_progrp(u)
           end if


       next


       'if session("mid") = 1 then
       'Response.write "prgKri: :"& prgKri & ":"
       'end if

       if prgKri = "(ProjektgruppeId = 0 " OR instr(prgKri, "-1") <> 0 then
       prgKri = prgKri & " OR (mid = "& medid &") "
       end if


       prgKri = prgKri & ")"



        'if session("mid") = 1 then
        'Response.write "<br>prgKri:" & prgKri
        'end if


       if len(trim(strSQLmansat)) <> 0 then
       strSQLmansat = strSQLmansat
       else
       
       strSQLmansat = "(m.mansat = 1 OR m.mansat = 3)" 
       
       end if
        
        
     if mtypesorter = 1 then
     odrBySQL = "mt.type, mnavn"
     else
     odrBySQL = "mnavn"
     end if
        

    prgKri = replace(prgKri, "ProjektgruppeId = 0 OR ", "")
   

    strSQLp = "SELECT Mnavn, Mid, ProjektgruppeId, MedarbejderId, teamleder, mnr, mansat, mt.type AS medarbejdertype, mt.sostergp, mts.type AS sosternavn, m.init, opsagtdato FROM "_
    &" progrupperelationer "_
    &" LEFT JOIN medarbejdere AS m ON (m.mid = MedarbejderId) "_
    &" LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) "_
    &" LEFT JOIN medarbejdertyper AS mts ON (mts.id = mt.sostergp) "_
    &" WHERE ("& prgKri & " AND "& strSQLmansat &")  GROUP BY mid ORDER BY "& odrBySQL
    
    
    'if session("mid") = 1 then
    'Response.Write strSQLp & "<br><br>"
    'Response.flush
    'end if
    
    'for q = 0 to UBOUND(intMids)
    'Response.write "q: "& intMids(q) & "<br>"  
    'next
   

    instrMedidProgrp = "#0#,"
    'strOptionsJqHd = "0"
    strOptionsJq = "<option value='0'>Alle</option>" 

    oRec3.Open strSQLp, oConn, 0, 1
    m = 0
    antalMedgrp = 0 
    lastMtype = ""			
    While Not oRec3.EOF
    
    if instr(instrMedidProgrp, "#"& oRec3("mid")  &"#") = 0 then
            
            medarbgrpIdSQLkri = medarbgrpIdSQLkri & " OR mid = " & oRec3("medarbejderid")
        
             instrMedidProgrpThisGrp = instrMedidProgrpThisGrp & "#"& oRec3("medarbejderid") &"#," 
    
            medarbgrpId(m) = oRec3("mid")

            if len(trim(oRec3("init"))) <> 0 then
            medarbgrpTxt(m) = oRec3("mnavn") & " ["& oRec3("init") &"]"
            else
            medarbgrpTxt(m) = oRec3("mnavn") & " ("& oRec3("mnr") &")"
            end if

            if cint(oRec3("teamleder")) = 1 then
            medarbgrpTxt(m) = medarbgrpTxt(m) & " - Teamleder"
            end if
    
            if oRec3("mansat") = "2" then
            medarbgrpTxt(m) = medarbgrpTxt(m) & " - De-aktiveret ("& oRec3("opsagtdato") &")"
            end if

            if oRec3("mansat") = "3" then
            medarbgrpTxt(m) = medarbgrpTxt(m) & " - Passiv"
            end if
    
            instrMedidProgrp = instrMedidProgrp & "#"& oRec3("mid") &"#,"   

            if mtypesorter = 1 AND (lastMtype <> oRec3("medarbejdertype")) then
            if m > 0 then
            strOptionsJq = strOptionsJq & "<option value='0' DISABLED></option>"
            end if

            sosNavn = ""
            if oRec3("sostergp") = -1 then
            sosNavn = ": (Skjult)"
            end if

            if oRec3("sostergp") > 0 then
            sosNavn = ": ("& oRec3("sosternavn") &")"
            end if

            strOptionsJq = strOptionsJq & "<option value='0' DISABLED>"& oRec3("medarbejdertype") &" "& sosNavn &"</option>" 
            end if



                mSel = ""

    
                if seloptions <> -1 then
                for t = 0 To UBOUND(intMids) 
                                
                    if mSel = "" then
                    
                        if cint(intMids(t)) = cint(medarbgrpId(m)) then
                        mSel = "SELECTED"
                        else
                        mSel = ""
                        end if
                    
                    end if
                    
                    
                
                next
                end if


            strOptionsJq = strOptionsJq & "<option value='"& oRec3("mid") &"' "& mSel &">"& medarbgrpTxt(m) &"</option>" 
            'strOptionsJqHd = strOptionsJqHd & ", "& oRec3("mid") 

            lastMtype = oRec3("medarbejdertype")
            antalMedgrp = m
            m = m + 1    

    end if
    oRec3.MoveNext
    wend
    oRec3.close
    
   
    
    
End Function

public antalMediPgrpX
function antalMediPgrp(pg)

antalMediPgrpX = 0

if len(trim(strSQLmansat)) <> 0 then
strSQLmansat = strSQLmansat
else
strSQLmansat = " m.mansat <> 2 "
end if

            '** antal aktive og passive medarbejdere i grp. ***'
            strSQLam = "SELECT COUNT(pgrel2.MedarbejderId) AS antalmed FROM medarbejdere AS m "_
            &" LEFT JOIN progrupperelationer AS pgrel2 ON (pgrel2.MedarbejderId = m.mid AND pgrel2.projektgruppeid = " & pg & ") WHERE "_
            &" "& strSQLmansat &" AND pgrel2.projektgruppeid = " & pg & " GROUP BY projektgruppeid"
            
            'Response.Write strSQLam
            'Response.flush
            
            oRec4.open strSQLam, oConn, 3
            if not oRec4.EOF then 
            antalMediPgrpX = oRec4("antalmed")
            end if
            oRec4.close

end function    


function progrupperIdNavn(id)

    strSQL = "SELECT id, navn FROM projektgrupper "_
    &" WHERE id <> 1 ORDER BY navn"
    
    'Response.Write strSQL & "<br>"
    'Response.flush
    
    oRec3.Open strSQL, oConn, 0, 3
    while Not oRec3.EOF 
    
    if cint(id) = cint(oRec3("id")) then
    pgSEL = "SELECTED"
    else
    pgSEL = ""
    end if

    %>
    <option value="<%=oRec3("id") %>" <%=pgSEL %>><%=oRec3("navn") %></option>
    <%
            
    
   
    oRec3.movenext
    wend
    oRec3.close

end function


'public progrpUsed
public strSQLmansat
sub progrpmedarb


'** Aktive, de-aktive, passive **'

if len(trim(request("FM_progrp"))) = 0 then 'Der er ikke søgt DEFAULT
visdeakmed = 0
visdeakmedCHK = ""
vispasmed = 1
vispasmedCHK = "CHECKED"
strSQLmansat = " (m.mansat = 1 OR m.mansat = 3) " 'viser aktive + passive 

else

visdeakmed = 0
vispasmed = 0
strSQLmansat = " (m.mansat = 1 " 'viser aktive 

if len(trim(request("FM_visdeakmed"))) <> 0 then
visdeakmed = 1
visdeakmedCHK = "CHECKED"

                if len(trim(request("FM_visdeakmed12"))) <> 0 then
                dd = now
                opsagtdatoKri = dateAdd("m", -12, dd)
                opsagtdatoKri = year(opsagtdatoKri) &"/"& month(opsagtdatoKri) &"/"& day(opsagtdatoKri) 

                strSQLmansat = strSQLmansat & " OR (m.mansat = 2" 'de-aktiverede
                strSQLmansat = strSQLmansat & " AND opsagtdato > '"& opsagtdatoKri &"')"

                visdeakmed12 = 1
                visdeakmed12CHK = "CHECKED"

                else
                opsagtdatoKri = ""
                strSQLmansat = strSQLmansat & " OR m.mansat = 2" 'de-aktiverede
                end if          


end if

if len(trim(request("FM_vispasmed"))) <> 0 then
vispasmed = 1
vispasmedCHK = "CHECKED"
strSQLmansat = strSQLmansat & " OR m.mansat = 3" 'viser de-aktiverede
end if


strSQLmansat = strSQLmansat & ")"

end if

%>
    <td valign=top style="padding-top:20px; width:426px;"><b>Projektgrupper:</b><br />
    <span style="font-size:10px; line-height:12px; color:#999999; padding-top:4px;">
        Admin.: alle projektgrupper, ellers dem du er teamleder for.
        </span><br />
       
       <% 
       
       pf = 0
       fo = 0

       level = session("rettigheder")


           'Response.write "HER " & progrp

       if (progrp = "0" AND level = 1) OR lto = "adra" then 'OR lto = "mi"
       
            progrp = 10 
            
           'select case lto
           'case "epi", "epi_cati", "epi_no", "epi_sta", "epi_ab", "intranet - local"
           '** Af performance hensyn henter vi en anden end "Alle" gruppen HVIS medarbejderen er medlem af andre grupper, når siden hentes første gang. 
           '** Hvis der ike fidnes medlemskaber bruges "Alle" gruppen 
           'strSQLp = "SELECT projektgruppeId FROM progrupperelationer WHERE medarbejderId = "& session("mid") &" AND projektgruppeId <> 10 LIMIT 1"
           
           'Response.write strSQLp
           'Response.flush
           'oRec6.open strSQLp, oConn, 3
           'if not oRec6.EOF then 

           'progrp = oRec6("projektgruppeId")

           'end if
           'oRec6.close

           'end select

       end if

       
       'Response.Write "<br><br><br><br><br><br><br>"& progrp & " level: "& level

       progrp = "0, " & progrp & ", 0"
       arr_progrp = split(progrp, ",")
       for u = 0 TO UBOUND(arr_progrp)
        call projgrp(trim(arr_progrp(u)),level,medarbid,0)
       next

     

       %>
     
        <select id="FM_progrp" name="FM_progrp" style="width:406px; font-size:11px;" multiple="multiple" size=9>
           
            
       <%

           
      

       for p = 0 to prgAntal 
       
       if prjGoptionsId(p) <> 0 then
        
        'if instr(progrp, ",") = 0 AND fo = 0 then
        'progrp = prjGoptionsId(p)
        '   end if
       
        if instr(progrp, ", " & prjGoptionsId(p) & ",") <> 0 then
        pgSEL = "SELECTED"
        fo = 1
        else
        pgSEl = ""
        end if%>
        <option value="<%=prjGoptionsId(p) %>" <%=pgSEl%>><%=prjGoptionsTxt(p) %> </option>

      
      <%
          

        pf = pf + 1
        
        end if
      
      next 
      
      if pf = 0 then
      %>
       <option value="-1" SELECTED>Ingen projektgruppe fundet</option>
       
      <%
      end if
      %>  
            </select>

      

            <br />
        <input id="FM_visdeakmed" name="FM_visdeakmed" type="checkbox" <%=visdeakmedCHK %> /> Vis de-aktiverede medarbejdere.  <input id="FM_visdeakmed12" name="FM_visdeakmed12" type="checkbox" <%=visdeakmed12CHK %> /> De-akt. opsagt indenfor senete 12 md.<br />
        <input id="FM_vispasmed" name="FM_vispasmed" type="checkbox" <%=vispasmedCHK %> /> Vis passive medarbejdere<br />&nbsp;
     
        <input type="hidden" id="jq_userid" value="<%=medarbid%>" />
         
        </td>
  
	<td valign=top style="padding-top:20px;"><b>Medarbejdere:</b> (<span id="antalmedarblist"><%=antalMedgrp+1 %></span>)
    <br /><img src="../ill/blank.gif" width="50" height="11"  border="0"/><br /> 
	<%
	mft = 0 
	mSel = ""
	
    if thisfile = "joblog_timetotaler" AND vis_medarbejdertyperChk = "CHECKED" then
    mTypeSort = 1
    else
    mTypeSort = 0
    end if 
    


        if cint(vis_medarbejdertyper_grp) = 1 then
            
            vlgtmtypgrp = 0
            call mtyperIGrp_fn(vlgtmtypgrp,1)    
            
            strOptionsJq = "<option value='0' DISABLED>Medarbejdertypegruppe(r)</option>"
            strOptionsJq = strOptionsJq & "<option value='0'>Alle</option>"
            'strOptionsJq = "<option value='0' DISABLED></option>"
            
            'for t = 1 to UBOUND(mtypgrpids)
                dim mSelMTypGRP
                redim mSelMTypGRP(50)    

                for mtgp = 1 to UBOUND(kunMtypgrp) 'mtypgrpids 
                 mSelMTypGRP(mtgp) = ""
                       for s = 0 To UBOUND(intMids) 
            
                        
                            if mSelMTypGRP(mtgp) = "" then
                   
                            if cint(intMids(s)) = cint(kunMtypgrp(mtgp)) then
                            mSelMTypGRP(mtgp) = "SELECTED"
                            else
                            mSelMTypGRP(mtgp) = ""
                            end if
                    
                            end if
                        
                              'strOptionsJq =  strOptionsJq & "<option value='"& kunMtypgrp(t) &"' "& mSelMTypGRP(mtgp) &">"& kunMtypgrpNavn(t) &" // "& intMids(s) &"</option>" 
                  

                        next

            
                if len(trim(kunMtypgrpNavn(mtgp))) <> 0 then 
                strOptionsJq =  strOptionsJq & "<option value='"& kunMtypgrp(mtgp) &"' "& mSelMTypGRP(mtgp) &">"& kunMtypgrpNavn(mtgp) &"</option>" 
                end if

                next


        else 
	        
            call medarbiprojgrp(progrp, medarbid, mTypeSort, 1)
        end if    


	'call medarbiprojgrp(progrp, medarbid, mTypeSort, 1)
	'Response.Write progrp &","& medarbid
	
        
        if thisfile = "bal_real_norm_2007.asp" then
        mselWidth = 250
        else
        mselWidth = 350
        end if


    %>
    <select name="FM_medarb" id="FM_medarb" multiple style="width:<%=mselWidth%>px; font-size:11px;" size=9>
    <%=strOptionsJq %>
    </select> 
      
        

     <%if thisfile = "joblog_timetotaler" then %>
             <br />
              <input type="checkbox" name="FM_vis_medarbejdertyper" id="FM_vis_medarbejdertyper" value="1" <%=vis_medarbejdertyperChk %> />Udspecificer på medarbejdertyper<br />

                <%if cint(bdgmtypon_val) = 1 AND cint(bdgmtypon_prgrp) > 1 then  %>
               <input type="checkbox" name="FM_vis_medarbejdertyper_grp" id="FM_vis_medarbejdertyper_grp" value="1" <%=vis_medarbejdertyper_grpChk %> />Udspecificer på medarbejdertype-grupper
               <br /> <span style="font-size:9px; color:#999999;">(ignorerer projektgrupper. Viser alle medarb. uanset status)</span><br /><br />  
            <%end if %>    

              <input type="checkbox" name="FM_visMedarbNullinier" id="FM_visMedarbNullinier" value="1" <%=vis_visMedarbNullinierChk %> />Vis medarbejdere uden real. timer i perioden
              
             
		    <%end if %>
	
	
	<%
    strFMmedarb_hd = "0"

    if cint(vis_medarbejdertyper_grp) <> 1 then
            
    for m = 0 to antalMedgrp 

    if len(trim(medarbgrpId(m))) <> 0 AND medarbgrpId(m) > 0 then 
    strFMmedarb_hd = strFMmedarb_hd & ", "& medarbgrpId(m) 
    end if%>
	
	<%next 
        
    end if%>
	<input id="FM_medarb_hidden" name="FM_medarb_hidden" type="hidden" value="<%=strFMmedarb_hd%>" />


       <br /><br /><img src="../ill/blank.gif" width="200" height="1" border="0" /><input id="Submit2" type="submit" value="Vis medarbejdere >> " style="font-size:9px;" />
	</td>

<% 
end sub



                  
        '**** tilføj Projektgrupper ****'

        public strProjektgr1, strProjektgr2, strProjektgr3, strProjektgr4
        public strProjektgr5, strProjektgr6, strProjektgr7, strProjektgr8, strProjektgr9, strProjektgr10
        public prgIsWrt1,prgIsWrt2,prgIsWrt3,prgIsWrt4,prgIsWrt5,prgIsWrt6,prgIsWrt7,prgIsWrt8,prgIsWrt9,prgIsWrt10
        public useFieldValIsWrt

        function tilfojProGrp(func,aj,jid,aid,nedfod,grp1,grp2,grp3,grp4,grp5,grp6,grp7,grp8,grp9,grp10,firstLoop)

        'grp1,grp2,... ER DE projektgrupper der er tilføjet på jobbet og som der skal sammelignes med ****'
        '** Uanset om der er valgt nedarv på aktiviteter eller fød job **'

        'Response.write "firstLoop" & firstLoop & "<br>"
        'Response.flush 

       
                 if cint(firstLoop) = 0 then
                 prgIsWrt1 = 0
                 prgIsWrt2 = 0
                 prgIsWrt3 = 0
                 prgIsWrt4 = 0
                 prgIsWrt5 = 0
                 prgIsWrt6 = 0
                 prgIsWrt7 = 0
                 prgIsWrt8 = 0
                 prgIsWrt9 = 0
                 prgIsWrt10 = 0

                 useFieldValIsWrt = ""
                 firstLoop = 1
                 end if



                'Response.write "<br>aj: " & aj &" fm_alle " & fm_alle  & " func  "& func  &" nedfod "& nedfod &" aid = "& aid &" jid: "& jid &"<br>"
                'Response.flush 

                  if jid <> 0 then 'jobids(j) <> 0 : er det en stamakt (=0)
                        
                        if cint(aj) = 1 then '1 = Job, finder aktiviet 
                        strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE id =" & aid
				        else '** 2 = Akt, finder job
                        strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id =" &jid
				        end if


                        'Response.write "<hr>"& strSQL & "<br>" 

                        oRec5.open strSQL, oConn, 3
				
				
					        if not oRec5.EOF then
					        
                            
                            strProjektgr1Tjk = oRec5("projektgruppe1")
					        strProjektgr2Tjk = oRec5("projektgruppe2")
					        strProjektgr3Tjk = oRec5("projektgruppe3")
					        strProjektgr4Tjk = oRec5("projektgruppe4")
					        strProjektgr5Tjk = oRec5("projektgruppe5")
					        strProjektgr6Tjk = oRec5("projektgruppe6")
					        strProjektgr7Tjk = oRec5("projektgruppe7")
					        strProjektgr8Tjk = oRec5("projektgruppe8")
					        strProjektgr9Tjk = oRec5("projektgruppe9")
					        strProjektgr10Tjk = oRec5("projektgruppe10")
                           
                            end if
                                
				            oRec5.close
                            
                            
                            

                            
                            '*** Hvis fød job er valgt, enten fra job opret/red eller fra  stamakt opret. ***
                            '*** Rediger akt eller opret akt kan ikke føde job da der kun kan vælges grupperder allerede er tilknyttet jobbet
                            'if (cint(aj) = 2 AND cint(nedfod) = 1 AND func = "dbopr") OR _ 
                            '(cint(aj) = 1 AND cint(nedfod) = 1)  then '*** Nedarv fra job, hvis opret, eller opret stam. Ellers hvis slået til ved rediger 
				            
                            ''(cint(aj) = 2 AND fm_alle <> "1" AND func = "dbred") OR
                            
                            '*** Fød job **'
                            if cint(nedfod) = 1 then

				            '*** Opretter gruppen på job hvis den ikke findes ****'
                            for p = 1 to 10

                            
                             

                             select case p
                             case 1
                             tjkGrp = grp1
                             case 2
                             tjkGrp = grp2
                             case 3
                             tjkGrp = grp3
                             case 4
                             tjkGrp = grp4
                             case 5
                             tjkGrp = grp5
                             case 6
                             tjkGrp = grp6
                             case 7
                             tjkGrp = grp7
                             case 8
                             tjkGrp = grp8
                             case 9
                             tjkGrp = grp9
                             case 10
                             tjkGrp = grp10
                             end select

                            
                             
                            'Response.write "<br>her grp:" & cint(tjkGrp) & ", strProjektgr1Tjk = "& strProjektgr1Tjk &"<br>"_
                            '&"strProjektgr2Tjk = "& strProjektgr2Tjk &"<br>"_
                            '&"strProjektgr3Tjk = "& strProjektgr3Tjk &"<br>"_
                            '&"strProjektgr4Tjk = "& strProjektgr4Tjk &"<br>"_
                            '&"strProjektgr5Tjk = "& strProjektgr5Tjk &"<br>"_
                            '&"strProjektgr6Tjk = "& strProjektgr6Tjk &"<br>"_
                            '&"strProjektgr7Tjk = "& strProjektgr7Tjk &"<br>"_
                            '&"strProjektgr8Tjk = "& strProjektgr8Tjk &"<br>"_
                            '&"strProjektgr9Tjk = "& strProjektgr9Tjk &"<br>"_
                            '&"strProjektgr10Tjk = "& strProjektgr10Tjk &"<br>"
                                
                             if cint(tjkGrp) <> cint(strProjektgr1Tjk) _
                             OR cint(tjkGrp) <> cint(strProjektgr2Tjk) _
                             OR cint(tjkGrp) <> cint(strProjektgr3Tjk) _
                             OR cint(tjkGrp) <> cint(strProjektgr4Tjk) _
                             OR cint(tjkGrp) <> cint(strProjektgr5Tjk) _
                             OR cint(tjkGrp) <> cint(strProjektgr6Tjk) _
                             OR cint(tjkGrp) <> cint(strProjektgr7Tjk) _
                             OR cint(tjkGrp) <> cint(strProjektgr8Tjk) _
                             OR cint(tjkGrp) <> cint(strProjektgr9Tjk) _ 
                             OR cint(tjkGrp) <> cint(strProjektgr10Tjk) then    
                             
                             

                           
                                            opdGrpNr = 0

                                            if cint(aj) = 2 then 'NEDARV akt fra job

                                           


                                            if cint(strProjektgr10Tjk) = 1 AND cint(prgIsWrt10) = 0 then
                                            opdGrpNr = 10
                                            end if

                                             if cint(strProjektgr9Tjk) = 1 AND cint(prgIsWrt9) = 0 then
                                            opdGrpNr = 9
                                            end if

                                             if cint(strProjektgr8Tjk) = 1 AND cint(prgIsWrt8) = 0 then
                                            opdGrpNr = 8
                                            end if

                                             if cint(strProjektgr7Tjk) = 1 AND cint(prgIsWrt7) = 0 then
                                            opdGrpNr = 7
                                            end if

                                             if cint(strProjektgr6Tjk) = 1 AND cint(prgIsWrt6) = 0 then
                                            opdGrpNr = 6
                                            end if


                                             if cint(strProjektgr5Tjk) = 1 AND cint(prgIsWrt5) = 0 then
                                            opdGrpNr = 5
                                            end if


                                            if cint(strProjektgr4Tjk) = 1 AND cint(prgIsWrt4) = 0 then
                                            opdGrpNr = 4
                                            end if


                                            if cint(strProjektgr3Tjk) = 1 AND cint(prgIsWrt3) = 0 then
                                            opdGrpNr = 3
                                            end if


                                            if cint(strProjektgr2Tjk) = 1 AND cint(prgIsWrt2) = 0 then
                                            opdGrpNr = 2
                                            end if

                                            if cint(strProjektgr1Tjk) = 1 AND cint(prgIsWrt1) = 0 then
                                            opdGrpNr = 1
                                            end if


                                            useFieldVal = tjkGrp

                                            else

                                            'Response.write "FØD<br>"


                                            if cint(grp10) = 1 AND cint(prgIsWrt10) = 0 then
                                            opdGrpNr = 10
                                            end if

                                             if cint(grp9) = 1 AND cint(prgIsWrt9) = 0 then
                                            opdGrpNr = 9
                                            end if

                                             if cint(grp8) = 1 AND cint(prgIsWrt8) = 0 then
                                            opdGrpNr = 8
                                            end if

                                             if cint(grp7) = 1 AND cint(prgIsWrt7) = 0 then
                                            opdGrpNr = 7
                                            end if

                                             if cint(grp6) = 1 AND cint(prgIsWrt6) = 0 then
                                            opdGrpNr = 6
                                            end if


                                             if cint(grp5) = 1 AND cint(prgIsWrt5) = 0 then
                                            opdGrpNr = 5
                                            end if


                                            if cint(grp4) = 1 AND cint(prgIsWrt4) = 0 then
                                            opdGrpNr = 4
                                            end if


                                            if cint(grp3) = 1 AND cint(prgIsWrt3) = 0 then
                                            opdGrpNr = 3
                                            end if


                                            if cint(grp2) = 1 AND cint(prgIsWrt2) = 0 then
                                            opdGrpNr = 2
                                            end if

                                            if cint(grp1) = 1 AND cint(prgIsWrt1) = 0 then
                                            opdGrpNr = 1
                                            end if



                                             select case p
                                             case 1
                                             useFieldVal = strProjektgr1Tjk
                                             case 2
                                             useFieldVal = strProjektgr2Tjk
                                             case 3
                                             useFieldVal = strProjektgr3Tjk
                                             case 4
                                             useFieldVal = strProjektgr4Tjk
                                             case 5
                                             useFieldVal = strProjektgr5Tjk
                                             case 6
                                             useFieldVal = strProjektgr6Tjk
                                             case 7
                                             useFieldVal = strProjektgr7Tjk
                                             case 8
                                             useFieldVal = strProjektgr8Tjk
                                             case 9
                                             useFieldVal = strProjektgr9Tjk
                                             case 10
                                             useFieldVal = strProjektgr10Tjk
                                             end select

                                            

                                            end if




                                                '*** Opdater progrp på job hvis ikke alle 10 er udfyldt.
                                                if cint(opdGrpNr) <> 0 AND useFieldVal <> 1 then


                                                select case opdGrpNr
                                                case 1
                                                prgIsWrt1 = 1
                                                case 2
                                                prgIsWrt2 = 1
                                                case 3
                                                prgIsWrt3 = 1
                                                case 4
                                                prgIsWrt4 = 1
                                                case 5
                                                prgIsWrt5 = 1
                                                case 6
                                                prgIsWrt6 = 1
                                                case 7
                                                prgIsWrt7 = 1
                                                case 8
                                                prgIsWrt8 = 1
                                                case 9
                                                prgIsWrt9 = 1
                                                case 10
                                                prgIsWrt10 = 1
                                                end select


                                                

                                                if instr(useFieldValIsWrt, "#"&useFieldVal&"#") = 0 then


                                                    

                                                projektgruppeFlt = "projektgruppe"&opdGrpNr 

                                                strSQLupdjobprgrp = "UPDATE job SET "& projektgruppeFlt &" = " & useFieldVal &" WHERE id =" &jid   
                                                oConn.execute(strSQLupdjobprgrp) 

                                                'Response.write "SQL: " & strSQLupdjobprgrp & "<br>"
                                                'Response.flush

                                                useFieldValIsWrt = useFieldValIsWrt & "#"&useFieldVal&"#" 

                                                end if

                                                end if
                             
                                 end if '** Opr grp hvis den ikke findes 


                            next

                           
                     end if '*** fød (nedfod = 1)
					    
                       


				
				
				end if


                 'Response.write "her slut"
                 'Response.end

                ' end if 'func / aj = 2
                


                '** Output **'
                'Response.write "aj:" & aj & "jid " & jid & "nedfod "& nedfod 
                'Response.end

                '*** A) Ved opret stamakt, tilføj til stamgrp
                '** B) Ved rediger akt / stam akt skal grupper altid følge de valgte på akt. (med mindre nedarv er valgt nedfod = 0)
                '** C) Opret stamakt tilføj til stamaktgrp  'OR (cint(aj) = 2 AND func = "dbopr" AND cint(nedfod) = 0)

                '** hvad bliver oprettet ***
                'aj = 2 Aktivitet
                'aj = 1 job


                if cint(aj) = 2 AND cint(jid) = 0 _
                 OR (cint(aj) = 2 AND cint(nedfod) = -1) then

                strProjektgr1 = grp1
				strProjektgr2 = grp2
				strProjektgr3 = grp3
				strProjektgr4 = grp4
				strProjektgr5 = grp5
				strProjektgr6 = grp6
				strProjektgr7 = grp7
				strProjektgr8 = grp8
				strProjektgr9 = grp9
				strProjektgr10 = grp10

                'Response.write "FFFF<br>"

                end if
                 
    

                '** Ved opret/rediger akt på job og nedarv fra job, bruse de progrp der blev fundet på jobbet **'
                '** Ved tilføj stamgrp til job og FØD job valgt | aj = 1 & nedfod = 1
                if (cint(aj) = 2 AND cint(nedfod) = 0 AND cint(jid) <> 0) OR cint(aj) = 1 AND cint(nedfod) = 1  then

                strProjektgr1 = strProjektgr1Tjk
				strProjektgr2 = strProjektgr2Tjk
				strProjektgr3 = strProjektgr3Tjk
				strProjektgr4 = strProjektgr4Tjk
				strProjektgr5 = strProjektgr5Tjk
				strProjektgr6 = strProjektgr6Tjk
				strProjektgr7 = strProjektgr7Tjk
				strProjektgr8 = strProjektgr8Tjk
				strProjektgr9 = strProjektgr9Tjk
				strProjektgr10 = strProjektgr10Tjk
                
                'Response.write "TTTT<br>"

                end if

                 'Response.flush

        end function




    public strPgrpSQLkri, antalpgrp, strSQLkri3, projektgruppeIds
	function hentbgrppamedarb(medarbid)
		
		f = 0
		dim bgrpId
		Redim bgrpId(f)
		
        projektgruppeIds = ""
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid= "& medarbid &" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid GROUP BY ProjektgruppeId" 
			oRec.Open strSQL, oConn, 0, 1 
			
			While Not oRec.EOF
			Redim preserve bgrpId(f) 
				bgrpId(f) = oRec("ProjektgruppeId")
				'Response.write oRec("ProjektgruppeId") & "<br>"
                projektgruppeIds = projektgruppeIds & oRec("ProjektgruppeId") & ", " 
               
				f = f + 1
				oRec.MoveNext
			Wend
		
		oRec.Close
		
        if len(trim(projektgruppeIds_len)) <> 0 then
        projektgruppeIds_len = len(projektgruppeIds)
        projektgruppeIds_left = left(projektgruppeIds, projektgruppeIds_len - 2)
        projektgruppeIds = projektgruppeIds_left
        end if

		antalpgrp = f
		
			'Rettigheds tjeck på job
			'***********************************************************************
	  	if antalpgrp > 0 then
		
		
				'********************************************************************
				'************ Job ***************************************************
				strPgrpSQLkri =  ""
				strPgrpSQLkri = strPgrpSQLkri &" AND ( "
				
				for intcounter = 0 to f - 1  
			  
			  	strPgrpSQLkri = strPgrpSQLkri &" j.projektgruppe1 = "&bgrpId(intcounter)&""_
				&" OR "_
				&" j.projektgruppe2 = "&bgrpId(intcounter)&""_
				&" OR "_
				&" j.projektgruppe3 = "&bgrpId(intcounter)&""_
				&" OR "_
				&" j.projektgruppe4 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe5 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe6 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe7 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe8 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe9 = "&bgrpId(intcounter)&""_
				&" OR "_
			  	&" j.projektgruppe10 = "&bgrpId(intcounter)&""_
				&" OR "
				
				next
			  	
				'** Trimmer de 2 sql states ***
				strPgrpSQLkri_len = len(strPgrpSQLkri)
				strPgrpSQLkri_left = strPgrpSQLkri_len - 3
				strPgrpSQLkri_use = left(strPgrpSQLkri, strPgrpSQLkri_left) 
			  	strPgrpSQLkri = strPgrpSQLkri_use & ") "
		
		
		
		'**************************** Rettigheds tjeck aktiviteter **************************
		  	strSQLkri3 = " a.job = j.id AND aktstatus = 1 AND ( "
			
			for intcounter = 0 to f - 1  
		  
			strSQLkri3 = strSQLkri3 &" a.projektgruppe1 = "& bgrpId(intcounter) &""_
			&" OR a.projektgruppe2 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe3 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe4 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe5 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe6 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe7 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe8 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe9 = "& bgrpId(intcounter) &" "_
			&" OR a.projektgruppe10 = "& bgrpId(intcounter) &" OR "
			
			next
		  	
			'** Trimmer sql states ***
			strSQLkri3_len = len(strSQLkri3)
			strSQLkri3_left = strSQLkri3_len - 3
			strSQLkri3_use = left(strSQLkri3, strSQLkri3_left)  
		  	strSQLkri3 = strSQLkri3_use &") "
		
		
		else

		
		
		
		'*** Job ***
		strPgrpSQLkri = strPgrpSQLkri &" AND ( "
		strPgrpSQLkri = strPgrpSQLkri &" j.projektgruppe1 = 0"_
		&" OR "_
		&" j.projektgruppe2 = 0"_
		&" OR "_
		&" j.projektgruppe3 = 0"_
		&" OR "_
		&" j.projektgruppe4 = 0"_
		&" OR "_
	  	&" j.projektgruppe5 = 0"_
		&" OR "_
	  	&" j.projektgruppe6 = 0"_
		&" OR "_
	  	&" j.projektgruppe7 = 0"_
		&" OR "_
	  	&" j.projektgruppe8 = 0"_
		&" OR "_
	  	&" j.projektgruppe9 = 0"_
		&" OR "_
	  	&" j.projektgruppe10 = 0"_
		&" )"
		
		
		
		'***  Aktiviteter ***
		strSQLkri3 = " a.job = j.id AND aktstatus = 1 AND ( "
		strSQLkri3 = strSQLkri3 &" a.projektgruppe1 = 0"_
		&" OR a.projektgruppe2 = 0"_
		&" OR a.projektgruppe3 = 0"_
		&" OR a.projektgruppe4 = 0"_
		&" OR a.projektgruppe5 = 0"_
		&" OR a.projektgruppe6 = 0"_
		&" OR a.projektgruppe7 = 0"_
		&" OR a.projektgruppe8 = 0"_
		&" OR a.projektgruppe9 = 0"_
		&" OR a.projektgruppe10 = 0)"
		
		end if
		
		
	end function
    
%>