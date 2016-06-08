
<%


    public afstemnul
	public meNormTimer, meNormTimerDag, meShowNormTimerdag, meNormTimerUge
    redim meNormTimer(1600), meNormTimerDag(1600), meShowNormTimerdag(1600), meNormTimerUge(1600), afstemnul(1600)
	
    public usedMeTypes 
    public strEksport, strEksportTxt, strEksportEmailTxt, headerwrt, ekspTxtA, ekspTxtB, headerwrtExp
    public realNormBal, normLontBal, realLontBal

	function medarbafstem(intMid, startDato, slutDato, visning, akttype_sel, m)

  
    'if session("mid") = 1 AND lto = "esn" then
    'response.Write "intMid: "& intMid &", startDato: "& startDato &", slutDato: "& slutDato &", visning: "& visning & "<br>"
    'end if

	
	x = m
	'Response.Write "akttype_sel" & akttype_sel & "<br>visning: "&visning

    


	select case visning 
	'case 5, 1
	'visTotalTimerAlltyp = 1
	case 1,2,3,4,5,7,14,6,77,50
	call akttyper2009(2) '7
	visTotalTimerAlltyp = 1
	case else
	visTotalTimerAlltyp = 0
	end select
	
	
	'**Viser kun tom igår ***'
	'slutdato = dateadd("d", -1, slutdato)
	
	
	startDato = year(startdato)&"/"&month(startdato)&"/"&day(startdato)
	slutDato = year(slutdato)&"/"&month(slutdato)&"/"&day(slutdato)
	
	stdato = day(startdato)&"/"&month(startdato)&"/"&year(startdato)
	sldato = day(slutdato)&"/"&month(slutdato)&"/"&year(slutdato)
	
  
	
	
	

    	 
	dim normTimerDag
	redim normTimerDag(1200)
	
	
	dim realTimer, medarbNavn, medarbNr, medarbInit, realIfTimer
	redim realTimer(1200), realIfTimer(1200), medarbNavn(1200), medarbNr(1200), medarbInit(1200)
	
	dim normTimerUge, normTimer
	redim normTimerUge(1200), normTimer(1200)
	
	 dim fakTimer, mfForbrug , resTimer
	 dim medarbEmail, fradragTimer, medarbId, realfTimer
	 
	 redim resTimer(1200), fakTimer(1200), mfForbrug(1200)
	 redim medarbEmail(1200), fradragTimer(1200), medarbId(1200), realfTimer(1200)

	 redim strEksport(1200)
	 
	
	if visTotalTimerAlltyp = 1 then
	
	intMid = trim(intMid) 
	'Response.Write intMid
	'Response.flush 
	  
	if cint(intMid) <> 0 then
	medarbSQL = " mid = " & intMid & ""
	else
	medarbSQL = " mid <> 0 "
	end if
	
	
	
	'*** Alle typer Arr ***'
	'dim akttype_sel_arr
	'redim akttype_sel_arr(100)
    'akttype_sel_arr = split(replace(aktiveTyper, "#", ""), ",") 'akttype_sel
	
	'*** nullinie tjek **'
	afstemnul(x) = 0
	
	'Response.Write "akttype_sel: " & akttype_sel & " visning: "& visning & "<br><br>"
	
	 if ((instr(akttype_sel, "#-1#") <> 0 OR instr(akttype_sel, "#-20#") <> 0) AND visning = 1) OR visning = 2 OR visning = 3 OR visning = 5 OR visning = 7 OR visning = 14 OR visning = 77 OR visning = 50 then
	'** Timer alle der tæller med i dagligt timeregnskab + korrektion overført 114 **'
	strSQL = "SELECT t.tid, sum(t.timer) AS realtimer, m.mnavn, m.mnr, m.init, m.email, "_
	&" t.tdato, m.mid, m.medarbejdertype, tfaktim FROM medarbejdere m "_
	&" LEFT JOIN timer t ON (t.tmnr = m.mid AND ("& aty_sql_realhours 

    'if (instr(akttype_sel, "#-1#") <> 0 OR instr(akttype_sel, "#-20#") <> 0) AND visning = 1 then
    if cint(mdoversigt_ultimo) = 1 then
    strSQL = strSQL &" OR tfaktim = 114"
    end if

    strSQL = strSQL &")"

	strSQL = strSQL &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"') "_
	&" WHERE "& medarbSQL &" AND mansat <> '2' "_
	&" GROUP BY m.mid ORDER BY m.mnavn"
	
    'OR tfaktim = 114
    'if session("mid") = 1 then
	'Response.Write strSQL &"<br>"
	'Response.flush
    'end if
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	     realTimer(x) = oRec("realtimer")
	     if len(trim(realTimer(x))) <> 0 then
	     realTimer(x) = realTimer(x)
	     else
	     realTimer(x) = 0
	     end if
	
	
	'x = x + 1
	oRec.movenext
    wend
	oRec.close
	
	if realTimer(x) <> 0 then
	afstemnul(x) = 100
	end if
	
	end if 'visning
	
	
	
	'Response.Write oRec("tid") & " - " & oRec("tdato") & " - " & oRec("realtimer") & " - " & oRec("mnavn") & "<br>"
	    
	  
	    
	    
	    call meStamdata(intMid)
	    
	    medarbEmail(x) = meEmail 'oRec("email")
	    medarbNavn(x) = meNavn '& "("& meType &")" 'oRec("mnavn")
	    medarbNr(x) = meNr 'oRec("mnr")
	    medarbInit(x) = meInit 'oRec("init")
	    medarbId(x) =  intMid 'oRec("mid")
	    
	    
	  

         '** DageDiff er altid hele perioden da normtimerPer() selv tjekker for ansættelses dato.  
         dageDiff = dateDiff("d", stdato, sldato, 2, 2)

         '** Er ansat daot efter statdato i interval ****'
         if cdate(meAnsatDato) <= cDate(stdato) then
                
               weekDiff = dateDiff("ww", stdato, sldato, 2, 2)

         else

                weekDiff = dateDiff("ww", meAnsatDato, sldato, 2, 2)
	
	     end if

        
        if len(weekDiff) <> 0 AND weekDiff <> 0 then
	    weekDiff = cint(weekDiff)
	    else
	    weekDiff = 1
	    end if
	     

         'Response.Write "weekDiff" & weekDiff

	     'if (instr(usedMeTypes, "#"&meType&"#") = 0 AND ( _
	     'cdate(meAnsatDato) <= cDate(startDato) OR meTypeHist = 1) OR len(trim(usedMeTypes)) = 0) OR visning <> 1 then
	     
	             if intMid <> 0 then
	         
	             call normtimerPer(intMid, stDato, dageDiff, 0)
	             normTimer(x) = ntimPer 
                 showNormTimerdag = normTimer(x) / antalDageMtimer
                 end if
        	    
                    'Response.Write "intMid: "& intMid &" stDato " & stDato & " sldato "& sldato &" dageDiff " & dageDiff & " normTimer(x) "& normTimer(x) &"<br>"
	            
        	    
	             if len(trim(normTimer(x))) <> 0 then
	             normTimer(x) = normTimer(x)
	             else
	             normTimer(x) = 0
	             end if
        	    
        	            '*** Fejl, denne skal tage højde for medarb. typen på det pågældende tidspunkt ***'
                        
                        '*** Normeret uge til (til bl.a ferie beregning + gns. fuld dag) ***'
	                    '** Standard normtimer / dag uanset helligdage. 7,4 (37 tim.) skal altid give 1 hel dag.
	                    'call nortimerStandardDag(meType)
	                    'call normtimerPer(intMid, stDato, dageDiff)
                        'normTimerDag(x) = formatnumber(normtimerStDag) 'formatnumber(ntimPer/(5*weekDiff),1)

                        '** ALtid = 1 efter at Ferie og sygdom pr. omreges i SQL Kaldet altid indtastetes sdom DAGE 1 eller 0,5
                        normTimerDag(x) = 1

                'if normTimerDag(x) <> 0 then
	            'normTimerDag(x) = normTimerDag(x)
	            'showNormTimerdag = normTimerDag(x)
	            'else
	            'normTimerDag(x) = 0
	            'showNormTimerdag = 0
	            'end if
	            
        	    
        	    '**** Bruges denne ? *************
	            normTimerUge(x) = ntimPer/weekDiff
        	    
        	   
	            if normTimerUge(x) <> 0 then
	            normTimerUge(x) = normTimerUge(x)
	            else
	            normTimerUge(x) = 0
	            end if
	     
	    
	    meNormTimer(meType) = normTimer(x)
	    meNormTimerDag(meType) = normTimerDag(x)
	    meShowNormTimerdag(meType) = showNormTimerdag
	    meNormTimerUge(meType) = normTimerUge(x)
	      
	    usedMeTypes = usedMeTypes & ",#" & meType & "#"
	    'else
	    
	    'normTimer(x) = meNormTimer(meType) 
	    'normTimerDag(x) = meNormTimerDag(meType)
	    'showNormTimerdag = meShowNormTimerdag(meType)
	    'normTimerUge(x) = meNormTimerUge(meType)
	    
	    'end if
	    
	    'Response.Write "meType: " & meType & " normtimer: "& meNormTimer(meType) & "<br>"
	    
	    
	    select case visning
	    case 1,5,7,6,4,14,77
	    call fordelpaaaktType(intMid, startDato, slutDato, visning, akttype_sel, x)
	    'case 14
	    'call fordelpaaaktType(intMid, startDato, slutDato, visning, "#-99#, "& aktiveTyper, x)
	    end select
	    
	    
	    
	  
        'Response.Write "meType: " & meType & " normtimer: "& meNormTimer(meType) & " akttype_sel = "& akttype_sel &"<br>"
	   
	   '*********************************'
	   '*** Kalder sum på samle-typer ***'
	   '*********************************'
       if session("stempelur") <> 0 then
            

       'response.write akttype_sel
	   if ((instr(akttype_sel, "#-1#") <> 0 OR instr(akttype_sel, "#-5#") <> 0 OR instr(akttype_sel, "#-10#") <> 0) AND visning = 1) OR visning = 2 OR visning = 3 OR visning = 5 OR visning = 7 OR visning = 14 OR visning = 77 OR visning = 50 then 
	   'if visning = 1 OR visning = 2 OR visning = 3 OR visning = 5 OR visning = 7 OR visning = 14 then
	    
	    
	     '*** Fradrag i løntimer ***'
	    
	    strSQLfradrag = "SELECT t.tid, sum(t.timer) AS fratimer, "_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& intMid &" AND ("& aty_sql_frawhours &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLfradrag & "<br><br>"
	   'Response.end
	   
	    oRec2.open strSQLfradrag, oConn, 3
	    while not oRec2.EOF
	    fradragTimer(x) = oRec2("fratimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    
	    strSQLtillag = "SELECT t.tid, sum(t.timer) AS tiltimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& intMid &" AND ("& aty_sql_tilwhours &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLtillag
	   'Response.flush
	    
	    tilTimer = 0
	    oRec2.open strSQLtillag, oConn, 3
	    while not oRec2.EOF
	    tilTimer = oRec2("tiltimer")
	    oRec2.movenext
	    wend
	    oRec2.close


        '** Tillæg omregnet til enheder DE TYPER DERER SAT til enheder
        strSQLtillagE = "SELECT t.tid, sum(t.timer * a.faktor) AS tilenheder,"_
	    &" t.tdato "_
        &" FROM timer t LEFT JOIN aktiviteter a ON (a.id = t.taktivitetid) WHERE t.tmnr = "& intMid &" AND ("& aty_sql_tilwhours2 &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLtillagE
	   'Response.flush
	    
	    tilenheder = 0
	    oRec2.open strSQLtillagE, oConn, 3
	    while not oRec2.EOF
	    tilenheder = oRec2("tilenheder")
	    oRec2.movenext
	    wend
	    oRec2.close


         '** Fradrag omregnet til enheder DE TYPER DERER SAT til enheder
         strSQLfraE = "SELECT t.tid, sum(t.timer * a.faktor) AS tilenheder,"_
	    &" t.tdato "_
        &" FROM timer t LEFT JOIN aktiviteter a ON (a.id = t.taktivitetid) WHERE t.tmnr = "& intMid &" AND ("& aty_sql_frawhours2 &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLtillag
	   'Response.flush
	    
	    fraenheder = 0
	    oRec2.open strSQLfraE, oConn, 3
	    while not oRec2.EOF
	    fraenheder = oRec2("tilenheder")
	    oRec2.movenext
	    wend
	    oRec2.close
        '***********************************************************************************

        
        tilenheder = (tilenheder - fraenheder)
        tilTimer = tilTimer + (tilenheder)

	    
	    '*** - minus (omvendt fortegn) tillæg da angivet timer i tilæg skal gøren fradrags tallet mindre
	     if len(trim(fradragTimer(x))) <> 0 then
         fradragTimer(x) = tilTimer - (fradragTimer(x))  
         else
         fradragTimer(x) = tilTimer
         end if
	    
	     if fradragTimer(x) <> 0 then 
	     afstemnul(x) = 106
	     end if
	     
	     
	    end if
	    
	    end if
	    
	    



	    
	    'if visning = 1 OR visning = 2 OR visning = 3 OR visning = 7 OR visning = 14 then
        'Response.write akttype_sel
	     if (instr(akttype_sel, "#-1#") <> 0 AND visning = 1) OR visning = 2 OR visning = 3 OR visning = 7 OR visning = 14 OR visning = 77 then
	  
	      
	    '***  Fakturerbare TOTAL ***'
	    strSQLf = "SELECT t.tid, sum(t.timer) AS faktimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& intMid &" AND ("& aty_sql_realHoursFakbar &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLf
	   'Response.flush
	   
	    oRec2.open strSQLf, oConn, 3
	    while not oRec2.EOF
	    realfTimer(x) = oRec2("faktimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	     
	    
	     if len(trim(realfTimer(x))) <> 0 AND realfTimer(x) <> 0 then
         realfTimer(x) = realfTimer(x)
         afstemnul(x) = 105
         else
         realfTimer(x) = 0
         end if
	    
	    'if visning = 1 then
	   
	    '*** Interne Timer / ikke fakturerbare TOTAL ***'
	    strSQLif = "SELECT t.tid, sum(t.timer) AS ikkefaktimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& intMid &" AND ("& aty_sql_realHoursIkFakbar &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLif
	   'Response.flush
	   
	    oRec2.open strSQLif, oConn, 3
	    while not oRec2.EOF
	    realIfTimer(x) = oRec2("ikkefaktimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	    if len(trim(realIfTimer(x))) <> 0 AND realIfTimer(x) <> 0 then
         realIfTimer(x) = realIfTimer(x)
         afstemnul(x) = 104
         else
         realIfTimer(x) = 0
         end if
	    
	    
	    
	    end if
	
	    
	    '*** Ressource Timer ***'
	    if (instr(akttype_sel, "#-2#") <> 0 AND visning = 1) then 
	    
	    strSQLres = "SELECT sum(r.timer) AS restimer FROM ressourcer_md r WHERE r.medid = "& intMid &" AND "_
	    &" (aar >= YEAR('"& startdato &"') AND md >= MONTH('"& startdato &"') "_
	    &" AND aar <= YEAR('"& slutdato &"') AND md <= MONTH ('"& slutdato &"'))"
	    
	    'Response.Write strSQLres & "<br>"
	    
	    oRec2.open strSQLres, oConn, 3
	    while not oRec2.EOF
	    resTimer(x) = oRec2("restimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    end if
	    
	    
	    '*** Faktureret ***'
	    'if instr(akttype_sel, "#-3#") <> 0 then 
	    if (instr(akttype_sel, "#-3#") <> 0 AND visning = 1) then
	    
	    strSQLfak = "SELECT sum(fms.fak) AS faktimer, f.fid, f.fakdato FROM fakturaer f "_
	    &" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.mid = "& intMid&")"_
	    &" WHERE f.fakdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY fms.mid"
	    
	    'Response.Write strSQLfak & "<br>"
	    'Response.Flush
	    
	    oRec2.open strSQLfak, oConn, 3
	    while not oRec2.EOF
	    fakTimer(x) = oRec2("faktimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    end if
	    
	   
	    
	    '*** Materiale forbrug ***'
	    if (instr(akttype_sel, "#-4#") <> 0 AND visning = 1) then 
	    
	    strSQLmf = "SELECT sum(mf.matantal * (mf.matsalgspris * (kurs/100))) AS mfforbrug FROM materiale_forbrug mf "_
	    &" WHERE mf.forbrugsdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND usrid = "& intMid &" GROUP BY mf.usrid"
	    
	    'Response.Write strSQLmf & "<br>"
	    'Response.Flush
	    
	    oRec2.open strSQLmf, oConn, 3
	    while not oRec2.EOF
	    mfForbrug(x) = oRec2("mfforbrug")/1
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    end if
	    
	     if len(trim(resTimer(x))) <> 0 AND resTimer(x) <> 0 then
	     resTimer(x) = resTimer(x)
	     afstemnul(x) = 103
	     else
	     resTimer(x) = 0
	     end if
    	 
	     if len(trim(fakTimer(x))) <> 0 AND fakTimer(x) <> 0 then
	     fakTimer(x) = fakTimer(x)
	     afstemnul(x) = 102
	     else
	     fakTimer(x) = 0
	     end if
	     
	     if len(trim(mfForbrug(x))) <> 0 AND mfForbrug(x) <> 0 then
	     mfForbrug(x) = mfForbrug(x)
	     afstemnul(x) = 101
	     else
	     mfForbrug(x) = 0
	     end if
	    
	    'end if
	   
	    
	
	
	
	
	
	
	
	
	end if 'visning
	
	
	
	
	
	'***************************************************************************************************************************
	'**** Præsentation ***'
	'***************************************************************************************************************************
	
	
	select case visning 

	case 1
	'****** Medarbejder afstemning og Løn HR listen *****'


     if cint(headerwrt) <> 1 then
	
    if endp = 0 then 'vis ikke på månedsopdeling
	strEksportPer = "Periode afgrænsning: "& formatdatetime(stdato, 1) & " - "&  formatdatetime(sldato, 1) & "xx99123sy#z"
	strEksportTxt = strEksportTxt & strEksportPer
	'strEksportEmailTxt = strEksportEmailTxt & strEksportPer
	end if	



	if instr(akttype_sel, "#-1#") <> 0 then
	cops = 6
	else
	cops = 0
	end if
	
	
	%>
    <table cellspacing=0 cellpadding=2 border=0><!-- #5C75AA -->
	



	 <tr bgcolor="#5582d2">
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_147%></b></td>
	 
	  <%if (instr(akttype_sel, "#-5#") <> 0 OR instr(akttype_sel, "#-10#") <> 0) AND stempelurOn = 1 then
      
        if instr(akttype_sel, "#-10#") <> 0 then
          
      
          cpsLt = 1  
        

          %>
	       <td class=alt valign="bottom" colspan=<%=cpsLt %> style="border-right:1px  #D6DfF5 solid;">&nbsp;</td>
           <%

        else

           if cint(showkgtil) = 1 then
           cpsLt = 7
           else
           cpsLt = 5
           end if


           %>
	       <td class=alt valign="bottom" colspan=<%=cpsLt %> style="border-right:1px  #D6DfF5 solid;"><b>Komme / Gå tid</b> (løntimer)<br />Stempelur<br />I valgt periode</td>
           <%
          end if

      
	 end if
	 
	 
	 if instr(akttype_sel, "#-1#") <> 0 OR instr(akttype_sel, "#-20#") <> 0 then
               
                     if instr(akttype_sel, "#-20#") = 0 then %>
	                 <td class=alt valign="bottom" colspan=<%=cops %> style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_148%></b> <br />
	                 <%=tsa_txt_149%> <br />I valgt periode</td>
                         <%else %>
                         <td style="border-right:1px  #D6DfF5 solid;">&nbsp;</td>

	                 <%end if %>
	 
	  <%end if %>
	 
	 
	 <%
	 
	 aktudspCP = 0
	 if instr(akttype_sel, "#1#") <> 0 then 
	 aktudspCP = aktudspCP + 1
	 end if %>
	 
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then
	 aktudspCP = aktudspCP + 1
	 end if %>
	
	 <%if instr(akttype_sel, "#6#") <> 0 then 
	 aktudspCP = aktudspCP + 1
	 end if %>
	 
	   <%if instr(akttype_sel, "#50#") <> 0 then 
	   aktudspCP = aktudspCP + 1
	   end if %>
	 
	  <%if instr(akttype_sel, "#51#") <> 0 then 
	  aktudspCP = aktudspCP + 1
	  end if %>
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then 
	 aktudspCP = aktudspCP + 1
	 end if %>
	 
	  <%if instr(akttype_sel, "#53#") <> 0 then 
	  aktudspCP = aktudspCP + 1
	  end if %>
	 
	  <%if instr(akttype_sel, "#54#") <> 0 then 
	  aktudspCP = aktudspCP + 1
	  end if %>
	 
	  <%if instr(akttype_sel, "#55#") <> 0 then
	  aktudspCP = aktudspCP + 1
	  end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then 
	  aktudspCP = aktudspCP + 1
	  end if %>
	 
	  <%if instr(akttype_sel, "#90#") <> 0 then 
	  aktudspCP = aktudspCP + 1
	  end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then 
	   aktudspCP = aktudspCP + 1 
	   end if %>

       <%if instr(akttype_sel, "#92#") <> 0 then 
	   aktudspCP = aktudspCP + 1 
	   end if %>
	 
	 <%if aktudspCP <> 0 then %>
	 <td class=alt valign="bottom" colspan=<%=aktudspCP %> style="border-right:1px  #D6DfF5 solid;"><b>Udspecificenring på akt. typer</b><br />
	 I valgt periode</td>
	 <%end if %>
	 
	 
	 
	 
	   <%if instr(akttype_sel, "#7#") <> 0 then %>
	<td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b>Flex</b></td>
	 <%end if %>
	
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_265%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#10#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_150%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b>Pause</b></td>
	 <%end if %>
	 
	 
	 <%
	 copsff = 1
     copsffPer = 1
	 if instr(akttype_sel, "#12#") <> 0 then 
	 copsff = copsff + 1
	 copsffPer = copsffPer
     end if
	 
	 if instr(akttype_sel, "#18#") <> 0 then 
	 copsff = copsff + 1
     copsffPer = copsffPer
	 end if
	 
	 if instr(akttype_sel, "#13#") <> 0 then 
	 copsff = copsff + 1
     copsffPer = copsffPer + 2
	 end if
	 
	 if instr(akttype_sel, "#17#") <> 0 then 
	 copsff = copsff + 1
     copsffPer = copsffPer + 2
	 end if
	 
	 if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then %>
	<td class=alt valign=bottom colspan=<%=copsff %> style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_174%></b><br />
	(Feriefriår: <%=ferieFriaarStart %>)<br />
    <span style="color:#FFFFFF; font-size:10px;"><%=formatdatetime(startdatoFerieFriLabel, 2)%> til <%=formatdatetime(slutdatoFerieFriLabel, 2) %></span></td>

    <% if (instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0) AND cint(visikkeFerieogSygiPer) <> 1 then %>
    <td class=alt valign=bottom colspan=<%=copsffPer %> style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_174%></b><br />(Periode)<br />
	<span style="color:#FFFFFF; font-size:10px;"><%=formatdatetime(startdato, 2)%> til <%=formatdatetime(slutdato, 2) %></span></td>
	<%end if %>

	 <%end if %>
	 
	 
	  <%
	 copsFe = 1
     copsFePer = 1
	 if instr(akttype_sel, "#11#") <> 0 then 
	 copsFe = copsFe + 1
     copsFePer = copsFePer 
	 end if
	 
	 if instr(akttype_sel, "#14#") <> 0 then 
	 copsFe = copsFe + 1
     copsFePer = copsFePer + 2
	 end if
	 
	 if instr(akttype_sel, "#19#") <> 0 then 
	 copsFe = copsFe + 1
     copsFePer = copsFePer + 2
	 end if
	 
	 if instr(akttype_sel, "#15#") <> 0 then 
	 copsFe = copsFe + 1
     copsFePer = copsFePer
	 end if

      if instr(akttype_sel, "#111#") <> 0 then 
	 copsFe = copsFe + 1
     copsFePer = copsFePer
	 end if


      if instr(akttype_sel, "#112#") <> 0 then 
	 copsFe = copsFe + 1
     copsFePer = copsFePer
	 end if
	 
	 if instr(akttype_sel, "#16#") <> 0 then 
	 copsFe = copsFe + 1
     copsFePer = copsFePer + 2
	 end if
	 
	 if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 OR instr(akttype_sel, "#111#") <> 0 OR instr(akttype_sel, "#112#") <> 0 then %>
	<td valign="bottom" class=alt colspan=<%=copsFe %> style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_152%></b><br />
	(Ferieår: <%=ferieaarStart %>)<br />
    <span style="color:#FFFFFF; font-size:10px;"><%=formatdatetime(startdatoFerieLabel, 2)%> til <%=formatdatetime(slutdatoFerieLabel, 2) %></span> 
    </td>

    <% if (instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#16#") <> 0 OR instr(akttype_sel, "#19#") <> 0) AND cint(visikkeFerieogSygiPer) <> 1 then %>

    <td valign="bottom" class=alt colspan=<%=copsFePer %> style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_152%></b><br /> (Periode)<br />
	<span style="color:#FFFFFF; font-size:10px;"><%=formatdatetime(startdato, 2)%> til <%=formatdatetime(slutdato, 2) %></span></td>
	<%end if %>
    <%end if %>

	  
	 <%if instr(akttype_sel, "#25#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b>Div. fri / 1 maj</b> timer</td>
	 <%end if %>


    	 <%if instr(akttype_sel, "#125#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b>Rejsedage</b></td>
	 <%end if %>


	 
	 <%
	 copsAf = 1
	 if instr(akttype_sel, "#30#") <> 0 then 
         if cint(visikkeFerieogSygiPer) <> 1 then 
	     copsAf = copsAf + 2
         else
         copsAf = copsAf + 1
         end if
	 end if
	 
	 if instr(akttype_sel, "#31#") <> 0 then 
           if cint(visikkeFerieogSygiPer) <> 1 then 
	        copsAf = copsAf + 3
           else
         copsAf = copsAf + 1
            end if
	 end if
	 
	 if instr(akttype_sel, "#32#") <> 0 then 
	 copsAf = copsAf + 1
	 end if
	 
	 if instr(akttype_sel, "#33#") <> 0 then 
	 copsAf = copsAf + 1
	 end if
	 
	 if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR _
	 instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	<td class=alt valign="bottom" colspan=<%=copsAf %> style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_151 %> / <%=global_txt_145 %></b>
    <span style="color:#FFFFFF; font-size:10px;">
	<br /> Total fra licens-start d. <%=formatdatetime(lisStDato,2) %> (eller fra ansættelsesdato) - <%=formatdatetime(slutdato, 2) %></span>
	</td>
	<%end if %>
	 
	 
	 <%
	 '** Sygdom ***'
	 'if instr(akttype_sel, "#20#") <> 0 then 
	 'copsS = 5
	 'else
	 'copsS = 0
	 'end if
	 
	
	 
	  %>
	 
	  <%if instr(akttype_sel, "#20#") <> 0 then %>
	 <td valign=bottom class=alt colspan=2 style="white-space:nowrap; border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_153%></b> (ÅTD)<br />
     <span style="color:#FFFFFF; font-size:10px; white-space:nowrap;">
	 <%=formatdatetime(sygaarSt, 2) %> til<br /> <%=formatdatetime(slutdato, 2) %></span>
	 </td>

    <%if cint(visikkeFerieogSygiPer) <> 1 then %>
     <td valign=bottom class=alt colspan=3 style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_153%></b> (Periode)<br />
     <span style="color:#FFFFFF; font-size:10px;">
    <%=formatdatetime(startdato, 2) %> til <%=formatdatetime(slutdato, 2) %></span>
	 </td>
        <%end if %>

	 <%end if 
     
     

     '**** Barn Syg ***'
     'if instr(akttype_sel, "#21#") <> 0 then 
	 'copsBS = 5
	 'else
	 'copsBS = 0
	 'end if
     
     
     
     %>

     <%if instr(akttype_sel, "#21#") <> 0 then %>
	 <td valign=bottom class=alt colspan=2 style="white-space:nowrap; border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_170%></b> (ÅTD)<br />
     <span style="color:#FFFFFF; font-size:10px; white-space:nowrap;">
	 <%=formatdatetime(sygaarSt, 2) %> til<br /> <%=formatdatetime(slutdato, 2) %></span>
	 </td>

        <%if cint(visikkeFerieogSygiPer) <> 1 then %>
     <td valign=bottom class=alt colspan=3 style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_170%></b> (Periode)<br />
      <span style="color:#FFFFFF; font-size:10px;">
    <%=formatdatetime(startdato, 2) %> til <%=formatdatetime(slutdato, 2) %></span>
	 </td>
	 <%end if
     end if %>
	 

     <%if (instr(akttype_sel, "#20#") <> 0 OR instr(akttype_sel, "#21#") <> 0) AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;">&nbsp;</td>
     
	 <%end if %>

	 

     <%if instr(akttype_sel, "#22#") <> 0 then %>
	 <td class=alt valign="bottom" colspan=2 style="border-right:1px  #D6DfF5 solid;"><b>Barsel</b> dage</td>
     
	 <%end if %>

     <%if instr(akttype_sel, "#23#") <> 0 then %>
	 <td class=alt valign="bottom" colspan=2 style="border-right:1px  #D6DfF5 solid;"><b>Omsorg</b> dage</td>
	 <%end if %>

     <%if instr(akttype_sel, "#24#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b>Senior</b> timer</td>
	 <%end if %>


      <!-- Alders reduktion -->
      <%if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;" colspan="4"><b><%=left(global_txt_185, 15) %></b> timer</td>
     
	 <%end if %>
    

     <!-- Alders reduktion, Omsprg 2, 10 K planlagt  -->
         <%if instr(akttype_sel, "#26#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=left(global_txt_188, 25) %></b> timer</td>
	 <%end if %>

         <%if instr(akttype_sel, "#120#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=left(global_txt_189, 25) %></b> timer</td>
	 <%end if %>

         <%if instr(akttype_sel, "#121#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=left(global_txt_190, 25) %></b> timer</td>
	 <%end if %>

     <%if instr(akttype_sel, "#122#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=left(global_txt_191, 35) %></b> timer</td>
	 <%end if %>

    <%if instr(akttype_sel, "#123#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=left(global_txt_192, 35) %></b> enh</td>
	 <%end if %>

    <%if instr(akttype_sel, "#124#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=left(global_txt_193, 35) %></b> enh</td>
	 <%end if %>
     

	  <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b>Sundh.</b> timer</td>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b>Læge</b> timer</td>
	 <%end if %>

       <%if instr(akttype_sel, "#115#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=global_txt_179%></b></td>
	 <%end if %>

	 
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_154%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_155%></b> timer</td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b>Stk.</b></td>
	 <%end if %>

      <%if instr(akttype_sel, "#113#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=global_txt_177%></b></td>
	 <%end if %>

     
      <%if instr(akttype_sel, "#114#") <> 0 then %>
	 <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=global_txt_178%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	  <td class=alt valign="bottom" style="border-right:1px  #D6DfF5 solid;"><b><%=tsa_txt_156%></b></td>
	  <%end if %>
	 </tr>


	 
	 
	 
	 
	 
	 
	 
	 <!-- Under overskrrifter --> <!--- Akt type navn -->
	 
	 <tr bgcolor="#8CAAe6">
	 
	 <td style="width:200px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;" class=alt>&nbsp;</td>
	 
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_147 & "; Nr.;Init;" %>

         <%   
         if (instr(akttype_sel, "#-5#") <> 0 AND instr(akttype_sel, "#-10#") = 0) then%>
             <td class=lille  valign=bottom bgcolor="#DCF5BD" style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=tsa_txt_158%></td>
	        <td class=lille  valign=bottom bgcolor="#DCF5BD" style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=tsa_txt_259%></td> 
         <%
             strEksportTxtHeader = strEksportTxtHeader & tsa_txt_158 &";"& tsa_txt_259 &";"
             end if 
         %>
	 
	  <%if (instr(akttype_sel, "#-5#") <> 0 OR instr(akttype_sel, "#-10#") <> 0) AND stempelurOn = 1 then %>

        <%if instr(akttype_sel, "#-10#") = 0 then %>

   
	 <td class=alt_lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">Komme/gå<br /><%=tsa_txt_148 %>:<%=tsa_txt_141 %></td>

     <%if cint(showkgtil) = 1 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">Till./Frad. +/-</td>
	 <td class=alt_lille  valign=bottom style="width:50px;border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">= Sum</td>
     <%end if %>

	 <td class=alt_lille  valign=bottom style="width:70px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">Balance <br />Kom./gå / Real.</td>

         <%end if %>


	 <td class=lille valign=bottom bgcolor="#DCF5BD" style="width:75px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
     <%select case lto
       case "kejd_pb", "intranet - local"
         %>
           Flekssaldo +/-
    
         <%
        case else
         %>
           Flekssaldo +/-<br />
           Kom./gå / Norm.
         <%end select %>
    </td>


	 <%if instr(akttype_sel, "#-10#") = 0 then %>
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_148 &":"& tsa_txt_141 &" (Omregnet til kommatal for kalkulation i excel);"
     
     if cint(showkgtil) = 1 then
     strEksportTxtHeader = strEksportTxtHeader &"Tillæg/Fradrag +/- ;Sum;"
     end if
     
     strEksportTxtHeader = strEksportTxtHeader &"Balance Lønt./Real;"
    
     end if
     
      
     strEksportTxtHeader = strEksportTxtHeader &"Flekssaldo +/-Akkumuleret Komme/gå tid / Normtid;" %>
	 
	 <%end if %>





	 
	 <%if (instr(akttype_sel, "#-1#") <> 0 OR instr(akttype_sel, "#-20#")) <> 0 then
         
         if instr(akttype_sel, "#-20#") = 0 then %>
	 <td class=lille  valign=bottom bgcolor="#cccccc" style="width:50px; border-right:1px #ffffff solid;"><%=tsa_txt_157%></td>
	 <td class=lille  valign=bottom style="width:50px; border-right: 1px #ffffff solid;" bgcolor="#cccccc"><%=tsa_txt_172 &" "& tsa_txt_179%></td>

      <td class=lille  valign=bottom style="width:50px; border-right: 1px #ffffff solid;" bgcolor="#cccccc"><%=tsa_txt_160%></td>
	 <td class=lille  valign=bottom style="width:50px; border-right: 1px #ffffff solid;" bgcolor="#cccccc"><%=tsa_txt_161%></td>
	 <td class=lille  valign=bottom style="width:50px; border-right: 1px #ffffff solid;" bgcolor="#cccccc"><%=tsa_txt_163%></td>

	
         <%end if %>

	 <td class=lille  valign=bottom bgcolor="pink" style="width:70px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
      
           Flekssaldo +/-<br />
           Real. / Norm.
     </td>
	
	 
	 <%
         if instr(akttype_sel, "#-20#") = 0 then 
         strEksportTxtHeader = strEksportTxtHeader & tsa_txt_157 &";"& tsa_txt_172 &" "& tsa_txt_179 &";" & tsa_txt_160 & ";" & tsa_txt_161 & ";" & tsa_txt_163 & ";"
         end if
         strEksportTxtHeader = strEksportTxtHeader & tsa_txt_159 & ";"  %>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#1#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(1, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(2, 1) %>
	 <%=akttypenavn%>
	 </td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	<%end if %>
	
	<%if instr(akttype_sel, "#50#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(50, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>


     <%if instr(akttype_sel, "#54#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(54, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#51#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(51, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(52, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#53#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(53, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  
	 
	  <%if instr(akttype_sel, "#55#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(55, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(60, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#90#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(90, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(91, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>

         
	   <%if instr(akttype_sel, "#92#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(92, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#6#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(6, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#7#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(7, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;" style="">
	 <%call akttyper(5, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#10#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(10, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(9, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 
	 <!--Ferie Fridage -->
	 <%if instr(akttype_sel, "#12#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(12, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#18#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(18, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %> <br> >> dd.</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(13, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#17#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(17, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 
	 
	 
	 <%if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then %>
	 <td bgcolor="#FFF000" class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;" style=""><%=tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_275%></td>
	 <!--<td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;" style=""><%=tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_276%></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#13#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(13, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %><br /> i periode</td>
     <td bgcolor="#EFf3ff" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%call akttyper(13, 1) %>
	 <%=akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") %><br /> i periode</td>
      <td bgcolor="#EFf3ff" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& replace(tsa_txt_275, "dage", "") %><br />Første dag i per.</td>
	<!-- <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &" i periode;"& akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") &" i periode;Første dag i per.;"%>
	 <%end if %>


       <%if instr(akttype_sel, "#17#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(17, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %><br /> i periode</td>
     <td bgcolor="#EFf3ff" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%call akttyper(17, 1) %>
	 <%=akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") %><br /> i periode</td>
      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &" i periode;"& akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") &" i periode;"%>
	 <%end if %>
	 
	
	 
	 <!-- Ferie -->
	 <%if instr(akttype_sel, "#15#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(15, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>


      <%if instr(akttype_sel, "#111#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(111, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>

        <%if instr(akttype_sel, "#112#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(112, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#11#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(11, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %> <br> >> dd.</td>
	 <!--<td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(14, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
    
	 <!--<td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#19#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(19, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#16#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(16, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <% if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 OR instr(akttype_sel, "#111#") <> 0 OR instr(akttype_sel, "#112#") <> 0 then %>
	  <td bgcolor="#6CA000" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;" style=""><%=tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_275%></td>
	 <!--<td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;" style=""><%=tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_276%></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader &" "& tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <!-- Ferie i periode -->

	  <%if instr(akttype_sel, "#14#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3ff"  class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(14, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %><br /> i periode</td>
      <td bgcolor="#EFf3ff"  class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(14, 1) %>
	 <%=akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") %><br /> i periode</td>
      <td bgcolor="#EFf3ff" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"> <%=akttypenavn &" "& replace(tsa_txt_275, "dage", "") %><br />Første dag i per.</td>
	 <!--<td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &" i periode;"& akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") &" i periode;Første dag i per.;"%>
	 <%end if %>


     <%if instr(akttype_sel, "#19#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3ff"  class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(19, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %><br /> i periode</td>
      <td bgcolor="#EFf3ff"  class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(19, 1) %>
	 <%=akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") %><br /> i periode</td>
 
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &" i periode;"& akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") &" i periode;"%>
	 <%end if %>


     <%if instr(akttype_sel, "#16#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3ff"  class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(16, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %><br /> i periode</td>
      <td bgcolor="#EFf3ff"  class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(16, 1) %>
	 <%=akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") %><br /> i periode</td>
 
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &" i periode;"& akttypenavn &" "& replace(tsa_txt_275, "dage", "timer") &" i periode;"%>
	 <%end if %>

     

    

     <!-- Divfritimer 1 maj timer -->
	  <%if instr(akttype_sel, "#25#") <> 0 then %>
	     <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(25, 1) %>
	     <%=akttypenavn%> 
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer;"%>
	     </td>
	  
      <%end if %>
	 

         <!-- Rejsedage -->
	  <%if instr(akttype_sel, "#125#") <> 0 then %>
	     <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(125, 1) %>
	     <%=akttypenavn%> 
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer;"%>
	     </td>
	  
      <%end if %>
	 
	 
	 <!-- Afspad --->
	 <%if instr(akttype_sel, "#30#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(30, 1) %>
	 <%=akttypenavn%><br />(Tim.) Enh.</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" Tim. ; " & akttypenavn &" Enh. ;"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(31, 1) %>
	 <%=akttypenavn%>.</td>
     <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &".;"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#32#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(32, 1) %>
	 <%=akttypenavn%>.</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &".;"%>
	 <%end if %>
	 
	<%if instr(akttype_sel, "#33#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(33, 1) %>
	 <%=akttypenavn%></td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	 <td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;" style=""><%=tsa_txt_283 &" "& tsa_txt_280%></td>
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_283 &" "& tsa_txt_280 &";"%>
	 <%end if %>
	 
	  <%
        if cint(visikkeFerieogSygiPer) <> 1 then 
    
          
      if instr(akttype_sel, "#30#") <> 0 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(30, 1) %>
	 <%=akttypenavn%><br />(Tim.) Enh. i periode</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" Timer i periode ; " & akttypenavn &" Enheder i periode ;"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(31, 1) %>
	 <%=akttypenavn%><br /> timer i periode</td>
     <td bgcolor="#EFf3FF" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">Første dag i per.</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" Timer i periode ;Første dag i per.;"%>
	 <%
         end if
         end if %>

      
	 
	 
	
	 <!-- Sygdom -->
	  <%if instr(akttype_sel, "#20#") <> 0 then %>
	     <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(20, 1) %>
	     <%=akttypenavn%> tim. ÅTD
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer ÅTD;"&  akttypenavn &" dage ÅTD;"%>
	     </td>
	      <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn%> dage ÅTD</td>

         <%if cint(visikkeFerieogSygiPer) <> 1 then %>
    	 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn%> tim. i per.  </td>
	     <td bgcolor="#EFf3FF" class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn%> dage i per</td>
           <td bgcolor="#EFf3FF" class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">Første dag i per.</td>
            <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer i per.;"&  akttypenavn &" dage i per.;Første dag i per.;"%>
          <%end if 
         end if %>
	 
	  <%if instr(akttype_sel, "#21#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(21, 1) %>
	 <%=akttypenavn%> tim. ÅTD
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer ÅTD;"&  akttypenavn &" dage ÅTD;"%>
	 </td>
	  <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn%> dage ÅTD</td>

         <%if cint(visikkeFerieogSygiPer) <> 1 then%>
    	 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn%> tim. i per.</td>
	      <td bgcolor="#EFf3FF" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=akttypenavn%> dage i per</td>
          <td bgcolor="#EFf3FF" class=lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">Første dag i per.</td>
           <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer i per.;"&  akttypenavn &" dage i per.;Første dag i per.;"%>
    	  
	 
	 <%end if
         end if %>

	 
	 <%if (instr(akttype_sel, "#20#") <> 0 OR instr(akttype_sel, "#21#") <> 0) AND cint(visikkeFerieogSygiPer) <> 1 then %>
	    <td bgcolor="lightpink" class=lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
        Sygefravær i periode i %<br />
           Incl. barsel & læge vis vist 
   
	     <%strEksportTxtHeader = strEksportTxtHeader &" Fravær i per. i %;"%>
	     </td>
         
	 <%end if %>


     <!-- Barsel -->
	  <%if instr(akttype_sel, "#22#") <> 0 then %>
	     <td class=alt_lille valign="bottom" style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(22, 1) %>
	     <%=akttypenavn%> 
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" dage;Første dag i per;"%>
	     </td>
         <td class=alt_lille valign="bottom" style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">Første dag i per.</td>
	  
      <%end if %>

         <!-- Omsorg -->
	  <%if instr(akttype_sel, "#23#") <> 0 then %>
	     <td class=alt_lille valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(23, 1) %>
	     <%=akttypenavn%> 
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" dage;Første dag i per;"%>
	     </td>
         <td class=alt_lille valign="bottom" style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">Første dag i per.</td>
	  
      <%end if %>

         <!-- Senior -->
	  <%if instr(akttype_sel, "#24#") <> 0 then %>
	     <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(24, 1) %>
	     <%=akttypenavn%> 
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer;"%>
	     </td>
	  
      <%end if %>

          <!-- Aldersreduktion -->
            <%if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then %>
	  <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(27, 1) %>
	     <%=left(akttypenavn, 10)%>.<br />
          <%=right(akttypenavn, 7)%>
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	     </td>
            <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(28, 1) %>
	    <%=left(akttypenavn, 10)%>.<br />
          <%=right(akttypenavn, 5)%>
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	     </td>
           <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(29, 1) %>
	    <%=left(akttypenavn, 10)%>.<br />
          <%=right(akttypenavn, 8)%>
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	     </td>


            <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	     <%call akttyper(29, 1) %>
	    <%=left(akttypenavn, 10)%>.<br />
          Saldo
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" saldo;"%>
	     </td>

	 <%end if %>

     

          <%if instr(akttype_sel, "#26#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(26, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>

          <%if instr(akttype_sel, "#120#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(120, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>

          <%if instr(akttype_sel, "#121#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(121, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>

          <%if instr(akttype_sel, "#122#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(122, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>


          <%if instr(akttype_sel, "#123#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(123, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>


          <%if instr(akttype_sel, "#124#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(124, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>


	 
	 <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(8, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(81, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>

         <%if instr(akttype_sel, "#115#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(115, 1) %>
	 <%=akttypenavn%> timer
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer;"%>
	 </td>
	 <%end if %>
	 
	
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=tsa_txt_154%></td>
	  <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_154 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=tsa_txt_155%>
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_155 &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(61, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>

       <%if instr(akttype_sel, "#113#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(113, 1) %>
	 <%=akttypenavn%><br />timer
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 

       <%if instr(akttype_sel, "#114#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;">
	 <%call akttyper(114, 1) %>
	 <%=akttypenavn%><br />timer
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>

	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	  <td class=alt_lille  valign=bottom style="width:50px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><%=tsa_txt_171%></td>
	  <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_171 &";"%>
	  <%end if %>
	 </tr>
	 
	 <%
	 
     strEksportTxtHeader = strEksportTxtHeader & ";Startdato;Slutdato;"
       
     if headerwrtExp = 0 then
	 strEksportTxt = strEksportTxtHeader & strEksportTxtMd
	 end if

     headerwrt = 1
     headerwrtExp = 1
	 end if 'x = 0 overskrifter
	 
	 
	 
	 
	 strEksportTxt = strEksportTxt & strEksport(x)
	 strEksportEmailTxtHeader = strEksportTxtHeader 'strEksport(x) ''
	 
	 
	 
	 'for x = 0 to x - 1 
	 
	 for x = m to x '+ 1
	 
	 
	 
	 '** Nulfilter **'
	 if afstemnul(x) <> 0 OR visnulfilter = 0 then
	 
	 select case right(x, 1)
	 case 0,2,4,6,8
	 bgcl = "#ffffff"
	 'bgcolorTr = "#cccccc"
	 case else
	 bgcl = "#EFF3ff" '"#D6DFf5"
	 'bgcolorTr = "#ffffff"
	 end select%>
	 
	 <%strEksport(x) = strEksport(x) &"xx99123sy#z"%>
	 
	 <tr bgcolor="<%=bgcl %>">
	 <td class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><b><%=medarbNavn(x) %></b> (<%=medarbNr(x) %>) 
	 <%if len(trim(medarbInit(x))) <> 0 then %>
	  - <%=medarbInit(x) %>
	  <%end if %>

	  </td>
	 
	 <%strEksport(x) = strEksport(x) & medarbNavn(x) &"; "& medarbNr(x) &";"& medarbInit(x) &";"
         
          if formatnumber(normTimer(x),2) <> 0 then
      normTimerTxt = formatnumber(normTimer(x),2) 
      normTimerTxtExp = normTimerTxt
      else
      normTimerTxt = ""
      normTimerTxtExp = 0
      end if

      if showNormTimerdag > 0 then
	 showNormTimerdagTXT = formatnumber(showNormTimerdag, 1) 
     else
     showNormTimerdagTXT = ""
     end if


         
          if (instr(akttype_sel, "#-5#") <> 0 AND instr(akttype_sel, "#-10#") = 0) then
         %>
         
          <td align=right style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;" class=lille><%=normTimerTxt%></td>
	     <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	     <%=showNormTimerdagTXT %> 
             
           
	     </td>
         
         <%
             strEksport(x) = strEksport(x) & normTimerTxtExp &";"& showNormTimerdagTxt & ";"
          end if
         
          if (instr(akttype_sel, "#-5#") <> 0 OR instr(akttype_sel, "#-10#") <> 0) AND stempelurOn = 1 then 

          ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	      call fLonTimerPer(ltimerStDato, dageDiff, 3, medarbId(x)) 



          'if cint(showkgtil) = 1 then 'Vis fradrag / tillæg
          lontimerTotMfradrag = ((totalTimerPer100/60) + (fradragTimer(x)))
          'else
	      'lontimerTotMfradrag = ((totalTimerPer100/60))
          'end if

          
          bal_lon_real = formatnumber(lontimerTotMfradrag,2) - formatnumber(realTimer(x),2)
            
          normtime_lontime = 0
            normtime_lontime = -((normTimer(x) - (lontimerTotMfradrag)) * 60)
	  
	        if normtime_lontime <> 0 then
	        normtime_lontime = cdbl(normtime_lontime)
	        else
	        normtime_lontime = normtime_lontime
	        end if


            fraDtimer = formatnumber(fradragTimer(x),2)
          
      if instr(akttype_sel, "#-10#") = 0 then
          %>
	 <td align=right style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;" class=lille>
	 <%
	 call timerogminutberegning(totalTimerPer100)
         %>
		<%=thoursTot &":"& left(tminTot, 2) %> 

	 <%
         strEksport(x) = strEksport(x) & thoursTot &","& left((tminTot*1.67), 2) &";"
	 %> 
	 
	</td>
	
	  <%
      
                         if cint(showkgtil) = 1 then 'Vis fradrag / tillæg
	                   
                                      %>
	                                  <td align=right style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;" class=lille>

                                          <%if fraDtimer <> 0 then %>
	                                          <%call timerogminutberegning(fraDtimer*60) %>
	                                          <%=thoursTot &":"& left(tminTot, 2) %>  
          
	  
	                                          <%strEksport(x) = strEksport(x) & formatnumber(fradragTimer(x), 2) &";"%>
	                                      <%else %>
       
	  
	                                      <%strEksport(x) = strEksport(x) & ";"%>
                                          <%end if %>
	                                 </td>
	
	
	                                      <td align=right style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;" class=lille>

                                          <%if lontimerTotMfradrag <> 0 then %>
	  
                                          <%call timerogminutberegning(lontimerTotMfradrag*60) %>
	                                      <%=thoursTot &":"& left(tminTot, 2) %> 
	  
	                                      <%strEksport(x) = strEksport(x) & formatnumber(lontimerTotMfradrag, 2) &";"%>
	  
                                          <%else %>
	                                      <%strEksport(x) = strEksport(x) & ";"%>
      
	                                      <%
                                          end if
                                          %>
                                          </td>

                          <%end if


	
	 %>
	  <td align=right style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;" class=lille>
	  <%call timerogminutberegning(bal_lon_real*60) %>
	  <%=thoursTot &":"& left(tminTot, 2) %></td>
	  
	  <%strEksport(x) = strEksport(x) & formatnumber(bal_lon_real, 2) &";"
          
          
      end if '-10
          
          normtime_lontimeBGCOl = ""
          if cdbl(normtime_lontime) < -10 then
          normtime_lontimeBGCOl = "red"
          end if

          if cdbl(normtime_lontime) > 40 then
          normtime_lontimeBGCOl = "green"
          end if
          %>
	  
	  
	                 <td align=right style=" white-space: nowrap; color:<%=normtime_lontimeBGCOl %>; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;" class=lille><b>
	                  <%
	                 
	                  call timerogminutberegning(normtime_lontime) %>
	                  <%=thoursTot &":"& left(tminTot, 2) %></b></td>
	  
	                  <%strEksport(x) = strEksport(x) & formatnumber(normtime_lontime/60, 2) &";"%>
	  
	 <%end if '-5, -10 +stempelur%>








	 
	  <%if (instr(akttype_sel, "#-1#") <> 0 OR instr(akttype_sel, "#-20#") <> 0)  then 
      
      if formatnumber(realTimer(x),2) <> 0 then
      realTimerTxt = formatnumber(realTimer(x),2)
      realTimerTxtExp = realTimerTxt
      else
      realTimerTxt = ""
      realTimerTxtExp = 0
      end if

      if formatnumber(realTimer(x)/weekDiff,2) <> 0 then
      realTimerprUgeTxt = formatnumber(realTimer(x)/weekDiff,2)
      realTimerprUgeTxtExp = realTimerprUgeTxt
      else
      realTimerprUgeTxt = ""
      realTimerprUgeTxtExp = 0
      end if

     
    
     balRealNormTxt = 0
     balRealNormTxt = formatnumber((realTimer(x) - normTimer(x)),2) 
     balRealNormTxtExp = balRealNormTxt


     if formatnumber(realfTimer(x),2)  <> 0 then
     realfTimerTxt = formatnumber(realfTimer(x),2)
     realfTimerTxtExp = realfTimerTxt
     else
     realfTimerTxt = ""
     realfTimerTxtExp = 0
     end if 
     
     
     if formatnumber(realIfTimer(x),2) <> 0 then
     realIfTimerTxt = formatnumber(realIfTimer(x),2)
     realIfTimerTxtExp = realIfTimerTxt
     else
     realIfTimerTxt = ""
     realIfTimerTxtExp = 0
     end if

     if realTimer(x) <> 0 AND realfTimer(x) <> 0 then
          iebal = (realfTimer(x)/realTimer(x)) * 100 
          ieBalTxt = formatnumber(iebal,0) & " %"
      else
	    iebal = 0
          ieBalTxt = ""
	  end if
      

     if instr(akttype_sel, "#-20#") = 0 then
     %>
	 <td align=right style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;" class=lille><b><%=realTimerTxt%></b></td>
	 <td align=right style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;" class=lille><%=realTimerprUgeTxt%></td>

     <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=realfTimerTxt%></td>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=realIfTimerTxt%></td>
     <td align=right class=lille style="  white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=ieBalTxt%></td>

	
	 
	 <%strEksport(x) = strEksport(x) & realTimerTxtExp &";" & realTimerprUgeTxtExp & ";"& realfTimerTxtExp &";"& realIfTimerTxtExp & ";"& formatnumber(iebal,0) & ";"%>
	
	
    <%end if %>

	 
         <% balRealNormBGCOl = ""
          if cdbl(balRealNormTxt) < -10 then
          balRealNormBGCOl = "red"
          end if

          if cdbl(balRealNormTxt) > 40 then
          balRealNormBGCOl = "green"
          end if %>

         <td align=right class=lille style="white-space:nowrap; color:<%=balRealNormBGCOl%>; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><b><%=balRealNormTxt%></b></td>
	
	 
	  
	  <%strEksport(x) = strEksport(x) & balRealNormTxtExp &";" %>
	  
	
	   <%end if %>





	 
	 <%if instr(akttype_sel, "#1#") <> 0 then 
     
        if formatnumber(fTimer(x),2) <> 0 then
        fTimerTxt = formatnumber(fTimer(x),2)
        fTimerTxtExp = fTimerTxt
        else
        fTimerTxt = ""
        fTimerTxtExp = 0
        end if%>

	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=fTimerTxt%></td>
	 <% strEksport(x) = strEksport(x) & fTimerTxtExp & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then 
     
     if formatnumber(ifTimer(x),2) <> 0 then
     ifTimerTxt = formatnumber(ifTimer(x),2)
     ifTimerTxtExp = ifTimerTxt
     else
     ifTimerTxt = ""
     ifTimerTxtExp = 0
     end if 

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=ifTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & ifTimerTxtExp & ";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#50#") <> 0 then
      
     if formatnumber(dagTimer(x),2) <> 0 then
     dagTimerTxt = formatnumber(dagTimer(x),2)
     dagTimerTxtExp = dagTimerTxt
     else
     dagTimerTxt = ""
     dagTimerTxtExp = 0
     end if 
      
       %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=dagTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & dagTimerTxtExp & ";"%>
	 <%end if %>

      <%if instr(akttype_sel, "#54#") <> 0 then 
      
      if formatnumber(aftenTimer(x),2) <> 0 then
     aftenTimerTxt = formatnumber(aftenTimer(x),2)
     aftenTimerTxtExp = aftenTimerTxt
     else
     aftenTimerTxt = ""
     aftenTimerTxtExp = 0
     end if 
      
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=aftenTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & aftenTimerTxtExp & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#51#") <> 0 then 
     
     if formatnumber(natTimer(x),2) <> 0 then
     natTimerTxt = formatnumber(natTimer(x),2)
     natTimerTxtExp = natTimerTxt
     else
     natTimerTxt = ""
     natTimerTxtExp = 0
     end if 
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=natTimerTxt%></td>
	 <%strEksport(x) = strEksport(x) & natTimerTxtExp & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then 
     
     if formatnumber(weekendTimer(x),2) <> 0 then
     weekendTimerTxt = formatnumber(weekendTimer(x),2)
     weekendTimerTxtExp = weekendTimerTxt
     else
     weekendTimerTxt = ""
     weekendTimerTxtExp = 0
     end if 
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=weekendTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & weekendTimerTxtExp & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#53#") <> 0 then 
     
     if formatnumber(weekendnatTimer(x),2) <> 0 then
     weekendnatTimerTxt = formatnumber(weekendnatTimer(x),2)
     weekendnatTimerTxtExp = weekendnatTimerTxt
     else
     weekendnatTimerTxt = ""
     weekendnatTimerTxtExp = 0
     end if 
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=weekendnatTimerTxt%></td>
	 <%strEksport(x) = strEksport(x) & weekendnatTimerTxtExp & ";"%>
	 <%end if %>
	 
	
	 <%if instr(akttype_sel, "#55#") <> 0 then 
     
     if formatnumber(aftenweekendTimer(x),2) <> 0 then
     aftenweekendTimerTxt = formatnumber(aftenweekendTimer(x),2)
     aftenweekendTimerTxtExp = aftenweekendTimerTxt
     else
     aftenweekendTimerTxt = ""
     aftenweekendTimerTxtExp = 0
     end if 
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=aftenweekendTimerTxt%></td>
	 <%strEksport(x) = strEksport(x) & aftenweekendTimerTxtExp & ";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then 
      
     if formatnumber(adhocTimer(x),2) <> 0 then
     adhocTimerTxt = formatnumber(adhocTimer(x),2)
     adhocTimerTxtExp = adhocTimerTxt
     else
     adhocTimerTxt = ""
     adhocTimerTxtExp = 0
     end if 
      
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=adhocTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & adhocTimerTxtExp & ";"%>
	 <%end if %>
	 
	 
     
     <%if instr(akttype_sel, "#90#") <> 0 then 
     
     if formatnumber(e1Timer(x),2) <> 0 then
     e1TimerTxt = formatnumber(e1Timer(x),2)
     e1TimerTxtExp = e1TimerTxt
     else
     e1TimerTxt = ""
     e1TimerTxtExp = 0
     end if 
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=e1TimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & e1TimerTxtExp & ";"%>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then 
       
     if formatnumber(e2Timer(x),2) <> 0 then
     e2TimerTxt = formatnumber(e2Timer(x),2)
     e2TimerTxtExp = e2TimerTxt
     else
     e2TimerTxt = ""
     e2TimerTxtExp = 0
     end if
       
       %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=e2TimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & e2TimerTxtExp & ";"%>
	 <%end if %>

      <%if instr(akttype_sel, "#92#") <> 0 then 
     
     if formatnumber(e3Timer(x),2) <> 0 then
     e3TimerTxt = formatnumber(e3Timer(x),2)
     e3TimerTxtExp = e3TimerTxt
     else
     e3TimerTxt = ""
     e3TimerTxtExp = 0
     end if 
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=e3TimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & e3TimerTxtExp & ";"%>
	 <%end if %>

	 
	 
	 <%if instr(akttype_sel, "#6#") <> 0 then 
     
     if formatnumber(sTimer(x),2) <> 0 then
     sTimerTxt = formatnumber(sTimer(x),2)
     sTimerTxtExp = sTimerTxt
     else
     sTimerTxt = ""
     sTimerTxtExp = 0
     end if
     
     %>
	  <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=sTimerTxt%></td>
	 <%strEksport(x) = strEksport(x) & sTimerTxtExp & ";"%>
	 <%end if %>
	  
	  <%if instr(akttype_sel, "#7#") <> 0 then 
      
     if formatnumber(flexTimer(x),2) <> 0 then
     flexTimerTxt = formatnumber(flexTimer(x),2)
     flexTimerTxtExp = flexTimerTxt
     else
     flexTimerTxt = ""
     flexTimerTxtExp = 0
     end if
      
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=flexTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & flexTimerTxtExp & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then 
     
     if formatnumber(km(x),2) <> 0 then
     kmTxt = formatnumber(km(x),2)
     kmTxtExp = kmTxt
     else
     kmTxt = ""
     kmTxtExp = 0
     end if
     
     %>
	   <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=kmTxt %></td>
	  <%strEksport(x) = strEksport(x) & kmTxtExp & ";"%>
	  <%end if %>
	  
	  <%if instr(akttype_sel, "#10#") <> 0 then 
      
     if formatnumber(fpTimer(x),2) <> 0 then
     fpTimerTxt = formatnumber(fpTimer(x),2)
     fpTimerTxtExp = fpTimerTxt
     else
     fpTimerTxt = ""
     fpTimerTxtExp = 0
     end if
      
      %>
	  <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=fpTimerTxt %></td>
	  <%strEksport(x) = strEksport(x) & fpTimerTxtExp & ";"%>
	  <%end if %>
	  
	  <%if instr(akttype_sel, "#9#") <> 0 then 
      
     if formatnumber(pausTimer(x),2) <> 0 then
     pausTimerTxt = formatnumber(pausTimer(x),2)
     pausTimerTxtExp = pausTimerTxt
     else
     pausTimerTxt = ""
     pausTimerTxtExp = 0
     end if
      
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=pausTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & pausTimerTxtExp & ";"%>
	 <%end if %>
	 
	 
	 
	  <!--Ferie Fridage -->
	 <%if instr(akttype_sel, "#12#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	
	 <%if normTimerDag(x) <> 0 then 
	 fefriVal = fefriTimer(x)/normTimerDag(x)
	 else
	 fefriVal = 0
	 end if
	

     if fefriVal <> 0 then
     fefriValTxt = formatnumber(fefriVal,2)
     fefriValTxtExp = fefriValTxt
     else
     fefriValTxt = ""
     fefriValTxtExp = 0
     end if
     %>

	 <%=fefriValTxt%>
	 
	 </td>
	 
	 <%strEksport(x) = strEksport(x) & fefriValTxtExp & ";"%>
	 <%end if %>
	 
	 
	  <%if instr(akttype_sel, "#18#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	
	 <%if normTimerDag(x) <> 0 then 
	 fefriplVal = fefriplTimer(x)/normTimerDag(x)
	 else
	 fefriplVal = 0
	 end if

     if fefriplVal <> 0 then
     fefriplValTxt = formatnumber(fefriplVal,2)
     fefriplValTxtExp = fefriplValTxt
     else
     fefriplValTxt = ""
     fefriplValTxtExp = 0
     end if
	 %>
	 <%=fefriplValTxt%>
	 
	 </td>
	 
	 <%strEksport(x) = strEksport(x) & fefriplValTxtExp & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	  <%if normTimerDag(x) <> 0 then 
	 fefribrVal = fefriTimerBr(x)/normTimerDag(x)
	 else
	 fefribrVal = 0
	 end if

     if fefribrVal <> 0 then
     fefribrValTxt = formatnumber(fefribrVal,2)
     fefribrValTxtExp = fefribrValTxt
     else
     fefribrValTxt = ""
     fefribrValTxtExp = 0
     end if
	 %>
	 
	 <%=fefribrValTxt%> </td> 
	 
	 <%strEksport(x) = strEksport(x) & fefribrValTxtExp & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#17#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 fefriudbVal = fefriTimerUdb(x)/normTimerDag(x)
	 else
	 fefriudbVal = 0
	 end if

     if fefriudbVal <> 0 then
     fefriudbValTxt = formatnumber(fefriudbVal,2)
     fefriudbValTxtExp = fefriudbValTxt
     else
     fefriudbValTxt = ""
     fefriudbValTxtExp = 0
     end if
	 %>
	 
	 <%=fefriudbValTxt%></td> 
	 
	 <%strEksport(x) = strEksport(x) & fefriudbValTxtExp & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then 
	 fefriBal = 0
	 fefriBal  = (fefriTimer(x) - (fefriTimerBr(x) + fefriTimerUdb(x)))
	 
	 if normTimerDag(x) <> 0 then 
	 fefriBalVal = fefriBal/normTimerDag(x)
	 else
	 fefriBalVal = 0
	 end if

     if fefriBalVal <> 0 then
     fefriBalValTxt = formatnumber(fefriBalVal,2)
     fefriBalValTxtExp = fefriBalValTxt
     else
     fefriBalValTxt = ""
     fefriBalValTxtExp = 0
     end if
	 
	 %>
	 <td align=right class=lille style="  white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 <%=fefriBalValTxt%></td>
	 <%strEksport(x) = strEksport(x) & fefriBalValTxtExp & ";"%>
	 <%end if %>
	 
	 <%
     if cint(visikkeFerieogSygiPer) <> 1 then    
         
     if instr(akttype_sel, "#13#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 fefriPerbrVal = fefriTimerPerBr(x)/normTimerDag(x)
     fefriTimerPerBrTimerVal = fefriTimerPerBrTimer(x)
	 else
	 fefriPerbrVal = 0
	 fefriTimerPerBrTimerVal = 0
     end if


     if fefriPerbrVal <> 0 then
     fefriPerbrValTxt = formatnumber(fefriPerbrVal,2)
     fefriPerbrValTxtExp = fefriPerbrValTxt
     fefriTimerPerBrTimerValTxt = formatnumber(fefriTimerPerBrTimerVal, 2)
     fefriTimerPerBrTimerValExp = fefriTimerPerBrTimerValTxt
     else
     fefriPerbrValTxt = ""
     fefriPerbrValTxtExp = 0
     fefriTimerPerBrTimerValTxt = ""
     fefriTimerPerBrTimerValExp = 0 
     end if
	 %>
	 
	 <%=fefriPerbrValTxt%></td> 
     <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=fefriTimerPerBrTimerValTxt %></td>


	 
	 <%strEksport(x) = strEksport(x) & fefriPerbrValTxtExp & ";"&fefriTimerPerBrTimerValExp&";"%>
	


     <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"> 
     <% 
     if len(trim(feriefriAFPerstDato(x))) AND feriefriAFPerstDato(x) <> "2001/12/31" then
     feriefriforsteDagIPerDato = formatdatetime(feriefriAFPerstDato(x), 2)
     else
     feriefriforsteDagIPerDato = ""
     end if
     %>

     <%=feriefriforsteDagIPerDato %>
    </td>
	 
	 <%strEksport(x) = strEksport(x) & feriefriforsteDagIPerDato&";"%>
     <%end if %>


      <%if instr(akttype_sel, "#17#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 fefriudbValPer = fefriTimerUdbPer(x)/normTimerDag(x)
	 else
	 fefriudbValPer = 0
	 end if

     if fefriudbValPer <> 0 then
     fefriudbValPerTxt = formatnumber(fefriudbValPer,2)
     fefriudbValPerTxtExp = fefriudbValPerTxt
     else
     fefriudbValPerTxt = ""
     fefriudbValPerTxtExp = 0
     end if
	 %>
	 
	 <%=fefriudbValPerTxt%></td> 

      <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 fefriudbValPerTimer = fefriTimerUdbPerTimer(x)/normTimerDag(x)
	 else
	 fefriudbValPerTimer = 0
	 end if

     if fefriudbValPerTimer <> 0 then
     fefriudbValPerTimerTxt = formatnumber(fefriudbValPerTimer,2)
     fefriudbValPerTimerTxtExp = fefriudbValPerTimerTxt
     else
     fefriudbValPerTimerTxt = ""
     fefriudbValPerTimerTxtExp = 0
     end if
	 %>
	 
	 <%=fefriudbValPerTimerTxt%></td> 

	 
	 <%strEksport(x) = strEksport(x) & fefriudbValPerTxtExp & ";" & fefriudbValPerTimerTxtExp & ";"%>
	 <%
     end if
             
     end if %>


	  
	
	 <!-- Ferie -->
	 
	  <%if instr(akttype_sel, "#15#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieVal = ferieOptjtimer(x)/normTimerDag(x)
	 else
	 ferieVal = 0
	 end if

     
     if ferieVal <> 0 then
     ferieValTxt = formatnumber(ferieVal,2)
     ferieValTxtExp = ferieValTxt
     else
     ferieValTxt = ""
     ferieValTxtExp = 0
     end if
	 %>
	 
	 <%=ferieValTxt %></td>
	 
	  <%strEksport(x) = strEksport(x) & ferieValTxtExp & ";"%>
	 <%end if %>


     <%'** Ferie overført ****'
     if instr(akttype_sel, "#111#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieValoverfort = ferieOptjOverforttimer(x)/normTimerDag(x)
	 else
	 ferieValoverfort = 0
	 end if

     
     if ferieValoverfort <> 0 then
     ferieValoverfortTxt = formatnumber(ferieValoverfort,2)
     ferieValoverfortTxtExp = ferieValoverfortTxt
     else
     ferieValoverfortTxt = ""
     ferieValoverfortTxtExp = 0
     end if
	 %>
	 
	 <%=ferieValoverfortTxt %></td>
	 
	  <%strEksport(x) = strEksport(x) & ferieValoverfortTxtExp & ";"%>
	 <%end if %>

    
      <%'** Ferie optjent u. løn **' 
     if instr(akttype_sel, "#112#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieValUlon = ferieOptjUlontimer(x)/normTimerDag(x)
	 else
	 ferieValUlon = 0
	 end if

     
     if ferieValUlon <> 0 then
     ferieValUlonTxt = formatnumber(ferieValUlon,2)
     ferieValUlonTxtExp = ferieValUlonTxt
     else
     ferieValUlonTxt = ""
     ferieValUlonTxtExp = 0
     end if
	 %>
	 
	 <%=ferieValUlonTxt %></td>
	 
	  <%strEksport(x) = strEksport(x) & ferieValUlonTxtExp & ";"%>
	 <%end if %>



	 
	 
	     <%if instr(akttype_sel, "#11#") <> 0 then %>
	     <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	     <%if normTimerDag(x) <> 0 then 
	     feriePlVal = feriePLTimer(x)/normTimerDag(x)
	     else
	     feriePlVal = 0
	     end if
	 
     
         if feriePlVal <> 0 then
         feriePlValTxt = formatnumber(feriePlVal,2)
         feriePlValTxtExp = feriePlValTxt
         else
         feriePlValTxt = ""
         feriePlValTxtExp = 0
         end if
         %>
	 
	 
	     <%=feriePlValTxt%></td>

	      <%strEksport(x) = strEksport(x) & feriePlValTxtExp & ";"%>
	 
	     <%end if %>
	 
	 <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieAftVal = ferieAFTimer(x)/normTimerDag(x)
	 else
	 ferieAftVal = 0
	 end if


     if ferieAftVal <> 0 then
     ferieAftValTxt = formatnumber(ferieAftVal,2)
     ferieAftValTxtExp = ferieAftValTxt
     else
     ferieAftValTxt = ""
     ferieAftValTxtExp = 0
     end if
	 %>
	 
	 <%=ferieAftValTxt %> </td>
	 
	 <%strEksport(x) = strEksport(x) & ferieAftValTxtExp & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#19#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieAftulonVal = ferieAFulonTimer(x)/normTimerDag(x)
	 else
	 ferieAftulonVal = 0
	 end if

     if ferieAftulonVal <> 0 then
     ferieAftulonValTxt = formatnumber(ferieAftulonVal,2)
     ferieAftulonValTxtExp = ferieAftulonValTxt
     else
     ferieAftulonValTxt = ""
     ferieAftulonValTxtExp = 0
     end if


	 %>

	 <%=ferieAftulonValTxt %> </td>
	 
	 <%strEksport(x) = strEksport(x) & ferieAftulonValTxtExp & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#16#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieUdbVal = ferieUdbTimer(x)/normTimerDag(x)
	 else
	 ferieUdbVal = 0
	 end if

     if ferieUdbVal <> 0 then
     ferieUdbValTxt = formatnumber(ferieUdbVal,2)
     ferieUdbValTxtExp = ferieUdbValTxt
     else
     ferieUdbValTxt = ""
     ferieUdbValTxtExp = 0
     end if
	 %>
	 
	 <%=ferieUdbValTxt%></td>
	 
	 <%strEksport(x) = strEksport(x) & ferieUdbValTxtExp & ";"%>
	 <%end if %>
	 
	 <% if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 OR instr(akttype_sel, "#111#") <> 0 OR instr(akttype_sel, "#112#") <> 0  then 
	 
	 ferieBal = 0
     call ferieBal_fn(ferieOptjtimer(x), ferieOptjOverforttimer(x), ferieOptjUlontimer(x), ferieAFTimer(x), ferieAFulonTimer(x), ferieUdbTimer(x))
	 
	 if normTimerDag(x) <> 0 then 
	 ferieBalVal = ferieBal/normTimerDag(x)
	 else
	 ferieBalVal = 0
	 end if
	 
     if ferieBalVal <> 0 then 
	 ferieBalValTxt = formatnumber(ferieBalVal,2)
     ferieBalValTxtExp = ferieBalValTxt
	 else
	 ferieBalValTxt = ""
     ferieBalValTxtExp = 0
	 end if

	 %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 <%=ferieBalValTxt %></td>
	
	<%strEksport(x) = strEksport(x) & ferieBalValTxtExp & ";"%>
	<%end if %>
	 


	 <!-- Ferie i periode -->
	  <%
      if cint(visikkeFerieogSygiPer) <> 1 then    
          
      if instr(akttype_sel, "#14#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	         <%if normTimerDag(x) <> 0 then 
	         ferieAftPerVal = ferieAFPerTimer(x)/normTimerDag(x)
             ferieAftPerValtim = formatnumber(ferieAFPerTimertimer(x),2)
	         else
	         ferieAftPerVal = 0
             ferieAftPerValtim = 0
	         end if

      
             if ferieAftPerVal <> 0 then 
	            
                ferieAftPerValTxt = formatnumber(ferieAftPerVal,2)
                ferieAftPerValTxtExp = ferieAftPerValTxt
                
                ferieAftPerValtimTxt = ferieAftPerValtim
                ferieAftPerValtimTxtExp = ferieAftPerValtimTxt

	         else
	         ferieAftPerValTxt = ""
             ferieAftPerValTxtExp = 0
             ferieAftPerValtimTxt = ""
             ferieAftPerValtimTxtExp = 0
	         end if
	         %>
	 
	 <%=ferieAftPerValTxt%></td>

     <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=ferieAftPerValtimTxt %></td>

      <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
     <% 
     if len(trim(ferieAFPerstDato(x))) AND ferieAFPerstDato(x) <> "2001/12/31" then
     ferieforsteDagIPerDato = formatdatetime(ferieAFPerstDato(x), 2)
     else
     ferieforsteDagIPerDato = ""
     end if
     %>

     <%=ferieforsteDagIPerDato %>
    </td>
	 
	 <%strEksport(x) = strEksport(x) & ferieAftPerValTxtExp&";"&ferieAftPerValtimTxtExp&";"&ferieforsteDagIPerDato&";"%>
	 <%end if %>




      <%if instr(akttype_sel, "#19#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	         <%if normTimerDag(x) <> 0 then 
	         ferieAftulonPerVal = ferieAFulonPer(x)/normTimerDag(x)
             ferieAftulonPerValtim = formatnumber(ferieAFulonPerTimer(x),2)
	         else
	         ferieAftulonPerVal = 0
             ferieAftulonPerValtim = 0
	         end if

      
             if ferieAftulonPerVal <> 0 then 
	            
                ferieAftulonPerValTxt = formatnumber(ferieAftulonPerVal,2)
                ferieAftulonPerValTxtExp = ferieAftulonPerValTxt
                
                ferieAftulonPerValtimTxt = ferieAftulonPerValtim
                ferieAftulonPerValtimTxtExp = ferieAftPerValtimTxt

	         else
	         ferieAftulonPerValTxt = ""
             ferieAftulonPerValTxtExp = 0
             ferieAftulonPerValtimTxt = ""
             ferieAftulonPerValtimTxtExp = 0
	         end if
	         %>
	 
	 <%=ferieAftulonPerValTxt%></td>

     <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=ferieAftulonPerValtimTxt %></td>
     <%strEksport(x) = strEksport(x) & ferieAftulonPerValTxtExp&";"&ferieAftulonPerValtimTxtExp&";"%>
     <%end if %>



         <%if instr(akttype_sel, "#16#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
	 
	         <%if normTimerDag(x) <> 0 then 
	         ferieUdbPerVal = ferieUdbPer(x)/normTimerDag(x)
             ferieUdbPerValtim = formatnumber(ferieUdbPerTimer(x),2)
	         else
	         ferieUdbPerVal = 0
             ferieUdbPerValtim = 0
	         end if

      
             if ferieUdbPerVal <> 0 then 
	            
                ferieUdbPerValTxt = formatnumber(ferieUdbPerVal,2)
                ferieUdbPerValTxtExp = ferieUdbPerValTxt
                
                ferieUdbPerValtimTxt = ferieUdbPerValtim
                ferieUdbPerValtimTxtExp = ferieUdbPerValtimTxt

	         else
	         ferieUdbPerValTxt = ""
             ferieUdbPerValTxtExp = 0
             ferieUdbPerValtimTxt = ""
             ferieUdbPerValtimTxtExp = 0
	         end if
	         %>
	 
	 <%=ferieUdbPerValTxt%></td>

     <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=ferieUdbPerValtimTxt %></td>
     <%strEksport(x) = strEksport(x) & ferieUdbPerValTxtExp&";"&ferieUdbPerValtimTxtExp&";"%>
     <%end if
     end if %>




   
	 
	 

        <%
     '** Div fri timer / 1 maj timer ***'
     if instr(akttype_sel, "#25#") <> 0 then 
     
     if divfritimer(x) <> 0 then
     divfritimerTxt = formatnumber(divfritimer(x),2) 
     divfritimerTxtExp = divfritimerTxt
     else
     divfritimerTxt = ""
     divfritimerTxtExp = 0
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=divfritimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & divfritimerTxtExp &";"%>
	 <%end if %>

       <%
     '** Rejsedage ***'
     if instr(akttype_sel, "#125#") <> 0 then 
     
     if rejseDage(x) <> 0 then
     rejseDageTxt = formatnumber(rejseDage(x),2) 
     rejseDageTxtExp = rejseDageTxt
     else
     rejseDageTxt = ""
     rejseDageTxtExp = 0
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=rejseDageTxt %></td>
	 <%strEksport(x) = strEksport(x) & rejseDageTxtExp &";"%>
	 <%end if %>


     <!-- Afspad --->
	 <%if instr(akttype_sel, "#30#") <> 0 then 
     
     if afspTimerTim(x) <> 0 then
     afspTimerTimTxt = "("& formatnumber(afspTimerTim(x), 2) & ")"
     afspTimerTimTxtExp = formatnumber(afspTimerTim(x), 2) 
     else
     afspTimerTimTxt = ""
     afspTimerTimTxtExp = 0
     end if


     if afspTimer(x) <> 0 then
     afspTimerTxt = formatnumber(afspTimer(x), 2)
     afspTimerTxtExp = afspTimerTxt
     else
     afspTimerTxt = ""
     afspTimerTxtExp = 0
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=afspTimerTimTxt &" "& afspTimerTxt%></td>
	 <%strEksport(x) = strEksport(x) & afspTimerTimTxtExp &";"& afspTimerTxtExp &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#31#") <> 0 then 
     
     if afspTimerBr(x) <> 0 then
     afspTimerBrTxt = formatnumber(afspTimerBr(x), 2)
     afspTimerBrTxtExp = afspTimerBrTxt
     else
     afspTimerBrTxt = ""
     afspTimerBrTxtExp = 0
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=afspTimerBrTxt%></td>

    


	 <%strEksport(x) = strEksport(x) & afspTimerBrTxtExp &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#32#") <> 0 then 
     
     if afspTimerUdb(x) <> 0 then
     afspTimerUdbTxt = formatnumber(afspTimerUdb(x), 2)
     afspTimerUdbTxtExp = afspTimerUdbTxt
     else
     afspTimerUdbTxt = ""
     afspTimerUdbTxtExp = 0
     end if
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=afspTimerUdbTxt%></td>
	 <%strEksport(x) = strEksport(x) & afspTimerUdbTxtExp &";"%>
	 <%end if %>

	 
	 <%if instr(akttype_sel, "#33#") <> 0 then 
	 afspadUdbBal = 0
	 afspadUdbBal = (afspTimerUdb(x) - afspTimerOUdb(x))
     
     if afspadUdbBal <> 0 then
     afspadUdbBalTxt = formatnumber(afspadUdbBal, 2)
     afspadUdbBalTxtExp = afspadUdbBalTxt
     else
     afspadUdbBalTxt = ""
     afspadUdbBalTxtExp = 0
     end if %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=afspadUdbBalTxt%></td>
	 <%strEksport(x) = strEksport(x) & afspadUdbBalTxtExp &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR _
	 instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	 
	 <%
     AfspadBal = 0 
	 AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))

     if AfspadBal <> 0 then
     AfspadBalTxt = formatnumber(AfspadBal, 2)
     AfspadBalTxtExp = AfspadBalTxt
     else
     AfspadBalTxt = ""
     AfspadBalTxtExp = 0
     end if
	 %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=AfspadBalTxt%></td>
	 <%strEksport(x) = strEksport(x) & AfspadBalTxtExp &";"%>
	 <%end if %>


	 
	 <%
     
       if cint(visikkeFerieogSygiPer) <> 1 then 

             
     if instr(akttype_sel, "#30#") <> 0 then 
     
     if afspTimerTimPer(x) <> 0 then
     afspTimerTimPerTxt = "("& formatnumber(afspTimerTimPer(x), 2) &")"
     afspTimerTimPerTxtExp = formatnumber(afspTimerTimPer(x), 2) 
     else
     afspTimerTimPerTxt = ""
     afspTimerTimPerTxtExp = 0
     end if
     

     if afspTimerPer(x) <> 0 then
     afspTimerPerTxt = formatnumber(afspTimerPer(x), 2)
     afspTimerPerTxtExp = afspTimerPerTxt
     else
     afspTimerPerTxt = ""
     afspTimerPerTxtExp = 0
     end if
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=afspTimerTimPerTxt &" "& afspTimerPerTxt%></td>
	 <%strEksport(x) = strEksport(x) & afspTimerTimPerTxtExp &";"& afspTimerPerTxtExp &";"%>
	 <%end if %>
	 
	 
     <%if instr(akttype_sel, "#31#") <> 0 then 
     
     if afspTimerBrPer(x) <> 0 then
     afspTimerBrPerTxt = formatnumber(afspTimerBrPer(x), 2)
     afspTimerBrPerTxtExp = afspTimerBrPerTxt
     else
     afspTimerBrPerTxt = ""
     afspTimerBrPerTxtExp = 0
     end if
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=afspTimerBrPerTxt%></td>

        <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;">
     <% 
     if len(trim(afspadPerstDato(x))) AND afspadPerstDato(x) <> "2001/12/31" then
     afspadforsteDagIPerDato = formatdatetime(afspadPerstDato(x), 2)
     else
     afspadforsteDagIPerDato = ""
     end if
     %>

     <%=afspadforsteDagIPerDato %>
    </td>
         <%strEksport(x) = strEksport(x) & afspTimerBrPerTxtExp &";"&afspadforsteDagIPerDato &";"%>
        <%end if %>


	 
	 <%end if %>
	 
	 
	 
	 <!-- Syg -->
	  
	  <%if instr(akttype_sel, "#20#") <> 0 then 
      
      if sygTimer(x) <> 0 then
      sygTimerTxt = formatnumber(sygTimer(x),2)
      sygTimerTxtExp = sygTimerTxt
      else
      sygTimerTxt = ""
      sygTimerTxtExp = 0
      end if
      
      if sygDage(x) <> 0 then
      sygDageTxt = formatnumber(sygDage(x),2)
      sygDageTxtExp = sygDageTxt
      else
      sygDageTxt = ""
      sygDageTxtExp = 0
      end if
      
      
      if sygTimerPer(x) <> 0 then
      sygTimerPerTxt = formatnumber(sygTimerPer(x),2)
      sygTimerPerTxtExp = sygTimerPerTxt
      else
      sygTimerPerTxt = ""
      sygTimerPerTxtExp = 0
      end if

      if sygDagePer(x) <> 0 then
      sygDagePerTxt = formatnumber(sygDagePer(x),2)
      sygDagePerTxtExp = sygDagePerTxt
      else
      sygDagePerTxt = ""
      sygDagePerTxtExp = 0
      end if
      %>

	 <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=sygTimerTxt %></td>
	 <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=sygDageTxt %></td>

     <%
     strEksport(x) = strEksport(x) & sygTimerTxtExp &";" & sygDageTxtExp & ";"    
         
     if cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=sygTimerPerTxt %></td>
	 <td align=right class=lille style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=sygDagePerTxt%></td>

       <%
              if len(trim(sygPerstDato(x))) AND sygPerstDato(x) <> "2001/12/31" then
             sygforsteDagIPerDato = formatdatetime(sygPerstDato(x), 2)
             else
             sygforsteDagIPerDato = ""
             end if
             %>

	         <td align=right class=lille  style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=sygforsteDagIPerDato %></td>

	
	 <%strEksport(x) = strEksport(x) & sygTimerPerTxtExp &";" & sygDagePerTxtExp & ";"& sygforsteDagIPerDato & ";"%>
	 
	 <%
     end if
     end if %>
	 
	 <%if instr(akttype_sel, "#21#") <> 0 then 
     
      if barnsygTimer(x) <> 0 then
      barnsygTimerTxt = formatnumber(barnsygTimer(x),2)
      barnsygTimerTxtExp = barnsygTimerTxt
      else
      barnsygTimerTxt = ""
      barnsygTimerTxtExp = 0
      end if

      if barnSygDage(x) <> 0 then
      barnSygDageTxt = formatnumber(barnSygDage(x),2)
      barnSygDageTxtExp = barnSygDageTxt
      else
      barnSygDageTxt = ""
      barnSygDageTxtExp = 0
      end if


      if barnSygTimerPer(x) <> 0 then
      barnSygTimerPerTxt = formatnumber(barnSygTimerPer(x),2)
      barnSygTimerPerTxtExp = barnSygTimerPerTxt
      else
      barnSygTimerPerTxt = ""
      barnSygTimerPerTxtExp = 0
      end if


      if barnsygDagePer(x) <> 0 then
      barnsygDagePerTxt = formatnumber(barnsygDagePer(x),2)
      barnsygDagePerTxtExp = barnsygDagePerTxt
      else
      barnsygDagePerTxt = ""
      barnsygDagePerTxtExp = 0
      end if
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=barnsygTimerTxt%></td>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=barnSygDageTxt%></td>

    <%
     strEksport(x) = strEksport(x) & barnsygTimerTxtExp &";" & barnSygDageTxtExp &";"   
        
        
     if cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=barnSygTimerPerTxt%></td>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=barnsygDagePerTxt%></td>

        <%
              if len(trim(barnsygPerstDato(x))) AND barnsygPerstDato(x) <> "2001/12/31" then
             barnsygforsteDagIPerDato = formatdatetime(barnsygPerstDato(x), 2)
             else
             barnsygforsteDagIPerDato = ""
             end if
             %>

	         <td align=right class=lille  style="white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=barnsygforsteDagIPerDato %></td>
	 
	 
	 <%strEksport(x) = strEksport(x) & barnSygTimerPerTxtExp &";" & barnsygDagePerTxtExp &";"& barnsygforsteDagIPerDato &";"%>
	 <%end if %>
	 
	 <%if (instr(akttype_sel, "#20#") <> 0 OR instr(akttype_sel, "#21#") <> 0) AND cint(visikkeFerieogSygiPer) <> 1 then 
	 
     if normTimer(x) <> 0 then
     '** Fravær beregnes udfra normtid / ikke optjent fravær, dvs. sygdom.
	 fravaeriper = formatnumber( ((BarnsygTimerPer(x) + sygTimerPer(x) + barsel(x) + lageTimer(x)) / normTimer(x)) * 100, 0) '+ fefriTimerPerBr(x) + ferieAFPerTimer(x)
	 else 
	 fravaeriper = 0
	 end if

     if fravaeriper <> 0 then
     fravaeriperTxt = fravaeriper
     fravaeriperTxtExp = fravaeriperTxt
     else
     fravaeriperTxt = ""
     fravaeriperTxtExp = 0
     end if
	 %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=fravaeriperTxt%> % </td>
	 <%strEksport(x) = strEksport(x) & fravaeriperTxtExp &";"%>
	 <%
         
     end if     


     end if%> 


	 

     <%
     '** Barsel ***
     if instr(akttype_sel, "#22#") <> 0 then 
     
     if barsel(x) <> 0 then
     barselTxt = formatnumber(barsel(x),2) 
     barselTxtExp = barselTxt
     else
     barselTxt = ""
     barselTxtExp = 0
     end if

      

     
     if len(trim(barselPerstDato(x))) AND  barselPerstDato(x) <> "2001/12/31" then
     barselforsteDagIPerDato = formatdatetime(barselPerstDato(x), 2)
     else
     barselforsteDagIPerDato = ""
     end if
     %>

     

      <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=barselTxt %></td>
      <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=barselforsteDagIPerDato %></td>
       
    <%strEksport(x) = strEksport(x) & barselTxtExp &";"& barselforsteDagIPerDato & ";"%>
	 <%end if %>



     <%
     '** Omsorg ***'
     if instr(akttype_sel, "#23#") <> 0 then 
     
     if omsorg(x) <> 0 then
     omsorgTxt = formatnumber(omsorg(x),2) 
     omsorgTxtExp = omsorgTxt
     else
     omsorgTxt = ""
     omsorgTxtExp = 0
     end if

       
     if len(trim(omsorgPerstDato(x))) AND  omsorgPerstDato(x) <> "2001/12/31" then
     omsorgforsteDagIPerDato = formatdatetime(omsorgPerstDato(x), 2)
     else
     omsorgforsteDagIPerDato = ""
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=omsorgTxt %></td>
        <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=omsorgforsteDagIPerDato %></td>
	 <%strEksport(x) = strEksport(x) & omsorgTxtExp &";"& omsorgforsteDagIPerDato &";"%>
	 <%end if %>



      <%
     '** Senior ***'
     if instr(akttype_sel, "#24#") <> 0 then 
     
     if senior(x) <> 0 then
     seniorTxt = formatnumber(senior(x),2) 
     seniorTxtExp = seniorTxt
     else
     seniorTxt = ""
     seniorTxtExp = 0
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=seniorTxt %></td>
	 <%strEksport(x) = strEksport(x) & seniorTxtExp &";"%>
	 <%end if %>




     <%
     '** Aldersrteduktion ***'
     if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then 
     
     if aldersreduktionOpj(x) <> 0 then
     aldersreduktionOpjTxt = formatnumber(aldersreduktionOpj(x),2) 
     aldersreduktionOpjTxtExp = aldersreduktionOpjTxt
     else
     aldersreduktionOpjTxt = ""
     aldersreduktionOpjTxtExp = 0
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=aldersreduktionOpjTxt %></td>
	 <%strEksport(x) = strEksport(x) & aldersreduktionOpjTxtExp &";"


     if aldersreduktionBr(x) <> 0 then
     aldersreduktionBrTxt = formatnumber(aldersreduktionBr(x),2) 
     aldersreduktionBrTxtExp = aldersreduktionBrTxt
     else
     aldersreduktionBrTxt = ""
     aldersreduktionBrTxtExp = 0
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=aldersreduktionBrTxt %></td>
	 <%strEksport(x) = strEksport(x) & aldersreduktionBrTxtExp &";"


     if aldersreduktionUdb(x) <> 0 then
     aldersreduktionUdbTxt = formatnumber(aldersreduktionUdb(x),2) 
     aldersreduktionUdbTxtExp = aldersreduktionUdbTxt
     else
     aldersreduktionUdbTxt = ""
     aldersreduktionUdbTxtExp = 0
     end if

     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=aldersreduktionUdbTxt %></td>
	 <%strEksport(x) = strEksport(x) & aldersreduktionUdbTxtExp &";"

         aldersreduktionSaldo = 0
         aldersreduktionSaldo = aldersreduktionOpj(x) - (aldersreduktionBr(x) + aldersreduktionUdb(x))

           
         if aldersreduktionSaldo <> 0 then
         aldersreduktionSaldoTxt = formatnumber(aldersreduktionSaldo, 2)
         aldersreduktionSaldoTxtExp = aldersreduktionSaldoTxt
         else
         aldersreduktionSaldoTxt = ""
         aldersreduktionSaldoTxtExp = 0
         end if 


         %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=aldersreduktionSaldoTxt%></td>
	 <%strEksport(x) = strEksport(x) & aldersreduktionSaldoTxtExp &";"

	 end if %>



        <!-- Aldersred, omsorg 2,10, K planglagt -->

      <%if instr(akttype_sel, "#26#") <> 0 then 
  
     if aldersreduktionPl(x) <> 0 then
     aldersreduktionPlTxt = formatnumber(aldersreduktionPl(x),2)
     aldersreduktionPlTxtExp = aldersreduktionPlTxt 
     else
     aldersreduktionPlTxt = ""
     aldersreduktionPlTxtExp = 0
     end if
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=aldersreduktionPlTxt %></td>
	 <%strEksport(x) = strEksport(x) & aldersreduktionPlTxtExp &";"%>
	 <%end if %>
    


            <%if instr(akttype_sel, "#120#") <> 0 then 
      
     if omsorg2pl(x) <> 0 then
     omsorg2plTxt = formatnumber(omsorg2pl(x),2)
     omsorg2plTxtExp = omsorg2plTxt 
     else
     omsorg2plTxt = ""
     omsorg2plTxtExp = 0
     end if
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=omsorg2plTxt %></td>
	 <%strEksport(x) = strEksport(x) & omsorg2plTxtExp &";"%>
	 <%end if %>


     <%if instr(akttype_sel, "#121#") <> 0 then 
      
     if omsorg10pl(x) <> 0 then
     omsorg10plTxt = formatnumber(omsorg10pl(x),2)
     omsorg10plTxtExp = omsorg10plTxt 
     else
     omsorg10plTxt = ""
     omsorg10plTxtExp = 0
     end if
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=omsorg10plTxt %></td>
	 <%strEksport(x) = strEksport(x) & omsorg10plTxtExp &";"%>
	 <%end if %>



     <%if instr(akttype_sel, "#122#") <> 0 then 
      
     if omsorgKpl(x) <> 0 then
     omsorgKplTxt = formatnumber(omsorgKpl(x),2)
     omsorgKplTxtExp = omsorgKplTxt 
     else
     omsorgKplTxt = ""
     omsorgKplTxtExp = 0
     end if
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=omsorgKplTxt %></td>
	 <%strEksport(x) = strEksport(x) & omsorgKplTxtExp &";"%>
	 <%end if %>



          <%if instr(akttype_sel, "#123#") <> 0 then 
      
     if ulempe1706udb(x) <> 0 then
     ulempe1706udbTxt = formatnumber(ulempe1706udb(x),2)
     ulempe1706udbTxtExp = ulempe1706udbTxt 
     else
     ulempe1706udbTxt = ""
     ulempe1706udbTxtExp = 0
     end if
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=ulempe1706udbTxt %></td>
	 <%strEksport(x) = strEksport(x) & ulempe1706udbTxtExp &";"%>
	 <%end if %>



           <%if instr(akttype_sel, "#124#") <> 0 then 
      
     if ulempeWudb(x) <> 0 then
     ulempeWudbTxt = formatnumber(ulempeWudb(x),2)
     ulempeWudbTxtExp = ulempeWudbTxt 
     else
     ulempeWudbTxt = ""
     ulempeWudbTxtExp = 0
     end if
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=ulempeWudbTxt %></td>
	 <%strEksport(x) = strEksport(x) & ulempeWudbTxtExp &";"%>
	 <%end if %>

 

	 
	 
	  <%if instr(akttype_sel, "#8#") <> 0 then 
      
     if sundTimer(x) <> 0 then
     sundTimerTxt = formatnumber(sundTimer(x),2)
     sundTimerTxtExp = sundTimerTxt 
     else
     sundTimerTxt = ""
     sundTimerTxtExp = 0
     end if
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=sundTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & sundTimerTxtExp &";"%>
	 <%end if %>
	 

         <!-- læge timer -->
	   <%if instr(akttype_sel, "#81#") <> 0 then 
       
     if lageTimer(x) <> 0 then
     lageTimerTxt = formatnumber(lageTimer(x),2) 
     lageTimerTxtExp = lageTimerTxt
     else
     lageTimerTxt = ""
     lageTimerTxtExp = 0
     end if
       %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=lageTimerTxt%></td>
	 <%strEksport(x) = strEksport(x) & lageTimerTxtExp &";"%>
	 <%end if %>
	 

         <!-- Tjenestefri --->

            <%if instr(akttype_sel, "#115#") <> 0 then 
       
     if tjenestefri(x) <> 0 then
     tjenestefriTxt = formatnumber(tjenestefri(x),2) 
     tjenestefriTxtExp = tjenestefriTxt
     else
     tjenestefriTxt = ""
     tjenestefriTxtExp = 0
     end if
       %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=tjenestefriTxt%></td>
	 <%strEksport(x) = strEksport(x) & tjenestefriTxtExp &";"%>
	 <%end if %>
	 



	 <!-- Ressource -->
	 <%if instr(akttype_sel, "#-2#") <> 0 then 
     
     if resTimer(x) <> 0 then
     resTimerTxt = formatnumber(resTimer(x),2)
     resTimerTxtExp = resTimerTxt 
     else
     resTimerTxt = ""
     resTimerTxtExp = 0
     end if
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=resTimerTxt %></td>
	 <%strEksport(x) = strEksport(x) & resTimerTxtExp &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then 
     
     if fakTimer(x) <> 0 then
     fakTimerTxt = formatnumber(fakTimer(x),2)
     fakTimerTxtExp = fakTimerTxt 
     else
     fakTimerTxt = ""
     fakTimerTxtExp = 0
     end if
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=fakTimerTxt%></td>
	 <%strEksport(x) = strEksport(x) & fakTimerTxtExp &";"%>
	 <%end if %> 
	  
	  <%if instr(akttype_sel, "#61#") <> 0 then 
      
     if stkAntal(x) <> 0 then
     stkAntalTxt = formatnumber(stkAntal(x),2)
     stkAntalTxtExp = stkAntalTxt 
     else
     stkAntalTxt = ""
     stkAntalTxtExp = 0
     end if
      
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=stkAntalTxt%></td>
	 <%strEksport(x) = strEksport(x) & stkAntalTxtExp &";"%>
	 <%end if %>
	 

       <%if instr(akttype_sel, "#113#") <> 0 then 
      
     if korrektionKomG(x) <> 0 then
     korrektionKomGTxt = formatnumber(korrektionKomG(x),2)
     korrektionKomGTxtExp = korrektionKomGTxt
     else
     korrektionKomGTxt = ""
     korrektionKomGTxtExp = 0
     end if
      
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=korrektionKomGTxt%></td>
	 <%strEksport(x) = strEksport(x) & korrektionKomGTxtExp &";"%>
	 <%end if %>


        <%if instr(akttype_sel, "#114#") <> 0 then 
      
     if korrektionReal(x) <> 0 then
     korrektionRealTxt = formatnumber(korrektionReal(x),2)
     korrektionRealTxtExp = korrektionRealTxt
     else
     korrektionRealTxt = ""
     korrektionRealTxtExp = 0
     end if
      
      %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=korrektionRealTxt%></td>
	 <%strEksport(x) = strEksport(x) & korrektionRealTxtExp &";"%>
	 <%end if %>


	 <%if instr(akttype_sel, "#-4#") <> 0 then 
     
     if mfForbrug(x) <> 0 then
     mfForbrugTxt = formatnumber(mfForbrug(x),2)
     mfForbrugTxtExp = mfForbrugTxt 
     else
     mfForbrugTxt = ""
     mfForbrugTxtExp = 0
     end if
     
     %>
	 <td align=right class=lille style=" white-space: nowrap; border-bottom: 1px  #D6DfF5 solid; border-right: 1px  #D6DfF5 solid;"><%=mfForbrugTxt%></td>
	 <%strEksport(x) = strEksport(x) & mfForbrugTxtExp &";"%>
	 <%end if %>
	 </tr>
	 
	 
	 
	 
	 <%

     

     strEksport(x) = strEksport(x) & ";"&startdato&";"&slutdato&";"

     '** FK opret eksport fil til SD løn ***'
     'if len(trim(request("sd_lon_fil"))) <> 0 then
     'strEksport(x) = strEksport(x) & ";"&forsteDagIPerDato&";"
     'end if
	 
	 strEksportTxt = strEksportTxt & strEksport(x)

     'Response.write "strEksportTxtstrEksportTxt: "& strEksportTxt & "<br>"
     'Response.flush
	 strEksportEmailTxt = strEksportEmailTxtHeader & strEksport(x)
	 
	 if sendemail = "j" then
	 
	 
	 
	 
	        '***** Oprettter Mail object ***
	        '**** Skal kun sende en linie til hver medarbejder ***'
	        '**** De må ikke se hinandens ****'
			if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\bal_real_norm_2007.asp" then
					
			
			
                txtEmailHtml = replace(strEksportEmailTxt, "xx99123sy#z", "</tr><tr><td class=lille align=right valign=bottom>")
			    txtEmailHtml = replace(txtEmailHtml, ";", "</td><td class=lille align=right valign=bottom>")
			   

                Set myMail=CreateObject("CDO.Message")
                myMail.Subject="TimeOut - Afstemningsrapport"
                myMail.From = "timeout_no_reply@outzource.dk"
				      
                'myMail.To=strEmail
                if len(trim(medarbEmail(x))) <> 0 then
                myMail.To= ""& medarbNavn(x) &"<"& medarbEmail(x) &">"
                end if

                           
               myMail.htmlBody= "<html><head><title></title>"_
			   &"<LINK rel=""stylesheet"" type=""text/css"" href=""http://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/style/timeout_style_print.css""></head>"_
			   &"<body><br>Afstemningsrapport<br>"&replace(strEksportPer, "xx99123sy#z", "<br>")&"<table width=100% cellpadding=1 cellspacing=0 border=1><tr><td class=lille valign=bottom>"_ 
			   & txtEmailHtml & "</table><br><br>Med venlig hilsen<br>" & session("user") & "</body></html>"

               'myMail.AddAttachment "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\log\data\"& file

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
                    
                if len(trim(medarbEmail(x))) <> 0 then
                myMail.Send
                end if
                set myMail=nothing

				
				'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			    '' Sætter Charsettet til ISO-8859-1 
			    'Mailer.CharSet = 2
			    '' Afsenderens navn 
			    'Mailer.FromName = "TimeOut Email Service"
			    '' Afsenderens e-mail 
			    'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
			    'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
    			
			    '' Mailens emne
			    'Mailer.Subject = "TimeOut Afstemningsrapport"
			    ' Modtagerens navn og e-mail
			    'Mailer.AddRecipient medarbNavn(x), medarbEmail(x)
			    'Mailer.ContentType = "text/html"
			    
			    'Mailer.AddAttachment "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\"& file 
			    
					'' Selve teksten
					'Mailer.BodyText = "<html><head><title></title>"_
					'&"<LINK rel=""stylesheet"" type=""text/css"" href=""https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/style/timeout_style_print.css""></head>"_
					'&"<body><br>Afstemningsrapport<br>"&replace(strEksportPer, "xx99123sy#z", "<br>")&"<table width=100% cellpadding=1 cellspacing=0 border=1><tr><td class=lille valign=bottom>"_ 
					'& txtEmailHtml & "</table><br><br>Med venlig hilsen<br>" & session("user") & "</body></html>"
					
					
				
					'If Mailer.SendMail Then
    				
					'Else
    				'	Response.Write "Fejl...<br>" & Mailer.Response
  					'End if
					
			
			end if
		
		
		strEksportEmailTxt = ""
		
	 
	 end if 'send email
	 
	 end if 'nulfilter
	 
	 
	 next
     
     
     
   
     %>
	 <!-- </table> -->

	<%
	


     '************************************************
     '** SALDO ***
     '************************************************
	 '************************************************
	 case 2,3 '**** minivisning på timereg siden ***
	 
	 %>
	
	 
	 
	 <%if visning = 2 then%>
	     	    
	   <br /><br />
	   
	     <h4> Måned >> Dato</h4> <br /><b><%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> </b> 
	     <%if year(slutDato) = year(now) AND month(slutDato) = month(now) then%>
	     (<%=tsa_txt_318 %>)
	     <%end if %>
	    
	 <%else%>
	    
	    
	    <%=tsa_txt_173%> pr. dag: <%=formatnumber(showNormTimerdag) %> 
	    <br /><%=tsa_txt_273%>
	    <br />
	    <br /><br />
	    
	    <h4>Uge <%=datepart("ww", stdato, 2, 2) %> - <%=datepart("yyyy", stdato, 2, 2) %></h4>
	    <b><%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %></b> 
	    <%if datepart("ww", stdato, 2, 2) = datepart("ww", now, 2, 2) AND datepart("yyyy", stdato, 2, 2) = datepart("yyyy", now, 2, 2) then%>
	    (<%=tsa_txt_318 %>)
	    <%end if %>
	 <%end if%>
	 
	 
	 
	 <%
	 x = 0
	 
	 %>
	 
	
	
	 <table cellspacing=0 cellpadding=2 border=0 width=100%>
	 <tr bgcolor="#EFF3FF">
	 <td class=lille style="border-bottom:1px silver dashed;">&nbsp;</td>
	 <td class=lille style="border-bottom:1px silver dashed;"><b>Timer<br /> realiseret</b></td>
	 <%if lto <> "cst" then %>
	 <td class=lille bgcolor="#cccccc" style="border-bottom:1px silver dashed;">(heraf<br />fak.bare)</td>
	 <%end if %>
	 <td align=right bgcolor="#DCF5BD" class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_173%></b></td>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Komme/gå</b><br />(løntimer)</td>

      <%if cint(showkgtil) = 1 then %>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Till. / <br />Frad. +/-</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Sum</b></td>
     <%end if %>

	 <%end if %>
	
	  <%if lto <> "cst" then %>
	 <td bgcolor="pink" align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b> (Real / Norm)</td>
	 <%end if %>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td bgcolor="#DCF5BD" align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b> (Komme/gå / Norm)</td>
	 
	     <%if lto <> "cst" then %>
	     <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_166%></b> (Real / Komme/gå)</td>
	     <%end if %>
	 
	 <%end if %>
	 </tr>
	 <tr>
	 <td class=lille style="border-bottom:1px silver dashed;">
      <%=tsa_txt_148%></td>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(realTimer(x),2) %></td>
	 
	  <%if lto <> "cst" then %>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">(<%=formatnumber(realfTimer(x),2)%>)</td>
	 <%end if %>
	 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimer(x),2)%></td>
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	  <% 
	  ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	 call fLonTimerPer(ltimerStDato, dageDiff, 2, intMid)
	 %>
	 </td>

       <%if cint(showkgtil) = 1 then %>

	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	  <%call timerogminutberegning(fradragTimer(x)*60) %>
	<%=thoursTot &":"& left(tminTot, 2) %>
	 </td>
	 <%
	 ltimerKorFrad = 0
	 ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
	 %>
	  <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	   <%call timerogminutberegning(ltimerKorFrad*60) %>
	<%=thoursTot &":"& left(tminTot, 2) %></td>
	 
     
     <%end if %>
     <%end if %>


	 
	 <%if lto <> "cst" then %>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber((realTimer(x) - normTimer(x)),2)%></b></td>
	 <%end if %>
	 
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%normtime_lontime = -((normTimer(x) - (ltimerKorFrad)) * 60) %>
	 <%call timerogminutberegning(normtime_lontime) %>
	 <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	 </td>
	 <%if lto <> "cst" then %>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber((realTimer(x) - (ltimerKorFrad)),2)%></td>
	 <%end if %>
	 
	 <%end if %>
	 </tr>
	 
	 <%if visning = 2 then %>
	 <!-- MD -Dato  -->
	  <tr>
	         <td class=lille style="border-bottom:1px silver dashed;">
              <%=tsa_txt_272%></td>
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(realTimer(x)/weekDiff,2)%></td>
	         <%if lto <> "cst" then %>
	         <td align=right class=lille style="border-bottom:1px silver dashed;">(<%=formatnumber(realfTimer(x)/weekDiff,2)%>)</td>
	         <%end if %>
	 
	         
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimerUge(x), 2)%></td>
	         <%if session("stempelur") <> 0 then %> 

             <%if cint(showkgtil) = 1 then %>
	         <td align=right style="border-bottom:1px silver dashed;" class=lille>
	          <% 
	          ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	         call fLonTimerPer(ltimerStDato, dageDiff, 4, intMid)
	         %>
	         </td>
	         <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	        <%call timerogminutberegning((fradragTimer(x)/weekDiff)*60) %>
	        <%=thoursTot &":"& left(tminTot, 2) %>
	         </td>


	         <%
	         ltimerKorFrad = 0
	         ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x))/weekDiff)
	         
             else

             ltimerKorFrad = 0
	         ltimerKorFrad = ((totalTimerPer100/60)/weekDiff)
             
             end if%>
	        <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	         
	         
	         <%call timerogminutberegning(ltimerKorFrad*60) %>
	        <%=thoursTot &":"& left(tminTot, 2) %></td>
	 
	         <%end if %>
	         
	         
	         <%if lto <> "cst" then %>
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber(((realTimer(x)/weekDiff) - normTimerUge(x)),2)%></b></td>
	        <%end if %>
	        
	        
	        
	        <%if session("stempelur") <> 0 then %> 
	        <td align=right style="border-bottom:1px silver dashed;" class=lille>
	       
	        <%
	        normtime_lontime = -((normTimerUge(x) - (ltimerKorFrad)) * 60)
	        call timerogminutberegning(normtime_lontime) %>
		    <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	        </td>
	         <%if lto <> "cst" then %>
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(((realTimer(x)/weekDiff) - (ltimerKorFrad)),2)%></td>
        	 <%end if %>
        	<%end if %>
	         </tr>
	         
	       <%end if %>  
	        
	         </table>
	
	    <%   
        
            
	 case 6   '***** ????   
	          
	          
	         
	          
	         '***** MD --> Dato Aktyper fordeling / udspec 
	         '** Tjekker om Afspadsering og / eller overarbejde er aktive ***'
	         'Response.Write "###"& afspTimer(x)
	         call akttyper2009(2)
	         'aktiveTyper = akttype_sel
        	  %> 
	         <!-- Afspad / Overarb --->
	         <%if instr(aktiveTyper, "#30#") <> 0 OR instr(aktiveTyper, "#31#") <> 0 then %>
	           <br /><br />
	           
	           
	       
	           
	          <table cellspacing=0 cellpadding=2 border=0 width=100%>
	           <tr>
	         <td colspan=6><b>Overarbejde</b><br />
	         <%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> 
	         <%if year(slutDato) = year(now) AND month(slutDato) = month(now) then%>
	         (<%=tsa_txt_318 %>)
	         <%end if %>
	         </td>
	         </tr>
	         <tr bgcolor="#EFF3FF">
	        <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b>Overarb.</b><br />(enh.)</td>
	          <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b>Afspads.</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Ønsk. Udbe.</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_283 &" "& tsa_txt_280 %></b></td>
	         </tr>
	          <tr>
        	 
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(afspTimer(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerBr(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerUdb(x), 2)%></td>
	     
	          <% 
	         afspadUdbBal = 0
	         afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspadUdbBal, 2)%></td>
	          <% 
	         AfspadBal = 0 
	         AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(AfspadBal, 2)%></td>
        	 
	         </tr>
	         </table>
	         <%end if %>
	         
	         
	         
        	 
        	 
            <!-- Syg -->
             <%if (level = 1 OR (session("mid") = usemrn)) OR (cint(erTeamlederForVilkarligGruppe) = 1) then%>
              <br /><br />
              <b>Sygdom & Sundhed</b><br />
               <%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> 
               
                 <%if year(slutDato) = year(now) AND month(slutDato) = month(now) then%>
	             (<%=tsa_txt_318 %>)
	             <%end if %>
               
	          <table cellspacing=0 cellpadding=2 border=0 width=100%>
	           <tr bgcolor="#EFF3FF">
	           <%if instr(akttype_sel, "#20#") <> 0 then %>
	        <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b>Syg</b><br />timer</td>
	        <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b>Syg</b><br />~ dage</td>
	        <%end if %>
	          <%if instr(akttype_sel, "#21#") <> 0 then %>
	          <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b>Barn syg</b><br />timer</td>
	          <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b>Barn syg</b><br />~ dage</td>
	 
	           <%end if %>
	           <%if instr(akttype_sel, "#8#") <> 0 then %>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Sundhed</b><br />timer</td>
	           <%end if %>
	           <%if instr(akttype_sel, "#81#") <> 0 then %>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Læge</b><br />timer</td>
	           <%end if %>
	          </tr>
	          <tr>
	          
	          
	          <%if instr(akttype_sel, "#20#") <> 0 then %>
	          <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(sygTimer(x),2) %></td>
	          
	          <%if sygDage(x) <> 0 then  'normTimerDag(x) <> 0 then 
	          sygDage(x) = sygDage(x) 'sygTimer(x) / normTimerDag(x) 
	          else 
	          sygDage(x) = 0
	          end if%> 
	           
	         <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(sygDage(x),2) %></td>

           
             <%end if %>
	         
	         
	         <%if instr(akttype_sel, "#21#") <> 0 then %>
	         <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(BarnsygTimer(x),2) %></td>
	         
	          <%if barnSygDage(x) <> 0 then 
	          barnSygDage(x) = barnSygDage(x) 'barnSygTimer(x) / normTimerDag(x) 
	          else 
	          barnSygDage(x) = 0
	          end if%> 
	           <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(barnSygDage(x),2) %></td>
	        
	         
	         <%end if %>
	        <%if instr(akttype_sel, "#8#") <> 0 then %>
	        <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(sundTimer(x),2) %></td>
	         <%end if %>
	        <%if instr(akttype_sel, "#81#") <> 0 then %>
	        <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(lageTimer(x),2) %></td>
	        <%end if %>
	        </tr>
	        </table>
	       
	        
	        <%end if%>
        	
        	
	         <br /><br />
	            <table cellspacing=0 cellpadding=2 border=0 width="100%">
        	 
	         <td style="border-bottom:1px silver dashed;">
                 <b><%=tsa_txt_265%></b> <br /><%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> 
                 
                 <%if year(slutDato) = year(now) AND month(slutDato) = month(now) then%>
	     (<%=tsa_txt_318 %>)
	     <%end if %>
                 </td>
        	 <%if len(trim(km(x))) <> 0 then
        	 km(x) = km(x)
        	 else
        	 km(x) = 0
        	 end if %>
	          <td align=right valign=bottom style="border-bottom:1px silver dashed;"><%=km(x) %> km.</td>
	         </tr>
	         </table>
	
	 
	 
	 
	 
	
	  <%
      
     
      case 5, 50 '** Status total på afstem_tot fra Timereg. JQuery
	
	  x = 0
	 
	
	     'Response.write "visning:" & visning
         'Response.end
	
	     ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)

          

	     call fLonTimerPer(ltimerStDato, dageDiff, 3, intMid)
	 
	     ltimerKorFrad = 0
	     ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
	
	     'call akttyper2009Prop(114) 'korrektion tælles med?

         'if cint(aty_real) = 1 then
	     realNormBal = formatnumber(((realTimer(x) + (korrektionReal(x))) - normTimer(x)),2)
         'else
         'realNormBal = formatnumber(((realTimer(x)) - normTimer(x)),2)
         'end if
  
	  
	     normtime_lontime = -((normTimer(x) - (ltimerKorFrad)) * 60) 
	     call timerogminutberegning(normtime_lontime)
        
         normLontBal = thoursTot &":"& left(tminTot, 2) 
         realLontBal = formatnumber((realTimer(x) - (ltimerKorFrad)),2)
        
         'if session("mid") = 1 AND lto = "esn" then
         'response.Write "<hr>stDato: "& ltimerStDato &", dageDiff: "& dageDiff &" = ltimerKorFrad: "& ltimerKorFrad &" =  ((totalTimerPer100: "& totalTimerPer100 &"/60) + (fradragTimer(x); "& fradragTimer(x) &")) realTimer(x): "& realTimer(x) &", korrektionReal(x) : "& korrektionReal(x) &" normTimer(x): "& normTimer(x) &"<br>"
         'response.Write "<br>normTimer(x): "& normTimer(x) &"ltimerKorFrad: "& ltimerKorFrad
         'response.write "<br>normtime_lontime: "& normtime_lontime & "<br>"
         'end if
	
      
      
      case 7, 14, 77 '**** afstem tot, år --> dato ***' / Ugesedlser 
	  
	  
	 'Response.write "akttype_sel:" & akttype_sel


     %>
     <!--#include file="global_func3_7_14_77_inc.asp" -->
   <%
	  
	  
	  
	  
	  
	         
	  
	  case 4 '**** Ferie ***'        
	         
	  %>
     <!--#include file="global_func3_4_inc.asp" -->
   <%
	 
	 
	 end select 
	 '** visning
	
	end function
	
	
	%>
	
	
	
	