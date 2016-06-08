
<%


    public afstemnul
	public meNormTimer, meNormTimerDag, meShowNormTimerdag, meNormTimerUge
    redim meNormTimer(1000), meNormTimerDag(1000), meShowNormTimerdag(1000), meNormTimerUge(1000), afstemnul(1000)
	
    public usedMeTypes 
    public strEksport, strEksportTxt, strEksportEmailTxt, headerwrt
	function medarbafstem(intMid, startDato, slutDato, visning, akttype_sel, m)
	
	
	x = m
	'Response.Write "akttype_sel" & akttype_sel
	
	select case visning 
	'case 5, 1
	'visTotalTimerAlltyp = 1
	case 1,2,3,4,5,7,14,6
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
	redim normTimerDag(1000)
	
	
	dim realTimer, medarbNavn, medarbNr, medarbInit, realIfTimer
	redim realTimer(1000), realIfTimer(1000), medarbNavn(1000), medarbNr(1000), medarbInit(1000)
	
	dim normTimerUge, normTimer
	redim normTimerUge(1000), normTimer(1000)
	
	 dim fakTimer, mfForbrug , resTimer
	 dim medarbEmail, fradragTimer, medarbId, realfTimer
	 
	 redim resTimer(1000), fakTimer(1000), mfForbrug(1000)
	 redim medarbEmail(1000), fradragTimer(1000), medarbId(1000), realfTimer(1000)

	 redim strEksport(1000)
	 
	
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
	
	'Response.Write "akttype_sel: " & akttype_sel
	
	 if (instr(akttype_sel, "#-1#") <> 0 AND visning = 1) OR visning = 2 OR visning = 3 OR visning = 5 OR visning = 7 OR visning = 14 then
	'** Timer alle der tæller med i dagligt timeregnskab **'
	strSQL = "SELECT t.tid, sum(t.timer) AS realtimer, m.mnavn, m.mnr, m.init, m.email, "_
	&" t.tdato, m.mid, m.medarbejdertype, tfaktim FROM medarbejdere m "_
	&" LEFT JOIN timer t ON (t.tmnr = m.mid AND ("& aty_sql_realhours &") "_
	&" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"') "_
	&" WHERE "& medarbSQL &" AND mansat <> '2' "_
	&" GROUP BY m.mid ORDER BY m.mnavn"
	
	'Response.Write strSQL &"<br>"
	'Response.flush
	
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
	
	end if
	
	
	
	'Response.Write oRec("tid") & " - " & oRec("tdato") & " - " & oRec("realtimer") & " - " & oRec("mnavn") & "<br>"
	    
	  
	    
	    
	    call meStamdata(intMid)
	    
	    medarbEmail(x) = meEmail 'oRec("email")
	    medarbNavn(x) = meNavn '& "("& meType &")" 'oRec("mnavn")
	    medarbNr(x) = meNr 'oRec("mnr")
	    medarbInit(x) = meInit 'oRec("init")
	    medarbId(x) =  intMid 'oRec("mid")
	    
	    
	    '*** Normeret periode (skal altid med til dags beregning under ferie, fefri, og afspad)  ***'
	    '** Beregner kun en gang for hver medarb. type, med mindre en medarbejder har skiftet type undervejs ***'
	    '** i det valgte interval, så må medarbejderen beregnes indiiduelt ***'
	    
	    '** Tjekker historik **'
	    'strSQLmth = "SELECT COUNT(mid) FROM medarbejdertyper_historik WHERE "_
        '&" mid = "& intMid &" GROUP BY mid"
        
        'meTypeHist = 0
        'oRec2.open strSQLmth, oConn, 3
        'if not oRec2.EOF then
        
        'meTypeHist = 1
      
        'end if
        'oRec2.close
	    
	     '''''ntimPer = 0

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
	             'Response.Write "stDato " & stDato & " dageDiff " & dageDiff
	             call normtimerPer(intMid, stDato, dageDiff)
	             normTimer(x) = ntimPer 
                 showNormTimerdag = normTimer(x) / antalDageMtimer
                 end if
        	    
	            
        	    
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
	    case 1,5,7,6,4,14
	    call fordelpaaaktType(intMid, startDato, slutDato, visning, akttype_sel, x)
	    'case 14
	    'call fordelpaaaktType(intMid, startDato, slutDato, visning, "#-99#, "& aktiveTyper, x)
	    end select
	    
	    
	    
	  
        'Response.Write "meType: " & meType & " normtimer: "& meNormTimer(meType) & " akttype_sel = "& akttype_sel &"<br>"
	   
	   '*********************************'
	   '*** Kalder sum på samle-typer ***'
	   '*********************************'
	   if (instr(akttype_sel, "#-1#") <> 0 AND visning = 1) OR visning = 2 OR visning = 3 OR visning = 5 OR visning = 7 OR visning = 14 then 
	   'if visning = 1 OR visning = 2 OR visning = 3 OR visning = 5 OR visning = 7 OR visning = 14 then
	    
	    
	     '*** Fradrag i løntimer ***'
	    
	    strSQLfradrag = "SELECT t.tid, sum(t.timer) AS fratimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& intMid &" AND ("& aty_sql_frawhours &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLfradrag & "<br><br>"
	   'Response.flush
	   
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
	    
	    
	    
	    
	    
	    'if visning = 1 OR visning = 2 OR visning = 3 OR visning = 7 OR visning = 14 then
	     if (instr(akttype_sel, "#-1#") <> 0 AND visning = 1) OR visning = 2 OR visning = 3 OR visning = 7 OR visning = 14 then
	  
	      
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
	   
	    
	
	
	
	
	
	
	
	
	end if
	
	
	
	
	
	
	'**** Præsentation ***'
	
	
	
	select case visning 
	case 1
	
	
	if cint(headerwrt) <> 1 then
	
	strEksportPer = "Periode afgrænsning: "& formatdatetime(stdato, 1) & " - "&  formatdatetime(sldato, 1) & "xx99123sy#z"
	strEksportTxt = strEksportTxt & strEksportPer
	'strEksportEmailTxt = strEksportEmailTxt & strEksportPer
		
	if instr(akttype_sel, "#-1#") <> 0 then
	cops = 8
	else
	cops = 0
	end if
	
	
	%>
	 <table cellspacing=1 cellpadding=2 border=0 bgcolor="#d6dff5"><!-- #5C75AA -->
	
	 <tr bgcolor="#5582d2">
	 <td class=alt valign="bottom"><b><%=tsa_txt_147%></b></td>
	 
	  <%if instr(akttype_sel, "#-5#") <> 0 AND stempelurOn = 1 then %>
	 <td class=alt valign="bottom" colspan=5><b>Løntimer</b><br />Stempelur<br />I valgt periode</td>
	 <%end if
	 
	 
	 if instr(akttype_sel, "#-1#") <> 0 then %>
	 <td class=alt valign="bottom" colspan=<%=cops %>><b><%=tsa_txt_148%></b> <br />
	 <%=tsa_txt_149%> <br />I valgt periode</td>
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
	 
	 <%if aktudspCP <> 0 then %>
	 <td class=alt colspan=<%=aktudspCP %>><b>Udspecificenring på akt. typer</b><br />
	 I valgt periode</td>
	 <%end if %>
	 
	 
	 
	 
	   <%if instr(akttype_sel, "#7#") <> 0 then %>
	<td class=alt valign="bottom"><b>Flex</b></td>
	 <%end if %>
	
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	 <td class=alt valign="bottom"><b><%=tsa_txt_265%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#10#") <> 0 then %>
	 <td class=alt valign="bottom"><b><%=tsa_txt_150%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td class=alt valign="bottom"><b>Pause</b></td>
	 <%end if %>
	 
	 
	 <%
	 copsff = 1
	 if instr(akttype_sel, "#12#") <> 0 then 
	 copsff = copsff + 1
	 end if
	 
	 if instr(akttype_sel, "#18#") <> 0 then 
	 copsff = copsff + 1
	 end if
	 
	 if instr(akttype_sel, "#13#") <> 0 then 
	 copsff = copsff + 2
	 end if
	 
	 if instr(akttype_sel, "#17#") <> 0 then 
	 copsff = copsff + 1
	 end if
	 
	 if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then %>
	<td class=alt valign=bottom colspan=<%=copsff %>><b><%=tsa_txt_174%></b><br />
	Ferieår: <%=ferieaarStart %>
    <br /><span style="font-size:10px; font-family:arial; color:#cccccc;"><%=formatdatetime(startdatoFerieLabel, 1)%> til<br /> <%=formatdatetime(slutdatoFerieLabel, 1) %></span></td>
	<%end if %>
	 
	 
	 
	  <%
	 copsFe = 1
	 if instr(akttype_sel, "#11#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#14#") <> 0 then 
	 copsFe = copsFe + 2
	 end if
	 
	 if instr(akttype_sel, "#19#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#15#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#16#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 then %>
	<td valign="bottom" class=alt colspan=<%=copsFe %>><b><%=tsa_txt_152%></b><br />
	Ferieår: <%=ferieaarStart %> 
    <br /><span style="font-size:10px; font-family:arial; color:#cccccc;"><%=formatdatetime(startdatoFerieLabel, 1)%> til <%=formatdatetime(slutdatoFerieLabel, 1) %></span></td>
	<%end if %>
	  
	 
	 
	 <%
	 copsAf = 1
	 if instr(akttype_sel, "#30#") <> 0 then 
	 copsAf = copsAf + 2
	 end if
	 
	 if instr(akttype_sel, "#31#") <> 0 then 
	 copsAf = copsAf + 2
	 end if
	 
	 if instr(akttype_sel, "#32#") <> 0 then 
	 copsAf = copsAf + 1
	 end if
	 
	 if instr(akttype_sel, "#33#") <> 0 then 
	 copsAf = copsAf + 1
	 end if
	 
	 if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR _
	 instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	<td class=alt colspan=<%=copsAf %>><b><%=tsa_txt_151 %> / <%=global_txt_145 %></b>
	<br /> Total fra licens-start d. <%=formatdatetime(lisStDato,1) %> (eller fra ansættelsesdato) - <%=formatdatetime(slutdato, 1) %><br />
	</td>
	<%end if %>
	 
	 
	 <%
	 '** Sygdom ***'
	 if instr(akttype_sel, "#20#") <> 0 then 
	 copsS = 5
	 else
	 copsS = 0
	 end if
	 
	 if instr(akttype_sel, "#21#") <> 0 then 
	 copsS = copsS + 4
	 else
	 copsS = copsS
	 end if
	 
	  %>
	 
	  <%if instr(akttype_sel, "#20#") <> 0 OR instr(akttype_sel, "#21#") <> 0 then %>
	 <td class=alt colspan=<%=copsS%>><b><%=tsa_txt_153%></b> (ÅTD)<br />
	 Fra <%=formatdatetime(sygaarSt, 1) %> til <%=formatdatetime(slutdato, 1) %> og <br />
	 I valgt periode
	 </td>
	 <%end if %>
	 

     <%if instr(akttype_sel, "#22#") <> 0 then %>
	 <td class=alt valign="bottom"><b>Barsel</b> dage</td>
	 <%end if %>


	  <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td class=alt valign="bottom"><b>Sundh.</b> timer</td>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td class=alt valign="bottom"><b>Læge</b> timer</td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td class=alt valign="bottom"><b><%=tsa_txt_154%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td class=alt valign="bottom"><b><%=tsa_txt_155%></b> timer</td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td class=alt valign="bottom"><b>Stk.</b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	  <td class=alt valign="bottom"><b><%=tsa_txt_156%></b></td>
	  <%end if %>
	 </tr>
	 
	 
	 
	 
	 
	 
	 
	 <!-- Under overskrrifter --> <!--- Akt type navn -->
	 
	 <tr bgcolor="#8CAAe6">
	 
	 <td style="border-right:1px #cccccc solid;" class=alt>&nbsp;</td>
	 
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_147 & "; Nr.;Init;" %>
	 
	  <%if instr(akttype_sel, "#-5#") <> 0 AND stempelurOn = 1 then %>
	 <td class=alt_lille  valign=bottom><%=tsa_txt_148 %>:<%=tsa_txt_141 %></td>
	 <td class=alt_lille  valign=bottom>Till./Frad. +/-</td>
	 <td class=lille  valign=bottom bgcolor="#cccccc">Sum</td>
	 <td class=lille  valign=bottom bgcolor="#EFF3FF">Bal Lønt. / Real</td>
	 <td class=lille  valign=bottom bgcolor="#DCF5BD">Bal Lønt. / Norm</td>
	 
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_148 &":"& tsa_txt_141 &" (Omregnet til kommatal for kalkulation i excel);Tillæg/Fradrag +/- ;Sum;Balance Lønt./Real;Balance Lønt./Norm;" %>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-1#") <> 0 then %>
	 <td class=lille  valign=bottom bgcolor="#cccccc" style=" "><%=tsa_txt_157%></td>
	 <td class=lille  valign=bottom bgcolor="#cccccc"><%=tsa_txt_172 &" "& tsa_txt_179%></td>
	 <td class=lille  valign=bottom bgcolor="#DCF5BD"><%=tsa_txt_158%></td>
	 <td class=lille  valign=bottom bgcolor="#DCF5BD"> ~ <%=tsa_txt_259%></td>
	 <td class=lille  valign=bottom bgcolor="pink"><%=tsa_txt_159%></td>
	 <td class=lille  valign=bottom bgcolor="#cccccc"><%=tsa_txt_160%></td>
	 <td class=lille  valign=bottom bgcolor="#cccccc"><%=tsa_txt_161%></td>
	 <td class=lille  valign=bottom bgcolor="#cccccc" style=""><%=tsa_txt_163%></td>
	 
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_157 & ";"& tsa_txt_172 &" "& tsa_txt_179 &";" & tsa_txt_158 & ";"& tsa_txt_259 & ";" & tsa_txt_159 & ";"& tsa_txt_160 & ";" & tsa_txt_161 & ";" & tsa_txt_163 & ";"  %>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#1#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(1, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(2, 1) %>
	 <%=akttypenavn%>
	 </td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	<%end if %>
	
	<%if instr(akttype_sel, "#50#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(50, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>


     <%if instr(akttype_sel, "#54#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(54, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#51#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(51, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(52, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#53#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(53, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  
	 
	  <%if instr(akttype_sel, "#55#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(55, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(60, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#90#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(90, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(91, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#6#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(6, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#7#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(7, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="">
	 <%call akttyper(5, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#10#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(10, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(9, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 
	 <!--Ferie Fridage -->
	 <%if instr(akttype_sel, "#12#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom>
	 <%call akttyper(12, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	<!-- <td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#18#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom>
	 <%call akttyper(18, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %> <br> >> dd.</td>
	<!-- <td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom>
	 <%call akttyper(13, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	<!-- <td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#17#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom>
	 <%call akttyper(17, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 
	 
	 
	 <%if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then %>
	 <td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style=""><%=tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_275%></td>
	 <!--<td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style=""><%=tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_276%></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom>
	 <%call akttyper(13, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %><br /> i periode</td>
	<!-- <td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &" i periode;"%>
	 <%end if %>
	 
	
	 
	 <!-- Ferie -->
	 <%if instr(akttype_sel, "#15#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom>
	 <%call akttyper(15, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#11#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom>
	 <%call akttyper(11, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %> <br> >> dd.</td>
	 <!--<td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom>
	 <%call akttyper(14, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#19#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom>
	 <%call akttyper(19, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#16#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom>
	 <%call akttyper(16, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <% if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 then %>
	  <td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style=""><%=tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_275%></td>
	 <!--<td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style=""><%=tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_276%></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader &" "& tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 
	  <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td bgcolor="#EFf3ff"  class=lille valign=bottom>
	 <%call akttyper(14, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %><br /> i periode</td>
	 <!--<td class=alt_lille  valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" "& tsa_txt_275 &" i periode;"%>
	 <%end if %>
	 
	 
	 
	 <!-- Afspad --->
	 <%if instr(akttype_sel, "#30#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(30, 1) %>
	 <%=akttypenavn%><br />(Tim.) Enh.</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" Tim. ; " & akttypenavn &" Enh. ;"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(31, 1) %>
	 <%=akttypenavn%>.</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &".;"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#32#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(32, 1) %>
	 <%=akttypenavn%>.</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &".;"%>
	 <%end if %>
	 
	<%if instr(akttype_sel, "#33#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(33, 1) %>
	 <%=akttypenavn%></td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	 <td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style=""><%=tsa_txt_283 &" "& tsa_txt_280%></td>
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_283 &" "& tsa_txt_280 &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#30#") <> 0 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom>
	 <%call akttyper(30, 1) %>
	 <%=akttypenavn%><br />(Tim.) Enh. i periode</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" Timer i periode ; " & akttypenavn &" Enheder i periode ;"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom>
	 <%call akttyper(31, 1) %>
	 <%=akttypenavn%><br /> timer i periode</td>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" Timer i periode ;"%>
	 <%end if %>
	 
	 
	
	 <!-- Sygdom -->
	  <%if instr(akttype_sel, "#20#") <> 0 then %>
	     <td class=alt_lille  valign=bottom>
	     <%call akttyper(20, 1) %>
	     <%=akttypenavn%> tim.
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer;"&  akttypenavn &" dage;"%>
	     </td>
	      <td class=alt_lille  valign=bottom><%=akttypenavn%> dage</td>
    	  
    	  
	     <td bgcolor="#EFf3FF" class=lille  valign=bottom>
    	 
	     <%=akttypenavn%> tim. i per.
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer i per.;"&  akttypenavn &" dage i per.;"%>
	     </td>
	      <td bgcolor="#EFf3FF" class=lille  valign=bottom><%=akttypenavn%> dage i per</td>
    	  
	    <%end if %>
	 
	  <%if instr(akttype_sel, "#21#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(21, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"&  akttypenavn &" dage;"%>
	 </td>
	  <td class=alt_lille  valign=bottom><%=akttypenavn%> dage</td>
    	 
	    <td bgcolor="#EFf3FF" class=lille  valign=bottom>
    	 
	     <%=akttypenavn%> tim. i per.
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" timer i per.;"&  akttypenavn &" dage i per.;"%>
	     </td>
	      <td bgcolor="#EFf3FF" class=lille  valign=bottom><%=akttypenavn%> dage i per</td>
    	  
	 
	 <%end if %>

	 
	 <%if instr(akttype_sel, "#20#") <> 0 then %>
	    <td bgcolor="lightpink" class=lille  valign=bottom>
     Fravær i periode i % *<br />
   
	      <%strEksportTxtHeader = strEksportTxtHeader &" Fravær i per. i %;"%>
	     </td>
	 <%end if %>


     <!-- Barsel -->
	  <%if instr(akttype_sel, "#22#") <> 0 then %>
	     <td class=alt_lille  valign=bottom>
	     <%call akttyper(22, 1) %>
	     <%=akttypenavn%> 
	      <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &" dage;"%>
	     </td>
	  
      <%end if %>

	 
	 <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(8, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(81, 1) %>
	 <%=akttypenavn%>
	  <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td class=alt_lille  valign=bottom><%=tsa_txt_154%></td>
	  <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_154 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td class=alt_lille  valign=bottom><%=tsa_txt_155%>
	 <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_155 &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td class=alt_lille  valign=bottom>
	 <%call akttyper(61, 1) %>
	 <%=akttypenavn%>
	 <%strEksportTxtHeader = strEksportTxtHeader & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	  <td style="width:90px;" class=alt_lille  valign=bottom><%=tsa_txt_171%></td>
	  <%strEksportTxtHeader = strEksportTxtHeader & tsa_txt_171 &";"%>
	  <%end if %>
	 </tr>
	 
	 <%
	 
     strEksportTxtHeader = strEksportTxtHeader & ";Startdato;Slutdato;"

	 strEksportTxt = strEksportTxtHeader
	 headerwrt = 1
	 end if 'x = 0 overskrifter
	 
	 
	 
	 
	 strEksportTxt = strEksportTxt & strEksport(x)
	 strEksportEmailTxtHeader = strEksport(x)
	 
	 
	 
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
	 <td class=lille style=" white-space: nowrap;"><b><%=medarbNavn(x) %></b> (<%=medarbNr(x) %>) 
	 <%if len(trim(medarbInit(x))) <> 0 then %>
	  - <%=medarbInit(x) %>
	  <%end if %>

	  </td>
	 
	 <%strEksport(x) = strEksport(x) & medarbNavn(x) &"; "& medarbNr(x) &";"& medarbInit(x) &";"%>
	 
	  <%if instr(akttype_sel, "#-5#") <> 0 AND  stempelurOn = 1 then %>
	 <td align=right style=" white-space: nowrap;" class=lille>
	 <%
	 
	 ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	 call fLonTimerPer(ltimerStDato, dageDiff, 2, medarbId(x)) 
	 strEksport(x) = strEksport(x) & thoursTot &","& left((tminTot*1.67), 2) &";"
	 %> 
	 
	</td>
	
	  <%
	 fraDtimer = formatnumber(fradragTimer(x),2)
	 %>
	  <td align=right style=" white-space: nowrap;" class=lille>
	  <%call timerogminutberegning(fraDtimer*60) %>
	  <%=thoursTot &":"& left(tminTot, 2) %> </td>
	  
	  <%strEksport(x) = strEksport(x) & formatnumber(lontimerTotMfradrag, 2) &";"%>
	
	
	
	<%
	lontimerTotMfradrag = ((totalTimerPer100/60) + (fradragTimer(x)))
    %>
	  <td align=right style=" white-space: nowrap;" class=lille>
	  <%call timerogminutberegning(lontimerTotMfradrag*60) %>
	  <%=thoursTot &":"& left(tminTot, 2) %> </td>
	  
	  <%strEksport(x) = strEksport(x) & formatnumber(lontimerTotMfradrag, 2) &";"%>
	  
	  
	  <%
	 bal_lon_real = formatnumber(lontimerTotMfradrag,2) - formatnumber(realTimer(x),2)
	 %>
	  <td align=right style=" white-space: nowrap;" class=lille>
	  <%call timerogminutberegning(bal_lon_real*60) %>
	  <%=thoursTot &":"& left(tminTot, 2) %> </td>
	  
	  <%strEksport(x) = strEksport(x) & formatnumber(bal_lon_real, 2) &";"%>
	  
	  
	 <td align=right style=" white-space: nowrap; " class=lille>
	  <%
	  normtime_lontime = -((normTimer(x) - (lontimerTotMfradrag)) * 60)
	  
	  if normtime_lontime <> 0 then
	  normtime_lontime = cdbl(normtime_lontime)
	  else
	  normtime_lontime = normtime_lontime
	  end if
	  call timerogminutberegning(normtime_lontime) %>
	  <%=thoursTot &":"& left(tminTot, 2) %></td>
	  
	  <%strEksport(x) = strEksport(x) & formatnumber(normtime_lontime/60, 2) &";"%>
	  
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#-1#") <> 0 then %>
	 <td align=right style=" white-space: nowrap;" class=lille><b><%=formatnumber(realTimer(x),2)%></b></td>
	 <td align=right style=" white-space: nowrap;" class=lille><%=formatnumber(realTimer(x)/weekDiff,2)%></td>
	 <td align=right style=" white-space: nowrap;" class=lille><%=formatnumber(normTimer(x),2)%></td>
	 
	 <%strEksport(x) = strEksport(x) & formatnumber(realTimer(x),2) &";" & formatnumber(realTimer(x)/weekDiff,2) & ";"& formatnumber(normTimer(x),2) &";"%>
	
	 
	 <td align=right class=lille style=" white-space: nowrap;">
	 <%if showNormTimerdag > 0 then %>
	 <%=formatnumber(showNormTimerdag, 1) %>
	 <%strEksport(x) = strEksport(x) & formatnumber(showNormTimerdag, 1) & ";" %>
	 <%else%>
	 <%=formatnumber(0, 2) %>
	 <%strEksport(x) = strEksport(x) & formatnumber(0, 2)  & ";" %>
	 <%end if %>
	 
	  
	 </td>
	 
	 <td align=right class=lille style="  white-space: nowrap;"><b><%=formatnumber((realTimer(x) - normTimer(x)),2)%></b></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(realfTimer(x),2)%></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(realIfTimer(x),2)%></td>
	 
	  <%
	  if realTimer(x) <> 0 AND realfTimer(x) <> 0 then
       
        iebal = (realfTimer(x)/realTimer(x)) * 100 
        
	  else
	  iebal = 0
	  end if%>
	  
	  <%strEksport(x) = strEksport(x) & formatnumber((realTimer(x) - normTimer(x)),2) &";"& formatnumber(realfTimer(x),2) &";"& formatnumber(realIfTimer(x),2) & ";"& formatnumber(iebal,2) & ";" %>
	  
	  <td align=right class=lille style="  white-space: nowrap;"><%=formatnumber(iebal,2)%>%</td>
	   <%end if %>
	 
	 <%if instr(akttype_sel, "#1#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fTimer(x),2) %></td>
	 <% strEksport(x) = strEksport(x) & formatnumber(fTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ifTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(ifTimer(x),2) & ";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#50#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(dagTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(dagTimer(x),2) & ";"%>
	 <%end if %>

      <%if instr(akttype_sel, "#54#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(aftenTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(aftenTimer(x),2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#51#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(natTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(natTimer(x),2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(weekendTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(weekendTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#53#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(weekendnatTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(weekendnatTimer(x),2) & ";"%>
	 <%end if %>
	 
	
	 <%if instr(akttype_sel, "#55#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(aftenweekendTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(aftenweekendTimer(x),2) & ";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(adhocTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(adhocTimer(x),2) & ";"%>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#90#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(e1Timer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(e1Timer(x),2) & ";"%>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(e2Timer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(e2Timer(x),2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#6#") <> 0 then %>
	  <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sTimer(x),2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(sTimer(x),2) & ";"%>
	 <%end if %>
	  
	  <%if instr(akttype_sel, "#7#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(flexTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(flexTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	   <td align=right class=lille style=" white-space: nowrap; "><%=formatnumber(km(x),2) %></td>
	  <%strEksport(x) = strEksport(x) & formatnumber(km(x),2) & ";"%>
	  <%end if %>
	  
	  <%if instr(akttype_sel, "#10#") <> 0 then %>
	  <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fpTimer(x),2) %></td>
	  <%strEksport(x) = strEksport(x) & formatnumber(fpTimer(x),2) & ";"%>
	  <%end if %>
	  
	  <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(pausTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(pausTimer(x),2) & ";"%>
	 <%end if %>
	 
	 
	 
	  <!--Ferie Fridage -->
	 <%if instr(akttype_sel, "#12#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	
	 <%if normTimerDag(x) <> 0 then 
	 fefriVal = fefriTimer(x)/normTimerDag(x)
	 else
	 fefriVal = 0
	 end if
	 %>
	 <%=formatnumber(fefriVal,2) %>
	 
	 </td>
	 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriVal,2) & ";"%>
	 <%end if %>
	 
	 
	  <%if instr(akttype_sel, "#18#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	
	 <%if normTimerDag(x) <> 0 then 
	 fefriplVal = fefriplTimer(x)/normTimerDag(x)
	 else
	 fefriplVal = 0
	 end if
	 %>
	 <%=formatnumber(fefriplVal,2) %>
	 
	 </td>
	 
	 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriplVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	  <%if normTimerDag(x) <> 0 then 
	 fefribrVal = fefriTimerBr(x)/normTimerDag(x)
	 else
	 fefribrVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(fefribrVal,2) %> </td> 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimerBr(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefribrVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#17#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 fefriudbVal = fefriTimerUdb(x)/normTimerDag(x)
	 else
	 fefriudbVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(fefriudbVal,2) %></td> 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimerUdb(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriudbVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then 
	 fefriBal = 0
	 fefriBal  = (fefriTimer(x) - (fefriTimerBr(x) + fefriTimerUdb(x)))
	 
	 if normTimerDag(x) <> 0 then 
	 fefriBalVal = fefriBal/normTimerDag(x)
	 else
	 fefriBalVal = 0
	 end if
	 
	 %>
	 <td align=right class=lille style="  white-space: nowrap;">
	 <%=formatnumber(fefriBalVal,2) %></td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriBal,2)%></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriBalVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	  <%if normTimerDag(x) <> 0 then 
	 fefriPerbrVal = fefriTimerPerBr(x)/normTimerDag(x)
	 else
	 fefriPerbrVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(fefriPerbrVal,2) %></td> 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimerBr(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriPerbrVal,2) & ";"%>
	 <%end if %>
	  
	
	 <!-- Ferie -->
	 
	  <%if instr(akttype_sel, "#15#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieVal = ferieOptjtimer(x)/normTimerDag(x)
	 else
	 ferieVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(ferieVal,2) %></td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieOptjtimer(x),2) %></td>-->
	  <%strEksport(x) = strEksport(x) & formatnumber(ferieVal,2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#11#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 feriePlVal = feriePLTimer(x)/normTimerDag(x)
	 else
	 feriePlVal = 0
	 end if
	 %>
	 
	 
	 <%=formatnumber(feriePlVal,2) %></td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(feriePLTimer(x),2) %></td>-->
	  <%strEksport(x) = strEksport(x) & formatnumber(feriePlVal,2) & ";"%>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieAftVal = ferieAFTimer(x)/normTimerDag(x)
	 else
	 ferieAftVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(ferieAftVal,2) %> </td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieAFTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(ferieAftVal,2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#19#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieAftulonVal = ferieAFulonTimer(x)/normTimerDag(x)
	 else
	 ferieAftulonVal = 0
	 end if
	 %>

	 <%=formatnumber(ferieAftulonVal,2) %> </td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieAFTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(ferieAftulonVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#16#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieUdbVal = ferieUdbTimer(x)/normTimerDag(x)
	 else
	 ferieUdbVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(ferieUdbVal,2) %></td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieUdbTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(ferieUdbVal,2) & ";"%>
	 <%end if %>
	 
	 <% if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 then 
	 
	 ferieBal = 0
     select case lto
     case "cst"
     ferieBal  = (ferieOptjtimer(x) - (ferieAFTimer(x) + ferieAFulonTimer(x) + ferieUdbTimer(x)))
     case "wwf"
	 ferieBal  = (ferieOptjtimer(x) - (ferieAFTimer(x) + ferieUdbTimer(x) - (ferieAFulonTimer(x))))
     case else
     ferieBal  = (ferieOptjtimer(x) - (ferieAFTimer(x) + ferieUdbTimer(x)))
     end select
	 
	 if normTimerDag(x) <> 0 then 
	 ferieBalVal = ferieBal/normTimerDag(x)
	 else
	 ferieBalVal = 0
	 end if
	 
	 %>
	 <td align=right class=lille style=" white-space: nowrap; ">
	 <%=formatnumber(ferieBalVal,2) %></td>
	<!-- <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieBal,2)%></td>-->
	<%strEksport(x) = strEksport(x) & formatnumber(ferieBalVal,2) & ";"%>
	<%end if %>
	 
	 <!-- Ferie i periode -->
	  <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieAftPerVal = ferieAFPerTimer(x)/normTimerDag(x)
	 else
	 ferieAftPerVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(ferieAftPerVal,2) %> </td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieAFTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(ferieAftPerVal,2) & ";"%>
	 <%end if %>
	 
	 
     <!-- Afspad --->
	 <%if instr(akttype_sel, "#30#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">(<%= formatnumber(afspTimerTim(x), 2) %>) <%=formatnumber(afspTimer(x), 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspTimerTim(x), 2) &";"& formatnumber(afspTimer(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(afspTimerBr(x), 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspTimerBr(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#32#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(afspTimerUdb(x), 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspTimerUdb(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#33#") <> 0 then 
	 afspadUdbBal = 0
	 afspadUdbBal = (afspTimerUdb(x) - afspTimerOUdb(x)) %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(afspadUdbBal, 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspadUdbBal, 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR _
	 instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	 
	 <%AfspadBal = 0 
	 
	 AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	 %>
	 <td align=right class=lille style=" white-space: nowrap; "><%=formatnumber(AfspadBal, 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(AfspadBal, 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#30#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">(<%=formatnumber(afspTimerTimPer(x), 2) %>) <%=formatnumber(afspTimerPer(x), 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspTimerTimPer(x), 2) &";"& formatnumber(afspTimerPer(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(afspTimerBrPer(x), 2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspTimerBrPer(x), 2) &";"%>
	 <%end if %>
	 
	 
	 
	 <!-- Syg -->
	  
	  <%if instr(akttype_sel, "#20#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sygTimer(x),2) %></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sygDage(x),2) %></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sygTimerPer(x),2) %></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sygDagePer(x),2) %></td>
	
	 <%strEksport(x) = strEksport(x) & formatnumber(sygTimer(x), 2) &";" & formatnumber(sygDage(x), 2) & ";"& formatnumber(sygTimerPer(x), 2) &";" & formatnumber(sygDagePer(x), 2) & ";"%>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#21#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(barnsygTimer(x),2) %></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(barnSygDage(x),2) %></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(barnSygTimerPer(x),2) %></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(barnsygDagePer(x),2) %></td>
	 
	 
	 <%strEksport(x) = strEksport(x) & formatnumber(BarnsygTimer(x), 2) &";" & formatnumber(BarnsygDage(x), 2) &";" & formatnumber(BarnsygTimerPer(x), 2) &";" & formatnumber(BarnsygDagePer(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#20#") <> 0 then 
	 if normTimer(x) <> 0 then
	 fravaeriper = ((BarnsygTimerPer(x) + sygTimerPer(x)) / (normTimer(x) + fefriTimerPerBr(x) + ferieAFPerTimer(x)) ) * 100
	 else 
	 fravaeriper = 0
	 end if
	 %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fravaeriper,2) %> %</td>
	 <%strEksport(x) = strEksport(x) & formatnumber(fravaeriper, 2) &";"%>
	 <%end if%> 
	 

      <%if instr(akttype_sel, "#22#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(barsel(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(barsel(x), 2) &";"%>
	 <%end if %>
	 
	 
	  <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sundTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(sundTimer(x), 2) &";"%>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(lageTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(lageTimer(x), 2) &";"%>
	 <%end if %>
	 
	 <!-- Ressource -->
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(resTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(resTimer(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fakTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(fakTimer(x), 2) &";"%>
	 <%end if %> 
	  
	  <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(stkAntal(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(stkAntal(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(mfForbrug(x)) & " DKK" %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(mfForbrug(x), 2) &";"%>
	 <%end if %>
	 </tr>
	 
	 
	 
	 
	 <%

     strEksport(x) = strEksport(x) & ";"&startdato&";"&slutdato&";"
	 
	 strEksportTxt = strEksportTxt & strEksport(x)
	 strEksportEmailTxt = strEksportEmailTxtHeader & strEksport(x)
	 
	 if sendemail = "j" then
	 
	 
	 
	 
	        '***** Oprettter Mail object ***
	        '**** Skal kun sende en linie til hver medarbejder ***'
	        '**** De må ikke se hinandens ****'
			if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\bal_real_norm_2007.asp" then
					
			
			
                txtEmailHtml = replace(strEksportEmailTxt, "xx99123sy#z", "</tr><tr><td class=lille align=right valign=bottom>")
			    txtEmailHtml = replace(txtEmailHtml, ";", "</td><td class=lille align=right valign=bottom>")
			   
				
				Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			    ' Sætter Charsettet til ISO-8859-1 
			    Mailer.CharSet = 2
			    ' Afsenderens navn 
			    Mailer.FromName = "TimeOut Email Service"
			    ' Afsenderens e-mail 
			    Mailer.FromAddress = "timeout_no_reply@outzource.dk"
			    Mailer.RemoteHost = "webmail.abusiness.dk"
    			
			    ' Mailens emne
			    Mailer.Subject = "TimeOut Afstemningsrapport"
			    ' Modtagerens navn og e-mail
			    Mailer.AddRecipient medarbNavn(x), medarbEmail(x)
			    Mailer.ContentType = "text/html"
			    
			    'Mailer.AddAttachment "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\"& file 
			    
					' Selve teksten
					Mailer.BodyText = "<html><head><title></title>"_
					&"<LINK rel=""stylesheet"" type=""text/css"" href=""https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/style/timeout_style_print.css""></head>"_
					&"<body><br>Afstemningsrapport<br>"&replace(strEksportPer, "xx99123sy#z", "<br>")&"<table width=100% cellpadding=1 cellspacing=0 border=1><tr><td class=lille valign=bottom>"_ 
					& txtEmailHtml & "</table><br><br>Med venlig hilsen<br>" & session("user") & "</body></html>"
					
					
				
					If Mailer.SendMail Then
    				
					Else
    					Response.Write "Fejl...<br>" & Mailer.Response
  					End if
					
			
			end if
		
		
		strEksportEmailTxt = ""
		
	 
	 end if 'send email
	 
	 end if 'nulfilter
	 
	 
	 next%>
	 <!-- </table> -->
	 <%
	 
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
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Løn timer</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Till. / <br />Frad. +/-</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Sum</b></td>
	 <%end if %>
	
	  <%if lto <> "cst" then %>
	 <td bgcolor="pink" align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b> (Real / Norm)</td>
	 <%end if %>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td bgcolor="#DCF5BD" align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b> (Lønt. / Norm)</td>
	 
	     <%if lto <> "cst" then %>
	     <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_166%></b> (Real / Lønt.)</td>
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
	 <%
	 end if %>
	 
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
	         %>
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
	 case 6      
	          
	          
	         
	          
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
             <%if level = 1 OR (session("mid") = usemrn) then %>
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
	          
	          <%if normTimerDag(x) <> 0 then 
	          sygDage(x) = sygTimer(x) / normTimerDag(x) 
	          else 
	          sygDage(x) = 0
	          end if%> 
	           
	         <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(sygDage(x),2) %></td>
	         <%end if %>
	         
	         
	         <%if instr(akttype_sel, "#21#") <> 0 then %>
	         <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(BarnsygTimer(x),2) %></td>
	         
	          <%if normTimerDag(x) <> 0 then 
	          barnSygDage(x) = barnSygTimer(x) / normTimerDag(x) 
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
	 
	 case 5 '** Status total på afstem_tot fra Timereg.
	
	 x = 0
	 
	 %>
	 
	 <b><%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %></b> (<%=tsa_txt_319 %>)
	 <table cellspacing=0 cellpadding=2 border=0 width=100%>
	 <tr bgcolor="#EFF3FF">
	 <%
	 select case lto
	 case "x"
	 %>
	 <td style="border-bottom:1px silver dashed;">&nbsp;</td>
	 <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_172%></b></td>
	  <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_173%></b></td>
	  
	  <%if session("stempelur") <> 0 then %> 
	 <td class=lille style="border-bottom:1px silver dashed;"><b>Løn timer</b></td>
	 <td class=lille style="border-bottom:1px silver dashed;"><b>Till. / <br />Frad. +/-</b></td>
	 <td class=lille style="border-bottom:1px silver dashed;"><b>Sum</b></td>
	 <%end if %>
	  
	 <%end select%> 
	  
	  <%if lto <> "cst" then %>
	  <td bgcolor="pink" class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b><br /> (Real / Norm)</td>
	  <%end if %>
	  
	 
	  <%if session("stempelur") <> 0 then %> 
	 <td bgcolor="#DCF5BD" class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b><br /> (Lønt. / Norm)</td>
	 
	  <%if lto <> "cst" then %>
	  <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_166%></b><br /> (Real / Lønt.)</td>
	  <%end if %>  
	    
	 <%end if %>
	 </tr>
	 
	 
	 <tr>
	 <%
	 
	
	 ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	 call fLonTimerPer(ltimerStDato, dageDiff, 3, intMid)
	 
	 ltimerKorFrad = 0
	 ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
	
	 select case lto
	 case "x"
	 %>
	 
	 <td class=lille style="border-bottom:1px silver dashed;">
      &nbsp;</td>
	
	<td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(realTimer(x),2)%></td>
	<td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimer(x),2)%></td>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%call timerogminutberegning(totalTimerPer100 ) %>
    <%=thoursTot &":"& left(tminTot, 2) %>
	</td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><%=formatnumber(fradragTimer(x), 2) %></td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><%=formatnumber(ltimerKorFrad, 2) %></td>
	 <%
	 end if %>
	 
	 
	 <%end select%> 
	 
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
	</table>
	
	       
	         <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	           <br /><br />
	          <table cellspacing=0 cellpadding=2 border=0 width=100%>
	           <tr>
	         <td colspan=6><b>Overarbejde</b><br />
	         <%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> (<%=tsa_txt_319 %>)</td>
	         </tr>
	         <tr bgcolor="#EFF3FF">
	        <td  class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b>Overarbejde <br /><%=tsa_txt_164%></b><br />(enh.)</td>
	          <td  style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b>Afspads.</b></td>
	           <td  style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b></td>
	           <td  style="border-bottom:1px silver dashed;" class=lille><b>Ønsk. Udbe.</b></td>
	           <td  style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_283 &" "& tsa_txt_280 %></b></td>
	         </tr>
	          <tr>
        	 
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(afspTimer(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerBr(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerUdb(x), 2)%></td>
	     
	          <% 
	         'afspadUdbBal = 0
	         'afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	         afspadUdbBal = 0
	         afspadUdbBal = (afspTimerUdb(x) - afspTimerOUdb(x))
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspadUdbBal, 2)%></td>
	          <% 
	         AfspadBal = 0 
	         AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(AfspadBal, 2)%></td>
        	 
	         </tr>
	         </table>
	         <br /><br />
	         <%end if %>
	         
	         
	  <%case 7, 14 '**** afstem tot, år --> dato ***' / Ugesedlser 
	  
	  
	   'x = 0
	   
	   '**** Beregninger til min/maks kriterie *****
	   ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	   call fLonTimerPer(ltimerStDato, dageDiff, 3, intMid)
	   
	   
	   ltimerKorFrad = 0
	   ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
	   normtime_lontime = -((normTimer(x) - (ltimerKorFrad)) * 60)
	   
	   if len(trim(normtime_lontime)) <> 0 then
	   normtime_lontime = normtime_lontime
	   else
	   normtime_lontime = 0
	   end if 
	   
	   balRealLontimer = realTimer(x) - (ltimerKorFrad)
	   if len(trim(balRealLontimer)) <> 0 then
	   balRealLontimer = balRealLontimer
	   else
	   balRealLontimer = 0
	   end if
	   
	   balRealNormtimer = (realTimer(x) - normTimer(x))
	   if len(trim(balRealNormtimer)) <> 0 then
	   balRealNormtimer = balRealNormtimer
	   else
	   balRealNormtimer = 0
	   end if
	    
	   showthisMedarb = 1 
	   
	   if visning = 14 then
	   '** ekstra søgekriterier ***'
	   call ekstrasogKri(useSogKri,moreorless,timeKri,saldoKri,visning)
	   end if
	   
	   if cint(showthisMedarb) = 1 then
	   
	   if len(trim(bgc)) <> 0 then
	   bgc = bgc
	   else
	   bgc = 0
	   end if
	   
	   select case right(bgc, 1)
	   case 0,2,4,6,8
	   bgthis = "#FFFFFF"
	   case else
	   bgthis = "#EFf3ff"
	   end select
	   
	   bgc = bgc + 1
	   
	   %>
	    <!-- st -->
	    <tr bgcolor="<%=bgthis %>">
	 <%
	 
	 strEksportTxt = strEksportTxt & "xx99123sy#z" & meNavn & ";" & meNr & ";"& meInit & ";"
	  
	   
	    if visning = 7 then%>
	    <td style="border-bottom:1px silver dashed; white-space:nowrap;"><b><%=monthname(datepart("m", startDato,2,2)) %></b></td>
	    <%
	    strEksportTxt = strEksportTxt & datepart("m", startDato,2,2) & ";"
	    else%>
	    <td style="border-bottom:1px silver dashed;" class=lille>
	         <%if media <> "print" then %>
	        <a href="weekpage_2010.asp?medarbid=<%=intMid %>&st_dato=<%=stDato%>&func=us" class="rmenu" target="_blank"><%=datepart("ww", startDato,2,2) %></a>
	        <%else %>
	        <%=datepart("ww", startDato,2,2) %>
	        <%end if %>
	        
	        - <%=datepart("d", startDato,2,2) & ". " & left(monthname(datepart("m", startDato,2,2)), 3) %>
	        
	        </td>
	    
	    <%
	    strEksportTxt = strEksportTxt & datepart("ww", startDato,2,2) & ";"
	    end if
	 
	 
	 
	 
	 %>
	 
	
	
	<td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(realTimer(x),2)%></td>
    <%strEksportTxt = strEksportTxt & formatnumber(realTimer(x),2) & ";" %>
     <%if lto <> "cst" then %>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">(<%=formatnumber(realfTimer(x),2)%>)</td>
	 <%arealfTimerTot = arealfTimerTot + (realfTimer(x)) %>
	 <%strEksportTxt = strEksportTxt & formatnumber(realfTimer(x),2) & ";" %>
	 <%end if %> 
    
    <td bgcolor="#DCF5BD" align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimer(x),2)%></td>
    <%strEksportTxt = strEksportTxt & formatnumber(normTimer(x),2) & ";" %>
	
	<%arealTimerTot = arealTimerTot + (realTimer(x)) %>
	<%anormTimerTot = anormTimerTot + (normTimer(x)) %>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%call timerogminutberegning(totalTimerPer100) %>
	 <%if visning = 14 then %>
	 
	 <%if media = "print" then %>
	 <%=thoursTot &":"& left(tminTot, 2) %>
	 <%else %>
        <a href="weekpage_2010.asp?medarbid=<%=intMid %>&st_dato=<%=stDato%>&func=lo" class="rmenu" target="_blank"><%=thoursTot &":"& left(tminTot, 2) %></a>
	 <%end if %>
	 <%else %>
	 <%=thoursTot &":"& left(tminTot, 2) %> 
	 <%end if %>
	 <%strEksportTxt = strEksportTxt & thoursTot &"."& left(tminTot, 2) & ";" %>
    
	</td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><%=formatnumber(fradragTimer(x), 2) %></td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><%=formatnumber(ltimerKorFrad, 2) %></td>
	 <%
	 strEksportTxt = strEksportTxt & formatnumber(fradragTimer(x), 2) & ";"& formatnumber(ltimerKorFrad, 2) & ";"
	 
	 
	 atotalTimerPer100 = atotalTimerPer100 + (totalTimerPer100)
	 afradragTimerTot = afradragTimerTot + (fradragTimer(x))
	 altimerKorFradTot = altimerKorFradTot + (ltimerKorFrad)
	 
	 end if %>
	 
	 
	 <%if lto <> "cst" then %>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber(balRealNormtimer,2)%></b></td>
	 <%strEksportTxt = strEksportTxt & formatnumber(balRealNormtimer, 2) & ";" %>
	 
	  <%
	 end if %>
	  

	  
	  <%if session("stempelur") <> 0 then %> 
	 <td bgcolor="#DCF5BD" align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%call timerogminutberegning(normtime_lontime) %>
		<b><%=thoursTot &":"& left(tminTot, 2) %></b>
	 </td>
	 <%strEksportTxt = strEksportTxt & thoursTot &"."& left(tminTot, 2) & ";" %>
	 
	 
	    <%if lto <> "cst" then %>
	   <td align=right style="border-bottom:1px silver dashed;" class=lille> <%=formatnumber(balRealLontimer,2)%></td>
	    <%strEksportTxt = strEksportTxt & formatnumber(balRealLontimer,2) & ";" %>
	 
	    
	    <%end if %>
	 <%end if %>
	 
	 
	     <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	           
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(afspTimer(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerBr(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerUdb(x), 2)%></td>
	     
	          <% 
	          aafspTimerTot = aafspTimerTot + afspTimer(x) 
	          aafspTimerBrTot = aafspTimerBrTot + afspTimerBr(x)
	          aafspTimerUdbTot = aafspTimerUdbTot + afspTimerUdb(x)
	          
	         afspadUdbBal = 0
	         afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	         
	         aafspadUdbBalTot = aafspadUdbBalTot + (afspadUdbBal)
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspadUdbBal, 2)%></td>
	          <% 
	         AfspadBal = 0 
	         AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	         aAfspadBalTot = aAfspadBalTot + (AfspadBal)
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(AfspadBal, 2)%></td>
        	 
	         
	         <%strEksportTxt = strEksportTxt & formatnumber(afspTimer(x), 2) &";"& formatnumber(afspTimerBr(x), 2) &";" & formatnumber(afspTimerUdb(x), 2) &";" & formatnumber(afspadUdbBal, 2) &";" & formatnumber(AfspadBal, 2) &";" %>
	 
	         
	         <%end if %>
	         
	        <%if visning = 14 then %>
	        <%
	        sidsteDag = year(dateadd("d", 6, startDato)) & "-"& month(dateadd("d", 6, startDato)) & "-"& day(dateadd("d", 6, startDato))
	        call erugeAfslutte(datepart("yyyy", startDato,2,2), datepart("ww", sidsteDag, 2, 2), intMid) 
	        
	            if cint(showAfsuge) = 0 then
	            showAfsugeTxt = "Ja"
	            else
	            showAfsugeTxt = "Nej"
	            end if
	            %>
	        <td class=lille style="border-bottom:1px silver dashed;" align=center><%=showAfsugeTxt %></td>
	        <%strEksportTxt = strEksportTxt & showAfsugeTxt &";" %>
	 
	        <%end if %>
	        
	        
	 <%if normTimerDag(x) <> 0 then
	 ferieAfVal_md = ferieAFPerTimer(x)/normTimerDag(x)
	 else
	 ferieAfVal_md = 0
	 end if
	 
	 ferieAfVal_md_tot = ferieAfVal_md_tot + ferieAfVal_md
	  %>
	 
	 <%if visning = 7 then %>
	 <td class=lille style="border-bottom:1px silver dashed;" align=right><%=formatnumber(ferieAfVal_md,2) %></td>
	        
	         <%if level = 1 OR (session("mid") = usemrn) then %>
	         
	          <%
              
              if normTimerDag(x) <> 0 then 
	          sygDage_barnSyg = (sygTimer(x) + barnSygTimer(x)) / normTimerDag(x) 
	          else 
	          sygDage_barnSyg = 0
	          end if
	          
	          sygDage_barnSyg_tot = sygDage_barnSyg_tot + sygDage_barnSyg
	          %> 
	           
	         <td align=right class=lille style="border-bottom:1px silver dashed; white-space: nowrap;"><%=formatnumber(sygDage_barnSyg,2) %></td>
	         <%end if %>
	         
	         
	       
	        
	  <%end if %>
	  </tr>
	  <%end if 'showmedarb %>
	  <!-- slut -->
	  
   
	  
	  
	  
	  
	  
	         
	  
	  <%case 4 '**** Ferie ***' %>       
	         
	   <%if lto <> "cst" AND instr(akttype_sel, "#13#") <> 0 then %>
	   <!-- Ferie fridage -->
	 <br /><br /><b>Feriefridage</b> (<%=year(startDato)%> - <%=year(slutDato)%>) 
	 <table cellspacing=0 cellpadding=2 border=0 width=100%>
	  <tr bgcolor="#FFFF99">
	 
	 <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_174 &" "& tsa_txt_164%></b><br />~ dage</td>
	 <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b>Planlagt <br> >> dd.</b><br />~ dage</td>
	 
	  <td valign=bottom align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_165%></b><br />~ dage</td>
	   <td valign=bottom align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b><br />~ dage</td>
	  <td valign=bottom align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_282 &" "& tsa_txt_280 %></b><br />~ dage</td>
	 </tr>
	  <tr>
	
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 
	  <%if normTimerDag(x) <> 0 then
	 fefriVal = fefriTimer(x)/normTimerDag(x)
	 else
	 fefriVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(fefriVal,2) %> 
	 <!--<br /><%=formatnumber(fefriTimer(x),2) %> t.--></td>
	 
	  <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 
	  <%if normTimerDag(x) <> 0 then
	 fefriplVal = fefriplTimer(x)/normTimerDag(x)
	 else
	 fefriplVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(fefriplVal,2) %> 
	 <!--<br /><%=formatnumber(fefriTimer(x),2) %> t.--></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	  <%if normTimerDag(x) <> 0 then
	 fefriBrVal = fefriTimerBr(x)/normTimerDag(x)
	 else
	 fefriBrVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(fefriBrVal,2) %> 
	 <!--<br /><%=formatnumber(fefriTimerBr(x),2) %> t.--></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	 <%if normTimerDag(x) <> 0 then
	 fefriUdbVal = fefriTimerUdb(x)/normTimerDag(x)
	 else
	 fefriUdbVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(fefriUdbVal,2) %>
	 <!--<br /><%=formatnumber(fefriTimerUdb(x),2) %> t.--></td>
	  
	  <% 
	 fefriBal = 0
	 fefriBal  = (fefriTimer(x) - (fefriTimerBr(x) + fefriTimerUdb(x)))
	 
	 if normTimerDag(x) <> 0 then
	 fefriBalVal = fefriBal/normTimerDag(x)
	 else
	 fefriBalVal = 0
	 end if
	 
	  %>
	  <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=formatnumber(fefriBalVal,2) %><br />
	 <!-- <%=formatnumber(fefriBal,2)%> t.</b>--></td>
	 </tr>
	 </table>



	 <%end if 'cst%>
	 
	 
	 <!-- Ferie -->
	 
     <%if instr(akttype_sel, "#14#") <> 0  then %>
	    <br /><br /><b><%=tsa_txt_281 %></b> (<%=year(startDato)%> - <%=year(slutDato)%>) 
	
	  <table cellspacing=0 cellpadding=2 border=0 width=100%>
	 <tr bgcolor="#6CAE1C">
	 <td valign=bottom align=right class=alt_lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_152%> Opt.</b><br />~ dage</td>
	 <td valign=bottom align=right class=alt_lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_317%><br> >> dd.</b><br />~ dage</td>
	 <td valign=bottom align=right class=alt_lille style="border-bottom:1px silver dashed;"><b>Afholdt</b><br />~ dage</td>
	   <td valign=bottom align=right class=alt_lille style="border-bottom:1px silver dashed;"><b>Afholdt u. løn</b><br />~ dage</td>
	   <td valign=bottom align=right style="border-bottom:1px silver dashed;" class=alt_lille><b>Udbetalt</b><br />~ dage</td>
	  <td  valign=bottom align=right style="border-bottom:1px silver dashed;" class=alt_lille><b><%=tsa_txt_281 &" "& tsa_txt_280 %></b><br />~ dage</td>
	 </tr>
	  <tr>
	
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 
	 <%if normTimerDag(x) <> 0 then
	 ferieOptjVal = ferieOptjtimer(x)/normTimerDag(x)
	 else
	 ferieOptjVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(ferieOptjVal,2) %>
	 <!--<br /><%=formatnumber(ferieOptjtimer(x),2) %> t.--></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	 <%if normTimerDag(x) <> 0 then
	 feriePlVal = feriePLTimer(x)/normTimerDag(x)
	 else
	 feriePlVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(feriePlVal,2) %>
	 <!--<br /><%=formatnumber(feriePLTimer(x),2) %> t.--></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	 
	 
	 <%if normTimerDag(x) <> 0 then
	 ferieAfVal = ferieAFTimer(x)/normTimerDag(x)
	 else
	 ferieAfVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(ferieAfVal,2) %>
	 <!--<br /><%=formatnumber(ferieAFTimer(x),2) %> t.--></td>
	 
	  <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	 
	 
	 <%if normTimerDag(x) <> 0 then
	 ferieAfulonVal = ferieAFulonTimer(x)/normTimerDag(x)
	 else
	 ferieAfulonVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(ferieAfulonVal,2) %>
	 <!--<br /><%=formatnumber(ferieAFTimer(x),2) %> t.--></td>
	 
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	  <%if normTimerDag(x) <> 0 then
	 ferieUdbVal = ferieUdbTimer(x)/normTimerDag(x)
	 else
	 ferieUdbVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(ferieUdbVal,2) %>
	 <!--<br /><%=formatnumber(ferieUdbTimer(x),2) %> t.--></td>
	  
	  <% 
	 ferieBal = 0
	 
   
     select case lto
     case "cst"
     ferieBal  = (ferieOptjtimer(x) - (ferieAFTimer(x) + ferieAFulonTimer(x) + ferieUdbTimer(x)))
     case "wwf"
	 ferieBal  = (ferieOptjtimer(x) - (ferieAFTimer(x) + ferieUdbTimer(x) - (ferieAFulonTimer(x))))
     case else
     ferieBal  = (ferieOptjtimer(x) - (ferieAFTimer(x) + ferieUdbTimer(x)))
     end select
	 
	 if normTimerDag(x) <> 0 then
	 ferieBalVal = ferieBal/normTimerDag(x)
	 else
	 ferieBalVal = 0
	 end if
	 
	  %>
	  <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(ferieBalVal,2) %><br />
	  <!--<%=formatnumber(ferieBal,2)%> t.--></td>
	 </tr>
	 </table>
	 
	 <% end if 'instr 14
	 
	 
	 end select 
	 '** visning
	
	end function
	
	
	%>
	
	
	
	