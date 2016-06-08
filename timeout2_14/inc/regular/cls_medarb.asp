 
 <%
 
  

 public meNavn, meNr, meInit, meTxt, meEmail, meType, meAnsatDato, meOpsagtdato, meforecaststamp, meBrugergruppe, meVisskiftversion, meMansat, timer_ststop
     'Public Shared Function meStamdata(medid)
    Function meStamdata(medid)  

       if medid <> 0 then
       medid = medid
       else
       medid = 0
       end if

       meVisskiftversion = 0
       
        strSQLmnavn = "SELECT mnavn, init, mnr, email, medarbejdertype, ansatdato, opsagtdato, forecaststamp, brugergruppe, visskiftversion, mansat, timer_ststop "_
        &" FROM medarbejdere WHERE mid = "& medid

	    oRec3.open strSQLmnavn, oConn, 3
	    if not oRec3.EOF then
    	
	    meNavn = oRec3("mnavn")
	    meNr = oRec3("mnr")
	    meInit = oRec3("init")
    	meEmail = oRec3("email")
    	meType = oRec3("medarbejdertype")
    	meAnsatDato = oRec3("ansatdato")
    	meOpsagtdato = oRec3("opsagtdato")
        meBrugergruppe = oRec3("brugergruppe")
        meVisskiftversion = oRec3("visskiftversion")'
        meMansat = oRec3("mansat")
    	
	    meTxt = meNavn ' & " ("& meNr & ")"
    	
	    if len(trim(meInit)) <> 0 then
	    meTxt = meTxt &" ["& meInit & "]"
	    end if
    	
        if IsNull(oRec3("forecaststamp")) <> true then
        meforecaststamp = formatdatetime(oRec3("forecaststamp"), 1)
        else
        meforecaststamp = ""
        end if

        timer_ststop = oRec3("timer_ststop")

	    end if
	    oRec3.close
       
   end function
   





   public antalmedarb
   function antalmedarb_fn(io)

   '1 = aktiv
   '2 = de-aktiv
   '3 = passiv

   if io = 2 then
   medStatKri = "mansat <> '2'" '** Aktiv + passive
   else
   medStatKri = "mansat <> '0'" 'Alle
   end if


            antalmedarb = 0
	        strSQL = "SELECT count(Mid) AS antusr FROM medarbejdere WHERE " & medStatKri
			oRec2.Open strSQL, oConn, 3
			if not oRec2.EOF then
			antalmedarb = oRec2("antusr")
			end if
			oRec2.close

   end function


     antMtypT = 150 'antalmtyper 
     '************* Medarbejdertyper i medarbejderhovedgrupper ****'
    public mtypgrpids, mtypgrpnavn, mtypgrpid, medarbimType, medarbSQlKriMtypGrp, mtypid, mtypnavn, kunMtypgrp, kunMtypgrpNavn, thisMtypGrpBelIg, thisMtypTimerBelob
    redim mtypgrpids(antMtypT), mtypgrpnavn(antMtypT), mtypgrpid(antMtypT), medarbimType(antMtypT), mtypid(antMtypT), mtypnavn(antMtypT)
    redim kunMtypgrp(10), kunMtypgrpNavn(10)
    redim thisMtypGrpBelIg(50), thisMtypTimerBelob(50)

    
     
   function mtyperIGrp_fn(vlgtmtypgrp,visning) 
     
     if visning = 0 then 'budget:0 / Realiseret:1
     sostergrpSQL = "mt.sostergp = 0" 
     else
     sostergrpSQL = "mt.sostergp <> -1"
     end if

     'Response.write "HER: "& vlgtmtypgrp
     'vlgtmtypgrp = 0
     if vlgtmtypgrp <> "0" then
    
        vlgtmtypgrpIds = split(vlgtmtypgrp, ", ")
        
        vlgtmtypgrpSQLkri = " (mtg.id = 0 "
        for v = 0 to UBOUND(vlgtmtypgrpIds) 
        vlgtmtypgrpSQLkri = vlgtmtypgrpSQLkri & " OR mtg.id = "& trim(vlgtmtypgrpIds(v))
        next
        
     vlgtmtypgrpSQLkri = vlgtmtypgrpSQLkri & ")"   

     'Response.write vlgtmtypgrpSQLkri
     'Response.end 

     else
     vlgtmtypgrpSQLkri = "mtg.id > 0" 'så vi undgår -1 gruppen, dvs dem der i søstergruppe er sat til ikke at skulle medregnes overhoved.
     end if

    'Response.write "THISFILE: "& thisfile & " FM_medarb. "& request("FM_medarb") & "<br>"
    'if thisfile = "joblog_timetotaler" AND 

    strSQLmtyp = "SELECT mtg.id AS mtypGrpId, mt.id AS mtypId, mtg.navn AS mtypgrpnavn, mt.type AS mtypenavn FROM medarbtyper_grp AS mtg "_
    &" LEFT JOIN medarbejdertyper AS mt ON (mt.mgruppe = mtg.id AND "& sostergrpSQL &")"_
    &" WHERE "& vlgtmtypgrpSQLkri &" AND ("& sostergrpSQL &") ORDER BY mtg.id, mt.id "
    
    'if session("mid") = 1 then
    'Response.write "xx" &  strSQLmtyp
    'Response.flush
    'end if
    
     tcnt = 0
     kmyp = 0


     mtypgrpidStr = "#0#"
     mtypgrpids(0) =  "typeid = 0 "
     LastMtypGrpId = 0
     LastMtypId = 0
     thisMtypid = 0
     medarbSQlKriMtypGrp = "(m.mid = 0 "
     medarbimType(tcnt) = " tmnr = 0 "
     
    
     oRec6.open strSQLmtyp, oConn, 3
     while not oRec6.EOF
    
    if isNull(oRec6("mtypId")) <> true then

     tcnt = tcnt + 1

            'if oRec6("mtypGrpId") <> LastMtypGrpId then
             if oRec6("mtypId") <> LastMtypId then

             'tcnt = tcnt + 1
             mtypgrpids(tcnt) = "typeid = 0 "
             medarbimType(tcnt) = " tmnr = 0 "
     
             end if
    
    
    mtypgrpids(tcnt) = mtypgrpids(tcnt) & " OR typeid = "& oRec6("mtypId")
    mtypgrpnavn(tcnt) = oRec6("mtypgrpnavn")
    mtypgrpid(tcnt) = oRec6("mtypGrpId")
    mtypid(tcnt) = oRec6("mtypId")
    mtypnavn(tcnt) = oRec6("mtypenavn")
   
    if instr(mtypgrpidStr, "#"&oRec6("mtypGrpId")&"#") = 0 then
    kmyp = kmyp + 1
    kunMtypgrp(kmyp) = oRec6("mtypGrpId")
    kunMtypgrpNavn(kmyp) = oRec6("mtypgrpnavn")
    mtypgrpidStr = mtypgrpidStr & ",#"& oRec6("mtypGrpId")&"#"
    end if

    if isNull(oRec6("mtypId")) <> true then
    thisMtypid = oRec6("mtypId")
    else
    thisMtypid = 0
    end if
            
            '** Medarbjedere af typen
            '*** Kan optimeres da medarbejerrtype_grp er kommet med på medarbejder niveau
            strSQLmedityp = "SELECT m.mid FROM medarbejdere AS m WHERE m.medarbejdertype = "& thisMtypid & " AND m.medarbejdertype_grp <> -1" 
            oRec4.open strSQLmedityp, oConn, 3
            while not oRec4.EOF

             medarbimType(tcnt) =  medarbimType(tcnt) & " OR tmnr = "& oRec4("mid")
             medarbSQlKriMtypGrp = medarbSQlKriMtypGrp & " OR m.mid = "& oRec4("mid")



             oRec4.movenext
             wend 
             oRec4.close

       
     
     
     end if 'mtypeID Null

     LastMtypGrpId = oRec6("mtypGrpId") 
     LastMtypId = oRec6("mtypId")
     

     oRec6.movenext
    Wend
    oRec6.close 

     medarbSQlKriMtypGrp = medarbSQlKriMtypGrp & ")"

    end function



     public mtypeids, mtypenavne, mtypesostergp, mtypeGp
     redim mtypeids(1), mtypenavne(1), mtypesostergp(1), mtypeGp(1)
     function fn_medarbtyper()

     t = 0
     '** kun dem der skal vises under stat. SosterGp > 0 tilføejs til dens søster
     strSQLmtyp = "SELECT id, type, sostergp, mgruppe FROM medarbejdertyper WHERE id <> 0 AND sostergp = 0 ORDER BY mtsortorder, type"

     oRec4.open strSQLmtyp, oConn, 3
     while not oRec4.EOF
     
     redim preserve mtypeids(t), mtypenavne(t), mtypesostergp(t), mtypeGp(t)
     
    
     mtypeids(t) = oRec4("id")
     mtypenavne(t) = oRec4("type")
     mtypesostergp(t) = ""
     mtypeGp(t) = oRec4("mgruppe") 
                
        strSQLSostyp = "SELECT id, type, sostergp FROM medarbejdertyper WHERE sostergp = "& oRec4("id") &" ORDER BY id"
        oRec5.open strSQLSostyp, oConn, 3
        while not oRec5.EOF
        
              mtypesostergp(t) = mtypesostergp(t) & " OR mtype = "& oRec5("id")

        oRec5.movenext
        wend
        oRec5.close
   
   
   
     
     t = t + 1
     oRec4.movenext
     wend 
     oRec4.close


     end function




     sub medarb_vaelgandre


     
					strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init FROM medarbejdere WHERE mansat <> 2 "& strSQLmids &" GROUP BY mid ORDER BY Mnavn"
					
                 

					%>
					<select name="usemrn" id="usemrn" style="width:250px;"><!-- onchange="submit(); -->
					<%
					
					oRec.open strSQL, oConn, 3
					while not oRec.EOF 
					
					if cint(oRec("Mid")) = cint(usemrn) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if
                    
                        if len(trim(oRec("init"))) <> 0 then
                        medTxt = oRec("mnavn") & " ["& oRec("init") &"]"
                        else
                        medTxt = oRec("mnavn") 
                        end if
                        
                    %>
					<option value="<%=oRec("Mid")%>" <%=rchk%>><%=medTxt%></option>
					<%
					
					
					oRec.movenext
					wend
					oRec.close
				%></select>


    <%

     end sub


    public mtypNoflex, mtypMtypeId 
    function mtyp_fn(mid)

        mtypMtypeId = 0
        mtypNoflex = 0
        strSQLmtyp = "SELECT m.medarbejdertype, mt.noflex, mt.id AS medarbtype FROM medarbejdere AS m "_
        &" LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) WHERE m.mid = "& mid &" ORDER BY id"
        
        'response.write strSQLmtyp
        'response.flush
        
        oRec5.open strSQLmtyp, oConn, 3
        while not oRec5.EOF
        
         mtypNoflex = oRec5("noflex")
         mtypMtypeId = oRec5("medarbtype")

        oRec5.movenext
        wend
        oRec5.close


    end function
   %>