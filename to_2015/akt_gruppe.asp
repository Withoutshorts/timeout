

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--include file="inc/budget_firapport_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<style>
    thead input {
        width: 100%;
        padding: 3px;
        box-sizing: border-box;
        font-weight: lighter;
    }
    .dataTables_filter {
display: none; 
}
</style>
<%

    func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function

    %>

        <!--#include file="to-menu.asp"-->

    <div id="wrapper">
        <div class="content">    

    <%
    select case func
    case "slet"

	    oskrift = "Stam-aktivitets gruppe" 
        slttxta = "Du er ved at <b>SLETTE</b> en stam-aktivitets gruppe. Er dette korrekt?"
        slttxtb = "" 
        slturl = "akt_gruppe.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)
    
    
    



    case "sletok"
	'*** Her slettes en akt gruppe ***
    

    '*** Indsætter i delete historik ****' 
        
        '** gruppen **'
        
        
        
        '** Aktivitteterne **'    
        strSQL = "SELECT id, navn FROM aktiviteter WHERE aktfavorit = "& id &" AND job = 0"

        oRec.open strSQL, oConn, 3
	    while not oRec.EOF
       
	    call insertDelhist("aktstam", id, 0, oRec("navn"), session("mid"), session("user"))

        oRec.movenext
        wend
        oRec.close


    oConn.execute("DELETE FROM aktiviteter WHERE aktfavorit = "& id &" AND job = 0")

	oConn.execute("DELETE FROM akt_gruppe WHERE id = "& id &"")


	Response.redirect "akt_gruppe.asp?menu=job&shokselector=1"


    case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
	
		<%
		errortype = 8
		call showError(errortype)
		
		else
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		forvalgt = request("FM_forvalgt")
        skabelonType = request("FM_type")
        aktgrp_account = 0 'request("FM_aktgrp_account")
		
        '** Nulstiller forvalgt **'
        oConn.execute("UPDATE akt_gruppe SET forvalgt = 0 WHERE id <> 0 AND skabelonType = "& skabelonType)
		

		if func = "dbopr" then
		oConn.execute("INSERT INTO akt_gruppe (navn, editor, dato, forvalgt, skabelontype, aktgrp_account) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& forvalgt &", "& skabelontype &", "& aktgrp_account &")")
		else
		oConn.execute("UPDATE akt_gruppe SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', forvalgt = "& forvalgt &", skabelontype = "& skabelontype &", aktgrp_account = "& aktgrp_account &" WHERE id = "&id&"")
		end if
		
		Response.redirect "akt_gruppe.asp?menu=job&shokselector=1"
		end if


    case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	intForvalgt = 0
    skabelontype = 0
    aktgrp_account = 0
	
	else
	strSQL = "SELECT navn, forvalgt, editor, dato, skabelontype, aktgrp_account FROM akt_gruppe WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intForvalgt = oRec("forvalgt")
    skabelontype = oRec("skabelontype")
    aktgrp_account = oRec("aktgrp_account")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if


    if cint(skabelontype) = 0 then
    skabelontypeSEL0 = "SELECTED"
    skabelontypeSEL1 = ""
    else
    skabelontypeSEL1 = "SELECTED"
    skabelontypeSEL0 = ""
    end if

	
        
%>

<!--aktivitest gruppe redigering-->
<div class="container">
    <div class="portlet">
            <h3 class="portlet-title">
            <u>Stam-aktivitestgrupper</u>
            </h3>

        <form action="akt_gruppe.asp?menu=tok&func=<%=dbfunc%>" method="post">
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
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2"><b>Aktivitetsgruppe:</b> (Navn)</div>
                <div class="col-lg-2">
                    <input name="FM_navn" type="text" class="form-control input-small" value="<%=strNavn %>" placeholder="Navn"/>
                </div>
            </div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2"><b>Forvalgt:</b></div>
                <div class="col-lg-2">
                    <%if cint(intForvalgt) <> 0 then
                    fvCHK1 = "CHECKED"
                    fvCHK0 = ""
                    else
                    fvCHK0 = "CHECKED"
                    fvCHK1 = ""
                    end if %>
                        <input id="Radio1" type="radio" name="FM_forvalgt" value="0" <%=fvCHK0 %> /> Nej
                        <br />
                        <input id="Radio2" type="radio" name="FM_forvalgt" value="1" <%=fvCHK1 %> /> Ja
                </div>
            </div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2"><b>Type:</b></div>
                <div class="col-lg-2">
                    <select class="form-control input-small" name="FM_type">
                        <option value="0" <%=skabelontypeSEL0 %>>Projekt</option>
                        <option value="1" <%=skabelontypeSEL1 %>>Budget</option>
                    </select>
                </div>
            </div>
                    
        </div>
         </section>
             <br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
        </div>
        </form>
    </div>
</div>



<% 
    case "kopier" 

    strSQL = "SELECT navn FROM akt_gruppe WHERE id =" & id
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	grpNavn = oRec("navn")
	
	oRec.movenext
	wend
	
	oRec.close
	
%>
            <!--Kopier sidens indhold-->
        <div class="container">
            <div class="portlet">
            <h3 class="portlet-title">
            <u>Kopier stam-aktivitestgruppe</u>
            </h3>

                <form method=post action="akt_gruppe.asp?func=kopierja&id=<%=id%>">
    
                    <div class="row">
                    <div class="col-lg-10">&nbsp</div>
                    <div class="col-lg-2 pad-b10">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b>Kopier gruppe</b></button>
                    </div>
                    </div>

            <div class="portlet-body">
                <div class="row">                 
                    <div class="col-lg-1">&nbsp</div>
                    <h4 class="col-lg-2">Stam-aktivitestgruppe:</h4>
                </div>

                <div class="row">                 
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><b>GruppeNavn:</b></div>
                    <div class="col-lg-3"><input type="text" name="FM_grpnavn" class="form-control input-small" id="FM_grpnavn" value="Kopi af: <%=grpNavn%>"></div>
                    <div class="col-lg-1 pad-b20">&nbsp</div>
                </div>

                <div class="row">                 
                    <div class="col-lg-1">&nbsp</div>
                    <h4 class="col-lg-2 pad-t20">Aktiviteterne:</h4>
                </div>

                <div class="row">                 
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><b>Aktivitetsnavn:</b></div>
                    <div class="col-lg-3"><input type="checkbox" name="FM_udskift_navn" id="FM_udskift_navn" value="1">&nbsp Ja, omdøb navn på aktiviteter.<br>
		            Udskift denne del af aktivitets navnet:<br>
                        <input type="text" name="FM_aktnavn_for" class="form-control input-small" id="FM_aktnavn_for" value="">
		                med dette:
		                <input type="text" name="FM_aktnavn_efter" class="form-control input-small" id="FM_aktnavn_efter" value="">
                    <br /></div>
                </div>

                <div class="row">                 
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><b>Fase:</b></div>
                    <div class="col-lg-2"><input type="checkbox" name="FM_udskift_fase" id="Checkbox1" value="1">&nbsp Ja, omdøb fase på aktiviteter.<br>
	                    Udskift denne del af fase navnet:<br> 
	                    <input type="text" name="FM_fase_for" class="form-control input-small" id="Text1" value="">
	                    med dette:
	                    <input type="text" name="FM_fase_efter" class="form-control input-small" id="Text2" value=""><br /></div>
                </div>

                <div class="row">                 
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><b>Faktor:</b></div>
                    <div class="col-lg-2"><input type="checkbox" name="FM_udskift_faktor" id="FM_udskift_faktor" value="1">Ja, udskift faktor.<br>
		                <input type="text" name="FM_faktor" id="FM_faktor" class="form-control input-small" value="" /><br /></div>
                </div>

                <div class="row">                 
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><b>Forretningsområder:</b></div>
                    <div class="col-lg-4"><input type="checkbox" name="FM_udskift_fomr" id="FM_udskift_fomr" value="1">Ja, udskift forretningsområder med:<br>
			            <select name="FM_fomr" class="form-control input-small" multible="multible" size="4">
		            <option value="0">Ingen</option>
		            <%
		            strSQL = "SELECT id, navn FROM fomr ORDER BY navn"
		            oRec.open strSQL, oConn, 3 
		            while not oRec.EOF 
		            %>
		            <option value="<%=oRec("id")%>"><%=oRec("navn")%></option>
		            <%
		            oRec.movenext
		            wend
		            oRec.close 
		            %>
		            </select><br /></div>
                   
                </div>

                <div class="row">                 
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><b>Tidslås:</b></div>
                    <div class="col-lg-4"><input type="checkbox" name="FM_udskift_tidslaas" id="FM_udskift_tidslaas" value="1">Ja, udskift tidslås.<br>
		            <input type="checkbox" name="FM_tidslaas" id="FM_tidslaas" value="1">Tidslås aktiv. Der skal <b>kun</b> kunne registreres timer mellem:<br> 
		            kl. <input type="text" name="FM_tidslaas_start" id="FM_tidslaas_start" class="form-control input-small" size="3" value="07:00"> og kl.
		            <input type="text" name="FM_tidslaas_slut" id="FM_tidslaas_slut" class="form-control input-small" size="3" value="23:30"> (Format: tt:mm)
		
		            <br />På følgende dage:<br />
                         Man 
                        <input id="FM_tidslaas_man" name="FM_tidslaas_man" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Tir
                        <input id="FM_tidslaas_tir" name="FM_tidslaas_tir" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Ons 
                        <input id="FM_tidslaas_ons" name="FM_tidslaas_ons" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Tor 
                        <input id="FM_tidslaas_tor" name="FM_tidslaas_tor" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Fre 
                        <input id="FM_tidslaas_fre" name="FM_tidslaas_fre" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Lør 
                        <input id="FM_tidslaas_lor" name="FM_tidslaas_lor" value="1" type="checkbox"/>&nbsp;&nbsp;
                         Søn 
                        <input id="FM_tidslaas_son" name="FM_tidslaas_son" value="1" type="checkbox"/>&nbsp;&nbsp;</div>
                </div>
            </div>
            </form>
            </div>
        </div>

<%
    case "kopierja"
    
    'Response.write "så er vi her"
	
			
			'** Faktor **
			if len(request("FM_udskift_faktor")) <> 0 then
			udskiftFaktor = 1
			else
			udskiftFaktor = 0
			end if
			aktFaktor = replace(request("FM_faktor"), ",", ".")
				
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(dblFaktor)
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
			
			<%
			errortype = 51
			call showError(errortype)
			isInt = 0
			response.end
			end if 
				
					
				
				'*** tidslaas validering ***
				'** Tidslaas **
				udskiftTidslaas = request("FM_udskift_tidslaas")
				
				if len(request("FM_tidslaas")) <> 0 then
				tidslaas = 1
				else
				tidslaas = 0
				end if
				
				tidslaas_st = request("FM_tidslaas_start") & ":00"
				tidslaas_sl = request("FM_tidslaas_slut") & ":00"
				
				if len(trim(request("FM_tidslaas_man"))) <> 0 then
				tidslaas_man = 1
				else 
				tidslaas_man = 0
				end if
				
				if len(trim(request("FM_tidslaas_tir"))) <> 0 then
				tidslaas_tir = 1
				else 
				tidslaas_tir = 0
				end if
				
				if len(trim(request("FM_tidslaas_ons"))) <> 0 then
				tidslaas_ons = 1
				else 
				tidslaas_ons = 0
				end if
				
				if len(trim(request("FM_tidslaas_tor"))) <> 0 then
				tidslaas_tor = 1
				else 
				tidslaas_tor = 0
				end if
				
				if len(trim(request("FM_tidslaas_fre"))) <> 0 then
				tidslaas_fre = 1
				else 
				tidslaas_fre = 0
				end if
				
				if len(trim(request("FM_tidslaas_lor"))) <> 0 then
				tidslaas_lor = 1
				else 
				tidslaas_lor = 0
				end if
				
				if len(trim(request("FM_tidslaas_son"))) <> 0 then
				tidslaas_son = 1
				else 
				tidslaas_son = 0
				end if
					
				function SQLBless3(s3)
				dim tmp3
				tmp3 = s3
				tmp3 = replace(tmp3, ":", "")
				SQLBless3 = tmp3
				end function
				
				
				idagErrTjek = day(now)&"/"&month(now)&"/"&year(now)
				
				for t = 1 to 2
				
				select case t
				case 1
				tdsl = tidslaas_st
				case 2
				tdsl = tidslaas_sl
				end select
				
				'Response.write SQLBless3(trim(tSttid(y))) & "<br>"
				'Response.flush
				
				call erDetInt(SQLBless3(trim(tdsl)))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
				'*** Punktum  i angivelse ved registrering af klokkeslet
				if instr(trim(tdsl), ".") <> 0 OR instr(trim(tdsl), ",") <> 0 then
					antalErr = 1
					errortype = 66
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				
				if len(trim(tdsl)) <> 0 then
				
				'Response.write idagErrTjek &" "& tSttid(y)&":00" &" - "& isdate(idagErrTjek &" "& tSttid(y)&":00") &"<br>"
					if isdate(idagErrTjek &" "& tdsl) = false then
						antalErr = 1
						errortype = 64
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
				
			
			next
				
				
			'*** Validering ok! *** 		
			
			
			'** Gruppe **
			intAktfavgp = id
			grpNavn = SQLBless(request("FM_grpnavn"))
			strEditor = session("user")
			strDato = day(now) &"-" & month(now) & "-" & year(now)
			
			'** Opretter gruppe ***
			strSQLopr = "INSERT INTO akt_gruppe (navn, editor, dato) VALUES ('"& grpNavn &"', '"& strEditor &"', '"& strDato &"')"
			oConn.execute strSQLopr
			
			
			'*** Henter gruppe ID ***
			strSQL = "SELECT id FROM akt_gruppe WHERE id <> 0 ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				
				newaktgrpId = oRec("id")
			
			end if
			oRec.close   
			
			'** Navn **
			if len(request("FM_udskift_navn")) <> 0 then
			udskiftNavn = 1
			else
			udskiftNavn = 0
			end if
			aktNavn_for = SQLBless(request("FM_aktnavn_for")) 
			aktNavn_efter = SQLBless(request("FM_aktnavn_efter"))
			
			
			'** Fase **'
			if len(request("FM_udskift_fase")) <> 0 then
			udskiftfase = 1
			else
			udskiftfase = 0
			end if
			fase_for = SQLBless(request("FM_fase_for")) 
			fase_efter = SQLBless(request("FM_fase_efter"))
			
			function SQLBless4(s, glnavn, nytnavn)
			tmpthis = s
			tmpthis = replace(lcase(tmpthis), ""&lcase(glnavn)&"", ""&nytnavn&"")
			SQLBless4 = tmpthis
			end function
			
			
			'** Forretningsområde ***
			if len(request("FM_udskift_fomr")) <> 0 then
			udskiftFomr = 1
			else
			udskiftFomr = 0
			end if

			aktFomr = 0 'request("FM_fomr") 


                  '************************************'
                  '***** Forrretningsområder **********'
                  '************************************'

                  if cint(udskiftFomr) = 1 then

                    fomrArr = split(request("FM_fomr"), ",")

                    for_faktor = 0
                    for afor = 0 to UBOUND(fomrArr)
                    for_faktor = for_faktor + 1 
                    next

                    if for_faktor <> 0 then
                    for_faktor = for_faktor
                    else
                    for_faktor = 1
                    end if

                    for_faktor = formatnumber(100/for_faktor, 2)
                    for_faktor = replace(for_faktor, ",", ".")
                  
                  
                  end if

                  '**************************************'
			
			
			
				
				
							'**** Henter aktiviteter fra oprindelig gruppe ***
							strSQL2 = "select id, navn, fakturerbar, budgettimer, fomr, faktor, "_
							&" aktbudget, aktstatus, tidslaas, tidslaas_st, tidslaas_sl, antalstk, "_
							&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
	                        &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg, brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr, kostpristarif "_
							&" FROM aktiviteter WHERE aktFavorit = "& intAktfavgp 
								
								    oRec2.open strSQL2, oConn, 3
									while not oRec2.EOF
									
									aktNavn = replace(oRec2("navn"), "'", "''")
									
									if cint(udskiftNavn) = 1 then
									aktNavn = SQLBless4(trim(aktNavn),aktNavn_for,trim(aktNavn_efter))
									end if
									
									aktFakbar = oRec2("fakturerbar")
									
									'if cint(udskiftFomr) = 0 then
									'aktFomr = oRec2("fomr")
									'else
									'aktFomr = aktFomr
									'end if
									
									if cint(udskiftFaktor) = 0 then
									aktFaktor = replace(oRec2("faktor"), ",", ".")
									else
									aktFaktor = aktFaktor
									end if
									
									if len(aktFaktor) <> 0 then
									aktFaktor = aktFaktor
									else
									aktFaktor = 0
									end if
									
									aktBudget = replace(oRec2("aktbudget"), ",", ".")
									aktbudgetsum = replace(oRec2("aktbudgetsum"), ",", ".")
									
									aktstatus = oRec2("aktstatus")
									
									if len(trim(oRec2("fase"))) <> 0 then
									strFase = replace(oRec2("fase"), "'", "''")
									else
									strFase = ""
									end if
									
									if cint(udskiftfase) = 1 then
									strFase = SQLBless4(trim(strFase),fase_for,trim(fase_efter))
									end if
									
									bgr = oRec2("bgr")
									
									sortorder = oRec2("sortorder")
									
									
									if cint(tidslaas) = 0 then 
									
										tidslaas = oRec2("tidslaas")
										tidslaas_st = oRec2("tidslaas_st")
										tidslaas_sl = oRec2("tidslaas_sl")
										
										
										if len(tidslaas) <> 0 then
										tidslaas = tidslaas
										else
										tidslaas = 0
										end if
										
										
										if len(trim(tidslaas_st)) <> 0 AND len(trim(tidslaas_sl)) <> 0 then
										tidslaas_st = left(formatdatetime(tidslaas_st, 3), 7)
										tidslaas_sl = left(formatdatetime(tidslaas_sl, 3), 7)
										else
										tidslaas_st = "07:00:00"
										tidslaas_sl = "23:30:00"
										end if
										
										
				                        tidslaas_man = oRec2("tidslaas_man")
				                        tidslaas_tir = oRec2("tidslaas_tir")
				                        tidslaas_ons = oRec2("tidslaas_ons")
				                        tidslaas_tor = oRec2("tidslaas_tor")
				                        tidslaas_fre = oRec2("tidslaas_fre")
				                        tidslaas_lor = oRec2("tidslaas_lor")
				                        tidslaas_son = oRec2("tidslaas_son")
				                       
									
									end if
									
									antalstk = replace(oRec2("antalstk"), ",", ".")
									aktBudgettimer = replace(oRec2("budgettimer"), ",", ".")
									
									easyreg = oRec2("easyreg")

                                    brug_fasttp = oRec2("brug_fasttp")
                                    brug_fastkp = oRec2("brug_fastkp")
                                    fasttp = replace(oRec2("fasttp"), ",", ".")
                                    fasttp_val = oRec2("fasttp_val")
                                    fastkp = replace(oRec2("fastkp"), ",", ".")
                                    fastkp_val = oRec2("fastkp_val")
                         
                                    avarenr = oRec2("avarenr")

                                    kostpristarif = oRec2("kostpristarif")

                                   

									
									
									strSQLins = "INSERT INTO aktiviteter "_
									&" (navn, dato, editor, job, fakturerbar, "_
									&" projektgruppe1, projektgruppe2, projektgruppe3, "_
									&" projektgruppe4, projektgruppe5, projektgruppe6, "_
									&" projektgruppe7, projektgruppe8, projektgruppe9, "_
									&" projektgruppe10, aktstartdato, aktslutdato, "_
									&" budgettimer, fomr, faktor, aktBudget, aktstatus, tidslaas, tidslaas_st, "_
									&" tidslaas_sl, antalstk, aktfavorit,"_
									&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
						            &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, sortorder, bgr, aktbudgetsum, easyreg, brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr, kostpristarif"
						            
						            if len(trim(strFase)) <> 0 then
						            strSQLins = strSQLins & ", fase"
						            end if
						             
						             
						             
									 strSQLins = strSQLins &") VALUES "_
									&"('"& aktNavn &"', "_
									&"'"& strDato &"', "_ 
									&"'"& strEditor &"', "_
									&"0, "_ 
									&""& aktFakbar &", "_
									&"10,1,1,1,1,1,1,1,1,1, "_ 
									&"'2001/1/1', "_ 
									&"'2044/1/1', "_
									&""& aktBudgettimer & ", "& aktFomr &", "_
									&""& aktFaktor &", "& aktBudget &", "& aktstatus &", "_
									&""&tidslaas&", '"&tidslaas_st&"', '"&tidslaas_sl&"', "& antalstk &", "& newaktgrpId &", "_
									&" "& tidslaas_man &", "& tidslaas_tir &", "& tidslaas_ons &", "_
						            &" "& tidslaas_tor &", "& tidslaas_fre &", "& tidslaas_lor &", "& tidslaas_son & ", " & sortorder & ", "& bgr &", " & aktbudgetsum & ", "& easyreg &", "& brug_fasttp &","& brug_fastkp &","& fasttp &","& fasttp_val &","& fastkp &","& fastkp_val &", '"& avarenr &"', '"& kostpristarif &"'"
						            
						            if len(trim(strFase)) <> 0 then
						            strSQLins = strSQLins & ", '"& strFase &"'"
						            end if
						            
						            strSQLins = strSQLins &")"
									
									
									'Response.write strSQLins
									'Response.flush
									oConn.execute(strSQLins)
									
									
									'*** Henter det netop oprettede akt-id ***
									strSQLid = "SELECT id FROM aktiviteter ORDER BY id DESC"
									oRec3.open strSQLid, oConn, 3
									if not oRec3.EOF then
									useAktid = oRec3("id")
									end if
									oRec3.close
									
									if len(useAktid) <> 0 then
									useAktid = useAktid
									else
									useAktid = 0
									end if
									
									
									'*** Timepriser ***
									strSQLtp1 = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris FROM timepriser WHERE aktid = " & oRec2("id")
									oRec3.open strSQLtp1, oConn, 3
									while not oRec3.EOF
									 	
										strjnr = oRec3("jobid")
										medarbid = oRec3("medarbid")
										timeprisalt = replace(oRec3("timeprisalt"), ",",".")
										this6timepris = replace(oRec3("6timepris"), ",",".")
										
										strSQLtp2 = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris) VALUES ("& strjnr &", "& useAktid &", "& medarbid &", "&timeprisalt&", "& this6timepris &")"
										oConn.execute (strSQLtp2)
										
										
									oRec3.movenext
									wend
									oRec3.close
									
									
                                                    
                                                    '********************************'
                                                    '***** Forretningsområder ******'
                                                    '********************************'
                    
                    
                                                    '**** Hvis udskift = 1 (ja), ellers kopier ***'
                                                    if cint(udskiftFomr) = 1 then

                                                            for afor = 0 to UBOUND(fomrArr)

                                                                    'Response.Write "her2" & afor & "<br>"
                                                                    'Response.Flush

                                                                    if fomrArr(afor) <> 0 then

                                                                    strSQLfomri = "INSERT INTO fomr_rel "_
                                                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                                                    &" VALUES ("& fomrArr(afor) &", "& strjnr &", "& useAktid &", "& for_faktor &")"

                                                                    oConn.execute(strSQLfomri)

                                                                   end if


                                                            next

                                                    else 'kopier
                                                            
                                                            strSQLafor = "SELECT for_fomr, for_jobid, for_aktid, for_faktor FROM fomr_rel WHERE for_aktid = "& oRec2("id")

                                                            oRec3.open strSQLafor, oConn, 3
                                                            while not oRec3.EOF

                                                                    strSQLfomri = "INSERT INTO fomr_rel "_
                                                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                                                    &" VALUES ("& oRec3("for_fomr") &", "& oRec3("for_fomr") &", "& useAktid &", "& oRec3("for_faktor") &")"

                                                                    oConn.execute(strSQLfomri)

                                                            oRec3.movenext
                                                            wend
                                                            oRec3.close


                                                    end if

                                                    '********************************' 

									
							
							oRec2.movenext
							wend
							oRec2.close
					
						
		
	Response.redirect "akt_gruppe.asp"
         
%>




<%
case else    
%>


<script src="js/akt_gruppe.js" type="text/javascript"></script>> 

<!--aktivitest grupper liste-->
<div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Stam-aktivitetsgrupper (skabeloner)</u>
        </h3>

          <form action="akt_gruppe.asp?menu=tok&func=opret" method="post">
          <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                          <button class="btn btn-sm btn-success pull-right"><b>Opret ny</b></button><br />&nbsp;
                          </div>
                 </div>
              </section>
         </form>

          <div class="porlet-body">
              <table id="example" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                  <thead>
                      <tr>
                          <th>Id</th>
                          <th>Stam aktivitest gruppe</th>
                          <th>Forvalgt</th>
                          <th>Se gruppe</th>
                          <th>Kopier</th>
                          <th>Slet</th>
                      </tr>
                  </thead>
                  <tbody>
                      <%
	                    strSQL = "SELECT id, navn, forvalgt, skabelontype FROM akt_gruppe WHERE id <> 2 ORDER BY navn" 'gl. kørsels gruppe
	                    oRec.open strSQL, oConn, 3
	                    while not oRec.EOF 
	                    %>
                      <tr>
                          <td><%=oRec("id") %></td>
                          <td><a href="akt_gruppe.asp?func=red&id=<%=oRec("id") %>"><%=oRec("navn") %></a></td>
                          <td><%=oRec("forvalgt") %>
                              <%if cint(oRec("skabelontype")) = 1 then 'budget
                                %>
                            - Budgetskabelon
                            <%
                            end if %>

                              <%
					            '** Henter aktiviteter i den aktuelle gruppe ****
					            strSQL2 = "SELECT count(id) AS antal FROM aktiviteter WHERE aktfavorit = "&oRec("id")&""
					            oRec2.open strSQL2, oConn, 3
					            if not oRec2.EOF then
					            intAntal = oRec2("antal")
					            end if
					            oRec2.close 
					          %>
                              
                          </td>
                          <td><a href="aktiv.asp?menu=job&func=favorit&id=<%=oRec("id")%>&stamakgrpnavn=<%=oRec("navn")%>">Se / Rediger Stam-aktiviteter i grp.&nbsp;</a>(<%=intAntal%>)</td>
                          <td><a href="akt_gruppe.asp?func=kopier&id=<%=oRec("id")%>" class=rmenu><span style="display: block; text-align: center;" class="fa fa-copy"></span></a></td>
                          <td><a href="akt_gruppe.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a></td>
                      </tr>
                      <%
                      oRec.movenext
                      wend
                      Orec.close   
                      %>
                  </tbody>
              </table>
          </div>
      </div>

</div>

<%end select  %>

</div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->