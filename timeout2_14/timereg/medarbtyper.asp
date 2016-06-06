<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/medarb_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
Session.LCID = 1030

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
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px;">
    <%
    
   
    slttxt = "<b>Slet medarbejdertype?</b><br />"_
	&"Du er ved at <b>slette</b> en medarbejdertype. Er dette korrekt?"
	slturl = "medarbtyper.asp?menu=medarb&func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,200,90)
    
     %>

	
	
	
	</div>
	<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM medarbejdertyper WHERE id = "& id &"")
	Response.redirect "medarbtyper.asp?menu=medarb"
	
	case "med"
	%>
	
	
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

	
	
      <%call menu_2014() %>	

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px;">

	<%

    'oimg = "ikon_medarb_48.png"
	

	
    if level = 1 then

        tTop = 20
	 tLeft = 0
	 tWdth = 600
	       
     call tableDiv(tTop,tLeft,tWdth)


        oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Medarbejdertyper"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)


    
    strSQL = "SELECT id, type FROM medarbejdertyper WHERE id=" & id 
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	gruppeNavn = oRec("type")
	end if
	oRec.close
	%>
	<br><br>
	Aktive og passive medarbejdere af typen <b><%=gruppeNavn%></b>:<br>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#EFF3FF">
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan=5 valign="top"><img src="../ill/blank.gif" width="587" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt style="width:200px;"><b>Navn</b></td>
        <td>Initialer</td>
        <td>Status</td>
		<td colspan=2>&nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT mnavn, mid, mansat, init, mnr FROM medarbejdere WHERE medarbejdertype = "&id&" AND mansat <> 2 ORDER BY mansat, mnavn"
	oRec.open strSQL, oConn, 3
	
    x = 0
	while not oRec.EOF 

    select case right(x, 1) 'if oRec("mansat") = 3 then
    case 0,2,4,6,8
    bgthis = "#EFF3FF"
    case else
    bgthis = "#FFFFFF"
    end select

	%>
	<tr>
		<td bgcolor="#CCCCCC" colspan="7"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgthis %>">
		<td valign="top"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
		<td height="20"><a href="medarb_red.asp?menu=medarb&func=red&id=<%=oRec("Mid")%>"><%=oRec("mnavn")%> (<%=oRec("mnr") %>)</a></td>
        <td><%=oRec("init") %></td>
		<td><%
        
        select case oRec("mansat") 
        case "1"
        mstatus = ""
        case "2"
        mstatus = "Lukket"
        case "3"
        mstatus = "Passiv"
        case else
        mstatus = ""
        end select%>
        
        <%=mstatus %>
        </td>
        <td colspan=2>&nbsp;</td>
		<td valign="top" align=right><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
    x = x + 1
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="5" valign="bottom"><img src="../ill/blank.gif" width="587" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
	</tr>	
	</table>
	<br><br>
    </div>
    <br /><br />
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br>
<br>

 <%else 
    
     tTop = 20
	 tLeft = 0
	 tWdth = 800
	       
     call tableDiv(tTop,tLeft,tWdth)

	
    
    %>
      <table><tr><td>Du har ikke adgang til at se denne side.</td></tr></table></div>

    <%end if %>


	<%
	
	case "dbopr", "dbred"
	%>
	<!--#include file="inc/isint_func.asp"-->
	<%
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
		<%
		errortype = 8
		call showError(errortype)
		
		else
		
				function SQLBlessDOT(s)
				dim tmpdot
				tmpdot = s
				tmpdot = replace(tmpdot, ".", ",")
				SQLBlessDOT = tmpdot
				end function
				
				
				call erDetInt(SQLBlessDOT(request("FM_Timepris")))
				if isInt > 0 then
				%>
				<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
				<%
				errortype = 39
				call showError(errortype)
				isInt = 0
						
				else
					
				call erDetInt(SQLBlessDOT(request("FM_Kostpris")))
				if isInt > 0 then
				%>
				<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
				<%
				errortype = 39
				call showError(errortype)
				isInt = 0
						
				else
					
				
					if len(request("FM_Kostpris")) = 0 OR len(request("FM_timepris")) = 0 then
					%>
					<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					
					<%
					errortype = 39
					call showError(errortype)
					
					else
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_son")))
						int1 = isInt 
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_man")))
						int2 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tir")))
						int3 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_ons")))
						int4 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tor")))
						int5 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_fre")))
						int6 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_lor")))
						int7 = isInt
						
						isInt = 0
						call erDetInt(request("FM_Kostpris")) 
						int8 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris"))
						int9 = isInt
						
						'isInt = 0
						'call erDetInt(request("FM_timepris1"))
						'int10 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris2"))
						int11 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris3"))
						int12 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris4"))
						int13 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris5"))
						int14 = isInt

                        isInt = 0
						call erDetInt(request("FM_kostpristarif_A"))
						int15 = isInt

                        sInt = 0
						call erDetInt(request("FM_kostpristarif_B"))
						int16 = isInt

                        sInt = 0
						call erDetInt(request("FM_kostpristarif_C"))
						int17 = isInt

                        sInt = 0
						call erDetInt(request("FM_kostpristarif_D"))
						int18 = isInt

                        

						
						if int1 > 0 OR int2 > 0 OR int3 > 0 OR int4 > 0 OR int5 > 0 OR int6 > 0 OR int7 > 0 OR int8 > 0 OR int9 > 0 OR int11 > 0 OR int12 > 0 OR int13 > 0 OR int14 > 0 _
                        OR int15 > 0 OR int16 > 0 OR int17 > 0 OR int18 > 0 then
						%>
						<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					
						<%
						errortype = 39
						call showError(errortype)
						isInt = 0
					
						else
							
							
							
							function SQLBless(s)
							dim tmp
							tmp = s
							tmp = replace(tmp, "'", "''")
							SQLBless = tmp
							end function
							
							function SQLBlessDOT2(s)
							dim tmpdot
							tmpdot = s
							tmpdot = replace(tmpdot, ",", ".")
							SQLBlessDOT2 = tmpdot
							end function
							
							strNavn = SQLBless(request("FM_navn"))
							strTimepris = SQLBlessDOT2(request("FM_Timepris"))
							strTimepris1 = strTimepris 'SQLBlessDOT2(request("FM_timepris1"))
							strTimepris2 = SQLBlessDOT2(request("FM_timepris2"))
							strTimepris3 = SQLBlessDOT2(request("FM_timepris3"))
							strTimepris4 = SQLBlessDOT2(request("FM_timepris4"))
							strTimepris5 = SQLBlessDOT2(request("FM_timepris5"))
							
							
							if len(strTimepris) <> 0 then
							strTimepris = strTimepris
							else
							strTimepris = 0
							end if
							
							if len(strTimepris1) <> 0 then
							strTimepris1 = strTimepris1
							else
							strTimepris1 = 0
							end if
							
							if len(strTimepris2) <> 0 then
							strTimepris2 = strTimepris2
							else
							strTimepris2 = 0
							end if
							
							if len(strTimepris3) <> 0 then
							strTimepris3 = strTimepris3
							else
							strTimepris3 = 0
							end if
							
							if len(strTimepris4) <> 0 then
							strTimepris4 = strTimepris4
							else
							strTimepris4 = 0
							end if
							
							if len(strTimepris5) <> 0 then
							strTimepris5 = strTimepris5
							else
							strTimepris5 = 0
							end if
							
							
							dubKostpris = replace(request("FM_kostpris"), ".", "")
                            dubKostpris = replace(dubKostpris, ",", ".")


                            kostpristarif_A = replace(request("FM_kostpristarif_A"), ".", "")
                            kostpristarif_A = replace(kostpristarif_A, ",", ".")

                            kostpristarif_B = replace(request("FM_kostpristarif_B"), ".", "")
                            kostpristarif_B = replace(kostpristarif_B, ",", ".")

                            kostpristarif_C = replace(request("FM_kostpristarif_C"), ".", "")
                            kostpristarif_C = replace(kostpristarif_C, ",", ".")

                            kostpristarif_D = replace(request("FM_kostpristarif_D"), ".", "")
                            kostpristarif_D = replace(kostpristarif_D, ",", ".")

                            



							normtimer_son = SQLBlessDOT2(request("FM_norm_son"))
							normtimer_man = SQLBlessDOT2(request("FM_norm_man"))
							normtimer_tir =	SQLBlessDOT2(request("FM_norm_tir"))
							normtimer_ons = SQLBlessDOT2(request("FM_norm_ons"))
							normtimer_tor = SQLBlessDOT2(request("FM_norm_tor"))
							normtimer_fre = SQLBlessDOT2(request("FM_norm_fre"))
							normtimer_lor = SQLBlessDOT2(request("FM_norm_lor"))
							
							
							if len(dubKostpris) <> 0 then
							dubKostpris = dubKostpris
							else
							dubKostpris = 0
							end if

                            if len(kostpristarif_A) <> 0 then
							kostpristarif_A = kostpristarif_A
							else
							kostpristarif_A = 0
							end if
							
                            if len(kostpristarif_B) <> 0 then
							kostpristarif_B = kostpristarif_B
							else
							kostpristarif_B = 0
							end if

                            if len(kostpristarif_C) <> 0 then
							kostpristarif_C = kostpristarif_C
							else
							kostpristarif_C = 0
							end if

                            if len(kostpristarif_D) <> 0 then
							kostpristarif_D = kostpristarif_D
							else
							kostpristarif_D = 0
							end if
							 
							
							if len(normtimer_man) <> 0 then
							normtimer_man = normtimer_man
							else
							normtimer_man = 0
							end if
							
							if len(normtimer_tir) <> 0 then
							normtimer_tir = normtimer_tir
							else
							normtimer_tir = 0
							end if
							
							if len(normtimer_ons) <> 0 then
							normtimer_ons = normtimer_ons
							else
							normtimer_ons = 0
							end if
							
							if len(normtimer_tor) <> 0 then
							normtimer_tor = normtimer_tor
							else
							normtimer_tor = 0
							end if
							
							if len(normtimer_fre) <> 0 then
							normtimer_fre = normtimer_fre
							else
							normtimer_fre = 0
							end if
							
							if len(normtimer_lor) <> 0 then
							normtimer_lor = normtimer_lor
							else
							normtimer_lor = 0
							end if
							
							if len(normtimer_son) <> 0 then
							normtimer_son = normtimer_son
							else
							normtimer_son = 0
							end if
							
							tp0_valuta = request("FM_valuta_0")
							tp1_valuta = tp0_valuta
							tp2_valuta = request("FM_valuta_2")
							tp3_valuta = request("FM_valuta_3")
							tp4_valuta = request("FM_valuta_4")
							tp5_valuta = request("FM_valuta_5")
							

                            mtsortorder = request("FM_sortorder")
                            sostergp = request("FM_soster") 

                            if sostergp > 0 then
                                    
                                    findHgrpSQL = "SELECT mgruppe FROM medarbejdertyper WHERE id = " & sostergp
                                    
                                    'Response.write findHgrpSQL
                                    'Response.end 
                                    oRec5.open findHgrpSQL, oConn, 3
                                    if not oRec5.EOF then
                                    
                                    mgruppe = oRec5("mgruppe")        

                                    end if
                                    oRec5.close 
                            else
                                mgruppe = request("FM_gruppe")
                                if cint(mgruppe) = -1 OR cint(mgruppe) = 0 then
                                mgruppe = 1
                                end if
							end if


                            if len(trim(request("FM_afslutugekri_proc"))) <> 0 then 
                            afslutugekri_proc = request("FM_afslutugekri_proc")
                            else
                            afslutugekri_proc = 0
                            end if

                            if len(trim(request("FM_afslutugekri"))) <> 0 then
                            afslutugekri = request("FM_afslutugekri")
                            else
                            afslutugekri = 0
                            end if
							
                            afslutugekri = replace(afslutugekri, ",", ".")

							if request("FM_Kostpris") < 0 OR strTimepris < 0 OR strTimepris1 < 0 OR strTimepris2 < 0 OR strTimepris3 < 0 OR strTimepris4 < 0 OR strTimepris5 < 0 then
							%>
							<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
							
							<%
							errortype = 39
							call showError(errortype)
							
							else
		

                if len(trim(request("FM_noflex"))) <> 0 then
                noflex = 1
                else
                noflex = 0
                end if
				
				strEditor = session("user")
				strDato = session("dato")
				
				if func = "dbopr" then
				oConn.execute("INSERT INTO medarbejdertyper (type, timepris, editor, dato, kostpris, normtimer_son, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, "_
				&" timepris_a1, timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
				&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, sostergp, mtsortorder, mgruppe, afslutugekri, afslutugekri_proc, noflex, "_
                &" kostpristarif_A, kostpristarif_B, kostpristarif_C, kostpristarif_D) VALUES"_
				&" ('"& strNavn &"', "& strTimepris &", '"& strEditor &"', '"& strDato &"', "& dubKostpris &", "_
				&" "& normtimer_son &", "& normtimer_man &", "& normtimer_tir &", "& normtimer_ons &", "_
				&" "& normtimer_tor &", "& normtimer_fre &", "& normtimer_lor &", "& strTimepris1 &", "_
				&" "& strTimepris2 &", "& strTimepris3 &", "& strTimepris4 &", "& strTimepris5 &", "_
				&" "& tp0_valuta &","& tp1_valuta &","& tp2_valuta &","& tp3_valuta &","& tp4_valuta &","& tp5_valuta &", "& sostergp &", "_
                &" "& mtsortorder &", "& mgruppe &","& afslutugekri &","& afslutugekri_proc &", "& noflex &", "_
                &" "& kostpristarif_A &","& kostpristarif_B &","& kostpristarif_C &","& kostpristarif_D &""_
				&" )")
				else
				oConn.execute("UPDATE medarbejdertyper SET type ='"& strNavn &"', timepris = "& strTimepris &", editor = '" &strEditor &"', dato = '" & strDato &"', kostpris = "& dubKostpris &", "_
				&" normtimer_son = "& normtimer_son &", normtimer_man="& normtimer_man &", normtimer_tir="& normtimer_tir &", "_
				&" normtimer_ons="& normtimer_ons &", normtimer_tor="& normtimer_tor &", normtimer_fre="& normtimer_fre &", normtimer_lor="& normtimer_lor &", "_
				&" timepris_a1 = "& strTimepris1 &", timepris_a2 = "& strTimepris2 &", timepris_a3 = "& strTimepris3 &", "_
				&" timepris_a4 = "& strTimepris4 &", timepris_a5 = "& strTimepris5 &", "_
				&" tp0_valuta = "& tp0_valuta &", tp1_valuta = "& tp1_valuta &", "_
				&" tp2_valuta = "& tp2_valuta &", tp3_valuta = "& tp3_valuta &", "_
				&" tp4_valuta = "& tp4_valuta &", tp5_valuta = "& tp5_valuta &", sostergp = "& sostergp &", mtsortorder = "& mtsortorder &", "_
                &" mgruppe = "& mgruppe &", afslutugekri = "& afslutugekri &", afslutugekri_proc = "& afslutugekri_proc &", noflex = "& noflex &", "_
                &" kostpristarif_A = "& kostpristarif_A &", kostpristarif_B = "& kostpristarif_B &", kostpristarif_C = "& kostpristarif_C &", kostpristarif_D = "& kostpristarif_D &""_
				&" WHERE id = "&id&"")
				end if

               

                strSQLm = "SELECT m.mid, m.mnr, mth.mtypedato, mth.id, mth.mtype FROM medarbejdere AS m "_
                &" LEFT JOIN medarbejdertyper_historik AS mth ON (mth.mid = m.mid AND mth.mtype = m.medarbejdertype) WHERE m.medarbejdertype = "& id

                'Response.write strSQLm & "<br>"
                'Response.flush

                oRec.open strSQLm, oConn, 3
                While Not oRec.EOF 
                        
                        if IsNull(oRec("mtypedato")) = true then
                        fromDt = "2000-1-1"
                        else
                        fromDt = year(oRec("mtypedato")) &"/"& month(oRec("mtypedato")) &"/"& day(oRec("mtypedato"))
                        end if


                         '*** Opdater åbnejob med internkostpris **'
		                if len(trim(request("FM_opd_intern"))) <> 0 then

                            '*** Ignorer projektgruppe rel, da hvis medarbejderen har været med på jobbet, 
                            '****må det være via sin projektgruppe
                            '*** Ellers bliver timer ikke opdateret ****
                            strSQLj = "SELECT jobnr, a.id AS aid, a.kostpristarif FROM job AS j "_
                            &"LEFT JOIN aktiviteter AS a ON (a.job = j.id) WHERE jobstatus = 1" 
                           
                            'response.write "strSQLj: " & strSQLj & "<br>"    
                                
                            oRec2.open strSQLj, oConn, 3
                            While Not oRec2.EOF 


                             '*** Overskrifer kostpris med kostpristarif fra medab.type
                             kostpristarif = oRec2("kostpristarif")
                             select case kostpristarif
                             case "0"
                             dubKostprisUse = dubKostpris
                             case "A"
                             dubKostprisUse = kostpristarif_A
                             case "B"
                             dubKostprisUse = kostpristarif_B
                             case "C"
                             dubKostprisUse = kostpristarif_C
                             case "D"
                             dubKostprisUse = kostpristarif_D
                             case else
                             dubKostprisUse = dubKostpris 
                             end select

                            if isNull(oRec2("aid")) <> true  then 'AND dubKostpris <> ""
                        
                            strSQLt = "UPDATE timer SET kostpris = "& dubKostprisUse &" WHERE tjobnr = '" & oRec2("jobnr") & "' AND tmnr = " & oRec("mid") & " AND taktivitetid = "& oRec2("aid") &" AND tdato >= '"& fromDt &"'"
                            'response.write "kostprisTarif: " & oRec2("kostpristarif") & "<br>"
                            'Response.Write strSQLt & "<br>"
                            'Response.flush
                            oConn.execute(strSQLt)

                            end if

                            oRec2.movenext
                            wend
                            oRec2.close

                        end if

                            '*** Opdaterer stamaktiviteter med de nye priser ****'
                            '*** kun hvis det er tilvalgt på den specifikekgruppe '****

                            if len(trim(request("FM_opd_stamgrp"))) <> 0 then
                            opd_stamgrp = 1
                            else
                            opd_stamgrp = 0
                            end if


                            if cint(opd_stamgrp) = 1 then

                            
                            if len(trim(request("FM_opd_stamgrp_ids"))) <> 0 then
                            opd_stamgrp_ids = split(request("FM_opd_stamgrp_ids"), ", ") 

                            for a = 0 to UBOUND(opd_stamgrp_ids)

                            'Response.write "<hr>opd_stamgrp_ids(a): "& opd_stamgrp_ids(a) & "<br>"
                            'Response.flush

                            if len(trim(opd_stamgrp_ids(a))) <> 0 then

                            ct = 0
                            '*** henter alle aktiviteter i valgte stam-grupper ***'
						    strSQLtpe = "SELECT tp.id, tp.timeprisalt, tp.6timepris, a.aktfavorit, a.job FROM aktiviteter AS a "_
                            &" LEFT JOIN timepriser AS tp ON (tp.aktid = a.id AND tp.jobid = 0 AND tp.medarbid = "& oRec("mid") &") WHERE a.job = 0 AND aktfavorit = "& trim(opd_stamgrp_ids(a)) & " AND tp.aktid = a.id AND tp.jobid = 0 AND tp.medarbid = "& oRec("mid")
						    
                            'Response.write "strSQLtpe: "& strSQLtpe & "<br>"
    						
						    oRec5.open strSQLtpe, oConn, 3 
						    while not oRec5.EOF 
    						    

                                select case oRec5("timeprisalt")
                                case 1
                                tpThis = strTimepris1
                                valThis = tp1_valuta
                                case 2
                                tpThis = strTimepris2
                                valThis = tp2_valuta
                                case 3
                                tpThis = strTimepris3
                                valThis = tp3_valuta
                                case 4
                                tpThis = strTimepris4
                                valThis = tp4_valuta
                                case 5
                                tpThis = strTimepris5
                                valThis = tp5_valuta
                                case 6
                                tpThis = strTimepris
                                valThis = tp0_valuta
                                case else
                                tpThis = strTimepris
                                valThis = tp0_valuta
                                end select

                                if isNull(oRec5("id")) <> true then

							    strSQLinsTp = "UPDATE timepriser SET 6timepris = "& tpThis &", 6valuta = "& valThis &" WHERE id = " & oRec5("id")

                                'Response.Write strSQLinsTp  & "<br>" '& oRec5("timeprisalt")
                                'Response.flush

                                oConn.execute(strSQLinsTp)

                                end if
    						    
                                ct = ct + 1
						    oRec5.movenext
						    wend
						    oRec5.close 


                            end if 'len

                            next

                            end if

                            end if'opddater stamgrp

                               
                oRec.movenext
                wend
                oRec.close

                
                '**** Opdater hovedgruppe på medarbejderlinier
                                if len(trim(request("FM_gruppe_opr"))) <> 0 then
                                oprHovedgruppe = request("FM_gruppe_opr")
                                else 
                                oprHovedgruppe = 0
                                end if
                


                nyHovedgruppe = mgruppe 'request("FM_gruppe")

                strSQlmed = "UPDATE medarbejdere SET medarbejdertype_grp = "& nyHovedgruppe &" WHERE medarbejdertype_grp = "& oprHovedgruppe
                
                'Response.write strSQlmed
                                
                oConn.execute(strSQlmed)
              


                'Response.end
				
                
                Response.redirect "medarbtyper.asp?menu=medarb"
				end if 'validering
				end if 'validering
				end if 'validering
				end if 'validering
			end if 'validering
		end if 'validering
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	strTimepris = 0
	dubKostpris = 0
	strTimepris1 = 0
	strTimepris2 = 0
	strTimepris3 = 0
	strTimepris4 = 0
	strTimepris5 = 0
	
	normtimer_son = 0
	normtimer_man = "7,5"
	normtimer_tir = "7,5"
	normtimer_ons = "7,5"
	normtimer_tor = "7,5"
	normtimer_fre = "7"
	normtimer_lor = 0
	
    sostergp = 0
    mtsortorder = 1000
    noflex = 0



    kostpristarif_A = 0
    kostpristarif_B = 0
    kostpristarif_C = 0
    kostpristarif_D = 0

	else
	strSQL = "SELECT type, editor, dato, timepris, kostpris, normtimer_son, normtimer_man, "_
	&" normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, timepris_a1, "_
	&" timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
	&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, sostergp, mtsortorder, mgruppe, afslutugekri, afslutugekri_proc, noflex, "_
    &" kostpristarif_A, kostpristarif_B, kostpristarif_C, kostpristarif_D "_
	&" FROM medarbejdertyper WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("type")
	strTimepris = oRec("timepris")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	dubKostpris = oRec("kostpris")
	normtimer_son = oRec("normtimer_son")
	normtimer_man = oRec("normtimer_man")
	normtimer_tir = oRec("normtimer_tir")
	normtimer_ons = oRec("normtimer_ons")
	normtimer_tor = oRec("normtimer_tor")
	normtimer_fre = oRec("normtimer_fre")
	normtimer_lor  = oRec("normtimer_lor")
	strTimepris1 = oRec("timepris_a1")
	strTimepris2 = oRec("timepris_a2")
	strTimepris3 = oRec("timepris_a3")
	strTimepris4 = oRec("timepris_a4")
	strTimepris5 = oRec("timepris_a5")
	
	tp0_valuta = oRec("tp0_valuta")
	tp1_valuta = oRec("tp1_valuta")
	tp2_valuta = oRec("tp2_valuta")
	tp3_valuta = oRec("tp3_valuta")
	tp4_valuta = oRec("tp4_valuta")
	tp5_valuta = oRec("tp5_valuta")

    sostergp = oRec("sostergp")
    mtsortorder = oRec("mtsortorder")

    mgruppe = oRec("mgruppe")

    afslutugekri_proc = oRec("afslutugekri_proc")
	afslutugekri = oRec("afslutugekri")

    noflex = oRec("noflex")

    kostpristarif_A = oRec("kostpristarif_A")
    kostpristarif_B = oRec("kostpristarif_B")
    kostpristarif_C = oRec("kostpristarif_C")
    kostpristarif_D = oRec("kostpristarif_D")

	end if
	oRec.close
	
	
    if cint(noflex) <> 0 then
    noflexCHK = "CHECKED"
    else
    noflexCHK = ""
    end if


	if len(trim(dubKostpris)) <> 0 then
	dubKostpris = dubKostpris
	else
	dubKostpris = 0
	end if
	
	
	ugetotal = formatnumber(normtimer_son + normtimer_man + normtimer_tir + normtimer_ons + normtimer_tor + normtimer_fre + normtimer_lor, 2)
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->


  <script src="inc/medarbtyper_jav.js"></script>
	
	
    <%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px;">
    
    <% 
    'oimg = "ikon_medarb_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Medarbejdertyper"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)

    if level = 1 then
    %>
	
	<form action="medarbtyper.asp?menu=medarb&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
        <br />
	<table cellspacing="0" cellpadding="0" border="0" width="500" bgcolor="#FFFFFF">
	
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan=2 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
    	<td colspan="2" class=alt><b><%=varbroedkrumme%> medarbejdertype</b><br />
		<%if dbfunc = "dbred" then%>
		Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b>
		<%end if%>
		</td>
	</tr>
	<tr>
		<td valign="top" ><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
		<td><br><b>Navn:</b></td>
		<td><br><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:350px;"></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="130" alt="" border="0"></td>
		<td valign="top" style="padding-top:3;"><b>Normeret tid: </b></td>
		<td valign="top">
		<table cellspacing=3 cellpadding=0 border=0><tr>
		<td>
		Mandag</td><td>
		Tirsdag</td><td>
		Onsdag</td><td>
		Torsdag</td><td>
		Fredag</td><td>
		Lørdag</td>
		<td>Søndag</td></tr>
		<tr>
		<td><input type="text" name="FM_norm_man" value="<%=normtimer_man%>" size="3" >
		</td><td><input type="text" name="FM_norm_tir" value="<%=normtimer_tir%>" size="3" >
		</td><td><input type="text" name="FM_norm_ons" value="<%=normtimer_ons%>" size="3" >
		</td><td><input type="text" name="FM_norm_tor" value="<%=normtimer_tor%>" size="3" >
		</td><td><input type="text" name="FM_norm_fre" value="<%=normtimer_fre%>" size="3" >
		</td><td><input type="text" name="FM_norm_lor" value="<%=normtimer_lor%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: darkred; border-style: solid;">
		</td>
		<td><input type="text" name="FM_norm_son" value="<%=normtimer_son%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: darkred; border-style: solid;">
		</td>
		</tr>
		<tr><td colspan=7>Timer&nbsp;&nbsp;<b>Total:</b>&nbsp;<u><%=ugetotal%></u><br />
            <input type="checkbox" name="FM_noflex" value="<%=noflex %>" <%=noflexChk %> /> Vis/beregn ikke flekssaldo for denne type.
            <br /><br />&nbsp;
		    </td></tr>
		</table></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="130" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
		<td valign="top"><b>Timepris 1:</b><br />
		(generel timepris)</td>
		<td valign="top"><input type="text" name="FM_timepris" value="<%=strTimepris%>" size="10" style="border:1px #86B5E4 solid; background-color:#FFFFFF;">&nbsp;
		<%call valutaKoder(0, tp0_valuta, 1) %>
	    </td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
	</tr>
	
	<!--
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
		<td valign="top" style="padding-top:5px;"><b>Timepris 1:</b></td>
		<td><input type="hidden" name="FM_timepris1" value="<%=strTimepris1%>" size="10" >&nbsp;
		<%'call valutaKoder(1, valuta, 1) %><br />
		 (sæt alt. timepris 1 = generel tp. pga. stamakt) &nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
	</tr>
	-->
    <tr>    <td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
        <td colspan="2">Angiv alternative timepriser på denne medarbejdertype, der kan for-vælges på aktivitets-skabeloner eller på job.</td>
        	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 2:</b></td>
		<td><input type="text" name="FM_timepris2" value="<%=strTimepris2%>" size="10" >&nbsp;
		<%call valutaKoder(2, tp2_valuta, 1) %>&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 3:</b></td>
		<td><input type="text" name="FM_timepris3" value="<%=strTimepris3%>" size="10" >&nbsp;
		<%call valutaKoder(3, tp3_valuta, 1) %>&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 4:</b></td>
		<td><input type="text" name="FM_timepris4" value="<%=strTimepris4%>" size="10" >&nbsp;
		<%call valutaKoder(4, tp4_valuta, 1) %>
		&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 5:</b></td>
		<td><input type="text" name="FM_timepris5" value="<%=strTimepris5%>" size="10" >&nbsp;
		<%call valutaKoder(5, tp5_valuta, 1) %>
		&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
    <%if func = "red" then %>
    <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="200" alt="" border="0"></td>
		<td valign="top"><br><br /><b>Opdater stam-aktivitetsgrupper</b></td>
		<td valign="top"><br /><br /><input id="Checkbox2" type="checkbox" name="FM_opd_stamgrp" value="1" /> Ja, opdater timepriserne for alle medarbejdere af denne type på alle aktiviteterne, i nedenstående stamaktivitets-grupper: 
        <br /><br />
        
        <select name="FM_opd_stamgrp_ids" multiple styke="width:250px;" size=5>
        
        <%
        strSQL = "SELECT navn, id FROM akt_gruppe WHERE id <> 0 AND skabelontype = 0 ORDER BY navn" 
        oRec3.open strSQL, oConn, 3
        while not oRec3.EOF 
        %>
        <option value="<%=oRec3("id") %>"><%=oRec3("navn") %></option>
        <%
        oRec3.movenext
        wend 
        oRec3.close
        %>
        </select><br /><br />&nbsp;
        </td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="200" alt="" border="0"></td>
	</tr>
    <%end if %>

        <%if func = "red" then
            kphgt = "95"
          else
            kphgt = "55"
          end if %>

	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
		<td valign="top"><br /><br /><b>Intern kostpris:</b></td>
		<td valign="top"><br /><br /><input type="text" name="FM_kostpris" value="<%=dubKostpris%>" size="10" >&nbsp;<%=basisValISO %></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
	</tr>
     <tr>    <td valign="top"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
        <td colspan="2"><br />Angiv kostpris tariffer der kan bruges som tillæg til kostpris på specielle aktiviteter, aften og nat aktiviteter mm.</td>
        	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
	</tr>
    <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
        <td valign="top">Kostpris tarif A  </td>
        <td valign="top"><input type="text" value="<%=kostpristarif_A %>" style="width:60px;" name="FM_kostpristarif_A" /> &nbsp;<%=basisValISO %></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
    <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
        <td valign="top">Kostpris tarif B  </td>
        <td valign="top"><input type="text" value="<%=kostpristarif_B %>" style="width:60px;" name="FM_kostpristarif_B" /> &nbsp;<%=basisValISO %></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
        <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
        <td valign="top">Kostpris tarif C  </td>
        <td valign="top"><input type="text" value="<%=kostpristarif_C %>" style="width:60px;" name="FM_kostpristarif_C" /> &nbsp;<%=basisValISO %></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
    <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="<%=kphgt %>" alt="" border="0"></td>
        <td valign="top">Kostpris tarif D  </td>
        <td valign="top"><input type="text" value="<%=kostpristarif_D %>" style="width:60px;" name="FM_kostpristarif_D" /> &nbsp;<%=basisValISO %>


            <%if func = "red" then %>
       <br /> <br /><input id="Checkbox1" type="checkbox" name="FM_opd_intern" value="1" /> Opdater interne kostpriser/tariffer, for denne type, <br />på alle eksisterende <u>åbne</u> job.
        <%end if %>
        </td>



		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="<%=kphgt %>" alt="" border="0"></td>
	</tr>


         <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
		<td valign="top" style="padding-right:10px;"><br /><br /><b>Søstertype:</b>
             
            
            <%
                qmtxt = "Skal denne medarbejdertype medregnes under en anden medarbejder-type i budget og omsætning?<br />Hovedgruppe følger evt. valgte søstergrp."
                qmtop = 600
                qmleft = 100
                qmwdt = 200
                qmid = "qmMT01"
            %>
            <span class="qmarkhelp" id="<%=qmid %>" style="font-size:11px; color:#999999;"><u>?</u></span>
            <%
               
                call qmarkhelpnote(qmtxt,qmtop,qmleft,qmid,qmwdt) %>

         
               
</td>
		<td valign="top"><br /><br /><select name="FM_soster" id="FM_soster" style="width:350px;">
        
        
        <%  if cint(sostergp) = 0 then
        mttypSEL = "SELECTED"
        else
        mttypSEL = ""
        end if  %>

        <option value="0" <%=mttypSEL  %>>Nej</option>

        <%
        
        if len(id) <> 0 then
        tyid = id
        else
        tyid = 0
        end if
        
        strSQLmt = "SELECT id, type FROM medarbejdertyper WHERE id <>  "& tyid & " ORDER BY type"
        oRec3.open strSQLmt, oConn, 3
        while not oRec3.EOF

        if cint(sostergp) = oRec3("id") then
        mttypSEL = "SELECTED"
        else
        mttypSEL = ""
        end if 
        %>
         <option value="<%=oRec3("id") %>" <%=mttypSEL %>><%=oRec3("type") %></option>
        <%
        oRec3.movenext
        wend
        oRec3.close  %>

        <%  if cint(sostergp) = -1 then
        mttypSEL = "SELECTED"
        else
        mttypSEL = ""
        end if  %>

        <option value="-1" <%=mttypSEL %>>Vis IKKE denne type under budget, statistik og forecast</option> <!-- -->

        </select>
		&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
	</tr>

      <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
		<td style="padding-right:10px; padding-top:20px;" valign=top><b>Hovedgruppe:</b>

           

                <%
                qmtxt = "Budget og omsætning kan samles på hovedgruppen."
                qmtop = 650
                qmleft = 100
                qmwdt = 200
                qmid = "qmMT02"
            %>
            <span class="qmarkhelp" id="<%=qmid %>" style="font-size:11px; color:#999999;"><u>?</u></span>
            <%
               
                call qmarkhelpnote(qmtxt,qmtop,qmleft,qmid,qmwdt) %>

        <td valign="top" style="padding-top:20px;"><select name="FM_gruppe" id="FM_gruppe" style="width:350px;">
        
        
       <%
        
        mttypSELid = 1
        strSQLmt = "SELECT id, navn FROM medarbtyper_grp WHERE id <> 0 ORDER BY navn"
        oRec3.open strSQLmt, oConn, 3
        while not oRec3.EOF

        if cint(mgruppe) = oRec3("id") then
        mttypSEL = "SELECTED"
        mttypSELid = oRec3("id")
        else
        mttypSEL = ""
        end if 
        %>
         <option value="<%=oRec3("id") %>" <%=mttypSEL %>><%=oRec3("navn") %></option>
        <%
        oRec3.movenext
        wend
        oRec3.close
            
            
            
        if cint(sostergp) > 0 then
        mttypSEL = "SELECTED"
        else
        mttypSEL = ""
        end if  %>

        <option value="-1" <%=mttypSEL %> style="background-color:#CCCCCC;">// Følger søstergruppe //</option> 

    
        </select> 
		&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>

          <%if sostergp <= 0 then
            mgruppe_opr = mgruppe
            else
                findHgrpSQL = "SELECT mgruppe FROM medarbejdertyper WHERE id = " & sostergp
                                    
                                    'Response.write findHgrpSQL
                                    'Response.end 
                                    oRec5.open findHgrpSQL, oConn, 3
                                    if not oRec5.EOF then
                                    
                                    mgruppe_opr = oRec5("mgruppe")        

                                    end if
                                    oRec5.close 
              
              
              end if %>


          <input type="hidden" value="<%=mgruppe_opr %>" name="FM_gruppe_opr" />
	</tr>


   

     <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Sortering:</b></td>
		<td><input type="text" style="width:40px;" value="<%=mtsortorder%>" name="FM_sortorder" id="FM_sortorder"></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
    </tr>

        <%
        call ersmileyaktiv() 
        call erStempelurOn()   
            
        if cint(smilaktiv) = 1 then 
          
            afslutugekri0 = ""
            afslutugekri1 = ""
            afslutugekri2 = ""

            select case afslutugekri
            case 1
            afslutugekri1 = "SELECTED"
            case 2
            afslutugekri2 = "SELECTED"
            case else
            afslutugekri0 = "SELECTED"
            end select
            
          %>
         <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
		<td valign="top" style="padding-top:26px;"><b>Afslut uge kriterie:</b>
		</td>
		<td><br />Minimumskrav til timeantal<br />Denne medarbejdertype kan afslutte sin uge når:<br />
             
        <select name="FM_afslutugekri">
                <option value="0" <%=afslutugekri0 %>>Ingen (kan altid afslutte)</option>
                <option value="1" <%=afslutugekri1 %>>Realiserede timer</option>
                <option value="2" <%=afslutugekri2 %>>Real. fakturerbare timer</option>
                
		    </select> udgør minimun <input type="text" style="width:40px;" value="<%=afslutugekri_proc %>" name="FM_afslutugekri_proc"> % 
            <%select case lto
               case "dencker"
                %>
                af komme/gå tid.
                <%
               case else  %>
            af normtid. 
            <%end select %>
		</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
    </tr>
        <%end if%>
    <tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="150" alt="" border="0"></td>
		<td valign="top" align=center colspan=2>
        <br />
        <br />
        <br /><br />
		
            
		<input type="submit" value="Opdater >>" /><br><br>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="150" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>	
	
	</table>
	</form>



	<div style="position:absolute; width:200px; left:560px; top:30px;">

     <% 
        uWdt = 300
        uTxt = "<b>Generel timepris</b> <br />"_
		&"Den generelle timepris er den timepris der automatisk bliver tildelt på et job "_
		&"ved joboprettelse. De andre timepriser kan vælges ved redigering af job eller "_
		&"de kan forvælges på stam-aktiviteter så den bliver forvalgt på den enkelte stam-aktivitet."
        uTxt = uTxt & "<br><br><b>Husk at tilpasse ferie optjent</b> for medarbejdere af denne type, hvis der ændres medarbejdertype / normtid."
        call infoUnisport(uWdt, uTxt)
        %>


		
       
	</div>
	
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>

    <%else 
    
     tTop = 20
	 tLeft = 0
	 tWdth = 800
	       
     call tableDiv(tTop,tLeft,tWdth)

	
    
    %>
      <table><tr><td>Du har ikke adgang til at se denne side.</td></tr></table></div>

    <%end if %>

	</div>
	<%case else%>
	
	

	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:52px;">
	
	
	
	<%
    

    if level = 1 then
	
	'oleftpx = 0
	'otoppx = 20
	'owdtpx = 190
	
	'call opretNy("medarbtyper.asp?menu=medarb&func=opret", "Opret ny medarbejder-type", otoppx, oleftpx, owdtpx) 
	%>

          



	<%
	tTop = 50
	tLeft = 0
	tWdth = 940
	
	
	call tableDiv(tTop,tLeft,tWdth)



        'oimg = "ikon_medarb_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Medarbejdertyper"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
	
	%><br />

          <span style="position:relative; left:747px; top:-30px;"><% 
                nWdt = 180
                nTxt = "Opret ny medarbejder-type"
                nLnk = "medarbtyper.asp?menu=medarb&func=opret"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %> 
	        </span>

	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan="10" valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td height="30" class=alt style="width:50px;"><a href="medarbtyper.asp?menu=medarb&sort=nr" class=alt><b>Id</b></a></td>
        <td class=alt style="width:70px;">Sortering</td>
		<td class=alt style="padding:2px 10px 2px 20px;"><a href="medarbtyper.asp?menu=medarb&sort=navn" class=alt><b>Navn</b></a></td>
        <td class=alt>Hovedgruppe</td>
        <td class=alt>Søstergrp?</td>
		<td align="left" class=alt><b>Medlemmer</b> (aktuelt antal, incl. passive) (historisk)</td>
		<td class=alt align=right>Generel timepris</td>
		<td class=alt align=right>Intern Kostpris</td>
		<td class=alt align=right>Normeret timer pr. uge<br /> (Gns. pr. dag)</td>
		<td>&nbsp;&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "nr" then
	odrBy = " ORDER BY id"
	else
	odrBy = " ORDER BY mgruppe, mtsortorder, type"
	end if
	
	strSQL = "SELECT mt.id, type, timepris, kostpris, normtimer_son, "_
	&" normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, "_
	&" normtimer_fre, normtimer_lor, tp0_valuta, v.valutakode, mtsortorder, sostergp, mgruppe, mtg.navn AS mtgnavn FROM medarbejdertyper mt "_
	&" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mt.mgruppe)"_
    &" LEFT JOIN valutaer v ON (v.id = tp0_valuta) "_
	&""& odrBy
	
	
	oRec.open strSQL, oConn, 3
	
	x = 0
	c = 0
	ugetotal = 0
	while not oRec.EOF
	ugetotal = formatnumber(oRec("normtimer_son") + oRec("normtimer_man") + oRec("normtimer_tir") + oRec("normtimer_ons") + oRec("normtimer_tor") + oRec("normtimer_fre") + oRec("normtimer_lor"), 2)
	
	'** Antal medab i type ***'
	strSQL2 = "SELECT Mid FROM medarbejdere WHERE Medarbejdertype = "& oRec("id") & " AND mansat <> 2 "
	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF 
	x = x + 1
	oRec2.movenext
	wend
	oRec2.close
	Antal = x
	
	
	t = 0
	'** Antal medarb i medarbtype historik ***'
	strSQL2 = "SELECT COUNT(mtype) AS antalhistorik FROM medarbejdertyper_historik WHERE mtype = "& oRec("id") & " GROUP BY mtype, mid"
	
	'Response.write strSQL2
	'Response.flush
	oRec2.open strSQL2, oConn, 3
	if not oRec2.EOF then
	t = cint(oRec2("antalhistorik"))
	end if
	oRec2.close
	
	if t <> 0 then
	t = t 
	else
	t = 0
	end if
	
	
	
	select case right(c,1)
	case 0,2,4,6,8
	trbg = "#EFF3FF"
	case else
	trbg = "#FFFFFF"
	end select%>
	<tr>
		<td bgcolor="#CCCCCC" colspan="12"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=trbg%>">
		<td valign="top"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
		<td><%=oRec("id")%></td>
		<td><%=oRec("mtsortorder")%></td>
        <td height="20" style="padding:2px 10px 2px 20px;">&nbsp;&nbsp;<a href="medarbtyper.asp?menu=medarb&func=red&id=<%=oRec("id")%>"><%=oRec("type")%> </a></td>
        <td>
        
           <%if oRec("mgruppe") <> 0 then %>
        <span style="font-size:9px;"><%=left(oRec("mtgnavn"), 20) %></span>
        <%end if %>
        </td>
        <td align=center>
        <%if oRec("sostergp") <> 0 then %>
        (id: <%=oRec("sostergp") %>)
        <%end if %>
        </td>
		<td><a href="medarbtyper.asp?menu=medarb&func=med&id=<%=oRec("id")%>" class=vmenuglobal>Se medlemmer</a> <b>(<%=antal%>)</b> (<%=t %>)</td>
		<td align=right><%=oRec("timepris")%>
            &nbsp;<%=oRec("valutakode") %></td>
		<td align=right><%=oRec("kostpris") &" "& basisValISO%></td>
		<td align=right><%=formatnumber(ugetotal)%> timer (<%=formatnumber(ugetotal/5, 1)%>)</td>
		<%if x > 0 OR t > 0 then%>
		<td>&nbsp;</td>
		<%else%>
		<td style="padding-left:10px;"><a href="medarbtyper.asp?menu=medarb&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" width="16" height="16" alt="" border="0"></a></td>
		<%end if%>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	c = c + 1
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="10" valign="bottom"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
	</tr>		
	</table>
	
	</div><!-- table div-->
	<br /><br />
	
	      <% 
       itop = 50
       ileft = 0 
       iwdt = 400
       call sideinfo(itop,ileft,iwdt)%>
	Gns. timer pr. dag er baseret på ugetotal / 5 arbejdsdage.<br />
	Bruges bl.a til at angive ferie dage.<br /> Gns. timer pr. dag = 1 ferie dag.
	<br><br>
	&nbsp;
	  <!-- side info slut -->
        </td></tr></table>
			</div>
			
		
	<br><br>
	<br><br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br><br>

    <%
else

	
	 tTop = 20
	 tLeft = 0
	 tWdth = 800
	       
     call tableDiv(tTop,tLeft,tWdth)

	
  
	%>
    <table><tr><td>Du har ikke adgang til at se denne side.</td></tr></table></div>

    <%

 end if 'level
 %>

	</div><!-- side div -->	
	<br>
	<br><br>
	<br>
&nbsp;
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
