<%response.buffer = true 
Session.LCID = 1030
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>


<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
 '** ER SESSION UDLØBET  ****
    if len(session("user")) = 0 then
    errortype = 5
	call showError(errortype)
    response.end
    end if
 %>


<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->







<%

    if media <> "print" then    
        call menu_2014
    end if
    

    %>

<div class="wrapper">
    <div class="content">
        

<%

'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
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
	<!--Slet sidens indhold-->
    <div class="container" style="width:500px;">
    <div class="porlet">
            
            <h3 class="portlet-title">
               Medarbejdertype
            </h3>
            
                <div class="portlet-body">
                    <div style="text-align:center;">Du er ved at <b>SLETTE</b> en medarbejdertype. Er dette korrekt?
                    </div><br />
                  
                        <div style="text-align:center;"><a class="btn btn-primary btn-sm" role="button" href="medarbtyper.asp?menu=tok&func=sletok&id=<%=id%>">&nbsp Ja &nbsp</a>&nbsp&nbsp&nbsp&nbsp<a class="btn btn-default btn-sm" role="button" href="Javascript:history.back()">Nej</a>
                    </div>
                    <br /><br />
                 </div>
    </div>
    </div>


<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM medarbejdertyper WHERE id = "& id &"")
	Response.redirect "medarbtyper.asp?menu=medarb"
	
    case "med"
%>
            <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

            <%
                if level = 1 then

                strSQL = "SELECT id, type FROM medarbejdertyper WHERE id=" & id 
	            oRec.open strSQL, oConn, 3
	            if not oRec.EOF then
	            gruppeNavn = oRec("type")
	            end if
	            oRec.close
	        %>

<!--Medlemmer-->
<div class="container">
<div class="portlet">
    <h3 class="portlet-title"><u>Medarbejdertyper</u></h3>

    <div>Aktive og passive medarbejdere af typen <b><%=gruppeNavn%></b></div><br />
    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
        <thead>
            <tr>
                <th>Navn</th>
                <th>Intialer</th>
                <th>Status</th>
            </tr>
        </thead>
        <tbody>
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
                <td><a href="medarb.asp?menu=medarb&func=red&id=<%=oRec("Mid")%>"><%=oRec("mnavn")%> (<%=oRec("mnr") %>)</a></td>
                <td><%=oRec("init") %></td>
                <td>
                    <%      
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
            </tr>
            <%
                oRec.movenext
                wend 
            %>
        </tbody>
    </table>
</div>
        <%else%> 
    	<div>Du har ikke adgang til denne side</div>      
        <%end if %>

	
    
      

<%

case "dbopr", "dbred"
	%>
	<!--#include file="../timereg/inc/isint_func.asp"-->
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
				
                
                        strSQlast = "SELECT id FROM medarbejdertyper WHERE id <> 0 ORDER BY id DESC"
                        oRec6.open strSQlast, oConn, 3
                        if not oRec6.EOF then

                                lastId = oRec6("id")

                        end if
                        oRec6.close
                                
                else

                strSQLupd = "UPDATE medarbejdertyper SET type ='"& strNavn &"', timepris = "& strTimepris &", editor = '" &strEditor &"', dato = '" & strDato &"', kostpris = "& dubKostpris &", "_
				&" normtimer_son = "& normtimer_son &", normtimer_man="& normtimer_man &", normtimer_tir="& normtimer_tir &", "_
				&" normtimer_ons="& normtimer_ons &", normtimer_tor="& normtimer_tor &", normtimer_fre="& normtimer_fre &", normtimer_lor="& normtimer_lor &", "_
				&" timepris_a1 = "& strTimepris1 &", timepris_a2 = "& strTimepris2 &", timepris_a3 = "& strTimepris3 &", "_
				&" timepris_a4 = "& strTimepris4 &", timepris_a5 = "& strTimepris5 &", "_
				&" tp0_valuta = "& tp0_valuta &", tp1_valuta = "& tp1_valuta &", "_
				&" tp2_valuta = "& tp2_valuta &", tp3_valuta = "& tp3_valuta &", "_
				&" tp4_valuta = "& tp4_valuta &", tp5_valuta = "& tp5_valuta &", sostergp = "& sostergp &", mtsortorder = "& mtsortorder &", "_
                &" mgruppe = "& mgruppe &", afslutugekri = "& afslutugekri &", afslutugekri_proc = "& afslutugekri_proc &", noflex = "& noflex &", "_
                &" kostpristarif_A = "& kostpristarif_A &", kostpristarif_B = "& kostpristarif_B &", kostpristarif_C = "& kostpristarif_C &", kostpristarif_D = "& kostpristarif_D &""_
				&" WHERE id = "&id&""

                'response.write strSQLupd
                'response.Flush

				oConn.execute(strSQLupd)


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
              
                '**********************************************************************'
                '*** Opdatrerer målsætninger på medarbejdertype **********************

                if func = "dbopr" then
                id = lastId
                end if

                'response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;..........................................."& request("FM_fomr_mal")
                'response.flush


                fomr_mal_arr = split(request("FM_fomr_mal"), ", ##,")
                fomr_mal_arr_id = split(request("FM_fomr_mal_fomrid"), ", ")   
                
                for f = 0 TO UBOUND (fomr_mal_arr_id)

                    if f = 0 then
                    strSQLDeleteFomrMal = "DELETE FROM mtype_mal_fordel_fomr WHERE mmff_mtype = "& id
                    oConn.execute(strSQLDeleteFomrMal)
                    end if
                    
                    fomr_mal_arr(f) = replace(fomr_mal_arr(f), "#", "")
                    if len(trim(fomr_mal_arr(f))) <> 0 AND fomr_mal_arr(f) > 0 then
                    
                    fomr_mal_arr(f) = replace(fomr_mal_arr(f), ",", ".") 
                    strSQLInsertFomrMal = "INSERT INTO mtype_mal_fordel_fomr SET mmff_mal = "& fomr_mal_arr(f) & ", mmff_mtype = "& id &", mmff_fomr = "& fomr_mal_arr_id(f)
                    oConn.execute(strSQLInsertFomrMal)

                    end if

                next
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

    if len(trim(request("mtypenavnforvlgt"))) <> 0 then 'mtype = 1:1 med medarbejder
    strNavn = request("mtypenavnforvlgt")
    else
	strNavn = ""
    end if
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
	

    <!--Rediger sidensinhold-->
            <%if level = 1 then %>
    <div class="container">
    <div class="portlet">
        <h3 class="portlet-title"><u>Medarbejdertype rediger</u></h3>

        <form action="medarbtyper.asp?menu=medarb&func=<%=dbfunc%>" method="post">
	        <input type="hidden" name="id" value="<%=id%>">
    
                <div class="row">
                <div class="col-lg-10">&nbsp</div>
                <div class="col-lg-2 pad-b10">
                <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                </div>
                </div>

        <div class="portlet-body">
            <div class="well well-white">
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2">Navn:</div>
                <div class="col-lg-2"><input class="form-control input-small" type="text" name="FM_navn" value="<%=strNavn%>" style="width:350px;"></div>
            </div>
             <div class="row"><div class="pad-b10"></div></div>
                <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">Nomeret tid:</div>
                    <div class="col-lg-1">Mandag <br />
                        <input type="text" class="form-control input-small" name="FM_norm_man" value="<%=normtimer_man%>">
                    </div>
                    <div class="col-lg-1">Tirsdag <br />
                        <input type="text" class="form-control input-small" name="FM_norm_tir" value="<%=normtimer_tir%>">
                    </div>
                    <div class="col-lg-1">Onsdag <br />
                        <input type="text" class="form-control input-small" name="FM_norm_ons" value="<%=normtimer_ons%>">
                    </div>
                    <div class="col-lg-1">Torsdag <br />
                        <input type="text" class="form-control input-small" name="FM_norm_tor" value="<%=normtimer_tor%>">
                    </div>
                    <div class="col-lg-1">Fredag <br />
                        <input type="text" class="form-control input-small" name="FM_norm_fre" value="<%=normtimer_fre%>">
                    </div>
                    <div class="col-lg-1">Lørdag <br />
                        <input type="text" class="form-control input-small" name="FM_norm_lor" value="<%=normtimer_lor%>">
                    </div>
                    <div class="col-lg-1">Søndag <br />
                        <input type="text" class="form-control input-small" name="FM_norm_son" value="<%=normtimer_son%>">
                    </div>
                     <div class="col-lg-1"><br /><b>Total:</b>&nbsp;<u><%=ugetotal%></u></div>
                </div>

                  <div class="row">
                        <div class="col-lg-8">&nbsp;</div>
                          <div class="col-lg-4"><input type="checkbox" name="FM_noflex" value="<%=noflex %>" <%=noflexChk %> /> Vis/beregn ikke flekssaldo for denne type.
                       </div>

                      </div>

            </div>
            <div class="row"><div class="pad-b10"></div></div>

            <section>
                <!-- Accordion -->
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <!-- PersonData -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">Time- og kost -priser</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">

                        <div class="panel-body">
                           
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">TimePris (Alt. Timepris 1:):</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris" value="<%=strTimepris%>"></div>
                                <div class="col-lg-2"> <%call valutaKoder(0, tp0_valuta, 1) %></div>
                            </div>
                            <div class="row"><div class="pad-b10"></div></div>

                            <div class="row">
                                <div class="col-lg-1 pad-t10">&nbsp</div>
                                <div class="col-lg-2 pad-t10"><h6>Alt. Timepris</h6></div>
                            </div>
                            <!--
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Alt. Timepris 1:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris1" value="<%=strTimepris1%>" disabled></div>
		                        <div class="col-lg-2"><%call valutaKoder(1, tp1_valuta, 1) %></div>
                            </div>
                            -->
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Alt. Timepris 2:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris2" value="<%=strTimepris2%>"></div>
		                        <div class="col-lg-2"><%call valutaKoder(2, tp2_valuta, 1) %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Alt. Timepris 3:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris3" value="<%=strTimepris3%>"></div>
		                        <div class="col-lg-2"><%call valutaKoder(3, tp3_valuta, 1) %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Alt. Timepris 4:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris4" value="<%=strTimepris4%>"></div>
		                        <div class="col-lg-2"><%call valutaKoder(4, tp4_valuta, 1) %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Alt. Timepris 5:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris5" value="<%=strTimepris5%>"></div>
		                        <div class="col-lg-2"><%call valutaKoder(5, tp5_valuta, 1) %></div>
                            </div>

                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Opdater stam-aktivitetsgrupper</div>
                                <div class="col-lg-4"><input id="Checkbox2" type="checkbox" name="FM_opd_stamgrp" value="1" />&nbsp Ja  <!--opdater timepriserne for alle medarbejdere af denne type på alle aktiviteterne, i nedenstående stamaktivitets-grupper--><br />
                                    <select name="FM_opd_stamgrp_ids" multiple class="form-control input-small">
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
                                   </select>
                                </div>
                            </div>

                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Intern kostpris:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_kostpris" value="<%=dubKostpris%>"></div>
                                <div class="col-lg-1"><%=basisValISO %></div>
                            </div>

                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-4"><h6>Angiv kostpris tariffer</h6><!--tariffer der kan bruges som tillæg til kostpris på specielle aktiviteter, aften og nat aktiviteter mm.---></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Kostpris tarif A:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=kostpristarif_A %>" name="FM_kostpristarif_A" /></div>
                                <div class="col-lg-4"><%=basisValISO %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Kostpris tarif B:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=kostpristarif_B %>" name="FM_kostpristarif_B" /></div>
                                <div class="col-lg-4"><%=basisValISO %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Kostpris tarif C:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=kostpristarif_C %>" name="FM_kostpristarif_C" /></div>
                                <div class="col-lg-4"><%=basisValISO %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Kostpris tarif D:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=kostpristarif_D %>" name="FM_kostpristarif_D" /></div>
                                <div class="col-lg-4"><%=basisValISO %></div>
                            </div>
                            <%if func = "red" then %>
                                <div class="row">
                                    <div class="col-lg-3">&nbsp</div>
                                    <div class="col-lg-4"><input id="Checkbox1" type="checkbox" name="FM_opd_intern" value="1" />&nbsp Opdater interne kostpriser/tariffer, for denne type, <br />på alle eksisterende <u>åbne</u> job.</div>
                                </div>
                           <%end if %>


                    
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
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo">Avanceret indstillinger</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">

                        <div class="panel-body">

                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Søstergruppe:</div>
                                <div class="col-lg-3">
                                    <select name="FM_soster" id="FM_soster" class="form-control input-small"> 
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
                                </div>
                            </div>

                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Hovedgruppe:</div>
                                <div class="col-lg-3">
                                    <select name="FM_gruppe" id="FM_gruppe" class="form-control input-small">              
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
                                        end if%>

                                        <option value="-1" <%=mttypSEL %> style="background-color:#CCCCCC;">// Følger søstergruppe //</option> 

    
                                      </select> 
                                </div>
                            </div>

                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Sortering:</div>
                                <div class="col-lg-1"><input type="text" class="form-control input-small" value="<%=mtsortorder%>" name="FM_sortorder" id="FM_sortorder"></div>
                             
                            </div>


                           
                                  
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

                              <div class="row pad-t20">
                                  <div class="col-lg-1">&nbsp</div>
                                 <div class="col-lg-2">
                           
                                       
                                       <h6>Afslut uge kriterie</h6>
		                 
		                  
                                </div>
                                 <div class="col-lg-6">Minimumskrav til timeantal<br />Denne medarbejdertype kan afslutte sin uge når:<br />             
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
                                    </div>


                                 </div>
                                </div>

                            <%end if %>


                             <div class="row pad-t20"><!-- RoW -->
                                  <div class="col-lg-1">&nbsp</div>
                                 <div class="col-lg-2">
                           
                                       
                                       <h6>Målsætning på <br />forretningsområde i %</h6>
		                 
		                  
                                </div>
                                 <div class="col-lg-6">
                                     <table><tr>



                                     <% afn = 0
                                        strSQLsel_fomr = "SELECT navn, id FROM fomr WHERE id <> 0" 
                                        oRec3.open strSQLsel_fomr, oConn, 3
                                        while not oRec3.EOF

                                           select case right(afn,1)
                                           case 3,6,9
                                          %>
                                            </tr><tr>
                                          <%
                                            case else
                                         
                                          end select %>
                                          <td style="padding:2px; width:100px;">
                                            <input type="hidden" name="FM_fomr_mal_fomrid" value="<%=oRec3("id")%>" />
                                             <span style="white-space:nowrap;"><%=left(oRec3("navn"), 20) %>:
                                            <%
                                            fomrfundet = 0
                                            strSQLsel_mmff = "SELECT mmff_id, mmff_mal FROM mtype_mal_fordel_fomr WHERE mmff_fomr = "& oRec3("id") 
                                            oRec4.open strSQLsel_mmff, oConn, 3
                                            if not oRec4.EOF then
                                            %>
                                             <input type="text" value="<%=oRec4("mmff_mal") %>" name="FM_fomr_mal" style="width:60px" class="form-control input-small">
                                            <%
                                            ' else
                                           
                                            fomrfundet = 1
                                            end if
                                            oRec4.close


                                            if cint(fomrfundet) = 0 then%>
                                             <input type="text" name="FM_fomr_mal" value="0" style="width:60px" class="form-control input-small">% &nbsp;&nbsp;&nbsp;&nbsp;
                                            <%end if

                                            %>
                                            <input type="hidden" name="FM_fomr_mal" value="##" /></span>
                                            
                                              
                                            </td>
                                            <%

                                            afn = afn + 1
                                        oRec3.movenext
                                        wend
                                        oRec3.close

                                       
                                      
                                   
                                       
                                         
                                    %></tr></table>
                                 </div><!-- RoW end -->


                                 </div>
                                </div>
                            </div>
                            </div>

                </section>
             <br /><br /><br />
            <%if func = "red" then %>    
            <div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
            <%end if %>
        
        </div>
             <div style="margin-top:15px; margin-bottom:15px;">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
             <div class="clearfix"></div>
        </div>

        </form>
        </div>

        <%else%> 
    	<div>Du har ikke adgang til denne side</div>      
        <%end if %>

<%case else %>



     <%if level = 1 then %>
    <script src="js/medarbtyper_liste.js" type="text/javascript"></script>

    <div class="container">
    <div class="portlet">
        <h3 class="portlet-title"><u>Medarbejdertyper</u></h3>
        
        <form action="medarbtyper.asp?menu=medarber&func=opret" method="post">
            <input type="hidden" name="lto" id="lto" value="<%=lto%>">
                <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                            <button class="btn btn-sm btn-success pull-right"><b>Opret ny +</b></button><br />&nbsp;
                            </div>
                        </div>
                </section>
         </form>

    </div>

    <div class="portlet-body">
        <table id="example" class="table dataTable table-striped table-bordered table-hover ui-datatable">
            <thead>
                <tr>
                    <th style="width: 2%">Id</th>
                    <th style="width: 5%">Sort.</th>
                    <th style="width: 22%">Navn</th>
                    <th style="width: 20%">Hovedgruppe</th>
                    <th style="width: 5%">Søstergrp?</th>
                    <th style="width: 20%">Medlemmer <br /><span style="font-size:10px;">(antal) (historisk)</span></th>
                    <th style="width: 8%">Timepris</th>
                    <th style="width: 8%">Kostpris</th>
                    <th style="width: 10%">Norm. timer</th>
                    
		            <th>Slet</th>
		          
                </tr>
            </thead>
            <tbody>
                <%
	                sort = Request("sort")
	                if sort = "nr" then
	                odrBy = " ORDER BY id"
	                else
	                odrBy = " ORDER BY mgruppe, mtsortorder, type"
	                end if

	                strSQL = "SELECT mt.id AS id, type, timepris, kostpris, normtimer_son, "_
	                &" normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, "_
	                &" normtimer_fre, normtimer_lor, tp0_valuta, v.valutakode, mtsortorder, sostergp, mgruppe, mtg.navn AS mtgnavn FROM medarbejdertyper mt "_
	                &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mt.mgruppe)"_
                    &" LEFT JOIN valutaer v ON (v.id = tp0_valuta) "_
	                &""& odrBy
	
	
	                oRec.open strSQL, oConn, 3
	
	               
	                'c = 0
	                
	                while not oRec.EOF

                    x = 0
                    antalx = 0
                    ugetotal = 0

	                ugetotal = formatnumber(oRec("normtimer_son") + oRec("normtimer_man") + oRec("normtimer_tir") + oRec("normtimer_ons") + oRec("normtimer_tor") + oRec("normtimer_fre") + oRec("normtimer_lor"), 2)
	
	                '** Antal medab i type ***'
	                strSQL2 = "SELECT count(Mid) AS mids FROM medarbejdere WHERE medarbejdertype = "& oRec("id") & " AND mansat <> 2 GROUP BY medarbejdertype"
	                
                   
                    
                    oRec2.open strSQL2, oConn, 3
	                if not oRec2.EOF then
	                x = oRec2("mids")
	                antalx = oRec2("mids")
                    'oRec2.movenext
	                end if
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
	
	            %>
                <tr>
                    <td><%=oRec("id") %></td>
                    <td><%=oRec("mtsortorder") %></td>
                    <td><a href="medarbtyper.asp?menu=medarb&func=red&id=<%=oRec("id")%>"><%=oRec("type")%> </a>

                    </td>
                    <td><%=oRec("mtgnavn") %></td>
                    <td><%if oRec("sostergp") <> 0 then %>
                    (id: <%=oRec("sostergp") %>)
                    <%end if %></td> 
                    <td><a href="medarbtyper.asp?menu=medarb&func=med&id=<%=oRec("id")%>">Se medlemmer</a> <b>(<%=antalx%>)</b> (<%=t %>)</td>
                    <td><%=oRec("timepris")%> &nbsp <%=oRec("valutakode") %></td>
                    <td><%=oRec("kostpris") &" "& basisValISO%></td>
                    <td><%=formatnumber(ugetotal)%> t. (<%=formatnumber(ugetotal/5, 1)%>)</td>
                    <%if x = 0 AND t = 0 then%>
		                    <td style="padding-left:10px; text-align:center;"><a href="medarbtyper.asp?menu=medarb&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
		              
		                <%else%>
                        <td>&nbsp;</td>
		                     <%end if%>
                </tr>
               <%
                   oRec.movenext
	               wend
               %>
            </tbody>
            <tfoot>

                    <th>Id</th>
                    <th>Sort.</th>
                    <th>Navn</th>
                    <th>Hovedgruppe</th>
                    <th>Søstergrp?</th>
                    <th>Medlemmer <br /><span style="font-size:10px;">(antal) (historisk)</span></th>
                    <th>Timepris</th>
                    <th>Kostpris</th>
                    <th>Norm. timer</th>
                   
		            <th>Slet</th>
                    

            </tfoot>
        </table>
    </div>
        <%else%> 
    	<div>Du har ikke adgang til denne side</div>      
        <%end if %>


<%end select %>

      </div>
    </div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->