<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<%



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	thisfile = "akt_typer"
	
	
	select case func
	case "opdliste"
	
	
	'Response.write "Opdaterer liste"
	
	'Response.Write request("FM_timer")
	'Response.end
	
	
	'Response.Write request("aty_id") & "<br>"
	'Response.Write request("aty_on") & "<br>"
	'Response.Write request("aty_on_realhours") & "<br>"
	'Response.Write request("aty_on_invoice") & "<br>"
	'Response.Write request("aty_on_invoice_chk") & "<br>"
	'Response.Write request("aty_on_workhours") & "<br>"
    'Response.Write request("aty_sort")
	'Response.Write request("aty_hide_on_treg")
	'Response.flush
	
	aty_id = split(request("aty_id"), ",")
	aty_on = split(request("aty_on"), "#,")
	aty_on_realhours = split(request("aty_on_realhours"), "#,")
	aty_on_invoice = split(request("aty_on_invoice"), "#,")
	aty_on_invoice_chk = split(request("aty_on_invoice_chk"), "#,")
	aty_on_workhours = split(request("aty_on_workhours"), "#,")
	aty_pre = split(request("aty_pre"), ", #, ")
    aty_sort = split(request("aty_sort"), ", #, ")
    aty_pre_dg = split(request("aty_pre_dg"), ", #, ")
    aty_pre_prg = split(request("aty_pre_prg"), ", #, ")

	aty_on_invoiceble = split(request("aty_on_invoiceble"), "#,")
    aty_hide_on_treg = split(request("aty_hide_on_treg"), "#,")

    aty_enh = split(request("FM_aty_enh"), ",")
	
	for t = 0 to UBOUND(aty_id)
	err = 0
	    
	    'Response.Write "aty_pre(t)"& aty_pre(t) & "<br>"
	    
	    aty_pre(t) = replace(aty_pre(t), " ", "")
	    aty_pre(t) = replace(aty_pre(t), ",", ".")

         aty_sort(t) = replace(aty_sort(t), " ", "")
	     aty_sort(t) = replace(aty_sort(t), ",", ".")
	      
	    if len(trim(aty_on(t))) = 0 then
	    aty_on(t) = 0
	    end if
	    
	    if len(trim(aty_on_realhours(t))) = 0 then
	    aty_on_realhours(t) = 0
	    end if
	    
	    if len(trim(aty_on_invoice(t))) = 0 then
	    aty_on_invoice(t) = 0
	    end if
	    
	    if len(trim(aty_on_invoice_chk(t))) = 0 then
	    aty_on_invoice_chk(t) = 0
	    end if
	    
	    if len(trim(aty_on_workhours(t))) = 0 then
	    aty_on_workhours(t) = 0
	    end if
	    
	    if len(trim(aty_on_invoiceble(t))) = 0 then
	    aty_on_invoiceble(t) = 0
	    end if
	    
        if len(trim(aty_hide_on_treg(t))) = 0 then
	    aty_hide_on_treg(t) = 0
        else
        aty_hide_on_treg(t) = trim(left(trim(aty_hide_on_treg(t)), 1))
            
            if aty_hide_on_treg(t) <> "#" then
            aty_hide_on_treg(t) = aty_hide_on_treg(t)
            else
            aty_hide_on_treg(t) = 0
            end if

	    end if
	    
        
	    
	     
	
	      'Response.Write "aty_on_invoice_chk(t)|"& aty_on_invoice_chk(t) & "|<br>"
	      'Response.Write "aty_on_workhours(t)|"& aty_on_workhours(t) & "|<br>"
	      
	      call erDetInt(aty_pre(t))
	    
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0


           call erDetInt(aty_sort(t))
	    
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	        

            '*** Trimemr ****
           aty_pre_dg(t) = replace(aty_pre_dg(t),", #", "")  
	       aty_pre_prg(t) = replace(aty_pre_prg(t),", #", "")  
	     
	    if cint(err) = 0 then
		
		strSQL = "UPDATE akt_typer SET aty_on = " & left(trim(aty_on(t)), 1) & ", "
		strSQL = strSQL &" aty_on_realhours = "& left(trim(aty_on_realhours(t)), 1) &", "
		strSQL = strSQL &" aty_on_invoiceble = "& left(trim(aty_on_invoiceble(t)), 1) &", "
		strSQL = strSQL &" aty_on_invoice = "& left(trim(aty_on_invoice(t)), 1) &", "
		strSQL = strSQL &" aty_on_invoice_chk = "& left(trim(aty_on_invoice_chk(t)), 1) &", "
		strSQL = strSQL &" aty_on_workhours = "& left(trim(aty_on_workhours(t)), 1) &", "
		strSQL = strSQL &" aty_pre = "& left(trim(aty_pre(t)), 4) &", "
        strSQL = strSQL &" aty_sort = "& left(trim(aty_sort(t)), 4) &", "
        strSQL = strSQL &" aty_pre_dg = '"& trim(aty_pre_dg(t)) &"', "
        strSQL = strSQL &" aty_pre_prg = '"& trim(aty_pre_prg(t)) &"', "
        strSQL = strSQL &" aty_hide_on_treg = "& aty_hide_on_treg(t) &", "
        strSQL = strSQL &" aty_enh = "& aty_enh(t) &" "
        strSQL = strSQL &" WHERE aty_id = "& aty_id(t)
		
		
		'Response.write "t" & t & "<br>"& strSQL & "<br>"
		'Response.flush
		oConn.execute(strSQL)
		
		end if
	
	next
	
	'Response.end
	Response.redirect "akt_typer.asp"
				
	
	
	
	case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:10px; width:1200px; visibility:visible;">
	
	<% 
	
	
	
	
	tTop = 0
	tLeft = 0
	tWdth = 1200
	
	
	call tableDiv(tTop,tLeft,tWdth)

        'oimg = "akt_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Aktivitets typer"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
	
	%>
	
	<br /><br />
	
	<table cellspacing="0" cellpadding="1" border="0" width="100%">
	<form method=post action="akt_typer.asp?func=opdliste">
	
	<tr bgcolor="#5582D2">
		<td class=alt valign=bottom align=right><b>Id</b></td>
		<td class=alt valign=bottom style="padding-left:10px;"><b>Type</b></td>
		<td class=alt valign=bottom><b>Aktiv</b></td>
		<td class=alt valign=bottom style="width:80px;"><b>Medregnes i dagligt timeregnskab</b><br />
		(real. timer)</td>
       
		<td class=alt valign=bottom style="width:90px;"><b>Fakturerbar</b><br />
		(er typen fakturerbar, intern ell. administrativ som f.eks ferie)</td>
		 <td class=alt valign=bottom style="width:80px;"><b>Vises i kalender</b></td>
        <td class=alt valign=bottom><b>Faktura</b> <br /> (med på)</td>
		<td class=alt valign=bottom><b>Faktura</b> <br /> (checked)</td>
		<td class=alt valign=bottom><b>Komme / Gå tid</b> <br />Løntimer <br /> (fradrag/tillæg - hvis enhed = "Enh" indregnes faktor på akt.)</td>
		<td class=alt valign=bottom><b>Pre. udfyldt</b><br />(60 min. = 1)</td>
        <td class=alt valign=bottom><b>Pre. udfyldt</b><br />Dage (man, tir, ons: 1,2,3, tom: man-fre)</td>
        <td class=alt valign=bottom><b>Pre. udfyldt</b><br />Projgrp (10 = alle grp.)</td>
		<td class=alt valign=bottom><b>Enhed</b></td>
		<td class=alt valign=bottom><b>Skjul på timereg. siden</b></td>
        <td class=alt valign=bottom><b>Kolonne & sortering</b><br />(På HR listen, tal før komma er kolonne)</td>
        <td class=alt valign=bottom><b>Vis kolonne på personlig afstemning</b><br />Måned | Dag</td>
        <td class=alt valign=bottom><b>Vis kolonne på Periode godkendelse</b></td>
	</tr>
	<%
	
	
	strSQL = "SELECT aty_id, aty_label, aty_desc, "_
    & "aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk, "_
    & "aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg, aty_pre_dg, aty_pre_prg, aty_on_calender"_
    &" FROM akt_typer ORDER BY aty_sort"
	
	x = 0
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	select case right(x, 1)
	case 0,2,4,6,8
	bgcol = "#ffffff"
	case else
	bgcol = "#EFF3FF"
	end select
	
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="17" style="padding:0px; height:1px;"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgcol %>">
        <input id="aty_id" name="aty_id" type="hidden" value="<%=oRec("aty_id")%>" />
		<td align=right><%=oRec("aty_id")%>

           </td>

         <%call akttyper(oRec("aty_id"), 1) %>
	

		<td style="padding-left:10px; white-space:nowrap;"><b> <%=akttypenavn%></b> <!-- <%=oRec("aty_desc")%> -->


             <% 'if oRec("aty_id") <> 1 then
                brugesAfantalAkt = 0
                strSQLantal = "SELECT COUNT(id) AS antal FROM aktiviteter WHERE fakturerbar = "& oRec("aty_id") 
               
                
                'response.write strSQLantal
                'response.flush 
                 oRec6.open strSQLantal, oConn, 3
                If not oRec6.EOF then
                
                brugesAfantalAkt = oRec6("antal")

                end if
                oRec6.close

                 'else

                 'brugesAfantalAkt = "> x"

                 'end if

                 response.flush
            
                 %>

		&nbsp;<span style="color:#999999; font-size:9px;">(<%=brugesAfantalAkt %>)</span>

		</td>
		<td>
		    <%if oRec("aty_on") = 1 then
		    aty_on_CHK = "CHECKED"
		    else
		    aty_on_CHK = ""
		    end if  %>
		
            <input id="aty_on" name="aty_on" type="checkbox" value="1" <%=aty_on_CHK %> />
            <input id="aty_on" name="aty_on" type="hidden" value="#" /></td>
         	<td>
         	<%if oRec("aty_on_realhours") = 1 then
         	aty_on_realhours_CHK = "CHECKED"
         	else
         	aty_on_realhours_CHK = ""
         	end if %>
         	
            <input id="aty_on_realhours" name="aty_on_realhours" type="checkbox" <%=aty_on_realhours_CHK %> value="1" />
            <input id="aty_on_realhours" name="aty_on_realhours" type="hidden" value="#" />
            </td>
            <td>
         	<%select case oRec("aty_on_invoiceble") 
         	case 1 
         	aty_on_invoiceble_SEL1 = "SELECTED"
         	aty_on_invoiceble_SEL2 = ""
         	aty_on_invoiceble_SEL0 = ""
         	case 2 
         	aty_on_invoiceble_SEL1 = ""
         	aty_on_invoiceble_SEL2 = "SELECTED"
         	aty_on_invoiceble_SEL0 = ""
         	case else
         	aty_on_invoiceble_SEL1 = ""
         	aty_on_invoiceble_SEL2 = ""
         	aty_on_invoiceble_SEL0 = "SELECTED"
            end select %>
                <select id="aty_on_invoiceble" name="aty_on_invoiceble" style="width:80px; font-size:9px; font-family:arial;">
                    <option value=1 <%=aty_on_invoiceble_SEL1 %>>Fakturerbar</option>
                    <option value=2 <%=aty_on_invoiceble_SEL2 %>>Ikke Fak.bar</option>
                    <option value=0 <%=aty_on_invoiceble_SEL0 %>>Administrativ</option>
                </select>
            <input id="aty_on_invoiceble" name="aty_on_invoiceble" type="hidden" value="#" />
            </td>

              <td>
         	<%select case oRec("aty_on_calender") 
         	case 1 
         	aty_on_calender_SEL1 = "SELECTED"
         	aty_on_calender_SEL2 = ""
         	aty_on_calender_SEL0 = ""
         	case 2 
         	aty_on_calender_SEL1 = ""
         	aty_on_calender_SEL2 = "SELECTED"
         	aty_on_calender_SEL0 = ""
         	case else
         	aty_on_calender_SEL1 = ""
         	aty_on_calender_SEL2 = ""
         	aty_on_calender_SEL0 = "SELECTED"
            end select %>
                <select id="aty_on_calender" name="aty_on_calender" style="width:80px; font-size:9px; font-family:arial;" DISABLED>
                    <option value=1 <%=aty_on_calender_SEL1 %>>Ja</option>
                    <option value=2 <%=aty_on_calender_SEL2 %>>Ja - fravær (orange)</option>
                    <option value=0 <%=aty_on_calender_SEL0 %>>Nej</option>
                </select>
            <input id="aty_on_invoiceble" name="aty_on_invoiceble" type="hidden" value="#" />
            </td>



           	<td>
           	
           	<%if oRec("aty_on_invoice") = 1 then
         	aty_on_invoice_CHK = "CHECKED"
         	else
         	aty_on_invoice_CHK = ""
         	end if %>
           	
            <input id="aty_on_invoice" name="aty_on_invoice" type="checkbox" <%=aty_on_invoice_CHK %> value="1" />
            <input id="aty_on_invoice" name="aty_on_invoice" type="hidden" value="#" />
            </td>
         	<td>
            <%if oRec("aty_on_invoice_chk") = 1 then
         	aty_on_invoice_chk_CHK = "CHECKED"
         	else
         	aty_on_invoice_chk_CHK = ""
         	end if %>
           	
            <input id="aty_on_invoice_chk" name="aty_on_invoice_chk" type="checkbox" <%=aty_on_invoice_chk_CHK %> value="1" />
            <input id="aty_on_invoice_chk" name="aty_on_invoice_chk" type="hidden" value="#" /></td>
         
            	
         	
         	<td>
         	<%select case oRec("aty_on_workhours") 
         	case 1 
         	aty_on_workhours_SEL1 = "SELECTED"
         	aty_on_workhours_SEL2 = ""
         	aty_on_workhours_SEL0 = ""
         	case 2 
         	aty_on_workhours_SEL1 = ""
         	aty_on_workhours_SEL2 = "SELECTED"
         	aty_on_workhours_SEL0 = ""
         	case else
         	aty_on_workhours_SEL1 = ""
         	aty_on_workhours_SEL2 = ""
         	aty_on_workhours_SEL0 = "SELECTED"
            end select %>
                <select id="aty_on_workhours" name="aty_on_workhours" style="width:60px; font-size:9px; font-family:arial;">
                    <option value=1 <%=aty_on_workhours_SEL1 %>>Fradrag</option>
                    <option value=2 <%=aty_on_workhours_SEL2 %>>Tillæg</option>
                    <option value=0 <%=aty_on_workhours_SEL0 %>>Ingen</option>
                </select>
            <input id="aty_on_workhours" name="aty_on_workhours" type="hidden" value="#" />
            </td>
         	
            	<td>
                    <input id="aty_pre" name="aty_pre" style="font-size:9px; width:50px;" id="Text1" value="<%=formatnumber(oRec("aty_pre"), 2) %>" type="text" />
                    <input id="aty_pre" name="aty_pre" type="hidden" value="#" />
           </td>

         <td>
                    <input id="aty_pre_dg" name="aty_pre_dg" style="font-size:9px; width:50px;" id="Text1" value="<%=oRec("aty_pre_dg")%>" type="text" />
                    <input id="aty_pre_dg" name="aty_pre_dg" type="hidden" value="#" />
           </td>
       <td>
                    <input id="aty_pre_prg" name="aty_pre_prg" style="font-size:9px; width:50px;" id="Text1" value="<%=oRec("aty_pre_prg")%>" type="text" />
                    <input id="aty_pre_prg" name="aty_pre_prg" type="hidden" value="#" />
           </td>

           <td>
            <%
            
                enhSEL01 = ""
                enhSEL0 = ""
                enhSEL1 = ""
                enhSEL2 = ""
                enhSEL3 = ""
            
            select case oRec("aty_enh")
            case -1
            enhSEL01 = "SELECTED"
	        case 1
	        'enh = "stk."
            enhSEL1 = "SELECTED"
	        case 2
	        'enh = "enh."
	        enhSEL2 = "SELECTED"
            case 3
            enh = "km."
            enhSEL3 = "SELECTED"
	        case else
	        'enh = "tim."
            enhSEL0 = "SELECTED"
	        end select%>
	      
               <select style="width:60px; font-size:9px;" name="FM_aty_enh">
                   <option value="0" <%=enhSEL0 %>>Timer</option>
                   <option value="1" <%=enhSEL1 %>>Stk.</option>
                   <option value="2" <%=enhSEL2 %>>Enheder</option>
                   <option value="3" <%=enhSEL3 %>>Km</option>
                   <option value="-1" <%=enhSEL01 %>>-</option>

               </select>

               
           </td>



           <td>
           <%if cint(oRec("aty_hide_on_treg")) = 1 then
         	aty_hide_on_treg_CHK = "CHECKED"
         	else
         	aty_hide_on_treg_CHK = ""
         	end if %>

            <input id="aty_hide_on_treg" name="aty_hide_on_treg" type="checkbox" value="1" <%=aty_hide_on_treg_CHK %> />
            <input id="aty_hide_on_treg" name="aty_hide_on_treg" type="hidden" value="#" />
           &nbsp;
           </td>
		
        <td><input type="text" name="aty_sort" value="<%=oRec("aty_sort")%>" style="width:40px; font-size:9px;" />
            <input id="Hidden1" name="aty_sort" type="hidden" value="#" />
        </td>
        <td>
              <input  name="" type="checkbox" value="1"  DISABLED />
            <input  name="" type="hidden" value="#" />

              <input name="" type="checkbox" value="1" DISABLED/>
            <input  name="" type="hidden" value="#" />

        </td>
        <td>

              <input name="" type="checkbox" value="1" DISABLED/>
            <input  name="" type="hidden" value="#" />

        </td>

	</tr>
	<%
	x = x + 1
	oRec.movenext
	wend
	%>	
	</table>
	<input id="Hidden1" name="aty_sort" type="hidden" value="#" />
	<input id="Hidden1" name="aty_on" type="hidden" value="#" />
	<input id="Hidden2" name="aty_on_realhours" type="hidden" value="#" />
	<input id="Hidden3" name="aty_on_invoiceble" type="hidden" value="#" />
	<input id="Hidden4" name="aty_on_invoice" type="hidden" value="#" />
	<input id="Hidden5" name="aty_on_invoice_chk" type="hidden" value="#" />
	<input id="Hidden6" name="aty_on_workhours" type="hidden" value="#" />
	<input id="Hidden7" name="aty_pre" type="hidden" value="#" />
	
	
	
	<br><br>&nbsp;</div>
	<table width=100%>
	<tr>
	    <td colspan=8 align=right style="padding-right:50px;"><br />
            <input id="Submit2" type="submit" value=" Opdater >> " /></td>
	</tr>
	
	</form>
	</table>
	
	<br><br><br>
	<a href="Javascript:window.close()" class=red>[Luk vindue]</a>
	
	
	<%
	
	itop = 50
	ileft = 0
	iwdt = 400
	
	call sideinfo(itop,ileft,iwdt) %>
	<b>Aktiv</b><br />
		Er denne aktivitetstype aktiv.<br /><br />
	
	    <b>Medregnes i dagligt timeregnskab</b> (realiserede timer) <br />
		Indgår denne type i de dagligt realiserede timer der udgør balancen mellem normeret tid pr. uge og realiseret tid.<br /><br />
		Hvis man angiver den normerede uge som 37 timer excl. frokost skal typen <b>"frokost"</b> sættes til <b>ikke</b> at tælle med i det daglige timeregnskab (de realiserede timer pr. dag). I dette tilfælde sættes den til "Administrativ".<br /><br />
	 Typen <b>"Afspadsering"</b> sættes normalt til Ja, så den tæller med i den realiserede tid pr. uge, således at man ikke går i minus på den løbende fleks saldo ved at afspadsere. normalt kan man kun afspadsere optjent overarbejde. <br /><br />
	 I det tilfælde hvor man har optjent et plus på norm/real tid saldoen bruges typen <b>"Fleks"</b> til at "nulle" saldoen. Derfor skal typen <b>"Fleks"</b> normalt ikke medregnes i de realiserede timer pr. dag.
	 <br /><br />
	    <b>Overarbejde</b><br />
	    Modregner "Afspadsering". <br /><br />
	    
        <b>Salg & newbizz</b><br />
        Denne type er som udgangspunkt den eneste man kan registrerere tid på tilbud på timereg. siden.<br /><br />


		<b>Fakturerbar</b><br />
		Skal denne type medregnes til fakturerbare timer.<br />
		Vælg mellem:<br /><br />
		<u>Fakturerbar</u><br />
		Fakturerbare arbejdsopgaver.<br />
		<br /><u>Ikke fakturerbar</u><br />
		Arbejdsopgaver der er "Ikke fakturerbare".<br /><br />
       
		<u>Administrativ</u><br />
		Ikke timebaserede aktiviteter, eller aktiviteter hvor man ikke er tilstede.
		F.eks Kørsel, ferie, stk, afspadsering etc.<br /><br />

        
		<b>Faktura</b> (med på) <br /> 
		Skal denne type vises på fakturaer.<br /><br />
		<b>Faktura (forvalgt)</b> <br />Skal denne type være forvalgt på faktura oprettelse.<br /><br />
		<b>Løn timer</b> <br /> 
		Skal denne type give fradrag i forhold til løn timer (Stempelur timer).<br /><br />
		<b>Pre. udfyldt</b><br />
		Skal denne type være forudfyldt på timereg. siden. (30 min. = 0,5). Timer vises automatisk 
		og indlæses første gang der registreres timer den uge.<br /><br />
		<b>Ferie</b><br />
		<b>"Ferie planlagt"</b> indlæser automatisk en <b>"ferie afholdt"</b> registrering når dagen oprinder. 
		Hvis der allerede findes en registrering på dagen
		med <b>"ferie afholdt"</b> bliver der ikke indlæst en ny registrering den pågældende dag.<br />
        <br />Ferie / Ferie fridage brugt / udbetalt modregner optjent.<br /><br />
        <b>Ferie / Ferie fri optjent</b><br />
        Ferie / Ferie fri optjent angives i timer ell. dage. Omregningen mellem timer og dage sker i forhold til normeret arbejdsuge i timer / 5 arb. dage. Så der derved bliver beregnet et dags-gennemsnit.<br /><br />
        Mandag - Søndag <br />
        7,5 - 7,5 - 7,5 - 7,5 - 7,0 - 0 - 0 giver et gennemsnit på 7,4 timer pr. dag.<br />
        0 - 0 - 7,5 - 7,5 - 7 - 0 - 0 giver 4,4 timer pr. dag i gennemsnit.<br />

	   
	   <!--<b>Drys timer</b><br />
        Denne type bruges til at "drysse" f.eks en time ud på forskellige job. TimeOut beregner selv hvor mange job der er på den
        "personlige aktiv liste" og drysser herefter den angivne tid, ud på de aktiviteter der findes af typen "drys".<br /><br />
	   -->
	   
	   <br /><b>Faste typer der modregner hinanden</b><br />
	   <b>"Afspadsering"</b> bliver automatisk altid modregnet typen <b>"Overarbejde"</b> og nulstiller typen <b>"Afspad. ønskes udb."</b><br />
	   
	    
	</td>
	</tr>
	</table>
	</div>
	<br />
	<br /><br /><br /><br />&nbsp;
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
