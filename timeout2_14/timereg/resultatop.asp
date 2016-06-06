<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/konto_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	dleft = 90
	dtop = 102
	
	
	call finddatoer()
	
	
	
	if print <> "j" then %>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
    <%call menu_2014() %>

	<%else%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if%>
	
	<div id="sindhold" style="position:absolute; left:<%=dleft%>px; top:<%=dtop%>px; visibility:visible;">
	
	<%pTxt = "Resultatopgørelse" %>
        

        <form action="resultatop.asp?menu=kon" method="POST">
        <%call filterheader_2013(0,0,800,pTxt)%>

       
            </td></tr>
            <tr><td>
			
		    <b>Periode:</b></td></tr>
			<tr>
			<!--#include file="inc/weekselector_b.asp"-->
                </td>
			
			<td align=right >
                <input id="Submit1" type="submit" value="Se resultatopg. >>" />
                </td>
			</tr>
            <input type="hidden" name="useper" value="j">
			

			</table>
			</div>
</form>
      
      
      
        <%
	
	
	
	periodefilter = " posteringsdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	
	
	'*** Resultat ***
	'*** 1 = drift
	'*** 2 = status
	
	
	'*** Drift, Udgifter ***
	salDoDriftUdg = 0
	strSQL = "SELECT k.kontonr, SUM(p.nettobeloeb) AS salDoDriftUdg FROM kontoplan k "_
	&" LEFT JOIN posteringer p ON (p.kontonr = k.kontonr AND "& periodefilter &" AND p.status = 1 AND debitkredit = 4) "_
	&" WHERE k.type = 1 GROUP BY k.type"
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	salDoDriftUdg = oRec("salDoDriftUdg") 
	end if
	oRec.close
	
	'*** Drift, indtægter ***
	salDoDriftInd = 0
	strSQL = "SELECT k.kontonr, SUM(p.nettobeloeb) AS salDoDriftInd FROM kontoplan k "_
	&" LEFT JOIN posteringer p ON (p.kontonr = k.kontonr AND "& periodefilter &" AND p.status = 1 AND debitkredit = 3) "_
	&" WHERE k.type = 1 GROUP BY k.type"
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	salDoDriftInd = oRec("salDoDriftInd")
	end if
	oRec.close
	
	
	'*** Status deb. ***
	saldoStatDeb = 0
	strSQL = "SELECT k.kontonr, SUM(p.nettobeloeb) AS saldoStatDeb FROM kontoplan k "_
	&" LEFT JOIN posteringer p ON (p.kontonr = k.kontonr AND "& periodefilter &" AND p.status = 1) "_
	&" WHERE k.type = 2 AND debitkredit = 2 GROUP BY k.type"
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	saldoStatDeb = oRec("saldoStatDeb") 
	end if
	oRec.close
	
	'*** Status kre. ***
	saldoStatKre = 0
	strSQL = "SELECT k.kontonr, SUM(p.nettobeloeb) AS saldoStatKre FROM kontoplan k "_
	&" LEFT JOIN posteringer p ON (p.kontonr = k.kontonr AND "& periodefilter &" AND p.status = 1) "_
	&" WHERE k.type = 2 AND debitkredit = 1 GROUP BY k.type"
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	saldoStatKre = oRec("saldoStatKre") 
	end if
	oRec.close
	
	
	'*** Moms I ****
	saldoMomsI = 0
	strSQL = "SELECT k.kontonr, SUM(p.moms) AS saldoMomsI FROM kontoplan k "_
	&" LEFT JOIN posteringer p ON (p.kontonr = k.kontonr AND "& periodefilter &" AND p.status = 1 AND debitkredit = 4) "_
	&" WHERE k.type = 1 GROUP BY k.type"
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	saldoMomsI = oRec("saldoMomsI") 
	end if
	oRec.close
	
	'*** Moms U ****
	saldoMomsU = 0
	strSQL = "SELECT k.kontonr, SUM(p.moms) AS saldoMomsU FROM kontoplan k "_
	&" LEFT JOIN posteringer p ON (p.kontonr = k.kontonr AND "& periodefilter &" AND p.status = 1 AND debitkredit = 3) "_
	&" WHERE k.type = 1 GROUP BY k.type"
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	saldoMomsU = oRec("saldoMomsU")
	end if
	oRec.close
	
	tilsvar = saldoMomsI + (saldoMomsU)
	
	
	function debitkredit(val)
	
    if len(trim(val)) <> 0 then
    val = val
    else
    val = 0
    end if


	if val < 0 then
	%>
	 <td>
         &nbsp;</td>
	        <td align=right><%=formatnumber(val, 2) %></td>
	<%
	else
	%>
	 
	        <td align=right><%=formatnumber(val, 2) %></td>
	        <td>
         &nbsp;</td>
	<%
	end if 
	
	end function
	%>
	
	
	
	
	
	<br /><br /><br />
	<table cellspacing=1 cellpadding=2 border=0 width="820" bgcolor="#cccccc">
	
	<!-- Drift --->
	    <tr bgcolor="#5582d2">
	        <td width=400 class=alt><b>Drift:</b></td>
	        <td width=200>&nbsp;</td>
	        <td class=alt width=100 align=right><b>Debit</b></td>
	        <td class=alt width=100 align=right><b>Kredit</b></td>
	    </tr>
	    <tr bgcolor="#ffffff">
	    <td>Indtægter</td>
	    <td>&nbsp;</td>
	        <%call debitkredit(salDoDriftInd) %>
	       
	     </tr>
	     <tr bgcolor="#ffffff">
	        <td>Omkostninger</td>
	        <td>&nbsp;</td>
	        <%call debitkredit(salDoDriftUdg) %>
	      </tr>
	    
	    <tr bgcolor="lightpink">
	        <td><b>Resultat, drift</b></td>
	        <%resultatExclMoms = (salDoDriftUdg + (salDoDriftInd)) %>
	        <td>&nbsp;</td>
	        <%call debitkredit(resultatExclMoms) %>
	    </tr>
	     
	     </table>
	     <br /><br /><br />
	     
	     <table cellspacing=1 cellpadding=2 border=0 width="820" bgcolor="#cccccc">
	     <%
	     aktiverTot = 0
	     %>
	     <!-- status -->
	     <tr bgcolor="#5582d2">
	        <td width=400 class=alt><b>Status:</b></td>
	        <td width=200>&nbsp;</td>
	         <td class=alt width=100 align=right><b>Debit</b></td>
	        <td class=alt width=100 align=right><b>Kredit</b></td>
	    </tr>
	     <tr bgcolor=#ffff99>
	        <td><b>Aktiver</b></td>
	        <td>&nbsp;</td>
	        <td>&nbsp;</td>
	        <td>&nbsp;</td>
	    </tr>
	    <tr bgcolor="#ffffff">
	        <td>Saldo - Debitorer:</td>
	        <td>&nbsp;</td>
	        
	        <%call debitkredit(saldoStatDeb) %>
	        
	    </tr>
	     
	        <%
	        
	        aktiverTot = aktiverTot + (saldoStatDeb)
	        
	        '*** Henter de forskellige konti der hører til aktiver ***
	        
	        saldoStatAktiver = 0
	       
        	strSQL = "SELECT k.kontonr, k.navn AS kontonavn, SUM(p.nettobeloeb) AS saldoStatAktiver, "_
        	&" nk.ap_type, nk.navn AS ap_navn"_
        	&" FROM nogletalskoder nk "_
        	&" LEFT JOIN kontoplan k  ON (k.keycode = nk.id) "_
        	&" LEFT JOIN posteringer p ON (p.kontonr = k.kontonr AND "& periodefilter &" AND p.status = 1)"_
        	&" WHERE nk.ap_type = 1 AND k.kontonr <> '' GROUP BY ap_type, k.kontonr" 
        	
        	'Response.Write strSQL
        	'Response.flush
        	
	        oRec.open strSQL, oConn, 3
	        while not oRec.EOF 
	           
	            if len(oRec("saldoStatAktiver")) <> 0 then
	            saldoStatAktiver = oRec("saldoStatAktiver")
	            else
	            saldoStatAktiver = 0
	            end if 
	            
	        %>
	        <tr bgcolor="#ffffff">
	            <td><%=oRec("kontonavn")%> (<%=oRec("kontonr") %>)</td>
	            <td class=lillegray><%=oRec("ap_navn") %> </td>
	            <%call debitkredit(saldoStatAktiver) %>
	            
	         </tr>
	        
	        <%
	        aktiverTot = aktiverTot + (saldoStatAktiver)
	        
	        oRec.movenext
	        wend
	        oRec.close
	        
	        %>
	          <tr bgcolor="lightpink">
	        <td><b>Aktiver total:</b></td>
	        <td>&nbsp;</td>
	         <%call debitkredit(aktiverTot) %>
	        
	    </tr>
	    </table>
	        
	   <br />
	     <table cellspacing=1 cellpadding=2 border=0 width="820" bgcolor="#cccccc">
	    
	    <!--- Passiver -->
	    <%
	    passiverTot = 0
	     %>
	   <tr bgcolor="#ffff99">
	        <td width=400><b>Passiver</b></td>
	        <td width=200>&nbsp;</td>
	        <td width=100>&nbsp;</td>
	        <td width=100>&nbsp;</td>
	    </tr>
	    <tr bgcolor="#ffffff">
	        <td>Saldo - Kreditorer:</td>
	        <td>&nbsp;</td>
	         <%call debitkredit(saldoStatKre) %>
	         <%passiverTot = passiverTot + (saldoStatKre) %>
	     </tr>
	    <tr bgcolor="#ffffff">
	        <td>Moms I:</td>
	        <td>&nbsp;</td>
	         <%call debitkredit(saldoMomsI) %>
	        
	    </tr>
	    <tr bgcolor="#ffffff">
	        <td>Moms U:</td>
	        <td>&nbsp;</td>
	         <%call debitkredit(saldoMomsU) %>
	        
	    </tr>
             <%if len(trim(tilsvar)) <> 0 then
               tilsvar = formatnumber(tilsvar,2)
               else
               tilsvar = 0  
               end if
                 %>

	     <tr bgcolor="#ffffe1">
	        <td><b>Tilsvar: <%=tilsvar%></b></td>
	        <td>&nbsp;</td>
	        <td>&nbsp;</td>
	        <td>&nbsp;</td>
	       
	    </tr>
	    
	    
	     <%
	     passiverTot = passiverTot + (tilsvar)
	        
	        
	        
	        '*** Henter de forskellige konti der hører til passiver (nøgletalskode) ***
	        
	        saldoStatPassiver = 0
	       
        	strSQL = "SELECT k.kontonr, k.navn AS kontonavn, SUM(p.nettobeloeb) AS saldoStatPassiver, "_
        	&" nk.ap_type, nk.navn AS ap_navn"_
        	&" FROM nogletalskoder nk "_
        	&" LEFT JOIN kontoplan k  ON (k.keycode = nk.id) "_
        	&" LEFT JOIN posteringer p ON (p.kontonr = k.kontonr AND "& periodefilter &" AND p.status = 1)"_
        	&" WHERE nk.ap_type = 2 AND k.kontonr <> '' GROUP BY ap_type, k.kontonr" 
        	
        	'Response.Write strSQL
        	'Response.flush
        	
	        oRec.open strSQL, oConn, 3
	        while not oRec.EOF 
	           
	            if len(oRec("saldoStatPassiver")) <> 0 then
	            saldoStatAktiver = oRec("saldoStatPassiver")
	            else
	            saldoStatAktiver = 0
	            end if 
	            
	        %>
	        <tr bgcolor="#ffffff">
	            <td><%=oRec("kontonavn")%> (<%=oRec("kontonr") %>)</td>
	            <td class=lillegray><%=oRec("ap_navn") %> </td>
	           <%call debitkredit(saldoStatPassiver) %>
	         </tr>
	        
	        <%
	        passiverTot = passiverTot + (saldoStatPassiver)
	        
	        oRec.movenext
	        wend
	        oRec.close
	        
	        %>
	        
	         <tr bgcolor="#ffffff">
	        <td>Resultat</td>
	        <td>&nbsp;</td>
	          <%call debitkredit(resultatExclMoms) %>
	       
	     </tr>
	        
	        <%
	        passiverTot = passiverTot + (resultatExclMoms)
	        %>
	          <tr bgcolor="lightpink">
	        <td><b>Passiver total:</b></td>
	        <td>&nbsp;</td>
	          <%call debitkredit(passiverTot) %>
	        
	    </tr>
	    
	</table>
	
	<br />
	</div>
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
