<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/sdsk_func.asp"-->
<%





if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
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
	
	thisfile = "sdsk_kontrolpanel"
	kview = "j"
	
	if len(request("lastedit")) <> 0 then
	lastedit = request("lastedit")
	else
	lastedit = 0
	end if
	
	select case func
	case "slet"
	%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
	
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>slette</b> en valuta.<br>"_
    &"Du vil samtidig nulstille alle fakturaer der er oprettet med denne valuta til at blive vist med grundvalutaen.<br>"
	
    slturl = "erp_valutaer.asp?menu=tok&func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	case "sletok"
	'*** Her slettes en valuta ***
	oConn.execute("DELETE FROM valutaer WHERE id = "& id &"")
	Response.redirect "erp_valutaer.asp?menu=tok"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		
		
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless2 = tmp
		end function
		
		err = 0
		
		
		
		if len(trim(request("FM_navn"))) <> 0 then
		strNavn = SQLBless(request("FM_navn"))
		else
		err = 110
		end if
		
        if len(trim(request("FM_grundvaluta"))) <> 0 then
        grundvaluta = 1
        else
        grundvaluta = 0
        end if 

		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		
		if len(trim(request("FM_kode"))) <> 0 then
		strKode = SQLBless(request("FM_kode"))
		else
		err = 111
		end if
		
		    %>
			<!--#include file="inc/isint_func.asp"-->
			<%


              '**Kurs altid = 100 på gerundvaluta ***'
                    if cint(grundvaluta) = 1 then
                    intKurs = 100
                    else
                    intKurs = request("FM_kurs") 
                    end if

			        call erDetInt(intKurs)

                    
			
					if isInt > 0 OR intKurs = "0" OR len(trim(intKurs)) = 0 then
					err = 112
					else
		            intKurs = SQLBless2(intKurs)
		            end if
		    
		            isInt = 0


                  
		
		if err = 0 then
		        
                '** nulstiller grundvaluta ***'
                if cint(grundvaluta) = 1 then
                strSQLg = "UPDATE valutaer SET grundvaluta = 0 WHERe id <> 0"
                oConn.execute(strSQLg)
                end if
		
		        if func = "dbopr" then
		        strSQLopr = "INSERT INTO valutaer (valuta, editor, dato, valutakode, kurs, grundvaluta) "_
		        &" VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "_
		        &" '"& strKode &"', "& intKurs &", "& grundvaluta &")"
        		
		        'Response.Write strSQLopr
		        'Response.flush
		        oConn.execute(strSQLopr)
        				
        				strSQL = "SELECT id FROM valutaer WHERE id <> 0 ORDER BY id DESC"
        				oRec.open strSQL, oConn, 3
        				if not oRec.EOF then
        				
        				lastedit = oRec("id") 
        				
        				end if
        				oRec.close
        				
        		else
		        oConn.execute("UPDATE valutaer SET valuta ='"& strNavn &"', editor = '" &strEditor &"', "_
		        &" dato = '" & strDato &"', valutakode = '"& strKode &"', kurs = "& intKurs &", grundvaluta = "& grundvaluta &""_
		        &" WHERE id = "& id &"")
        		
		        lastedit = id
        		
		        end if
		        
		        strSQLvalutahist = "INSERT INTO valuta_historik (valutaid, kurs, dato, editor) "_
		        &" VALUES ("& lastedit &", "& intKurs &", '" & strDato &"', '" &strEditor &"')"
		        
		        
		        oConn.execute(strSQLvalutahist)



        '**** opdateterer eksisterende *****' 
		if request("FM_opdater") = "1" then

                   
                    '*** Henter aktive job i valuta og opdaterer kurs
                    'strSQLopdajob = "UPDATE job SET kurs = "& intKurs &" WHERE jobstatus = 1 AND valuta = "& id &""  
                    'oConn.execute(strSQLopdajob)

                   

                    '*** Henter aktive job i valuta og opdaterer timer
                    
                     '*** SPECIAL LTO SETTINGS *****'
                    select case lto
                    case "nt", "intranet - local"
                    strSQLopdajobtimer = "SELECT id, jobnr, orderqty, tax_pc, cost_price_pc, sales_price_pc, freight_pc, sales_price_pc_valuta, "_
                    &" cost_price_pc_valuta, comm_pc, fastpris FROM job WHERE jobstatus = 1 AND cost_price_pc_valuta = "& id & ""
                    case else
                    strSQLopdajobtimer = "SELECT id, jobnr FROM job WHERE jobstatus = 1 AND valuta = "& id &""
                    end select
                   
                    
                    'response.write strSQLopdajobtimer & "<br>"
                 
                        j = 0
                    	oRec.open strSQLopdajobtimer, oConn, 3
        				while not oRec.EOF
        				            
                


                                    'if cint(j) = 0 then

                        
                                        '*** SPECIAL LTO SETTINGS *****'
                                        select case lto
                                        case "nt", "intranet - local"


                                        if oRec("tax_pc") <> 0 then
                                        tax = 1 + (oRec("tax_pc")/100)
                                        else
                                        tax = 1
                                        end if

                                        orderqty = oRec("orderqty")
                                        
                                        freight_pc = oRec("freight_pc")
                                            
                                        cost_price_pc = oRec("cost_price_pc")
                                        
                                        sales_price_pc = oRec("sales_price_pc")

                                        comm_pc = oRec("comm_pc")
                                        
                                        intKursUse = intKurs/100
                                        
                                        'response.write "<br><br>((qty: "&  orderqty &" * cost "& cost_price_pc &" * "& tax &") + ("& freight_pc &" * "& orderqty & ")) * kurs: "& (intKursUse)  & " # comm_pc: "& comm_pc
                                        'response.write "<br>sales_price_pc: "& sales_price_pc
                                       
                                        'if oRec("fastpris") = 2 then 'COMI

                                        jo_udgifter_intern_ny = (orderqty * cost_price_pc * tax)
                                        'response.write "jo_udgifter_intern_ny 1: "& jo_udgifter_intern_ny & "<br>"
                                        jo_udgifter_intern_ny = jo_udgifter_intern_ny  + (freight_pc * orderqty)
                                        'response.write "jo_udgifter_intern_ny 2: "& jo_udgifter_intern_ny & "<br>"
                                        jo_udgifter_intern_ny = jo_udgifter_intern_ny  * (intKursUse)
                                        'jo_udgifter_intern_ny = formatnumber(jo_udgifter_intern_ny, 2)
                                        'response.write "jo_udgifter_intern_ny 3: "& jo_udgifter_intern_ny & "<br>"

                                        'jo_udgifter_intern_ny = replace(jo_udgifter_intern_ny, ".", "")
                                        'jo_udgifter_intern_ny = replace(jo_udgifter_intern_ny, ",", ".")
                                        'response.write "jo_udgifter_intern_ny 4: "& jo_udgifter_intern_ny & "<br>"

                                        'else 'salesprice

                                     

                                        'end if                        

                                        jo_bruttooms_ny = (orderqty * sales_price_pc)

                                        if cint(oRec("sales_price_pc_valuta")) = cint(oRec("cost_price_pc_valuta")) then 'salgspris har samme valuta som kost
                                        jo_bruttooms_ny = jo_bruttooms_ny * (intKursUse)
                                        
                                        else
                                                            
                                                                intSalgsPcKurs = 100
                                                                strSQLv = "SELECT kurs FROM valutaer WHERE id = " & oRec("sales_price_pc_valuta")
	                                                            oRec3.open strSQLv, oConn, 3
	
	                                                            if not oRec3.EOF then
	
	                                                         
	                                                            intSalgsPcKurs = oRec3("kurs")/100
	                                                            
	
	                                                            end if
	                                                            oRec3.close

                                        jo_bruttooms_ny = jo_bruttooms_ny * (intSalgsPcKurs)                                    
                                        end if
                                        
                                        'jo_bruttooms_ny = formatnumber(jo_bruttooms_ny, 2)
                                        
                                        'jo_bruttooms_ny = replace(jo_bruttooms_ny, ".", "")
                                        'jo_bruttooms_ny = replace(jo_bruttooms_ny, ",", ".")

                                        strSQLopdajob = "UPDATE job SET jo_udgifter_intern = "& jo_udgifter_intern_ny &", jo_bruttooms = "& jo_bruttooms_ny &" WHERE id = "& oRec("id")
                                        
                                        'response.write "<br>"& strSQLopdajob
                                        'response.end
                                        oConn.execute(strSQLopdajob)

                                        end select
                    


                                    'end if


                
                                    '*** Opdaterer alle valutaer på timeregistreringer 
                                    strSQLopdtimer = "UPDATE timer SET kurs = "& intKurs &" WHERE tjobnr = '"& oRec("jobnr") &"' AND valuta = "& id 

                                    'response.write "<br>" & strSQLopdtimer
                                    oConn.execute(strSQLopdtimer)            

                                    
                        j = j + 1
        				oRec.movenext
        				wend
        				oRec.close

                    
                    response.flush

                    %>
                    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
                    <div style="position:absolute; left:20px; top:20px; padding:20px;">
                    <h4>Opdatering af valuta</h4>

                    Opdatering af timer og beløb er klar.<br /><br />

                    <a href="erp_valutaer.asp?menu=tok&shokselector=1&lastedit="<%=lastedit%>">Ok - videre --></a>
                        </div>
                    <%

        else

                	Response.redirect "erp_valutaer.asp?menu=tok&shokselector=1&lastedit="&lastedit

        end if

        'response.write "<br><br>OK"
        'response.end

	
		
		
		else
		
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		errortype = err
		call showError(errortype)
		
		
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	rsptid = 0
	intkunweekend = 1
    grundvaluta = 0
	
	else
	strSQL = "SELECT valuta, editor, dato, valutakode, kurs, grundvaluta FROM valutaer WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("valuta")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strKode = oRec("valutakode")
	intKurs = oRec("kurs")
	grundvaluta = oRec("grundvaluta")
	
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if


    if cint(grundvaluta) = 1 then
    grundvalutaCHK = "CHECKED"
    else
	grundvalutaCHK = ""
    end if
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:62; visibility:visible;">
	
	<br><br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="erp_valutaer.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><h3><img src="../ill/ac0044-24.png" />  Valutaer - <%=varbroedkrumme%></h3></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Valuta:</b></td>
		<td style="padding-top:10px;"><input type="text" name="FM_navn" value="<%=strNavn%>" size="40" style="border: 1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Valuta ISO kode:</b> </td>
		<td style="padding-top:10px;"><input type="text" name="FM_kode" value="<%=strKode%>" size="5" style="border: 1px #86B5E4 solid;"> (DKR, SEK, EUR, USD etc.)</td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Kurs:</b><br />
		 Afrundes til 2 decimaler</td>
		<td style="padding-top:10px;"><input type="text" name="FM_kurs" value="<%=intKurs %>" size=10 style="border: 1px #86B5E4 solid;"> (89,34)</td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Grundvaluta:</b><br />
		 Angiv om denne valuta er systemets grundvaluta.</td>
		<td style="padding-top:10px;"><input type="checkbox" name="FM_grundvaluta" value="1" <%=grundvalutaCHK %>></td>
	</tr>

    <tr>
		<td style="padding-top:10px;"><b>Opdater eksisterende:</b><br />
        Opdater eksisterende <b>aktive job</b> og tilhørende timeregistreringer, på de job der er oprettet i denne valuta.
		</td>
		<td style="padding-top:10px;"><input type="checkbox" name="FM_opdater" value="1"></td>
	</tr>
	
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="100" height="1" alt="" border="0"><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case else
	
	
	%>
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:62px; visibility:visible; width:600px;">
	
	<h3>
        <img src="../ill/ac0044-24.png" /> Valutaer</h3>
	
	<%
	
	oleftpx = 0
	otoppx = 20
	owdtpx = 120
	
	call opretNy("erp_valutaer.asp?menu=tok&func=opret&prio_grp="&prio_grp&"", "Opret ny valuta", otoppx, oleftpx, owdtpx) 
	
	
	%>
	<br /><br /><br />
	<%
	
	
	tTop = 0
	tLeft = 0
	tWdth = 600
            	
            	
	call tableDiv(tTop,tLeft,tWdth)
	%>
   
	
	<table cellspacing="0" cellpadding="2" border="0" width="100%">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=6 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td width="8" align=right valign=top rowspan=2><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="50" class=alt><b>Id</b></td>
		<td class=alt><b>Valuta</b> (grundvaluta)</td>
		<td class=alt><b>Valuta ISO kode</b></td>
		<td class=alt align=right><b>Kurs</b></td>
		<td class=alt align=right><b>Antal fakturaer<br /> i valuta</b>
            &nbsp;</td>
            <td>
                &nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT v.id, v.valuta, v.valutakode, v.kurs, v.grundvaluta, COUNT(f.fid) AS antalfak "_
	&" FROM valutaer v "_
	&" LEFT JOIN fakturaer f ON (f.valuta = v.id) "_
	&" WHERE v.id <> 0 "_
	&" GROUP BY f.valuta, v.id ORDER BY v.valuta"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	if cint(oRec("id")) = cint(lastedit) then
	bgCol = "#ffff99"
	else
	bgCol = "#ffffff"
	end if%>

        <tr><td colspan="8">&nbsp;<br><br /></br></td></tr>
	
	<tr bgcolor="<%=bgCol%>">
		<td style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
		<td style="border-bottom:1px #CCCCCC solid;"><%=oRec("id")%></td>
		<td style="border-bottom:1px #CCCCCC solid;">
		
		<a href="erp_valutaer.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("valuta")%> </a>
        <%if cint(oRec("grundvaluta")) = 1 then%>
	    <br /><span style="color:darkred;">(Grundvaluta)</span>
		<%end if %>
		
		</td>
		<td style="border-bottom:1px #CCCCCC solid;"><b><%=oRec("valutakode")%></b></td>
		<td style="border-bottom:1px #CCCCCC solid;" align=right><b><%=oRec("kurs")%></b>
        <%if cint(oRec("grundvaluta")) = 1 then %>
        <br />kurs altid = 100 på grundvaluta
        <%end if %>
        </td>
		<td style="border-bottom:1px #CCCCCC solid;" align=right><%=oRec("antalfak") %>
		&nbsp;</td>
		<td align=right style="padding-top:2px;">
		<%if cint(oRec("antalfak")) = 0 then %>
		<a href="erp_valutaer.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" alt="" border="0"></a>
		<%end if %>
		&nbsp;</td>
		<td>
            &nbsp;</td>
	</tr>
	        
	        
	        <tr bgcolor="<%=bgCol%>">
	        
	        <td colspan=2>
                &nbsp;</td>
	        <td colspan=6>Valuta historik</td>
	        </tr>
	        <!-- '*** Historik *** -->
	        <%
	        strSQLvh = "SELECT vh.dato, vh.kurs, vh.editor FROM valuta_historik vh WHERE vh.valutaid = "& oRec("id") &" ORDER BY vh.id DESC" 
	        
	        oRec2.open strSQLvh, oConn, 3
	        while not oRec2.EOF 
	        %>
	        <tr bgcolor="<%=bgCol%>">
	
	        <td colspan=3>&nbsp;</td>
	        <td><%=oRec2("dato") %></td>
	        <td align=right><%=oRec2("kurs") %></td>
	        <td colspan=3>&nbsp;</td>
	        </tr>
	        
	        <%
	        oRec2.movenext
	        wend 
	        oRec2.close
	        %>
	        
	
	<%
	x = 0
	oRec.movenext
	wend
	%>
	
	</table>
    </div>
	<br><br><br>
	<br><br><br>
&nbsp;
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
