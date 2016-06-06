<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/budget_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
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
	
	if len(Request("FM_budget_brutto_id")) <> 0 then
	budget_brutto_id = Request("FM_budget_brutto_id")
	else
	budget_brutto_id = 0
	end if
	
	thisfile = "budget_medarb.asp"
	
	menu = "erp"
	
	
	function top()%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%call menu_2014() %>
	<div style="position:absolute; left:90px; top:102px;">
	<h3>Budget medarbejdere</h3>
	<% end function
	
	
	public thisAar, strNavn
	function vaelgAar()
	%>
	
	<%call filterheader(0,0,600,pTxt)%>
	<form method=post action=budget_medarb.asp>
	Vælg år: 
	
	<%
	if len(request("FM_aar")) <> 0 then
	vlgtAar = request("FM_aar")
	else
	vlgtAar = year(now)
	end if 
	%>
	
	
        <select id="FM_aar" name="FM_aar">
        <%
        for x = -5 to 5
        
        if x = -5 then
        thisAar = year(now) - 5
        else
        thisAar = cint(thisAar) + 1
        end if
        
        if cint(vlgtAar) = cint(thisAar) then
        arrCHK = "SELECTED"
        else 
        arrCHK = ""
        end if
        %>
            <option value="<%=thisAar %>" <%=arrCHK %>><%=thisAar %></option>
        
        <%
        next
        %>
        </select>
        &nbsp;<input id="Submit1" type="submit" value="Vis år >> " />
	</form>
	</td></tr></table>
	</div>
	
	<h4>Gemte medarbejder budgetter: (<%=vlgtAar %>)</h4>
	
	
	<%
    
    tTop = 0
	tLeft = 0
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	<table cellpadding=3 cellspacing=0 width="100%">
	<tr bgcolor="#8cAAe6">
	<td height=30 class=alt><b>År</b></td>
	<td class=alt><b>Navn</b></td>
	<td class=alt align=center>Slet?</td>
	</tr>
	<%
	thisAar = vlgtAar
	strNavn = ""
	
	strSQL = "SELECT id, navn, budget_aar, budget_brutto_id FROM budget_medarb WHERE budget_aar = " & thisAar
	oRec.open strSQL, oConn, 3 
	x = 0
	while not oRec.EOF 
	
	    if cint(id) = oRec("id") then
	    bgThis = "#ffff99"
	    strNavn = oRec("navn")
	    else
	        select case right(x,1)
	        case 0,2,4,6,8
	        bgThis = "#ffffff"
	        case else
	        bgThis = "#EFf3FF"
	        end select
	    end if
	
	%>
	<tr bgcolor="<%=bgThis %>">
	<td style="border-bottom:1px silver dashed;"><%=oRec("budget_aar") %></td>
	<td style="border-bottom:1px silver dashed;"><a href="budget_medarb.asp?func=red&id=<%=oRec("id") %>&FM_aar=<%=thisaar %>&FM_budget_brutto_id=<%=oRec("budget_brutto_id") %>" class=vmenu><%=oRec("navn") %></a></td>
	<td style="border-bottom:1px silver dashed;" align=center><a href="budget_medarb.asp?func=slet&id=<%=oRec("id") %>&FM_aar=<%=thisaar %>" class=vmenu>
        <img src="../ill/slet_16.gif" border="0" /></a></td>
	</tr>
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close
	
	if x = 0 then%>
	<tr bgcolor="#ffffff"><td colspan=3>(Ingen)</td></tr>
	<%end if %>
	</table>
	</div>
	
	<br />
	
	
	
	<%
	
	oleftpx = 0
	otoppx = 0
	owdtpx = 200
	
	call opretNy("budget_medarb.asp?func=opr&FM_aar="&thisaar&"", "Opret nyt medarb. budget", otoppx, oleftpx, owdtpx) 
	%>
	
	<%if len(strNavn) <> 0 then %>
        &nbsp;
	<h4><%=strNavn%>
	            
	            <%if request("fromupdate") = "1" then
	            Response.Write " - <font color=darkred>Opdateret!</font>"
	            end if
	             %>
	</h4>
	<%
	else
	%>
	
	<br /><br />
	<%
	end if 
	end function
	
	
	function tjkIsInt(val)
    call erDetInt(val) 
    if isInt <> 0 then
    val = 0
    end if  
    
        
        if isInt <> 0 OR len(trim(val)) = 0 then
        %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <%
	    errortype = 113
	    call showError(errortype)
	
	    Response.End 
        end if
    
    isInt = 0
	end function
	
	
	select case func
	case "slet"
	call top()
	
	slttxt = "<b>Slet medarbejder budget?</b><br />"_
	&"Skema kan ikke gen-skabes."
	slturl = "budget_medarb.asp?func=sletja&id="&id&"&FM_aar="&request("FM_aar")
	
	slttxtalt = ""
    slturlalt = ""
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	

	
	case "sletja"
	
	strSQLdel = "DELETE FROM budget_medarb WHERE id = "& id
	oConn.execute(strSQLdel)
	
	Response.Redirect "budget_medarb.asp?FM_aar="&request("FM_aar")
	
	case "dbopr", "reddb"
	
	strNavn = replace(request("FM_navn"), "'", "''")
	if len(trim(strNavn)) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	
	            errortype = 109
				call showError(errortype)
	
	Response.End 
	end if
	
	budget_brutto_id = request("FM_budget_brutto_id")
	budgetAar = request("FM_aar")
	editor = session("user")
	
	if func = "dbopr" then
	            
	            
	            strSQL = "INSERT INTO budget_medarb "_
	            &" (navn, budget_aar, editor, budget_brutto_id) "_
	            &" VALUES ('"& strNavn &"', "& budgetAar &", '"& editor &"', "& budget_brutto_id &")"
            	
	            'Response.Write strSQL
	            'Response.flush
            	
	            oConn.execute(strSQL)
            	
	            strSQL = "SELECT id FROM budget_medarb WHERE id <> 0 ORDER BY id DESC"
	            oRec.open strSQL, oConn, 3
	            if not oRec.EOF then
	            lastid = oRec("id")
	            end if
	            oRec.close
	
	                    
	                    
	else
	
	strSQL = "UPDATE budget_medarb SET "_
	&" navn = '"& strNavn &"', budget_aar = "& budgetAar &", editor = '"& editor &"', budget_brutto_id = "& budget_brutto_id &""_
	&" WHERE id = "& id
    
    'Response.Write strSQL
	'Response.flush
	
	oConn.execute(strSQL)
   
    lastid = id
    
    strSQL = "DELETE FROM budget_medarb_rel WHERE budget_id = "& lastid
    oConn.execute(strSQL)
    
    end if
	
	
	                    
	                    
	                    '*** Opdaterer hver enkelt medarbejder ***'
	
	                    if len(request("antal_m")) <> 0 then
	                    antal_m = split(request("antal_m"),",")
	                    else
	                    antal_m = 0
	                    end if
	                    
	                 
                    	
	                    for m = 0 to UBOUND(antal_m)
	                    
	                        val_1 = 0
                    	    val_2 = 0
                    	    val_3 = 0
                    	    val_4 = 0
                    	    val_5 = 0
                    	    val_6 = 0
                    	    val_7 = 0
                    	    val_8 = 0
                    	    val_9 = 0
                    	    val_10 = 0
                    	    val_11 = 0
                    	    val_12 = 0
	                        
	                        if len(request("FM_masseopdater")) <> 0 then
	                        
	                        masseVal = replace(request("FM_masseopdater"), ",", ".")
	                        
	                        val_1 = masseVal
                    	    val_2 = masseVal
                    	    val_3 = masseVal
                    	    val_4 = masseVal
                    	    val_5 = masseVal
                    	    val_6 = masseVal
                    	    val_7 = masseVal
                    	    val_8 = masseVal
                    	    val_9 = masseVal
                    	    val_10 = masseVal
                    	    val_11 = masseVal
                    	    val_12 = masseVal
	                        else
	                        val_1 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_1"), ",", ".")
                    	    val_2 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_2"), ",", ".")
                    	    val_3 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_3"), ",", ".")
                    	    val_4 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_4"), ",", ".")
                    	    val_5 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_5"), ",", ".")
                    	    val_6 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_6"), ",", ".")
                    	    val_7 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_7"), ",", ".")
                    	    val_8 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_8"), ",", ".")
                    	    val_9 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_9"), ",", ".")
                    	    val_10 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_10"), ",", ".")
                    	    val_11 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_11"), ",", ".")
                    	    val_12 = replace(request("FM_bd_medarb_"&trim(antal_m(m))&"_12"), ",", ".")
                    	    end if
                    	    
                    	    if len(trim(request("FM_masseopdater_tpris"))) <> 0 then
                    	    timepris = replace(request("FM_masseopdater_tpris"), ",", ".")
                    	    else 
                    	    timepris = replace(request("FM_tpris_medarb_"& trim(antal_m(m)) &""), ",", ".")
                    	    end if
                    	    
                    	    if len(trim(request("FM_masseopdater_nt"))) <> 0 then
                    	    nt = replace(request("FM_masseopdater_nt"), ",", ".")
                    	    else 
                    	    nt = replace(request("FM_nt_medarb_"& trim(antal_m(m)) &""), ",", ".")
                    	    end if
                    	    
                    	    
                    	    
                    	    
                    	    call tjkIsInt(val_1)
                    	    call tjkIsInt(val_2)
                    	    call tjkIsInt(val_3)
                    	    call tjkIsInt(val_4)
                    	    call tjkIsInt(val_5)
                    	    call tjkIsInt(val_6)
                    	    call tjkIsInt(val_7)
                    	    call tjkIsInt(val_8)
                    	    call tjkIsInt(val_9)
                    	    call tjkIsInt(val_10)
                    	    call tjkIsInt(val_11)
                    	    call tjkIsInt(val_12)
                    	    call tjkIsInt(timepris)
                    	    call tjkIsInt(nt)
                    	       
                    	    
                    	    
                            strSQL = "INSERT INTO budget_medarb_rel (editor, medid, budget_id,"_
                            &" jan, feb, mar, apr, maj, jun, jul, aug, sep, okt, nov, des, timepris, ntimerpruge)"_
	                        &" VALUES ('"& editor &"', "& antal_m(m) &", "& lastid &","_
	                        &""&val_1&","&val_2&","&val_3&","&val_4&","&val_5&","_
	                        &""&val_6&","&val_7&","&val_8&","&val_9&","&val_10&","&val_11&","&val_12&","&timepris&", "&nt&")"
	                        
	                        
	                        
	                       'Response.Write strSQL & "<br>"
	                       'Response.flush
	                       
	                       oConn.execute(strSQL)
	                    
	                    next
	
	
	medarbfilter = request("medarbfilter")
	
	'Response.end
	Response.redirect "budget_medarb.asp?func=red&id="&lastid&"&FM_aar="&budgetAar&"&FM_budget_brutto_id="&budget_brutto_id&"&fromupdate=1&medarbfilter="&medarbfilter
	
	case "opr", "red"
	
	
	            call top()
	            call vaelgAar()
            	
	            if func = "opr" then
	            dbfunc = "dbopr"
	            else
	            dbfunc = "reddb"
	            end if
	          
	
	%>
	
	
	<%
	
	call filterheader(0,0,600,pTxt)%>
	 <form method=post action=budget_medarb.asp?func=<%=dbfunc%>>
	 
	 <table cellpadding=3 cellspacing=1 border=0>
	 <tr><td>
	 <b>Vælg Brutto / Netto Dage & Timer grundlag:</b> <select name="FM_budget_brutto_id" style="width:300px; font-size:10px;">
	<%
	
	
	'thisAar = vlgtAar
	'strNavn = ""
	
	strSQL = "SELECT id, navn, budget_aar FROM budget_bruttonetto_dt WHERE budget_aar = "& thisAar &" ORDER BY navn"
	oRec.open strSQL, oConn, 3 
	x = 0
	while not oRec.EOF 
	
	    if cint(budget_brutto_id) = oRec("id") then
	    selThis = "SELECTED"
	    else
	    selThis = ""
	    end if
	
	%>
	<option value="<%=oRec("id") %>" <%=selThis %>><%=oRec("budget_aar")%> - <%=oRec("navn") %></option>
	
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close
	
	if x = 0 then%>
	<option value="0">Ingen (Opret først årets Brutto/Netto timer & dage grundlag)</option>
	<%end if %>
	</select></td></tr>
	
	 
	 <tr><td>
    <b>Angiv belpægnings % pr. medarbejder i heltal herunder, </b><br /> eller masseopdater alle medarbejdere her:</td>
    </tr><tr>
    <td>
    <br />
    <b>Norm. Timer pr. uge: <input id="FM_masseopdater_nt" name="FM_masseopdater_nt" type="text" style="width:50px; font-size:10px;" /> </b>
        &nbsp;   <b>Timepris:</b> <input id="FM_masseopdater_tpris" name="FM_masseopdater_tpris" type="text" style="width:50px; font-size:10px;" /> &nbsp;  <b>Belægningsprocent:</b> <input id="FM_masseopdater" name="FM_masseopdater" type="text" style="width:30px; font-size:10px;" /> % </td>
	   </tr>
	   <tr><td>   
	
	<%
	
	        if len(request("medarbfilter")) <> 0 then
            medarbfilter = request("medarbfilter")
            else
            medarbfilter = 1
            end if
            
            select case medarbfilter
            case 1
            medarbSQLkri  = " m.mansat <> '2' "
            mCHK1 = "CHECKED"
            mCHK2 = ""
            mCHK3 = ""
            case 2
            medarbSQLkri  = " m.mansat <> '0' "
            mCHK2 = "CHECKED"
            mCHK1 = ""
            mCHK3 = ""
            case 3
            medarbSQLkri  = " YEAR(m.ansatdato) <= '"& thisAar &"' AND YEAR(m.opsagtdato) >= '"& thisAar &"'"
            mCHK3 = "CHECKED"
            mCHK2 = ""
            mCHK1 = ""
            end select
	
	%>
	
	<br />
	 <b>Medarbejder filter:</b> (Opdater budget for at se medarbejdere)<br />
     <input id="Radio1" name="medarbfilter" type="radio" value=1 <%=mCHK1%> />Vis KUN aktive medarbejdere <br />
	 <input id="Radio1" name="medarbfilter" type="radio" value=2 <%=mCHK2%> />Vis alle <br />
	 <input id="Radio1" name="medarbfilter" type="radio" value=3 <%=mCHK3%> />Vis alle der var / er ansat i det valgte år <b>(<%=thisAar%>)</b> <br />
	 (ansat- og opsagt -dato sættes på den enkelte medarbejder i TSA delen)
      <br />
      <br /><input id="Submit1" type="submit" value=" Opdater >> " />        
        </td></tr>  </table> 
	    
	    
	    <!-- filter header sLut -->
	</td></tr></table>
	</div>
	   <br />   <br />    
	    <%
    
    tTop = 0
	tLeft = 0
	tWdth = 1004
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	    
	            
	            <table cellspacing=1 cellpadding=2 border=0 bgcolor="#CCCCCC" width=100%>
	           
                    <input id="id" name="id" value="<%=id %>" type="hidden" />
                    <input id="FM_aar" name="FM_aar" value="<%=thisaar%>" type="hidden" />
                    
	        <tr bgcolor="#8cAAe6">
                <td width=100 class=alt height=20><b>Medarb./Måned</b></td>
                <td class=lille bgcolor="#CCCCCC"><b>Norm. uge</b><br />
                (fra medarb. type)
                <br />37,5 = 100%</td>
                <td class=alt><b>Budg. Timepris</b></td>
                <td class=alt><b>Jan</b></td>
                <td class=alt><b>Feb</b></td>
                <td class=alt><b>Mar</b></td>
                <td class=alt><b>Apr</b></td>
                <td class=alt><b>Maj</b></td>
                <td class=alt><b>Jun</b></td>
                <td class=alt><b>Jul</b></td>
                <td class=alt><b>Aug</b></td>
                <td class=alt><b>Sep</b></td>
                <td class=alt><b>Okt</b></td>
                <td class=alt><b>Nov</b></td>
                <td class=alt><b>Dec</b></td>
            </tr>
            
            
           
           
            <% 
            
            
            
            m = 0
            strSQL = "SELECT m.mnavn, m.mnr, m.mid, m.init, SUM(normtimer_man+ "_
            &" normtimer_tir+normtimer_ons+normtimer_tor+normtimer_fre+normtimer_lor+normtimer_son) AS normtimer "_
            &" FROM medarbejdere m "_
            &" LEFT JOIN medarbejdertyper mt ON (mt.id = m.medarbejdertype) WHERE "& medarbSQLkri &" GROUP BY m.mid ORDER BY m.mnavn "
            
            'Response.Write strSQL
            'Response.flush
            
            oRec.open strSQL, oConn, 3
            While not oRec.EOF
             m = m + 1
             
             ntimerpruge = oRec("normtimer")
             
             select case right(m, 1)
             case 0,2,4,6,8
             bgt = "#EFF3FF"
             case else 
             bgt = "#ffffff"
             end select%>
            <tr bgcolor="<%=bgt %>">
            <td><b><%=oRec("mnavn") %>   </b> (<%=oRec("mnr") %>)
                
                <%if len(oRec("init")) <> 0 then %>
                &nbsp;- <%=oRec("init") %>
                <%end if %>
              
            </td>
           
             <input id="antal_m" name="antal_m" value="<%=oRec("mid")%>" type="hidden" />
            <%
            
            
            janval = 0
            febval = 0
            marval = 0
            aprval = 0
            majval = 0
            junval = 0
            julval = 0
            augval = 0
            sepval = 0
            oktval = 0
            novval = 0
            desval = 0
            timepris = 0
            
            if func = "red" then
            strSQL2 = "SELECT jan,feb,mar,apr,maj,jun,jul,aug,sep,okt,nov,des,timepris,ntimerpruge FROM budget_medarb_rel WHERE budget_id = "& id &" AND "_
            &" medid = "& oRec("mid")
            
            oRec2.open strSQL2, oConn, 3
            if not oRec2.EOF then
            
            janval = oRec2("jan")
            febval = oRec2("feb")
            marval = oRec2("mar")
            aprval = oRec2("apr")
            majval = oRec2("maj")
            junval = oRec2("jun")
            julval = oRec2("jul")
            augval = oRec2("aug")
            sepval = oRec2("sep")
            oktval = oRec2("okt")
            novval = oRec2("nov")
            desval = oRec2("des")
            timepris = oRec2("timepris")
            ntimerpruge = oRec2("ntimerpruge")
            
            end if
            oRec2.close
            
           end if
            
       
            
            %>
            <td>
                <input id="FM_nt_medarb_<%=oRec("mid")%>" name="FM_nt_medarb_<%=oRec("mid")%>" size=4 value="<%=ntimerpruge%>" type="text" /></td>
             <td>
                <input id="FM_tpris_medarb_<%=oRec("mid")%>" name="FM_tpris_medarb_<%=oRec("mid")%>" size=8 type="text" value="<%=timepris %>" /></td>
            <%
            
            for x = 1 to 12 
            
                select case x
                case 1
                valthis = janval
                case 2
                valthis = febval
                case 3
                valthis = marval
                case 4
                valthis = aprval
                case 5
                valthis = majval
                case 6
                valthis = junval
                case 7
                valthis = julval
                case 8
                valthis = augval
                case 9
                valthis = sepval
                case 10
                valthis = oktval
                case 11
                valthis = novval
                case 12
                valthis = desval
                end select
            %>
            <td align=right class=lille>
            <input id="FM_bd_medarb_<%=oRec("mid")%>_<%=x%>" name="FM_bd_medarb_<%=oRec("mid")%>_<%=x%>" type="text" style="width:35px; font-size:9px;" value="<%=valthis%>" /> %</td>
           
            <%next %>
            </tr>
            <% 
            oRec.movenext
            wend
            oRec.close
            %>
            
            </table>
            </div>
            
            
            <br />
            <b>Gem medarbejder budget som...</b><br />
             
            * Navn: <input id="FM_navn" name="FM_navn" type="text" value="<%=strNavn %>" size="40" />
            <input id="Submit1" type="submit" value="Gem / Opdater >>" />
            </form>	
            	
	 <%
	case else
		
	 call top()
	 call vaelgAar()
	


	
	end select
	%>

<br /><br /><br />&nbsp;
	</div>
	<%
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
