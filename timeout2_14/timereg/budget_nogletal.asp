<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/budget_func.asp"-->
<!--#include file="inc/helligdage_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<SCRIPT Language=JavaScript>
	function BreakItUp()
	{
	  //Set the limit for field size.
	  var FormLimit = 302399 
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm.BigTextArea.value
	
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm.appendChild(objTEXTAREA)
	
	      TempVar = TempVar.substr(FormLimit)
	    }
	  }
	}
	</script>
	
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
	
	if len(request("FM_budget")) <> 0 then
	budgetid = request("FM_budget")
	else
	budgetid = 0
	end if
	
	thisfile = "budget_aar_dato.asp"
	print = request("print")
	
	
	
	function top()
	
	if print <> "j" then
	tp = 132
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	    <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	    <%call erpmainmenu(4)%>
	    </div>
	    <div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	    <%
	    if showonejob <> 1 then
		    call budgettopmenu(1)
	    end if
	    %>
	    </div>
	
	<%else 
	
	tp = 20%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if %>
	
	
	<div style="position:relative; top:80px; left:20px; border:0px #000000 solid;">
	
	<h3>Budget Nøgletal</h3>
	<% end function
	
	
	public thisAar, strNavn, tdoskrift, vis
	function vaelgBudget(budgetid)
	
	call filterheader(0,0,600)
	
	if print <> "j" then%>
	<form method=post action=budget_aar_dato.asp?func=se>
	<b>Vælg medarbejder budget:</b>
	
	<select id="FM_budget" name="FM_budget" style="width:200px; font-size:10px;">
	<% 
    end if
	
	strSQL = "SELECT id, navn, budget_aar FROM budget_medarb WHERE id <> 0 ORDER BY budget_aar DESC "
	oRec.open strSQL, oConn, 3 
	x = 0
	while not oRec.EOF 
	
	if oRec("id") = cint(budgetid) then
	thisSel = "SELECTED"
	else
	thisSel = ""
	end if
	
	if print <> "j" then%>
	<option value="<%=oRec("id") %>" <%=thisSel %>><%=oRec("budget_aar") %> - <%=oRec("navn") %></option>
	<%else
	    if thisSel = "SELECTED" then%>
	    Valgt budget: <b><%=oRec("budget_aar") %> - <%=oRec("navn") %></b>
	    <%end if
	end if
	
	oRec.movenext
	wend 
	oRec.close
	
	if print <> "j" then%>
    </select>
    
    <br /><br />
    <b>Vis:</b><br />
    <%end if %>
    
    
    <%
    if len(request("FM_vis")) <> 0 then
    vis = request("FM_vis")
    else
    vis = 1
    end if
    
    visTxt1 = "Budget / Realiseret fakturerbare timer (Afsætning)"
    visTxt2 = "Budget / Faktureret (Omsætning)"
    visTxt3 = "Budget / Ressource forecast (Belægning)"
    visTxt4 = "Budget / Normeret tid (Normering) "
    
    select case vis
    case 1
    CHKvis1 = "CHECKED"
    CHKvis2 = ""
    CHKvis3 = ""
    CHKvis4 = ""
    tdoskrift = "Real."
    
    ptxt = visTxt1
    case 2
    CHKvis1 = ""
    CHKvis2 = "CHECKED"
    CHKvis3 = ""
    CHKvis4 = ""
    tdoskrift = "Fak."
    
    ptxt = visTxt2
    case 3
    CHKvis1 = ""
    CHKvis2 = ""
    CHKvis3 = "CHECKED"
    CHKvis4 = ""
    tdoskrift = "Ress."
    
    ptxt = visTxt3
    case 4
    CHKvis1 = ""
    CHKvis2 = ""
    CHKvis3 = ""
    CHKvis4 = "CHECKED"
    tdoskrift = "Norm."
    
    ptxt = visTxt4
    end select  
    
    if print <> "j" then%>
    <input id="FM_vis1" name="FM_vis" value="1" type="radio" <%=CHKvis1 %> /><%=visTxt1 %> <br /> 
    <input id="FM_vis2" name="FM_vis" value="2" type="radio" <%=CHKvis2 %> /><%=visTxt2 %> <br />
    <input id="FM_vis3" name="FM_vis" value="3" type="radio" <%=CHKvis3 %> /><%=visTxt3 %> <br />  
    <input id="FM_vis4" name="FM_vis" value="4" type="radio" <%=CHKvis4 %> /><%=visTxt4 %> <br />  
        
        <img src="../ill/blank.gif" width=1 height=5 /><br />
        &nbsp;<input id="submit" type="submit" value="Vis budget >> " />
	</form>
    <%
    else
    %>
    <br />Viser: <b><%=ptxt %></b>
    <%
    end if
    
    %>
    </td></tr></table>
	</div>
    <%
	end function
	
	
	
	
	
	select case func
	case "se"
	
	
	            call top()
	            call vaelgbudget(budgetid)
            	
	            if func = "opr" then
	            dbfunc = "dbopr"
	            else
	            dbfunc = "reddb"
	            end if
	            
	            dagsDato = date()
	
	
	expTxt = "Medarb.;Mnr.;Init;Jan;;;Feb;;;Mar;;;Apr;;;Maj;;;Jun;;;Jul;;;Aug;;;Sep;;;Okt;;;Nov;;;Dec;;;ÅTD;;;"
	expTxt = expTxt &"xx99123sy#z;;;"
	
	
	
	
	
    
    tTop = 20
	tLeft = 0
	tWdth = 1204
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>         
	
	
            	
	        <table cellspacing=0 cellpadding=1 border=0 width=100%>
	         <tr bgcolor="#8cAAe6">
                <td width=120 class=alt height=20><b>Medarb./Måned</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Jan</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Feb</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Mar</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Apr</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Maj</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Jun</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Jul</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Aug</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Sep</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Okt</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Nov</b></td>
                <td class=alt colspan=3 style="border-right:1px #FFFFFF solid;"><b>Dec</b></td>
                
                <%if vis = 1 OR vis = 2 then
                cps = 5
                else
                cps = 3
                end if %>
                
                <td align=center bgcolor="pink" colspan="<%=cps %>" style="border-right:1px silver solid;"><b>ÅTD</b></td>
            </tr>
            <tr bgcolor="#ffffff">
                <td>
                    &nbsp;</td>
                <%for t = 1 to 13 %>
                <td align=right class=lille style="padding-right:2px;">Bgt.</td>
                <td align=right class=lille style="padding-right:2px;"><%=tdoskrift %></td>
                <td align=right class=lille style="padding-right:2px; border-right:1px silver solid;">Bal.</td>
                <%
                
                expTxt = expTxt &"Bgt.;"&tdoskrift&";Bal.;"
                
                if t = 13 AND (vis = 1 OR vis = 2) then
                %>
                <td align=right class=lille style="padding-right:2px;">Bgt. tpris</td>
                <td align=right class=lille style="padding-right:2px;">Gns. tpris</td>
                <%
                
                expTxt = expTxt &"Bgt. tpris; Gns. tpris.;"
                
                end if
                
               
                
                next %>
            </tr>
            
            
           
           
            <% 
            m = 0
            strSQL = "SELECT bm.budget_aar, budget_brutto_id, budget_id, jan, feb, mar, "_
            &" apr, maj, jun, jul, aug, sep, okt,"_
            &" nov, des, m.mnavn, m.mnr, m.mid, m.init,"_
            &" dt.netto_1, dt.netto_2, dt.netto_3, dt.netto_4, dt.netto_5, dt.netto_6, "_
            &" dt.netto_7, dt.netto_8, dt.netto_9, dt.netto_10, dt.netto_11, dt.netto_12, "_
            &" bml.ntimerpruge, bml.timepris "_
            &" FROM budget_medarb bm "_
            &" LEFT JOIN budget_medarb_rel bml ON (budget_id = bm.id) "_
            &" LEFT JOIN budget_bruttonetto_dt dt ON (dt.id = bm.budget_brutto_id)"_
            &" LEFT JOIN medarbejdere m ON (m.mid = medid) "_
            &" WHERE bm.id = "& budgetid &" ORDER BY m.mnavn "
            
            'Response.Write strSQL
            'Response.flush
            
            oRec.open strSQL, oConn, 3
            While not oRec.EOF
            m = m + 1
            valthisTot = 0
            
             select case right(m, 1)
             case 0,2,4,6,8
             bgt = "#ffffff"
             case else 
             bgt = "#Eff3ff"
             end select
             
             %>
            <tr bgcolor="<%=bgt %>">
            
           
            <td height=30 class=lille style="border-bottom:1px silver dashed;"><b><%=oRec("mnavn") %></b> (<%=oRec("mnr") %>)
                
                <%if len(oRec("init")) <> 0 then %>
                 - <%=oRec("init") %>
                <%end if %>
                
            </td>
             
            <%
            expTxt = expTxt &"xx99123sy#z"
            expTxt = expTxt & oRec("mnavn") & ";"& oRec("mnr") &";"& oRec("init") &";"
            
           
            '*** Nortimer i forhold til 37,5 pr. uge = 100% ****
            ntimKvo = formatnumber((oRec("ntimerpruge")/37.5), 2) 
            'ntimKvo = 1   
            
                  
          
            janval = (oRec("jan")/100) * oRec("netto_1") * ntimKvo
            febval = (oRec("feb")/100) * oRec("netto_2") * ntimKvo
            marval = (oRec("mar")/100) * oRec("netto_3") * ntimKvo
            aprval = (oRec("apr")/100) * oRec("netto_4") * ntimKvo
            majval = (oRec("maj")/100) * oRec("netto_5") * ntimKvo
            junval = (oRec("jun")/100) * oRec("netto_6") * ntimKvo
            julval = (oRec("jul")/100) * oRec("netto_7") * ntimKvo
            augval = (oRec("aug")/100) * oRec("netto_8") * ntimKvo
            sepval = (oRec("sep")/100) * oRec("netto_9") * ntimKvo
            oktval = (oRec("okt")/100) * oRec("netto_10") * ntimKvo
            novval = (oRec("nov")/100) * oRec("netto_11") * ntimKvo
            desval = (oRec("des")/100) * oRec("netto_12") * ntimKvo
            
           
            
            for x = 1 to 13 
            
                select case x
                case 1
                valthis = janval
                stDato = oRec("budget_aar") &"/1/1"
                slDato = dateadd("d", -1, "1/2/"&oRec("budget_aar")) 
                case 2
                valthis = febval
                stDato = oRec("budget_aar") &"/2/1"
                slDato = dateadd("d", -1, "1/3/"&oRec("budget_aar"))
                case 3
                valthis = marval
                stDato = oRec("budget_aar") &"/3/1"
                slDato = dateadd("d", -1, "1/4/"&oRec("budget_aar"))
                case 4
                valthis = aprval
                stDato = oRec("budget_aar") &"/4/1"
                slDato = dateadd("d", -1, "1/5/"&oRec("budget_aar"))
                case 5
                valthis = majval
                stDato = oRec("budget_aar") &"/5/1"
                slDato = dateadd("d", -1, "1/6/"&oRec("budget_aar"))
                case 6
                valthis = junval
                stDato = oRec("budget_aar") &"/6/1"
                slDato = dateadd("d", -1, "1/7/"&oRec("budget_aar"))
                case 7
                valthis = julval
                stDato = oRec("budget_aar") &"/7/1"
                slDato = dateadd("d", -1, "1/8/"&oRec("budget_aar"))
                case 8
                valthis = augval
                stDato = oRec("budget_aar") &"/8/1"
                slDato = dateadd("d", -1, "1/9/"&oRec("budget_aar"))
                case 9
                valthis = sepval
                stDato = oRec("budget_aar") &"/9/1"
                slDato = dateadd("d", -1, "1/10/"&oRec("budget_aar"))
                case 10
                valthis = oktval
                stDato = oRec("budget_aar") &"/10/1"
                slDato = dateadd("d", -1, "1/11/"&oRec("budget_aar"))
                case 11
                valthis = novval
                stDato = oRec("budget_aar") &"/11/1"
                slDato = dateadd("d", -1, "1/12/"&oRec("budget_aar"))
                case 12
                valthis = desval
                stDato = oRec("budget_aar") &"/12/1"
                slDato = oRec("budget_aar") &"/12/31"
                case 13
                valthis = valthisTot
                stDato = oRec("budget_aar") &"/1/1"
                
                    if year(dagsDato) = year(stDato) then
                    slDato = dagsDato
                    else
                    slDato = oRec("budget_aar") &"/12/31"
                    end if
                    
                end select
                
                if len(valthis) <> 0 then
                valthis = valthis
                else
                valthis = 0
                end if
                
                
                
                '*** ÅTD vis måned? ****' 
                if cdate(dagsDato) > cDate(slDato) OR x = 13 OR month(dagsDato) = month(slDato) OR vis = 3 then
                slDato = year(slDato) &"/"& month(slDato) &"/"& day(slDato)
                vismd = 1
                else
                slDato = stDato
                vismd = 0
                end if
                
                if x <> 13 AND vismd = 1 then
                valthisTot = valthisTot + valthis
                end if
                
            %>
            <!-- Budget -->
            <td align=right class=lille style="border-bottom:1px silver dashed;">
            <%=formatnumber(valthis,0)%>
            </td>
            
            <%expTxt = expTxt & formatnumber(valthis,0) &";" %>
            
            <!-- Realiseret timer/fak/forecast -->
            <%
            antalTimer = 0
            
            if vismd = 1 then
            
            select case vis
            case 1
            nstprisThisVaegtet = 0
            antalrealTimer = 0
            gnstpris = 0
            strSQLtim = "SELECT t.timer, t.timepris, "_
            &" j.budgettimer, j.jobtpris, j.fastpris FROM timer t "_
            &" LEFT JOIN job j ON (j.jobnr = t.tjobnr) WHERE t.tmnr = "& oRec("mid") &" AND t.tdato "_
            &" BETWEEN '"&stDato&"' AND '"&slDato&"' AND t.tfaktim = 1"
            
            'Response.Write strSQLtim
            'Response.flush
            
            oRec2.open strSQltim, oConn, 3
            while not oRec2.EOF 
            
            antalrealTimer = antalrealTimer + oRec2("timer")
            
            if oRec2("fastpris") = 1 then 'fastpris 
            
            
                if oRec2("budgettimer") > 0 then
				timeprisThis = (oRec2("jobtpris") / oRec2("budgettimer"))
				else
				timeprisThis = (oRec2("jobtpris") / 1)
				end if
            
            gnstprisThisVaegtet = gnstprisThisVaegtet + (timeprisThis * oRec2("timer"))
            
            else
            
            gnstprisThisVaegtet = gnstprisThisVaegtet + (oRec2("timepris") * oRec2("timer"))
          
            end if
            
            oRec2.movenext
            wend 
            oRec2.close
            
            if antalrealTimer <> 0 then
            gnstpris = (gnstprisThisVaegtet / antalrealTimer)
            antalTimer = antalrealTimer 
            else
            gnstpris = 0
            antalTimer = 0
            end if
            
            case 2
            
            '**** Fakturaer ****'
            antalfaktimer = 0
            gnstpris = 0
            strSQL0 = "SELECT sum(fak) AS faktimer, sum(fms.fak * fms.enhedspris) AS faktpris FROM fakturaer f"_
            &" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.mid = "& oRec("mid") &")"_
            &" WHERE fakdato "_
            &" BETWEEN '"&stDato&"' AND '"&slDato&"' AND faktype = 0 AND fak > 0 GROUP BY fms.mid"
            
            oRec2.open strSQL0, oConn, 3
            if not oRec2.EOF then
            
            antalfaktimer = oRec2("faktimer")
            gnstpris = (oRec2("faktpris") / antalfaktimer)
            
            end if
            oRec2.close
            
            
            '*** Minus Kredit notaer ****'
            faktimerKre = 0
            strSQL1 = "SELECT sum(fak) AS faktimerKre FROM fakturaer f"_
            &" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.mid = "& oRec("mid") &")"_
            &" WHERE fakdato "_
            &" BETWEEN '"&stDato&"' AND '"&slDato&"' AND faktype = 1 AND fak > 0 GROUP BY fms.mid"
            
            oRec2.open strSQL1, oConn, 3
            if not oRec2.EOF then
            
            faktimerKre = oRec2("faktimerKre")
            
            end if
            oRec2.close
            
            antalfaktimer = antalfaktimer -(faktimerKre)
            antalTimer = antalfaktimer 
            
            case 3
            
            strSQLr = "SELECT sum(rmd.timer) AS restimer FROM ressourcer_md rmd "_
			&" LEFT JOIN job j ON (j.id = jobid AND j.jobnavn IS NOT NULL)"_
			&" WHERE medid = "& oRec("mid") &" AND  "_
            &" ((rmd.md >= "& month(stDato) &" AND rmd.aar = "& year(stDato) &") "_
            &" AND (rmd.md <= "& month(slDato) &" AND rmd.aar = "& year(slDato) &"))"_
            &" AND j.jobnavn IS NOT NULL GROUP BY medid "
        	
            
            'Response.Write strSQLr &"<br>"
        	
        	resTimer = 0
			oRec2.open strSQLr, oConn, 3
            if not oRec2.EOF then
            
                resTimer = oRec2("restimer") 
            
            end if
            oRec2.close
            
            antalTimer = resTimer
            
            
            case 4
            
            interval = dateDiff("d",stDato, slDato, 2,2)
            
            call normtimerPer(oRec("mid"), stDato, interval)
            antalTimer = nTimPer
            
            
            end select
            
            
            end if
            
            if len(antalTimer) <> 0 then
            antalTimer = antalTimer
            else
            antalTimer = 0
            end if
            
            
            
            
            expTxt = expTxt & formatnumber(antalTimer,0) &";" 
             
             %>
            
            <td align=right class=lille style="border-bottom:1px silver dashed;">
            <%if vismd = 1 then %>
            <%=formatnumber(antalTimer,0)%>
            <%end if %>
            
            <%if month(dagsDato) = month(slDato) AND x <> 13 AND year(dagsDato) = year(slDato) AND vis <> 3 then %>
            ..
            <%end if %>
         
            </td>
            
            <!-- Bal -->
            <%
            
            if vismd = 1 then
            
            bal = 0
            bal = (antalTimer-(valthis)) 
            
                if vis = 4 then
                bgBal = "#8cAAe6"
                else
                
                if bal > 0 then
                bgBal = "yellowgreen"
                else
                    if bal = 0 then
                    bgBal = bgt
                    else
                    bgBal =  "#ff6666" '"#ff1e5f" '"#ff3300"
                    end if
                end if
            
            end if
            
            else
            bal = 9999999999
            bgBal = bgt
            end if
            %>
            <td bgcolor=<%=bgBal %> align=right class=lille style="border-bottom:1px silver dashed; border-right:1px silver solid;">
            <%if bal <> 9999999999 then %>
            <%=formatnumber(bal,0)%>
            <%expTxt = expTxt & formatnumber(bal,0) &";" %>
            <%else %>
            <%expTxt = expTxt & formatnumber(0,0) &";" %>
            <%end if %>
            
            </td>
            
            <%if x = 13 AND (vis = 1 OR vis = 2) then %>
            <td class=lille align=right style="border-bottom:1px silver dashed;"><%=formatnumber(oRec("timepris"), 0) %></td>
            <td class=lille align=right style="border-bottom:1px silver dashed;"><%=formatnumber(gnstpris, 0) %></td>
            <%
            expTxt = expTxt & formatnumber(oRec("timepris"), 0) &";"
            expTxt = expTxt & formatnumber(gnstpris, 0) &";"
            end if %>
            
            <%next %>
            </tr>
            <% 
            oRec.movenext
            wend
            oRec.close
            %>
            
            </table>
            </div>
            
            
           
          
            <%
            itop = 40
            ileft = 0
            iwdt = 300
            
            call sideinfo(itop, iwdt, ileft) 
            %>
            
                * Budgettimer er baseret på netto timer<br />
            	* Kun fakturerbare aktiviter er medregnet i realiserede timer.<br />
            	* Alle tal er afrundet med 0 decimaler.<br /><br />
            	* Gns. Timepris i realiserede timer er beregnet udfra en vægtet timepris på antal <u>timer * timepris</u>. <br />
            	Hvis job står til at være et fastpris job er <u>timepris = (jobbduget / antal forkalk. fakturerbare timer)</u>.<br />
	            <br />
	            * Gns. faktureret timepris er ikke fratrukket evt. kreditnotaer.
	            </td>
	            </tr>
	            </table>
	            </div>
	            <br /><br />
                &nbsp;
            
            	
	 <%
	case else
		
	 call top()
	 call vaelgbudget(budgetid)
	

    end select
    
    
	
			if print <> "j" then
            
            ptop = 40
            pleft = 770
            pwdt = 120
            
            call eksportogprint(ptop, pleft, pwdt)
            %>

            <form action="budget_adt_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="datointerval" id="datointerval" value="<%=thisAar%>">
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=expTxt%>">
			<input type="hidden" name="txt20" id="txt20" value="">
			
            <tr>
                <td align=center><input type="image" src="../ill/export1.png">
                </td>
                <td>.txt fil eksport</td>
               </tr>
                </form>
                
                <tr>
                <td align=center>
                
			    <a href="budget_aar_dato.asp?print=j&func=se&FM_budget=<%=budgetid %>&FM_vis=<%=vis %>" target="_blank"><img src="../ill/printer3.png" border=0></a>
    			
		       </td><td>Print version</td>
               </tr>
               </form>
	            
               </table>
            </div>
            
            <%
            else
            Response.Write("<script language=""JavaScript"">window.print();</script>")
            end if%>
	
	
	
    </div>
    <br /><br /><br /><br /><br /><br />
&nbsp;
	
	<%
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
