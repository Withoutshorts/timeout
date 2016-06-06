<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/budget_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

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
	
	if len(request("FM_budget")) <> 0 then
	budgetid = request("FM_budget")
	else
	budgetid = 0
	end if
	
	thisfile = "budget_aar_dato.asp"
	print = request("print")
	
	menu = "erp"
	
	
	'*** Akt. typer der tælles med i ugereg. ***'
	call akttyper2009(2)
	
	function top()
	
	if print <> "j" then
	tp = 132
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <%
    
        call menu_2014()    
        
        else 
	
	tp = 20%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if %>
	
	
	<div style="position:relative; top:80px; left:20px; border:0px #000000 solid;">
	
	<h3>Budget År / Dato</h3>
	<% end function
	
	
	
	public expTxt
	function bal(tal1, tal2)
	
	tal1 = formatnumber(tal1, 0)
	tal2 = formatnumber(tal2, 0)
	
	'if vismd = 1 then
	  
            bal = 0
            bal = (tal1-(tal2)) 
            
    
                if bal > 0 then
                tdcls = "lillegron"
                else
                    if bal = 0 then
                 
                    tdcls = "lille"
                    else
                  
                    tdcls = "lillerod"
                    end if
                end if
    'else       
           
            
    'bal = 9999999999
    
    'end if
    
    if len(bal) <> 0 then
    bal = bal
    else
    bal = 9999999999
    end if
    
    bgBal = bgt
    
    
            %>
            <td bgcolor=<%=bgBal %> align=right class=<%=tdcls %> style="border-bottom:1px silver dashed; border-right:1px silver solid;">
            <%if bal <> 9999999999 then %>
            <%=formatnumber(bal,0)%>
            <%expTxt = expTxt & formatnumber(bal,0) &";" %>
            <%else %>
            <%expTxt = expTxt & formatnumber(0,0) &";" %>
            <%end if %>
            </td>
            
           <%
            
	
	end function
	
	
	function timerOtd(ctspan) 
	%>
    <td align=center colspan=<%=ctspan%> class=lille style="padding-right:2px; border-left:2px silver solid; border-right:1px silver solid;">Timer</td>
    <%
    
    expTxt = expTxt &"Timer"
    
    for z = 1 to ctspan
    expTxt = expTxt &";"
    next
           
	end function
	
	
	function tprisOtd(ctspan) 
	%>
    <td align=center colspan=<%=ctspan%> class=lille style="padding-right:2px;">Gns. Timepriser</td>
    <%
    
    expTxt = expTxt &"Gns. Timepriser"
    
    for z = 1 to ctspan
    expTxt = expTxt &";"
    next
    
	end function
	
	
	sub oskriftL2A
	
	 %>
               <td align=right class=lille style="padding-right:2px; border-left:2px silver solid; border-right:1px silver solid;">Bgt.</td>
                
                <%
                expTxt = expTxt &"Bgt;"
                
                if vis <> 5 then %>
                <td align=right class=lille style="padding-right:2px;"><%=tdoskrift %></td>
                <%
                
                expTxt = expTxt & tdoskrift &";"
                
                else %>
                <td align=right class=lille style="padding-right:2px;">Real.</td>
                <td align=right class=lille style="padding-right:2px; border-right:1px silver solid;">Bal.</td>
                <td align=right class=lille style="padding-right:2px;">Fak.</td>
                <%
                
                expTxt = expTxt &"Real.;Bal.;Fak.;"
                
                end if %>
                <td align=right class=lille style="padding-right:2px; border-right:1px silver solid;">Bal.</td>
                <%
                
                expTxt = expTxt &"Bal.;"
	
	end sub
	
	sub oskriftL2B
	
	
              
                    
                    %>
                    <td align=right class=lille style="padding-right:2px;">Bgt.</td>
                    <%
                    
                    expTxt = expTxt &"Bgt.;"
                    
                    if vis <> 5 then%>
                    <td align=right class=lille style="padding-right:2px;">Gns.</td>
                    <%
                    
                    expTxt = expTxt &"Gns.;"
                    
                    else
                    %>
                    <td align=right class=lille style="padding-right:2px;">Real.</td>
                    <td align=right class=lille style="padding-right:2px;">Fak.</td>
                    <%
                    
                    expTxt = expTxt &"Real.;Fak.;"
                    
                    end if
                
              
	
	end sub
	
	sub gnsTpriserA
	
	 %>
            <td class=lille align=right style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(budgetGnsTpris, 0) %>
            
            <%
            expTxt = expTxt &""& formatnumber(budgetGnsTpris, 0) &";"
            %>
            
            </td>
            <td class=lille align=right style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(gnstRpris, 0) %>
            
            <%
            expTxt = expTxt &""& formatnumber(gnstRpris, 0) &";"
            %>
            
           </td>
            <td class=lille align=right style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(gnstFpris, 0) %>
            
            <%
            expTxt = expTxt &""& formatnumber(gnstFpris, 0) &";"
            %>
            
            </td>
            <% 
            
    end sub
	
	
	sub gnsTpriserB
	
	  '**** Total i periode gns timepriser ****'
            
            
            '*** Budget i periode ***'
            budgetGnsTprisTot = (budgetGnsTprisTot / 3)
            
            '** Real i periode *'
            if antalrealTimerTot <> 0 then
            tRealgnsTpris = (gnstprisThisVaegtetTot / antalrealTimerTot)
            else
            tRealgnsTpris = 0
            end if
            
            '*** Fak i periode **'
            if tFaktimerTot <> 0 then
            tTotFakgnsTpris = (tTotFakgnsTpris / tFaktimerTot) 
            '* Kun timer på fakturaer. kreditnota timer er IKKE med *'
            else
            tTotFakgnsTpris = 0
            end if
            
            
             %>
            <td class=lille align=right style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(budgetGnsTprisTot, 0) %>
            
            <%
            expTxt = expTxt &""& formatnumber(budgetGnsTprisTot, 0) &";"
            %>
            
            </td>
            <td class=lille align=right style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(tRealgnsTpris, 0) %>
            
            <%
            expTxt = expTxt &""& formatnumber(tRealgnsTpris, 0) &";"
            %>
            
            </td>
            <td class=lille align=right style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(tTotFakgnsTpris, 0) %>
            
            <%
            expTxt = expTxt &""& formatnumber(tTotFakgnsTpris, 0) &";"
            %>
            
            </td>
            <%
            
           
	end sub
	
	sub realfakbalTimerA
	%> 
            
           <td align=right class=lille style="padding-right:2px; border-bottom:1px silver dashed;">
           <%=formatnumber(tRealtimer,0)%>
           </td>
            
            <%
            expTxt = expTxt &""& formatnumber(tRealtimer, 0) &";"
            %>
            
            <%
            
            tal1 = tRealtimer
            tal2 = valthis
            
            
            call bal(tal1, tal2) %>
            
            <td align=right class=lille style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(tFaktimer,0)%>
            </td>
           
            <%
            expTxt = expTxt &""& formatnumber(tFaktimer, 0) &";"
            %>
           
            
            <%
            tal1 = tFaktimer
            tal2 = valthis
            
            call bal(tal1, tal2) %>
            
            <%
	
	
	end sub
	
	
	
	sub realfakbalTimerB
	
	%> 
            <td align=right class=lille style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(tTotalTimer,0)%>
            </td>
            
            <%
            expTxt = expTxt &""& formatnumber(tTotalTimer, 0) &";"
            %>
            
            <%
            tal1 = tTotalTimer
            tal2 = valthisTot
            
            
            call bal(tal1, tal2) 
            %>
            <td align=right class=lille style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(tTotalFaktimer,0)%>
            </td>
            
            <%
            expTxt = expTxt &""& formatnumber(tTotalFaktimer, 0) &";"
            %>
            
            <%
            tal1 = tTotalFaktimer
            tal2 = valthisTot
           
            call bal(tal1, tal2)
	
	end sub
	
	
	public thisAar, strNavn, tdoskrift, vis, t5m, headerOskrift
	function vaelgBudget(budgetid)
	
	call filterheader(0,0,600,pTxt)
	
	if print <> "j" then%>
	<form method=post action=budget_aar_dato.asp?func=se>
	<br /><b>Vælg medarbejder budget:</b>
	
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
	
	x = x + 1
	oRec.movenext
	wend 
	oRec.close
	
	if x = 0 then
	%>
	<option value="0">Ingen (Opret først medab. budget)</option>
	<%
	end if
	
	if print <> "j" then%>
    </select>
    
    <br /><br />
    <b>Vis 12 måneder md/md + ÅTD for:</b><br />
    <%end if %>
    
    
    <%
    if len(request("FM_vis5month")) <> 0 then
    t5m = request("FM_vis5month")
    else
    t5m = 0
    end if
    
    if len(request("FM_vis")) <> 0 then
    vis = request("FM_vis")
    else
    vis = 1
    end if
    
    visTxt1 = "Budget / Realiseret fakturerbare timer (Afsætning)"
    visTxt2 = "Budget / Faktureret (Omsætning)"
    visTxt3 = "Budget / Ressource forecast (Belægning)"
    visTxt4 = "Budget / Normeret tid (Normering) "
    visTxt5 = "Nøgletal"
    
    select case vis
    case 1
    CHKvis1 = "CHECKED"
    CHKvis2 = ""
    CHKvis3 = ""
    CHKvis4 = ""
    CHKvis5 = ""
    headerOskrift = "Realiserede timer"
    tdoskrift = "Real."
    
    ptxt = visTxt1
    case 2
    CHKvis1 = ""
    CHKvis2 = "CHECKED"
    CHKvis3 = ""
    CHKvis4 = ""
    CHKvis5 = ""
    headerOskrift = "Faktureret"
    tdoskrift = "Fak."
    
    ptxt = visTxt2
    case 3
    CHKvis1 = ""
    CHKvis2 = ""
    CHKvis3 = "CHECKED"
    CHKvis4 = ""
    CHKvis5 = ""
    headerOskrift = "Ressource timer"
    tdoskrift = "Ress."
    
    ptxt = visTxt3
    case 4
    CHKvis1 = ""
    CHKvis2 = ""
    CHKvis3 = ""
    CHKvis4 = "CHECKED"
    CHKvis5 = ""
    headerOskrift = "Normerede timer"
    tdoskrift = "Norm."
    
    ptxt = visTxt4
    
    case 5
    CHKvis1 = ""
    CHKvis2 = ""
    CHKvis3 = ""
    CHKvis4 = ""
    CHKvis5 = "CHECKED"
    tdoskrift = ""
    headerOskrift = "Sammenlign Budget / Real. / Fakturerede timer i periode"
    ptxt = visTxt4
    
    end select  
    
    if print <> "j" then%>
    <input id="FM_vis1" name="FM_vis" value="1" type="radio" <%=CHKvis1 %> /><%=visTxt1 %> <br /> 
    <input id="FM_vis2" name="FM_vis" value="2" type="radio" <%=CHKvis2 %> /><%=visTxt2 %> <br />
    <input id="FM_vis3" name="FM_vis" value="3" type="radio" <%=CHKvis3 %> /><%=visTxt3 %> <br />  
    <input id="FM_vis4" name="FM_vis" value="4" type="radio" <%=CHKvis4 %> /><%=visTxt4 %> <br />  
    
    <br /><b>Sammenlign Budget / Real. / fakturerede timer og timepriser for: </b><br />
    <input id="Radio1" name="FM_vis" value="5" type="radio" <%=CHKvis5 %> />  
     
  
     
        Vælg periode: <select name="FM_vis5month" id="FM_vis5month" style="font-size:9px;">
        
            
            <%for m = 1 to 16
            
            if m < 13 then
            nameLabel = monthname(m)
            else
                select case m
                case 13
                nameLabel = "1 Kvartal"
                case 14
                nameLabel = "2 Kvartal"
                case 15
                nameLabel = "3 Kvartal"
                case 16
                nameLabel = "4 Kvartal"
                'case 17
                'nameLabel = "1 ½år Januar - Juni"
                'case 18
                'nameLabel = "2 ½år Juli - December"
                end select 
            end if 
            
            if cint(m) = cint(t5m) then
            mSEL = "SELECTED"
            else
            mSEL = ""
            end if
            
            %>
            <option value="<%=m %>" <%=mSEL %> ><%=nameLabel%></option>
            <%
            
            next%>
            
            
          
        </select>
        
        <img src="../ill/blank.gif" width=1 height=5 /><br /><br />
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
	
	
	
	
	
	%>
	<br />
&nbsp;<br />
	<h4><%=headerOskrift%></h4>
	
	
	<%
	
	
    
    
    
    tTop = 0
	tLeft = 0
	tWdth = 1204
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	if vis <> 5 then
	t1 = 1
	t2 = 12
	else
	
	select case t5m
	case 1,2,3,4,5,6,7,8,9,10,11,12
	t1 = t5m
	t2 = t5m 
	mtname = monthname(t1)
	mspan = 8
	case 13
	t1 = 1
	t2 = 3
	mtname = "1 kvartal"
	mspan = 24
	case 14
	t1 = 4
	t2 = 6
	mspan = 24
	mtname = "2 kvartal"
	case 15 
	t1 = 7
	t2 = 9
	mspan = 24
	mtname = "3 kvartal"
	case 16
	t1 = 10
	t2 = 12
	mtname = "4 kvartal"
	mspan = 24
	
	'case 17
	't1 = 1
	't2 = 6
	'mtname = "1 ½år Januar - Juni"
	'mspan = 48
	'case 18
	't1 = 7
	't2 = 12
	'mtname = "2 ½år Juli - December"
	'mspan = 48
	'case 91
	't1 = 1
	't2 = 12
	'mspan = 96
	'mtname = "Januar - December"
	end select
	
	end if
	%>         
	
	
            
	        <table cellspacing=0 cellpadding=1 border=0 width=100%>
	         <tr bgcolor="#8cAAe6">
                <td width=120 class=alt height=30 style="border-bottom:1px #5582d2 solid; border-right:1px #FFFFFF solid;"><b>Medarb./Måned</b></td>
                
                <%select case vis
                case 5
                %>
                <td class=alt colspan=<%=mspan %> align=center style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b><%=mtname %></b></td>
                <%
                
                expTxt = "Medarb.;Mnr.;Init;"& mtname &";"
                
                if t5m < 12 then
                expTxt = expTxt &";;;;;;;"
                else
                expTxt = expTxt &";;;;;;;;;;;;;;;;;;;;;;;"
                'expTxt = expTxt &"xx99123sy#z;;;"
                end if
                
                
	            
	
                
                case else %>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Jan</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Feb</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Mar</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Apr</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Maj</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Jun</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Jul</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Aug</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Sep</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Okt</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Nov</b></td>
                <td align=center class=alt colspan=3 style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Dec</b></td>
                <%
                
                expTxt = "Medarb.;Mnr.;Init;Jan;;;Feb;;;Mar;;;Apr;;;Maj;;;Jun;;;Jul;;;Aug;;;Sep;;;Okt;;;Nov;;;Dec;;;"
	            
	
                
                end select %>
                
                <%
                
               
                
                select case vis
                case 1,2
                cps = 5
                case 5
                cps = 8
                case else
                cps = 3
                end select 
                
                
                if vis = 5 AND t5m > 12 then
                %>
                <td align=center bgcolor="#c4c4c4" colspan="<%=cps %>" style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>Total i periode</b></td>
                
                <%
                expTxt = expTxt &"Total i periode;;;;;;;;"
                end if%>
                
                <td align=center bgcolor="pink" colspan="<%=cps %>" style="border-right:1px #FFFFFF solid; border-bottom:1px #5582d2 solid;"><b>ÅTD (Jan - Dec)</b></td>
                <%
                expTxt = expTxt &"ÅTD (Jan - Dec);;;;;;;;"
                expTxt = expTxt &"xx99123sy#z;;;"
                %>
               
            
            </tr>
            
            <tr bgcolor="#ffffff">
                <td>
                    &nbsp;</td>
                    
                    
                <%for t = t1 to (t2+1) %>
               
                <%
                
                
                if vis <> 5 then 
                ctspan = 3 
                else 
                ctspan = 5 
                end if 
                
                call timerOtd(ctspan) 
                
                if (vis = 3 OR vis = 4) AND t = (t2+1) then
                expTxt = expTxt &"xx99123sy#z;;;"
                end if
                
                
                if (t = (t2+1) AND (vis = 1 OR vis = 2)) OR vis = 5 then
                
                if vis <> 5 then
                ctspan = 2 
                else
                ctspan = 3
                end if
                
                call tprisOtd(ctspan)
                
                if vis = 5 AND t = t2 AND t5m > 12 then
                call timerOtd(5) 
                call tprisOtd(ctspan)
                end if
                
                if t = (t2+1) then
                expTxt = expTxt &"xx99123sy#z;;;"
                end if
                
                end if
                
                
               
                
                next %>
                
                
                 </tr>
            
            <tr bgcolor="#ffffff">
                <td>&nbsp;</td>
                    
                    
                <%for t = t1 to (t2+1) 
                
                call oskriftL2A
                
                
                if (t = (t2+1) AND (vis = 1 OR vis = 2)) OR vis = 5 then
                call oskriftL2B
                end if
                
                if vis = 5 AND t = t2 AND t5m > 12 then
                call oskriftL2A
                call oskriftL2B
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
            
            
            '*** Alle **'
            antalTimer = 0
            
            '*** Budget **'
            valthis = 0
            valthisTot = 0
            budgetGnsTprisTot = 0
            budgetGnsTpris = 0
            valthisTot12md = 0
            
            '** Real **'
            antalrealTimerTot = 0
            tTotalTimer = 0
            tRealgnsTpris = 0
            gnstprisThisVaegtetTot = 0
            
            '** fak **'
            tTotFakgnsTpris = 0
            fakantal = 0
            tTotalFaktimer = 0
            tFaktimerTot = 0
            
          
            
             select case right(m, 1)
             case 0,2,4,6,8
             bgt = "#ffffff"
             case else 
             bgt = "#Eff3ff"
             end select
             
             %>
            <tr bgcolor="<%=bgt%>">
            
           
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
            
            if janval <> "" then
            janval = janval
            else
            janval = 0
            end if
            
             if febval <> "" then
            febval = febval
            else
            febval = 0
            end if
            
             if marval <> "" then
            marval = marval
            else
            marval = 0
            end if
            
             if aprval <> "" then
            aprval = aprval
            else
            aprval = 0
            end if
            
             if majval <> "" then
            majval = majval
            else
            majval = 0
            end if
            
             if junval <> "" then
            junval = junval
            else
            junval = 0
            end if
            
             if julval <> "" then
            julval = julval
            else
            julval = 0
            end if
            
             if augval <> "" then
            augval = augval
            else
            augval = 0
            end if
            
             if sepval <> "" then
            sepval = sepval
            else
            sepval = 0
            end if
            
             if oktval <> "" then
            oktval = oktval
            else
            oktval = 0
            end if
            
             if novval <> "" then
            novval = novval
            else
            novval = 0
            end if
            
            if desval <> "" then
            desval = desval
            else
            desval = 0
            end if
            
            valthisTot12md = (cint(janval) + cint(febval) + cint(marval) + cint(aprval) + cint(majval) + cint(junval) + cint(julval) + cint(augval) + cint(sepval) + cint(oktval) + cint(novval) + cint(desval))
            
            for x = t1 to (t2+1)
            
            
            if x <> t2+1 then
            
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
                slDato = "31/12/"&oRec("budget_aar")
               
               
                'case t2+1
                'valthis = valthisTot
                'stDato = oRec("budget_aar") &"/1/1"
                
                '    if year(dagsDato) = year(stDato) then
                '    slDato = dagsDato
                '    else
                '    slDato = oRec("budget_aar") &"/12/31"
                '    end if
                    
                end select
                
              
              else
                
                
                valthis = valthisTot
                stDato = oRec("budget_aar") &"/1/1"
                slDato = oRec("budget_aar") &"/12/31"
                   
              end if
              
                
                if len(valthis) <> 0 then
                valthis = valthis
                else
                valthis = 0
                end if
                
                if len(oRec("timepris")) <> 0 then
                budgetGnsTpris = oRec("timepris")
                else
                budgetGnsTpris = 0
                end if
                
                if x <> (t2+1) then
                budgetGnsTprisTot = budgetGnsTprisTot + budgetGnsTpris
                else
                budgetGnsTprisTot = budgetGnsTprisTot  
                end if
                
                '*** ÅTD vis måned? ****' 
                'if (cdate(dagsDato) > cDate(slDato) OR x = (t2+1) OR month(dagsDato) = month(slDato)) OR vis = 3 OR vis = 4 then
                slDato = year(slDato) &"/"& month(slDato) &"/"& day(slDato)
                'vismd = 1
                'else
                'slDato = stDato
                'vismd = 0
                'end if
                
                '*** Undgå at beløb for dec. bliver lagt til 2 gange. ***'
                '*** valthisTot = valthisTot i periode **'
                
                if x <> (t2+1) then
                valthisTot = valthisTot + valthis
                end if
                
            %>
            <!-- Budget -->
            <td align=right class=lille style="padding-right:2px; border-bottom:1px silver dashed; border-right:1px silver solid; border-left:2px silver solid;">
            <%if x <> (t2+1) then %>
            <%=formatnumber(valthis,0)%>
            
            <%expTxt = expTxt & formatnumber(valthis,0) &";" %>
            
            <%else %>
            <%=formatnumber(valthisTot12md,0)%>
            
            <%expTxt = expTxt & formatnumber(valthisTot12md,0) &";" %>
            
            <%end if %>
            </td>
            
            
            
            <!-- Realiseret timer/fak/forecast -->
            <% 
            
            
            if vis = 1 OR vis = 5 then 'AND (x <> (t2+1))
            
            antalrealTimer = 0
            gnstprisThisVaegtet = 0
            gnstpris = 0
            gnstRpris = 0
            
            strSQLtim = "SELECT t.timer, t.timepris, "_
            &" j.budgettimer, j.jobtpris, j.fastpris, j.id AS jid FROM timer t "_
            &" LEFT JOIN job j ON (j.jobnr = t.tjobnr) WHERE t.tmnr = "& oRec("mid") &" AND t.tdato "_
            &" BETWEEN '"&stDato&"' AND '"&slDato&"' AND ("& aty_sql_realhours &")"
            
            'if oRec("mid") = 1 then
            'Response.Write strSQLtim
            'Response.flush
            'end if
            
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
            
            timeprisThis = oRec2("timepris")
            gnstprisThisVaegtet = gnstprisThisVaegtet + (oRec2("timepris") * oRec2("timer"))
          
            end if
            
            'Response.Write "oRec2(jobtpris): " & oRec2("jobtpris") & " ("& oRec2("jid") &") | "& oRec2("timer") &" * tp "& timeprisThis &" = gnsTprisVægt: "& gnstprisThisVaegtet  &"<br>"
            
            oRec2.movenext
            wend 
            oRec2.close
            
           
            
            if x <> (t2+1) AND vis = 5 AND t5m > 12 then
            antalrealTimerTot = antalrealTimerTot + antalrealTimer
            else
            antalrealTimerTot = antalrealTimer
            end if
            
            
            if antalrealTimerTot <> 0 then
            gnstRpris = gnstprisThisVaegtet / antalrealTimerTot
            else
            gnstRpris = gnstRpris
            end if
            
            if antalrealTimer <> 0 then
            antalTimer = antalrealTimer 
            else
            antalTimer = 0
            end if
            
            
            
            
            gnstprisThisVaegtetTot = gnstprisThisVaegtetTot + gnstprisThisVaegtet
         
            tRealtimer = antalrealTimer 'antalrealTimerTot 
            tTotalTimer = tTotalTimer + tRealtimer
            gnstpris = gnstRpris 'tRealgnsTpris 'gnstRpris
           
           
            end if
            
           
           
            
            if vis = 2 OR vis = 5 then 'AND (x <> (t2+1))
            
            '**** Fakturaer ****'
            
            
            antalfaktimer = 0
            gnstFpris = 0
            gnstpris = 0
            faktimerKre = 0
            gnstFprisSum = 0
           
            
            strSQL0 = "SELECT sum(fak) AS faktimer, sum(fms.fak * fms.enhedspris) AS faktpris FROM fakturaer f"_
            &" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.mid = "& oRec("mid") &")"_
            &" WHERE fakdato "_
            &" BETWEEN '"&stDato&"' AND '"&slDato&"' AND faktype = 0 AND fak > 0 GROUP BY fms.mid"
            
            
            
            'if oRec("mid") = 1 then
            'Response.Write strSQL0 & "<br>"
            'Response.flush
            'end if
            
            oRec2.open strSQL0, oConn, 3
            if not oRec2.EOF then
            
            antalfaktimer = oRec2("faktimer")
            gnstFpris = (oRec2("faktpris") / antalfaktimer)
            gnstFprisSum = oRec2("faktpris")
            
            end if
            oRec2.close
            
            
            
            tFaktimer = antalfaktimer
            gnstpris = gnstFpris
            
            tFaktimerTot = tFaktimerTot + tFaktimer 
            tTotFakgnsTpris = tTotFakgnsTpris + gnstFprisSum
            
            
            '*** Minus Kredit notaer ****'
            faktimerKre = 0
            strSQL1 = "SELECT sum(fak) AS faktimerKre FROM fakturaer f"_
            &" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.mid = "& oRec("mid") &")"_
            &" WHERE fakdato "_
            &" BETWEEN '"&stDato&"' AND '"&slDato&"' AND faktype = 1 AND fak > 0 GROUP BY fms.mid"
            
            'if oRec("mid") = 1 then
            'Response.Write strSQL1
            'Response.flush
            'end if
            
            oRec2.open strSQL1, oConn, 3
            if not oRec2.EOF then
            
            faktimerKre = oRec2("faktimerKre")
            
            end if
            oRec2.close
            
            antalfaktimer = antalfaktimer - (faktimerKre)
            tFaktimer = antalfaktimer
            tTotalFaktimer = tTotalFaktimer + (antalfaktimer)
           
            
            'if oRec("mid") = 1 then
            'Response.Write "<br><br>tFaktimer"& tFaktimer & "faktimerKre: "& faktimerKre &" antalfaktimer: "& antalfaktimer  &"  tTotalFaktimer: " & tTotalFaktimer & "<hr>"
            'end if
            
            antalTimer = antalfaktimer 
           
            
            end if
            
            
            
            
            
            
            if vis = 3 then
            
            strSQLr = "SELECT sum(rmd.timer) AS restimer FROM ressourcer_md rmd "_
			&" LEFT JOIN job j ON (j.id = jobid AND j.jobnavn IS NOT NULL)"_
			&" WHERE medid = "& oRec("mid") &" AND  "_
            &" ((rmd.md >= "& month(stDato) &" AND rmd.aar = "& year(stDato) &") "_
            &" AND (rmd.md <= "& month(slDato) &" AND rmd.aar = "& year(slDato) &"))"_
            &" AND j.jobnavn IS NOT NULL GROUP BY medid "
        	
            
            'Response.Write strSQLr &"<br>"
        	
        	resTimerThisA = 0
			oRec2.open strSQLr, oConn, 3
            if not oRec2.EOF then
            
                resTimerThisA = oRec2("restimer") 
            
            end if
            oRec2.close
            
            antalTimer = resTimerThisA
            
            end if
            
            
            
            
            if vis = 4 then
            
            interval = dateDiff("d",stDato, slDato, 2,2)
            
            'if oRec("mid") = 1 then
            'Response.Write "<br>interval" & interval &" Datoer: "& stDato &","& slDato
            'end if
            
            call normtimerPer(oRec("mid"), stDato, interval, 0)
            antalTimer = nTimPer
            
            end if
            
            
            if len(antalTimer) <> 0 then
            antalTimer = antalTimer
            else
            antalTimer = 0
            end if
            
            'expTxt = expTxt & formatnumber(antalTimer,0) &";" 
             
            
            
            select case vis 
            case 5
            
            call realfakbalTimerA
            call gnsTpriserA
            
            
            '*** Total i valgte periode ***'
            if x = t2 AND t5m > 12 then
            budgetTimeriPeri = valthisTot
            %>
            <td align=right class=lille style="padding-right:2px; border-bottom:1px silver dashed; border-right:1px silver solid; border-left:2px silver solid;">
            <%=formatnumber(budgetTimeriPeri,0)%>
           </td>
            <%
            
            expTxt = expTxt & formatnumber(budgetTimeriPeri, 0) &";"
            
            call realfakbalTimerB
            call gnsTpriserB
            
            
            end if
            
            
            case else
            %>
            
            <td align=right class=lille style="padding-right:2px; border-bottom:1px silver dashed;">
            <%=formatnumber(antalTimer,0)%>
            </td>
            <%
            
            expTxt = expTxt & formatnumber(antalTimer, 0) &";"
            
            tal1 = formatnumber(antalTimer, 0)
            
            if x <> (t2+1) then
            tal2 = formatnumber(valthis, 0)
            else
            tal2 = valthisTot12md
            end if
            
            call bal(tal1, tal2)
            
                    if x = (t2+1) AND (vis = 1 OR vis = 2) then %>
                    <td class=lille align=right style="padding-right:2px; border-bottom:1px silver dashed;"><%=formatnumber(budgetGnsTpris, 0) %></td>
                    <td class=lille align=right style="padding-right:2px; border-bottom:1px silver dashed;"><%=formatnumber(gnstpris, 0) %></td>
                    <%
                    
                    expTxt = expTxt & formatnumber(budgetGnsTpris, 0) &";"
                    expTxt = expTxt & formatnumber(gnstpris, 0) &";"
                    
                    end if
            
            end select 
            
            
            
            next 
            
            
            
            %>
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
            
            call sideinfo(itop, ileft, iwdt) 
            %>
            
                * Budgettimer er baseret på netto timer<br />
            	* Kun timer registreret på fakturerbare aktiviter er medregnet i realiserede timer og timepriser.<br />
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
            
            ptop = 37
            pleft = 622
            pwdt = 220
            
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
                <td>.csv fil eksport</td>
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
