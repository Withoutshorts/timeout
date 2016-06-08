<%

	
	
	function jobtopmenu()
		%>
		
		<br><a href='jobs.asp?menu=job&shokselector=1&fromvemenu=j' class='rmenu'>Joboversigt</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='jobs.asp?menu=job&func=opret&id=0&int=1' class='rmenu'>Opret nyt job</a>
		<%if level <= 2 OR level = 6 then %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='job_print.asp?menu=job&kid=0&id=0' class='rmenu'>Print / PDF Center</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='projektgrupper.asp?menu=job' class='rmenu'>Projektgrupper</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='akt_gruppe.asp?menu=job&func=favorit' class='rmenu'>Stam-aktivitets grupper (Job skabeloner)</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='ulev_gruppe.asp?menu=job&func=favorit' class='rmenu'>Underleverandør grp. / Salgsomkostninger</a>

		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='pipeline.asp?menu=job&FM_kunde=0&FM_progrupper=10' class='kal_g'>Pipeline >> </a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='webblik_joblisten21.asp' class='kal_g'>Jobplanner (Gantt årsoversigt) >> </a>
		<%end if %>
		<br>&nbsp;
		
		<%
	end function
	

    public belobTotThisGrp, antalOpenCalc, antalMtypGrp
    function hovedgrpsum(lastGruppenavnTxt, gruppenavnTxt, timerTotThisGrp, timeprisTotThisGrp, faktorTotThisGrp, belobTotThisGrp, basisValISO, visning, grpid, opencalc, cnt, kostbelobTotThisGrp)

                if lastGruppenavnTxt <> gruppenavnTxt OR visning = 1 then
                
                if belobTotThisGrp = 1 then
                belobTotThisGrp = 0
                end if

                lastGruppenavnTxt = replace(lastGruppenavnTxt, "(", "")
                lastGruppenavnTxt = replace(lastGruppenavnTxt, ")", "")
                
                if len(trim(opencalc)) <> 0 then
                opencalc = opencalc
                else
                opencalc = -1
                end if

                if len(trim(grpid)) <> 0 then
                grpid = grpid
                else
                grpid = 0
                end if

                if opencalc = 0 then '** Der KAN KUN være 1 opencalc.
                mxlngt = 0
                bgcolor = "#CCCCCC"
                bdr = "border:0px;"
                antalOpenCalc = antalOpenCalc + 1
                else
                mxlngt = 20
                bgcolor = "#FFFFFF"
                bdr = ""
                antalOpenCalc = antalOpenCalc
                end if

                antalMtypGrp = antalMtypGrp + 1

                if timerTotThisGrp > 0 then
                timeprisTotThisGrp = (timeprisTotThisGrp/timerTotThisGrp)
                faktorTotThisGrp = (faktorTotThisGrp/timerTotThisGrp)
                belobffTotThisGrp = (timerTotThisGrp * timeprisTotThisGrp)
                else
                timeprisTotThisGrp = timeprisTotThisGrp
                faktorTotThisGrp = faktorTotThisGrp
                belobffTotThisGrp = 0    
                end if

                if kostbelobTotThisGrp <> 0 then
                kostbelobTotThisGrp = kostbelobTotThisGrp
                else
                kostbelobTotThisGrp = 0
                end if

                if faktorTotThisGrp = 0 then
                faktorTotThisGrp = 1
                end if
                %>
                <tr><td style="height:20px; font-size:9px; white-space:nowrap; border:0px;"><b><%=lastGruppenavnTxt %> ialt:</b></td>
                <td style="font-size:9px; white-space:nowrap; padding-left:3px; border:0px;"><input type="text" id="FM_mt_timer_totgrp_<%=grpid %>" value="<%=formatnumber(timerTotThisGrp, 2) %>" name="" style="width:40px; font-size:9px; border:0px;" maxlength="0" /></td>
                <td style="font-size:9px; white-space:nowrap; padding-left:3px; border:0px;"><input type="text" id="FM_mt_timepris_totgrp_<%=grpid %>" value="<%=formatnumber(timeprisTotThisGrp, 2) %>" name="" style="width:75px; font-size:9px; border:0px;" maxlength="0" /></td>
                <td style="font-size:9px; white-space:nowrap; padding-left:9px; border:0px;"><input type="text" id="FM_mt_belobff_totgrp_<%=grpid %>" value="<%=formatnumber(belobffTotThisGrp, 2) %>" name="" style="width:40px; font-size:9px; border:0px;" maxlength="0" /></td>
                <td style="font-size:9px; white-space:nowrap; padding-left:9px; border:0px;"><input type="text" id="FM_mt_faktor_totgrp_<%=grpid %>" value="<%=formatnumber(faktorTotThisGrp, 2) %>" name="" style="width:40px; font-size:9px; border:0px;" maxlength="0" /></td>
                <td style="font-size:9px; white-space:nowrap; border:0px;">= <input type="text" class="mt_belob_totgrp" id="FM_mt_belob_totgrp_<%=grpid %>" value="<%=formatnumber(belobTotThisGrp, 2) %>" name="" style="width:75px; font-size:9px; <%=bdr%>;" maxlength="<%=mxlngt %>" />

                    <input type="hidden" id="FM_mt_kostbelob_totgrp_<%=grpid %>" value="<%=formatnumber(kostbelobTotThisGrp, 2) %>" name=""/>

                </td>
                </tr>
                <tr><td style="border:0px;" colspan=6 align="right">&nbsp;
                <%if opencalc <> 0 then 'AND func <> "red" %>
                
                    
                    <input type="checkbox" id="FM_laasbelob_<%=grpid %>" class="FM_laasbelob" value="1" /> Lås beløb
                    
                   
                <%else %>
                  <input type="checkbox" id="FM_laasbelob_<%=grpid %>" class="FM_laasbelob" value="1" CHECKED/> Lås beløb
                <%end if %>
                     </td></tr>
                 <tr><td style="border:0px;" colspan=6 align="right">&nbsp;</td></tr>

                <input type="hidden" value="<%=opencalc %>" name="FM_mt_opencalc_<%=grpid %>" id="FM_mt_opencalc_<%=grpid %>" />
                 <input type="hidden" id="FM_mtype_slutno_<%=grpid%>" name="FM_mtype_slutno" value="<%=cnt%>"/>
                
                <%

               belobTotThisGrp = 0
               timerTotThisGrp = 0
               timeprisTotThisGrp = 0
               faktorTotThisGrp = 0
                    kostBelobTotThisGrp = 0
                end if


    end function
	


function jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)
 
%>
<td class="jobmenu" align=center id="<%=lnkid%>" style="white-space:nowrap; width:<%=lnkwdt%>px; padding:4px; background-color:<%=bgcol%>; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="<%=lnk %>" id="<%=lnkid%>_A" class=alt target="<%=lnktgt%>"><%=lnktxt %></a></td>
<%
end function

function jobfanebladjv(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)
 
%>
<td class="jobmenu" align=center id="<%=lnkid%>" style="white-space:nowrap; width:<%=lnkwdt%>px; padding:4px; background-color:<%=bgcol%>; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" onclick="Javascript:window.open('<%=lnk %>', '', 'width=1104,height=800,resizable=yes,scrollbars=yes')" class="alt"><%=lnktxt %></a>
</td>
<%
end function


public jo_gnstpris, gnsSalgsogKostprisTxt, gnsinttpris
function gnstp_fn(antalM, virkgnskostpris)

                     if antalM <> 0 AND virkgnskostpris <> 0 then  
			           
                        gnsSalgsogKostprisTxt = formatnumber(virkgnssalgspris/antalM, 2) &" / "& formatnumber(virkgnskostpris/antalM, 2)  
						    
					
					       
					    select case lto
                        case "epi", "epi_no", "epi_sta", "epi_ab"
                        gnsinttpris = virkgnssalgspris/antalM
                        case else 
                        gnsinttpris = virkgnskostpris/antalM
                        end select
					

                    else 
                        gnsSalgsogKostprisTxt = "na/na"
					    gnsinttpris = 0
                    
                    end if 
                    
                    if func <> "red" then
                    jo_gnstpris = gnsinttpris 
                    end if

end function





public visstamdata_top, vistpris_top, tpris_vzb, tpris_dsp, stdata_vzb, stdata_dsp
public fVzb, fDsp, oVzb, oDsp
function faneblade(side, div)
            

            call antalmedarb_fn(2)

'Response.Write "side " & side &" & div "& div
'Response.end

select case side
case "job"
strNavn = strNavn
id = id
rdir = rdir	
	
	
		
	select case div
	case "tpriser"
	visstamdata_top = 159
	vistpris_top = 161
	
	tpris_vzb = "visible"
	tpris_dsp = ""
	
	stdata_vzb = "visible"
	stdata_dsp = ""

    fVzb = "hidden"
    fDsp = "none"

    oVzb = "hidden"
    oDsp = "none"
	
    case "forkalk"
    visstamdata_top = 161
	vistpris_top = 159

    tpris_vzb = "hidden"
	tpris_dsp = "none"
    
    stdata_vzb = "visible"
	stdata_dsp = ""

    fVzb = "visible"
    fDsp = ""

    oVzb = "hidden"
    oDsp = "none"

	case else
	visstamdata_top = 161
	vistpris_top = 159
	
	tpris_vzb = "hidden"
	tpris_dsp = "none"
	
	stdata_vzb = "visible"
	stdata_dsp = ""

    oVzb = "visible"
    oDsp = ""

    fVzb = "hidden"
    fDsp = "none"
    

    end select
	
	stamdatLink = "#" 'onClick='showstamdata()	

    if antalmedarb < 25 then
    tprisLink = "#"
    else
	tprisLink = "jobs.asp?menu=job&func=red&id="&id&"&int=1&rdir="&rdir&"&showdiv=tpriser&FM_mtype="&mtype
    end if

    fkalkLink = "#"
    oblikLink = "#"
    milepLink = "#"

case "aktiv"
    
        strNavn = jobnavnThis
        id = jobid
        rdir = rdir	
        strKnr = kid
    
    select case div
	case "tpriser"
	visstamdata_top = 159
	vistpris_top = 161
	
	tpris_vzb = "visible"
	tpris_dsp = ""
	
	stdata_vzb = "hidden"
	stdata_dsp = "none"
	
	case else
	visstamdata_top = 161
	vistpris_top = 159
	
	tpris_vzb = "hidden"
	tpris_dsp = "none"
	
	stdata_vzb = "visible"
	stdata_dsp = ""
	end select
	
	stamdatLink = "jobs.asp?menu=job&func=red&id="&id&"&int=1&rdir="&rdir	
	tprisLink = "jobs.asp?menu=job&func=red&id="&id&"&int=1&rdir="&rdir&"&showdiv=tpriser"
    fkalkLink = "jobs.asp?menu=job&func=red&id="&jid&"&int=1&rdir="&rdir&"&showdiv=fkalk"
    oblikLink = stamdatLink
	milepLink = "jobs.asp?menu=job&func=red&id="&jid&"&int=1&rdir="&rdir&"&showdiv=milep"
	
		
case else
	 'f.eks milepæle
		
	
	visstamdata_top = 159
	vistpris_top = 159
	
	tpris_vzb = "visible"
	tpris_dsp = ""
	
	stdata_vzb = "hidden"
	stdata_dsp = ""
	
	stamdatLink = "jobs.asp?menu=job&func=red&id="&jid&"&int=1&rdir="&rdir	
	tprisLink = "jobs.asp?menu=job&func=red&id="&jid&"&int=1&rdir="&rdir&"&showdiv=tpriser"
    fkalkLink = "jobs.asp?menu=job&func=red&id="&jid&"&int=1&rdir="&rdir&"&showdiv=fkalk"
    oblikLink = stamdatLink
    milepLink = "milepale.asp?menu=job&jid="&id&"&func=list&rdir="&rdir&"&kundeid="&strKnr

    tVzb = "hidden"
    tDsp = "none"

	
strNavn = jobnavn
id = jid
rdir = rdir
end select






'*** Jobfaneblade **'
%>
 <div id=jobmenuDiv style="position:absolute; left:100px; top:132px; background-color:#ecebe8;">
    <!--showakt:<%=showakt %>-->
    <table cellpadding=0 cellspacing=0 border=0>
        <tr bgcolor="#ecebe8">


<%


'*** stamdata ***'
'lnkidtd = "tdvisstd"
lnkid = "visstd"
lnktxt = "Jobstamdata" 'left(strNavn, 15) 
lnk = stamdatLink
lnkwdt = 100
lnktgt = ""
bgcol = "#8CAAe6"
call jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)


select case lto
case "oko", "intranet - local"
    
    select case level
    case 1

    showjobfaneblad_budget = 1
    showjobfaneblad_mile = 1
    showjobfaneblad_gantt = 1
    showjobfaneblad_filer = 1
    showjobfaneblad_timep = 1

    case 2,6

    showjobfaneblad_budget = 1
    showjobfaneblad_mile = 0
    showjobfaneblad_gantt = 0
    showjobfaneblad_filer = 0
    showjobfaneblad_timep = 0

    case else

    showjobfaneblad_budget = 0
    showjobfaneblad_mile = 0
    showjobfaneblad_gantt = 0
    showjobfaneblad_filer = 0
    showjobfaneblad_timep = 0

    end select

case else

    showjobfaneblad_budget = 1
    showjobfaneblad_mile = 1
    showjobfaneblad_gantt = 1
    showjobfaneblad_timep = 1
    showjobfaneblad_filer = 1

end select


if side <> "aktiv" AND side <> "mile" then

'*** Spcae ***'
'lnkidtd = "tdvisgantt"
lnkid = "sp"
lnktxt = ""
lnk = "#"
lnkwdt = 370
lnktgt = "_blank"
bgcol = "#ecebe8" '#d6dff5"


call jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)



'*** Minivising ***'
'lnkidtd = "tdvisgantt"
lnkid = "oblik"
lnktxt = "Joboverblik"
lnk = oblikLink
lnkwdt = 85
lnktgt = ""
bgcol = "#8CAAe6"
call jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)

'*** forkalk ***'
'lnkidtd = "tdvisgantt"
lnkid = "forkalk"
lnktxt = "Job & Aktivitet. (Budget)"
lnk = fkalkLink
lnkwdt = 150
lnktgt = ""
bgcol = "#8CAAe6"

if cint(showjobfaneblad_budget) = 1 then
call jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)
end if


'*** gantt ***'
'lnkidtd = "tdvisgantt"
lnkid = "visgantt"
lnktxt = "Gantt"
lnk = "webblik_joblisten21.asp?nomenu=1&jobnr_sog="&fbjobnr&"&FM_kunde="&fbKnr
lnkwdt = 55
lnktgt = "_blank"
bgcol = "#8CAAe6"

'Response.write "strJobnr & strKnr: "& strJobnr & strKnr
if cint(showjobfaneblad_gantt) = 1 then
call jobfanebladjv(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)
end if

'*** aktiviteter ***'
'lnkidtd = "tdvisakt"
lnkid = "visakt"
lnktxt = "Aktiviteter"
lnk = "aktiv.asp?menu=job&jobid="&id&"&jobnavn="&strNavn&"&rdir="&rdir
lnkwdt = 100
lnktgt = ""
if side = "aktiv" then
bgcol = "#5582d2"
else
bgcol = "#8CAAe6"
end if

'call jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)


'*** milepæle ***'
'lnkidtd = "tdvismp"
lnkid = "vismp"
lnktxt = "Milepæle & Term."
'lnk = "milepale.asp?menu=job&jid="&id&"&func=list&rdir="&rdir&"&kundeid="&strKnr
lnk = milepLink
lnkwdt = 105
lnktgt = ""
bgcol = "#8CAAe6"
if cint(showjobfaneblad_mile) = 1 then
call jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)
end if
'*** timepriser ***'
'lnkidtd = "tdvistp"
lnkid = "vistp"
if antalmedarb < 25 then
lnkwdt = 125
lnktxt = "Medarb. timepriser"
else
lnkwdt = 155
lnktxt = "Medarb. timepriser (reload)"
end if 'onClick="showtpris()"
lnk = tprisLink

lnktgt = ""
bgcol = "#8CAAe6"

if cint(showjobfaneblad_timep) = 1 then
call jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)
end if

'*** Filarkiv ***'
'lnkidtd = "tdvisfil"
lnkid = "visfil"
lnktxt = "Filarkiv" 
lnk = "filer.asp?kundeid="&fbKnr&"&jobid="&id&"&nomenu=1" 'strKnr
lnkwdt = 65
lnktgt = "_blank"
bgcol = "#8CAAe6"

if cint(showjobfaneblad_filer) = 1 then
call jobfaneblad(lnkid, lnktxt, lnk, lnkwdt, lnktgt, bgcol)
end if

end if

%>
</tr>
</table>
</div>




<%
end function











'******************************************************************************
'*** div projekt beregner ****'
'********************************************************************************  
sub projektberegner


	            
end sub %>	
