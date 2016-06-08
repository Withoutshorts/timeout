
<%

 public wipHistdato, restestimat, stade_tim_proc
 function wip_historik_fn(jobid,slutdato)
 wipHistdato = "1-1-2001"

             strSQLSeljWiphist = "SELECT dato, restestimat, stade_tim_proc, medid, jobid FROM wip_historik WHERE jobid = "& oRec("id") & " AND dato <= '"& sqlDatoslut &"' ORDER BY dato DESC, id DESC LIMIT 1"
             
             'response.write strSQLSeljWiphist & "<br>"
            
             oRec6.open strSQLSeljWiphist, oConn, 3
             if not oRec6.EOF then

                    restestimat = oRec6("restestimat")
                    stade_tim_proc = oRec6("stade_tim_proc")
                    wipHistdato = oRec6("dato")

             end if
             oRec6.close


end function


public stexptxt, mth_1, mth_2, mth_3
sub stadeindm(usemrn, visning, dtuse)

'** visning = 1, popup fra tiemreg. siden
'** Visning = 2, fra link (stat_opdater_igv.asp)

if datepart("d", now, 2,2) > 22 OR (datepart("m", now, 2,2) = 1 AND datepart("d", now, 2,2) < 15) OR visning <> 1 then




'call erERPaktiv()

stexptxt = ""
jobasnvigv = 0
strSQLigv = "SELECT jobasnvigv FROM licens WHERE id = 1"
oRec.open strSQLigv, oConn, 3
if not oRec.EOF then
jobasnvigv = oRec("jobasnvigv")
end if
oRec.close

if jobasnvigv = 1 then 'AND(session("mid") = 59 OR session("mid") = 1) then 'DW + support 


''*** if usemrn = 0 / Admin view ***'
'if usemrn = 0 then
'strSQLallm = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"




'** Admin view **'
if usemrn <> 0 then


'*** Er md afsluttet for medarbejer ***'
yNow = year(dtuse)
mNow = month(dtuse)


strSQLigvst = "SELECT medid, maaned, aar FROM job_igv_status "_
&" WHERE medid = " & usemrn & " AND aar = "& yNow & " AND maaned = "& mNow & " GROUP BY medid"
mdAngivet = 0
mdAngivetCHK = ""
oRec.open strSQLigvst, oConn, 3
if not oRec.EOF then
mdAngivet = 1
mdAngivetCHK = "CHECKED"
end if
oRec.close


strSQLaj = "SELECT COUNT(j.id) AS antaljob FROM job AS j WHERE jobans1 = " & usemrn & " AND (jobstatus = 1 OR jobstatus = 3) AND risiko >= 0"
aJob = 0
oRec.open strSQLaj, oConn, 3
if not oRec.EOF then

    if isNULL(oRec("antaljob")) <> true then
    aJob = trim(oRec("antaljob"))
    else
    aJob = 0
    end if

end if
oRec.close


else

mdAngivet = 0
aJob = 1

end if 'usemrn




if (cint(mdAngivet) = 0 AND cint(aJob) >= 1) OR visning <> 1 then


if media <> "export" then

select case visning
case 1 'fra timeregsiden enkelt medarb.
dvtp = 300
fmaction = "timereg_akt_2006.asp?func=opdigv&mthuse="& dtuse
case 2' fra stat siden - enkelt medaerb
dvtp = 20
fmaction = "stat_opdater_igv.asp?func=opdaterdb&usemrn=" & usemrn & "&mthuse="& dtuse
case 3 'admin
dvtp = 20
fmaction = "stat_opdater_igv.asp?func=opdaterdb&usemrn=" & usemrn & "&mthuse="& dtuse
case else
dvtp = 300
fmaction = "timereg_akt_2006.asp?func=opdigv&mthuse="& dtuse
end select


if media = "print" then
dv_bdr = 0
else
dv_bdr = 10
end if

if cint(visning) = 1 then
dv_hgt = "height:700px;"
dv_wdt = "width:700px;"
else
dv_hgt = ""
dv_wdt = "width:840px;"
end if


 %>
<div id="jobstatusigv" style="position:absolute; top:<%=dvtp%>px; left:40px; <%=dv_wdt%> <%=dv_hgt%> display:; visibility:visible; border:<%=dv_bdr%>px #CCCCCC solid; padding:20px; overflow:auto; background-color:#FFFFFF; z-index:20000000000;">

<form action="<%=fmaction%>" method="post">
<table cellspacing=0 cellpadding=3 border=0 width=100%>
<tr><td colspan=7>

<%end if


if usemrn <> 0 then
call meStamdata(usemrn)
else
meTxt = ".."
end if

mth_1 = dateAdd("m", 1, dtuse)
mth_2 = dateAdd("m", 2, dtuse)
mth_3 = dateAdd("m", 3, dtuse)


if usemrn <> 0 then
lmt = 100
mnavnfld = ""
else
lmt = 500
mnavnfld = ", m.mnavn" 
end if



if media <> "export" then
%>



<h4>Stade, igangværende- job og -tilbud<br /><span style="font-size:11px; color:#5C75AA;">Job hvor <%=meTxt %> er <u>jobansvarlig</u> (maks. <%=lmt%>)</span></h4>

<%if media <> "print" then %>

        <%if usemrn <> 0 then %>
        Angiv stade og forventet fakturering, på de job hvor du er jobansvarlig.<br /> Vises fra d. 25 i hver måned og indtil du angiver at Du er <b>klar med din stade-indmelding.</b><br />  
        <input type="checkbox" name="i_per_afsl" value="1" <%=mdAngivetCHK %> /> Stade-indmelding afsluttet for <b><%=monthname(month(now)) & " - "& year(now) %></b>

                <%if level = 1 AND cdate(now) <= cdate("1-8-2012") AND (lto <> "epi" AND lto <> "epi_no" AND lto <> "epi_ab" AND lto <> "epi_sta") then %>
                <br /><br />
                <div style="border:2px #999999 solid; padding:5px; background-color:#FFFFFF; width:350px;">
                <input type="checkbox" name="i_jobasnvigv" value="1" /> Angiv her hvis I IKKE ønsker at bruge Stade-indmeldings funktionen. (Kan også slås til/fra i kontrolpanelet)
                </div>
                <%end if %>

        <%end if %>

<input type="hidden" value="<%=usemrn %>" name="i_mid" />

</td><td valign=top style="width:20px;">

<%if visning = 1 then %>
<span style="color:red; font-size:11px;" id="s_luk_igv">[X]</span>
<%end if %>

<%if (visning = 2 OR visning = 3) then 'AND level = 1  

if level <> 1 then
mtNow = date()
mthuse_minus = mthuse_minus

mthuse_minusDiff = dateDiff("m", mtNow, mthuse_minus)
if mthuse_minusDiff >= -1 then
showTilbagePil = 1
else
showTilbagePil = 0
end if

else
showTilbagePil = 1
end if
%>

                            <!-- skift uge pile -->
	                        <table cellpadding=0 cellspacing=0 border=0 width=80>
	                        <tr>
	                        <td valign=top align=right>
                            <%if showTilbagePil = 1 then %>
                            <a href="stat_opdater_igv.asp?func=<%=func%>&usemrn=<%=usemrn%>&mthuse=<%=mthuse_minus%>"><img src="../ill/nav_left_blue.png" border="0" /></a>
                            <%else%>&nbsp;
                            <%end if%></td> <!-- jobid=<=jobid%>&usemrn=<=usemrn%>& &fromsdsk=<=fromsdsk%> -->
                           <td style="padding-left:20px;" valign=top align=right><a href="stat_opdater_igv.asp?func=<%=func%>&usemrn=<%=usemrn%>&mthuse=<%=mthuse_plus%>"><img src="../ill/nav_right_blue.png" border="0" /></a></td>
	                        </tr>
	                        </table>

<%end if %>

<%else 'print%>
</td><td>&nbsp;

<%end if %>

</td></tr>

<%if media <> "print" then %>
<tr><td colspan=8 align=right style="padding:20px 20px 20px 0px;"><input type="submit" value="Opdater >>" /></td></tr>
<%end if %>

<tr><td valign=bottom>Job</td><td valign=bottom>Bruttooms <%=basisValISO%></td><td valign=bottom>WIP<br /><span style="font-size:9px;">Job: Afslutt. %<br />
Tilb.: Sandsy. %</span></td><td valign=bottom align=center>Faktureret <%=basisValISO%></td>
<td valign=bottom align=right>Forv. fakturering<br /><br /><%=left(monthname(datepart("m", mth_1, 2,2)), 3) &" "& datepart("yyyy", mth_1, 2,2) %></td>
<td valign=bottom><%=left(monthname(datepart("m", mth_2, 2,2)), 3) &" "& datepart("yyyy", mth_2, 2,2) %></td>
<td valign=bottom><%=left(monthname(datepart("m", mth_3, 2,2)), 3) &" "& datepart("yyyy", mth_3, 2,2) %></td>
<td valign=bottom align=center>Status</td></tr>

<%

end if 'media

'&" if(f.faktype <> 1, SUM(f.beloeb * (f.kurs / 100)), SUM(f.beloeb * -1 * (f.kurs / 100))) AS faktureret, "
strSQLj = "SELECT j.id AS jid, jobstartdato, jobnavn, jobnr, kkundenavn, kkundenr, jobstatus, jo_bruttooms, "_
& "SUM( "_
& "if (f.faktype = 0, f.beloeb * (f.kurs / 100), f.beloeb * -1 * (f.kurs / 100)) "_
& ") AS faktureret, "

if level = 1 then 'faktureret pr. md
'strSQLj = strSQLj &" if(f1.faktype <> 1, SUM(f1.beloeb * (f1.kurs / 100)), SUM(f1.beloeb * -1 * (f1.kurs / 100))) AS faktureret_1, "
strSQLj = strSQLj & " SUM( "_
& "if (f1.faktype = 0, f1.beloeb * (f1.kurs / 100), f1.beloeb * -1 * (f1.kurs / 100)) "_
& ") AS faktureret_1, "

'strSQLj = strSQLj &" if(f2.faktype <> 1, SUM(f2.beloeb * (f2.kurs / 100)), SUM(f2.beloeb * -1 * (f2.kurs / 100))) AS faktureret_2, "
strSQLj = strSQLj & " SUM( "_
& "if (f2.faktype = 0, f2.beloeb * (f2.kurs / 100), f2.beloeb * -1 * (f2.kurs / 100)) "_
& ") AS faktureret_2, "

'strSQLj = strSQLj &" if(f3.faktype <> 1, SUM(f3.beloeb * (f3.kurs / 100)), SUM(f3.beloeb * -1 * (f3.kurs / 100))) AS faktureret_3, "
 strSQLj = strSQLj & " SUM( "_
& "if (f3.faktype = 0, f3.beloeb * (f3.kurs / 100), f3.beloeb * -1 * (f3.kurs / 100)) "_
& ") AS faktureret_3, "
end if

strSQLj = strSQLj &" restestimat, stade_tim_proc, sandsynlighed "& mnavnfld &" FROM job AS j "_
&" LEFT JOIN kunder AS k ON (k.kid = jobknr) "

strSQLj = strSQLj & " LEFT JOIN fakturaer AS f ON (f.jobid = j.id AND f.shadowcopy <> 1 AND f.medregnikkeioms <> 1 AND f.medregnikkeioms <> 2) "

if level = 1 then

strSQLj = strSQLj & " LEFT JOIN fakturaer AS f1 ON (f1.jobid = j.id) "_
&" AND ((YEAR(f1.fakdato) = '"& year(mth_1) &"' AND MONTH(f1.fakdato) = '"& month(mth_1) &"' AND f1.brugfakdatolabel = 0) "_
&" OR (f1.brugfakdatolabel = 1 AND YEAR(f1.labeldato) = '"& year(mth_1) &"' AND MONTH(f1.labeldato) = '"& month(mth_1) &"')) AND f1.shadowcopy <> 1 AND f1.medregnikkeioms <> 1 AND f.medregnikkeioms <> 2"

strSQLj = strSQLj & " LEFT JOIN fakturaer AS f2 ON (f2.jobid = j.id) "_
&" AND ((YEAR(f2.fakdato) = '"& year(mth_2) &"' AND MONTH(f2.fakdato) = '"& month(mth_2) &"' AND f2.brugfakdatolabel = 0) "_
&" OR (f2.brugfakdatolabel = 2 AND YEAR(f2.labeldato) = '"& year(mth_2) &"' AND MONTH(f2.labeldato) = '"& month(mth_2) &"')) AND f2.shadowcopy <> 1 AND f2.medregnikkeioms <> 1 AND f.medregnikkeioms <> 2"

strSQLj = strSQLj & " LEFT JOIN fakturaer AS f3 ON (f3.jobid = j.id) "_
&" AND ((YEAR(f3.fakdato) = '"& year(mth_3) &"' AND MONTH(f3.fakdato) = '"& month(mth_3) &"' AND f3.brugfakdatolabel = 0) "_
&" OR (f3.brugfakdatolabel = 3 AND YEAR(f3.labeldato) = '"& year(mth_3) &"' AND MONTH(f3.labeldato) = '"& month(mth_3) &"')) AND f3.shadowcopy <> 1 AND f3.medregnikkeioms <> 1 AND f.medregnikkeioms <> 2"

end if


if usemrn <> 0 then
strSQLj = strSQLj & " WHERE (jobans1 = "& usemrn &") AND risiko >= 0 AND (jobstatus = 1 OR jobstatus = 3) GROUP BY j.id ORDER BY jobstartdato LIMIT "& lmt
else
strSQLj = strSQLj & " LEFT JOIN medarbejdere AS m ON (m.mid = j.jobans1) "
strSQLj = strSQLj & " WHERE risiko >= 0 AND (jobstatus = 1 OR jobstatus = 3) GROUP BY j.id ORDER BY jobnavn LIMIT "& lmt
end if

'OR jobans2 = "& session("mid"

'if session("mid") = 1 then
'Response.write strSQLj
'Response.flush
'end if

r = 0
faktureretTot = 0

oRec.open strSQLj, oConn, 3
while not oRec.EOF 

'select case oRec("jobstatus")
'case 0
'jStatus = "Lukket"
'jStatusFntCol = "darkred" 
'case 1
'jStatus = "Aktiv"
'jStatusFntCol = "#6CAE1C" 
'case 2
'jStatus = "Passiv"
'jStatusFntCol = "#CCCCC" 
'case 3
'jStatus = "Tilbud"
'jStatusFntCol = "#999999" 
'case 4
'jStatus = "Gennemsyn"
'jStatusFntCol = "#CCCCCC" 
'end select
if level = 1 then
faktureret_1 = oRec("faktureret_1")
faktureret_2 = oRec("faktureret_2")
faktureret_3 = oRec("faktureret_3")
else
faktureret_1 = 0
faktureret_2 = 0
faktureret_3 = 0
end if

if cint(oRec("stade_tim_proc")) = 0 then 'rest estimat angivet i timer
afsl_tim_proc = "t."
else	
afsl_tim_proc = "%" 
end if

if oRec("jobstatus") = 3 then 'tilbud
estVZBtyp = "Hidden"
sanVZBtyp = "Text"
else
estVZBtyp = "Text"
sanVZBtyp = "Hidden"
end if

sandsynlighed = oRec("sandsynlighed")

restestimat = oRec("restestimat")


if oRec("jo_bruttooms") <> 0 then
btoms = formatnumber(oRec("jo_bruttooms"), 2)
else
btoms = ""
end if


        select case oRec("jobstatus")
        case 0
        jobStatusTxt = "Lukket"
        case 1
        jobStatusTxt = "Aktiv"
        case 2
        jobStatusTxt = "Passiv / Til fakturering"
        case 3
        jobStatusTxt = "Tilbud"
        case 4
        jobStatusTxt = "Gennemsyn"
        end select


if media <> "export" then
%>

<tr><td style="border-bottom:1px #CCCCCC solid; white-space:nowrap;"><span style="font-size:11px; color:#5C75AA;"><%=left(oRec("kkundenavn"), 25) %></span><br />
<b><%=left(oRec("jobnavn"), 20) &"</b> ("& oRec("jobnr") &")"%> 

<%if cint(usemrn) = 0 then

if oRec("mnavn") <> "" then
jobbansThis = oRec("mnavn")
else
jobbansThis = "(ej angivet)"
end if%>
<%else 

jobbansThis = meNavn

end if 


%>

<br /><span style="font-size:9px; color:#999999;"><%=jobbansThis%></span>
</td>

<td style="border-bottom:1px #CCCCCC solid; white-space:nowrap;">
<%if media <> "print" then %>
<input type="hidden" value="<%=oRec("jid") %>" name="i_jobids" /><input type="hidden" value="<%=oRec("jobstatus") %>" name="i_status" />
<input type="text" name="i_brutto_<%=oRec("jid") %>" value="<%=btoms%>" style="font-size:10px; width:80px;" />
<%else %>
&nbsp;<%=btoms%>
<%end if %>
</td>
<td style="border-bottom:1px #CCCCCC solid; white-space:nowrap;">
<%if media <> "print" then %>
<input type="<%=estVZBtyp %>" name="i_restestimat_<%=oRec("jid") %>" value="<%=restestimat%>" style="font-size:10px; font-family:arial; width:30px;" />
		
		<%
       


        if estVZBtyp <> "Hidden" then 'hidden = tilbud
        
            select case oRec("stade_tim_proc")
		    case 0
                if restestimat <> 0 then '** timer forvalgt
		        stade_timproc_SEL0 = "SELECTED"
		        stade_timproc_SEL1 = ""
                
                else
                stade_timproc_SEL0 = ""
		        stade_timproc_SEL1 = "SELECTED"
                
                end if
		    case 1
		    stade_timproc_SEL0 = ""
		    stade_timproc_SEL1 = "SELECTED"
            
		    end select %>

		    <select name="i_stade_tim_proc_<%=oRec("jid") %>" style="font-size:10px; font-family:arial;">
		        <option value=0 <%=stade_timproc_SEL0 %>>t.</option>
		        <option value=1 <%=stade_timproc_SEL1 %>>%</option>
		    </select>
        
        <%else%>
        <input type="hidden" name="i_stade_tim_proc_<%=oRec("jid") %>" value="<%=oRec("stade_tim_proc") %>" />
        <input type="<%=sanVZBtyp %>" name="i_sandsynlighed_<%=oRec("jid") %>" value="<%=sandsynlighed%>" style="font-size:10px; font-family:arial; width:30px;" /> %
		
        <%end if 
        
        
        
       
  else      
   
   if estVZBtyp <> "Hidden" then 'timer
    
            select case oRec("stade_tim_proc")
		    case 0
            enh = " t."  
		    case 1
		    enh = " %"
            end select
            %>
            <%=restestimat&enh%>
            <%
    else
    %>
    <%=sandsynlighed%> %
    <%
    end if

   %>
  
  <%end if %>

</td>

    <%if isNull(oRec("faktureret")) <> true then
        erfaktureretBel = formatnumber(oRec("faktureret"), 2) 
      else
        erfaktureretBel = 0
    end if %>

<td style="border-bottom:1px #CCCCCC solid; white-space:nowrap;" align=right><%=erfaktureretBel%></td>

<%


faktureretTot = faktureretTot + (erfaktureretBel) 

call forvfak(media, jid, yUse, mUse)

%>

<td style="border-bottom:1px #CCCCCC solid;" align=center>

<%if media <> "print" then

stSEL0 = ""
stSEL1 = ""
stSEL2 = ""
stSEL3 = ""
stSEL4 = ""

select case oRec("jobstatus")
case 0
stSEL0 = "SELECTED"
case 1
stSEL1 = "SELECTED"
case 2
stSEL2 = "SELECTED"
case 3
stSEL3 = "SELECTED"
case 4
stSEL4 = "SELECTED"
end select


 %>


<select name="i_luk_<%=oRec("jid") %>" style="width:50px; font-size:9px;">
<option value="1" <%=stSEL1%>>Aktiv</option>
<option value="0" <%=stSEL0%>>Lukket</option>
<option value="2" <%=stSEL2%>>Passiv</option>
<option value="3" <%=stSEL3%>>Tilbud</option>
<option value="4" <%=stSEL4%>>Gennesyn</option>
</select>

<%
else
Response.write jobStatusTxt
end if

%></td></tr>
<%

else 'media = export
stexptxt = stexptxt & oRec("kkundenavn") &";"& oRec("kkundenr") &";"& oRec("jobnavn") &";"& oRec("jobnr") & ";"
stexptxt = stexptxt & meNavn &";"

 
stexptxt = stexptxt & jobStatusTxt &";"& btoms &";"

stexptxt = stexptxt & sandsynlighed  &";"& restestimat &";"

select case oRec("stade_tim_proc")
case 0
rtimproc ="t."
case 1
rtimproc ="%"
end select


stexptxt = stexptxt & rtimproc &";"
stexptxt = stexptxt & formatnumber(oRec("faktureret"), 2) &";"

call forvfak(media, jid, yUse, mUse)

stexptxt = stexptxt &";xx99123sy#z"
end if

r = r + 1
oRec.movenext
wend
oRec.close 


if media <> "export" then
 %>
 <tr><td colspan="3" align=right>Total:</td>
    <td align=right><%=formatnumber(faktureretTot, 2) %></td>
    <td align=right><%=formatnumber(faktureretThisMthTot_1, 2) %></td>
    <td align=right><%=formatnumber(faktureretThisMthTot_2, 2) %></td>
    <td align=right><%=formatnumber(faktureretThisMthTot_3, 2) %></td>
    <td>&nbsp;</td>
 </tr>
<tr><td colspan=8 style="padding:20px 20px 20px 0px;">Der er <b><%=r %></b> job og tilbud på <b><%=meTxt %></b> liste.</td></tr>

<%if media <> "print" then %>
<tr><td colspan=8 align=right style="padding:20px 20px 20px 0px;"><input type="submit" value="Opdater >>" /></td></tr>
<%end if %>
</table>
</form>

</div>
<%
end if 'media

end if 'aJob

end if 'session(mid)


end if
end sub



function stadeopdater()


    jobids = split(request("i_jobids"), ", ")
    dddato = day(now) &" "& left(monthname(month(now)), 3) &". "& year(now) 

    iMedid = request("i_mid")

    for i = 0 to UBOUND(jobids)
    ibrutto = request("i_brutto_"& jobids(i)) 
    irest = request("i_restestimat_"& jobids(i))
    isand = request("i_sandsynlighed_"& jobids(i))

    call erDetInt(ibrutto)
    call erDetInt(irest)
    call erDetInt(isand)
    if isInt = 0 then 

    ibrutto = replace(ibrutto, ".","")
    ibrutto = replace(ibrutto, ",",".")

    if len(trim(ibrutto)) <> 0 then
    ibrutto = ibrutto
    else
    ibrutto = 0
    end if
   
    irest = replace(irest, ".","")
    irest = replace(irest, ",",".")

    if len(trim(irest)) <> 0 then
    irest = irest
    else
    irest = 0
    end if

    isand = replace(isand, ".","")
    isand = replace(isand, ",",".")

    if len(trim(isand)) <> 0 then
    isand = isand
    else
    isand = 0
    end if

    ist_tim_proc = request("i_stade_tim_proc_"& jobids(i))

    'if len(trim(request("i_luk_"& jobids(i)))) <> 0 then
    i_luk = request("i_luk_"& jobids(i))
    'else
    'i_luk = 0
    'end if

    'if cint(i_luk) = 1 then
    strSQLjobst = ", jobstatus = " & i_luk
    'else 
    'strSQLjobst = ""
    'end if

    strSQL = "UPDATE job SET editor = '"& session("user") &"', dato = '"& dddato &"', jo_bruttooms = "& ibrutto &", restestimat = "& irest &", stade_tim_proc = "& ist_tim_proc &", sandsynlighed = '"& isand &"' "& strSQLjobst &" WHERE id = "& jobids(i)
    'Response.write strSQL & "<br><br>"
    
    oConn.execute(strSQL) 

    i_md1_bel = request("i_md1_"& jobids(i))

    call erDetInt(i_md1_bel)
    if isInt = 0 then 
    
    i_md1_bel = replace(i_md1_bel, ".","")
    i_md1_bel = replace(i_md1_bel, ",",".")

    if len(trim(i_md1_bel)) <> 0 then
    i_md1_bel = i_md1_bel
    else
    i_md1_bel = 0
    end if

    milepal_dato_1 = request("i_md1_dt_"& jobids(i)) 
    milepal_dato_1 = year(milepal_dato_1) &"/"& month(milepal_dato_1) &"/15"
    
   findes_1 = request("i_findes1_"& jobids(i))

    if cdbl(findes_1) = 0 then
        if cdbl(i_md1_bel) <> 0 then
        strSQL = "INSERT INTO milepale (navn, type, beskrivelse, editor, milepal_dato, jid, belob) VALUES "_
        &" ('Faktura', 1, 'Faktura', '"& session("user") &"', '"& milepal_dato_1 &"', "& jobids(i) &", "& i_md1_bel &")"
        end if
    else
        if cdbl(i_md1_bel) <> 0 then
        strSQL = "UPDATE milepale SET milepal_dato = '"& milepal_dato_1 &"', belob = "& i_md1_bel &" WHERE id = "& findes_1
        else 'slet
        strSQL = "DELETE FROM milepale WHERE id = "& findes_1
        end if
        
    end if

    oconn.execute(strSQL)
    end if
    isInt = 0

    i_md2_bel = request("i_md2_"& jobids(i))

    call erDetInt(i_md2_bel)
    if isInt = 0 then 
    
    i_md2_bel = replace(i_md2_bel, ".","")
    i_md2_bel = replace(i_md2_bel, ",",".")

    if len(trim(i_md2_bel)) <> 0 then
    i_md2_bel = i_md2_bel
    else
    i_md2_bel = 0
    end if

    milepal_dato_2 = request("i_md2_dt_"& jobids(i)) 
    milepal_dato_2 = year(milepal_dato_2) &"/"& month(milepal_dato_2) &"/15"
    
    findes_2 = request("i_findes2_"& jobids(i))

    if cdbl(findes_2) = 0 then
        if cdbl(i_md2_bel) <> 0 then
        strSQL = "INSERT INTO milepale (navn, type, beskrivelse, editor, milepal_dato, jid, belob) VALUES "_
        &" ('Faktura', 1, 'Faktura', '"& session("user") &"', '"& milepal_dato_2 &"', "& jobids(i) &", "& i_md2_bel &")"
        end if
    else
        if cdbl(i_md2_bel) <> 0 then
        strSQL = "UPDATE milepale SET milepal_dato = '"& milepal_dato_2 &"', belob = "& i_md2_bel &" WHERE id = "& findes_2
        else 'slet
        strSQL = "DELETE FROM milepale WHERE id = "& findes_2
        end if
        
    end if

    oconn.execute(strSQL)
    end if
    isInt = 0


    i_md3_bel = request("i_md3_"& jobids(i))

    call erDetInt(i_md3_bel)
    if isInt = 0 then 
    
    i_md3_bel = replace(i_md3_bel, ".","")
    i_md3_bel = replace(i_md3_bel, ",",".")

    if len(trim(i_md3_bel)) <> 0 then
    i_md3_bel = i_md3_bel
    else
    i_md3_bel = 0
    end if

    milepal_dato_3 = request("i_md3_dt_"& jobids(i)) 
    milepal_dato_3 = year(milepal_dato_3) &"/"& month(milepal_dato_3) &"/15"
    
    findes_3 = request("i_findes3_"& jobids(i))

    if cdbl(findes_3) = 0 then
        if cdbl(i_md3_bel) <> 0 then
        strSQL = "INSERT INTO milepale (navn, type, beskrivelse, editor, milepal_dato, jid, belob) VALUES "_
        &" ('Faktura', 1, 'Faktura', '"& session("user") &"', '"& milepal_dato_3 &"', "& jobids(i) &", "& i_md3_bel &")"
        end if
    else
        if cdbl(i_md3_bel) <> 0 then
        strSQL = "UPDATE milepale SET milepal_dato = '"& milepal_dato_3 &"', belob = "& i_md3_bel &" WHERE id = "& findes_3
        else 'slet
        strSQL = "DELETE FROM milepale WHERE id = "& findes_3
        end if
        
    end if

    oconn.execute(strSQL)
    end if
    isInt = 0


    end if
    isInt = 0

    next


    '** Opdaterer status ****'s

    ddUse = day(now) &"/"& month(now) &"/"& year(now)
    yUse = year(ddUse)
    mUse = month(ddUse)
    ddUseSQL = year(ddUse) &"/"& month(ddUse) &"/"& day(ddUse)



    if len(trim(request("i_per_afsl"))) <> 0 then
                
                igv_findes_id = 0
                igv_findes = 0
                strSQLigv_findes = "SELECT id FROM job_igv_status WHERE medid = "& iMedid &" AND aar = "& yUse &" AND maaned = "& mUse
                oRec.open strSQLigv_findes, oConn, 3
                if not oRec.EOF then
                igv_findes = 1
                igv_findes_id = oRec("id")
                end if
                oRec.close

                if igv_findes = 0 then
                strSQLigv_stat = "INSERT INTO job_igv_status (editor, dato, medid, aar, maaned) VALUES ('"& session("user") &"', '"& ddUseSQL &"', "& iMedid &", "& yUse &", "& mUse &")"
                oConn.execute(strSQLigv_stat)
                else
                strSQLigv_stat = "UPDATE job_igv_status SET editor = '"& session("user") &"', dato = '"& ddUseSQL &"' WHERE id = "& igv_findes_id
                oConn.execute(strSQLigv_stat)
                end if
    else
    
    strSQLigv_stat = "DELETE FROM job_igv_status WHERE medid = "& iMedid &" AND aar = "& yUse &" AND maaned = "& mUse
    oConn.execute(strSQLigv_stat)

    end if 

    '*** Opdaterer vis funktion *****'
    if len(trim(request("i_jobasnvigv"))) <> 0 then
    strSQligv = "UPDATE licens SET jobasnvigv = 0 WHERE id = 1"
    oConn.execute(strSQligv)
    end if   

end function








public faktureretThisMthTot_1, faktureretThisMthTot_2, faktureretThisMthTot_3, bThisUse
function forvfak(media, jid, yUse, mUse)


for i = 1 to 3 
   
   
   bThisUse = 0 
   fFindes = 0
   bThis = ""

   select case i
   case 1
   iMtUse = mth_1
   faktureretThisMth = faktureret_1 
   faktureretThisMthTot_1 = faktureretThisMthTot_1 + faktureret_1 
   case 2
   iMtUse = mth_2
   faktureretThisMth = faktureret_2
   faktureretThisMthTot_2 = faktureretThisMthTot_2 + faktureret_2
   case 3
   iMtUse = mth_3
   faktureretThisMth = faktureret_3
   faktureretThisMthTot_3 = faktureretThisMthTot_3 + faktureret_3
   end select

   yUSe = year(iMtUse)
   mUse = month(iMtUse)

  strSQLmp = "SELECT belob, id FROM milepale WHERE jid = " & oRec("jid") &" AND YEAR(milepal_dato) = '"& yUSe &"' AND MONTH(milepal_dato) = '"& mUse &"' " 
  
  'Response.write strSQLmp
  'Response.flush
  oRec3.open strSQLmp, oConn, 3
  if not oRec3.EOF then

  if oRec3("belob") <> 0 then
  bThis = formatnumber(oRec3("belob"), 2)
  fFindes = oRec3("id")
  else
  bThis = ""
  fFindes = 0
  end if
  
  end if
  oRec3.close

  if bThis <> "" then
  bThisUse = bThis 
  else
  bThisUse = 0
  end if

   select case i
   case 1
  faktureretThisMthTot_1 = faktureretThisMthTot_1 + bThisUse 
   case 2
   faktureretThisMthTot_2 = faktureretThisMthTot_2 + bThisUse
   case 3
    faktureretThisMthTot_3 = faktureretThisMthTot_3 + bThisUse
   end select


  Response.flush

if media <> "export" AND media <> "print" then 
%>
<td style="border-bottom:1px #CCCCCC solid;" align=right>
<input type="hidden" value="<%=fFindes%>" name="i_findes<%=i%>_<%=oRec("jid") %>" />
<input type="hidden" value="<%=iMtUse%>" name="i_md<%=i%>_dt_<%=oRec("jid") %>" />
<input type="text" name="i_md<%=i%>_<%=oRec("jid") %>" value="<%=bThis %>" style="font-size:10px; width:50px;" />
<%if level = 1 AND faktureretThisMth <> 0 then %>
<br /><span style="font-size:9px; color:#999999;">f: <%=formatnumber(faktureretThisMth, 2)%></span>
<%end if %>
</td>


<%
else
    if media <> "print" then 'export 
    stexptxt = stexptxt & bThis &";"
    else
    %>
<td style="border-bottom:1px #CCCCCC solid;" align=right>&nbsp;<%=bThis %></td>
<%
    end if
end if

next 


end function




sub ekportogprint_fn()

    
    if media = "export" then 

    strEkspHeader = "Kontakt;Kontakt id;Job;Job.nr;Jobansvarlig;Status;Brutto.Oms.(Budget);Sandsynlighed %;Restestimat;t./%;Faktureret;"_
    &"Forv. fak. "& left(monthname(datepart("m", mth_1, 2,2)), 3) &" "& datepart("yyyy", mth_1, 2,2) &";"_
    &"Forv. fak. "& left(monthname(datepart("m", mth_2, 2,2)), 3) &" "& datepart("yyyy", mth_2, 2,2) &";"_
    &"Forv. fak. "& left(monthname(datepart("m", mth_3, 2,2)), 3) &" "& datepart("yyyy", mth_3, 2,2) &";"
    
    strEksportTxt = stexptxt
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	call TimeOutVersion()
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\stat_opdater_igv.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\stadeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\stadeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\stadeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\stadeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "stadeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
           

                

				'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
              
                 strEkspHeader = replace(strEkspHeader, "xx99123sy#z", vbcrlf)
	            ekspTxt = replace(strEksportTxt, "xx99123sy#z", vbcrlf)

				objF.WriteLine(strEkspHeader)
				objF.WriteLine(ekspTxt)
				
                %>

               
	            <table border=0 cellspacing=1 cellpadding=0 width="200">
	            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	            <img src="../ill/outzource_logo_200.gif" />
	            </td>
	            </tr>
	            <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Din CSV. fil er klar >></a>
	            </td></tr>
	            </table>

               
               
	            <%
                Response.end
                 'Response.redirect "../inc/log/data/"& file &""	
	

            end if 













 if media <> "print" AND media <> "export" then
    ptop = 20
    pleft = 960
    pwdt = 190

call eksportogprint(ptop,pleft, pwdt)
%>




<form action="javascript:popUp('stat_opdater_igv.asp?media=export&func=<%=func%>','400','200','250','120');" method="post" target="_self">
<tr> 

    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit3" type="submit" value="Eksportér til .csv >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>


<form action="stat_opdater_igv.asp?media=print&func=<%=func%>" method="post" target="_blank">
<tr>

    <td valign=top align=center>
   <input type=image src="../ill/printer3.png" />
    </td><td class=lille><input id="Submit6" type="submit" value="Print venlig" style="font-size:9px; width:130px;" /></td>
</tr>
</form>

   
	
   </table>
</div>

<%else 

 Response.Write("<script language=""JavaScript"">window.print();</script>")

end if 




end sub


%>