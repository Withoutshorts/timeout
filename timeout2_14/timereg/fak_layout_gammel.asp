<%

 '**** til ikke printbar medarbjeder timeoversigt (højre side) *** 
								if lastaktid <> oRec("aktid") then 
								strMedarbFakdet = strMedarbFakdet &"<tr><td colspan=4><b>"& oRec("beskrivelse") &"</b></td></tr>"
								end if
								
								mf = 0
								
								strSQL2 = "SELECT fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, showonfak, medrabat FROM fak_med_spec WHERE fakid = "& id &" AND aktid = "& oRec("aktid") &" AND enhedspris = "& ehpris &" AND fak <> 0"
								oRec2.open strSQL2, oConn, 3 
								while not oRec2.EOF 
								
								if kurs <> 0 then
								medEhpris = oRec2("enhedspris") * (100/kurs)
								medBelob = oRec2("beloeb") * (100/kurs)
								else
								medEhpris = oRec2("enhedspris")
								medBelob = oRec2("beloeb") 
								end if
								
								if intFaktype = 1 then 'kreditnota
								medBelob = -(medBelob)
								medEhpris = -(medEhpris)
								end if
								
								medrabatThis = oRec2("medrabat")
								
								strMedarbFakdet = strMedarbFakdet &"<tr>"_
								&"<td valign=top align=right style='padding-right:2px;'><font class=megetlillesort>"& formatnumber(oRec2("fak"), 2) &"</td>"_
								&"<td valign=top style='padding-left:5px;'><font class=megetlillesort>" & left(oRec2("tekst"), 10) &"</td>"_
								&"<td valign=top align=right style='padding-right:2px;'><font class=megetlillesort>" & formatnumber(medEhpris, 2) &"</td>"
								
								if cint(visrabatkol) = 1 then
								strMedarbFakdet = strMedarbFakdet &"<td valign=top align=right style='padding-right:10px;'>"
								
								    if cdbl(medrabatThis) <> 0 then
								    strMedarbFakdet = strMedarbFakdet &"<font class=megetlillesort>" & (medrabatThis * 100) &" %</td>"
								    else
								    strMedarbFakdet = strMedarbFakdet &"&nbsp;</td>"
								    end if
								
								end if
								
								strMedarbFakdet = strMedarbFakdet &"<td valign=top align=right><font class=megetlillesort>"& formatnumber(medBelob,2) &"</td>"_
								&"</tr>"
								
								medarbejdertimer = medarbejdertimer + oRec2("fak")
								medarbejdersum = medarbejdersum + medBelob 'oRec2("beloeb")
								
								mf = 1
								oRec2.movenext
								wend
								oRec2.close 
								
								
								if mf = 0 then
								strMedarbFakdet = strMedarbFakdet &"<tr><td colspan=4 style='padding-left:20;'>-</td></tr>"
								end if
 %>